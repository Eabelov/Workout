#requires -Version 5.1
<#
.SYNOPSIS
    Removes Exchange mailbox database copies (passive first), then the database
    itself, and deletes leftover .db / .log folders on each DAG node.

.DESCRIPTION
    Reads database names from a text file (one per line) and for each database:
      1. Aborts that database if any mailbox still lives on it (user, archive,
         arbitration, audit, migration, public folder, monitoring). Re-checks
         immediately before removing copies and before removing the last copy.
      2. Verifies EDB/log paths on each passive copy (UNC + optional WinRM),
         removes the copy, polls until it disappears, then deletes files.
      3. Requires double confirmation, then dismounts and removes the last
         (active) copy, polls for completion, and deletes its files.

    Supports -WhatIf / -Confirm via SupportsShouldProcess. Errors are collected
    instead of calling exit from the loop; -ContinueOnError keeps going.

.PARAMETER ContinueOnError
    If set, a failure on one database does not stop the remaining databases.
    All failures are still reported at the end (exit code 1 if any occurred).

.PARAMETER Force
    Skip the interactive double confirmation for the last copy. Does not
    override -WhatIf.

.EXAMPLE
    .\Remove-ExchangeDatabaseCopies.ps1 -WhatIf

.EXAMPLE
    .\Remove-ExchangeDatabaseCopies.ps1 -ContinueOnError -PollIntervalSeconds 5

.EXAMPLE
    .\Remove-ExchangeDatabaseCopies.ps1 -Force

.NOTES
    Run from Exchange Management Shell (or a session with EMS snap-in loaded)
    as an Organization Management admin that can reach C$/D$/... on DAG nodes.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [string]$DatabaseListPath = 'C:\BEA\dag4databases.txt',
    [string]$HistoryLogPath   = 'C:\BEA\dbcopystatushistory.txt',

    [int]$PollIntervalSeconds = 5,
    [Alias('CopyRemovalWaitSeconds')]
    [int]$CopyRemovalTimeoutSeconds = 120,
    [Alias('LastCopyDismountWaitSeconds')]
    [int]$DismountTimeoutSeconds = 60,
    [int]$DatabaseRemovalTimeoutSeconds = 120,
    [int]$FileDeleteRetryWaitSeconds = 60,
    [int]$FileDeleteMaxRetries = 10,

    [switch]$ContinueOnError,
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:OperationErrors = New-Object 'System.Collections.Generic.List[object]'
$script:SucceededDatabases = New-Object 'System.Collections.Generic.List[string]'

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )
    $line = '{0:yyyy-MM-dd HH:mm:ss} [{1}] {2}' -f (Get-Date), $Level, $Message
    Write-Host $line
    $logDir = Split-Path -Parent $HistoryLogPath
    if ($logDir -and -not (Test-Path -LiteralPath $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    Add-Content -LiteralPath $HistoryLogPath -Value $line -Encoding UTF8
}

function Add-OperationError {
    param(
        [Parameter(Mandatory)][string]$Database,
        [Parameter(Mandatory)][string]$Stage,
        [Parameter(Mandatory)][string]$Message
    )
    $item = [pscustomobject]@{
        Timestamp = Get-Date
        Database  = $Database
        Stage     = $Stage
        Message   = $Message
    }
    [void]$script:OperationErrors.Add($item)
    Write-Log "${Database} [${Stage}]: $Message" 'ERROR'
}

function Initialize-ExchangeShell {
    if (Get-Command -Name Get-MailboxDatabase -ErrorAction SilentlyContinue) {
        return
    }

    $registered = @(Get-PSSnapin -Registered -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -eq 'Microsoft.Exchange.Management.PowerShell.SnapIn' -or
                       $_.Name -eq 'Microsoft.Exchange.Management.PowerShell.E2010' })
    if (-not $registered) {
        throw 'Exchange Management snap-in is not registered. Run this script from EMS.'
    }
    foreach ($snap in $registered) {
        if (-not (Get-PSSnapin -Name $snap.Name -ErrorAction SilentlyContinue)) {
            Add-PSSnapin -Name $snap.Name -ErrorAction Stop
        }
    }

    if (-not (Get-Command -Name Get-MailboxDatabase -ErrorAction SilentlyContinue)) {
        throw 'Exchange cmdlets are not available after adding the snap-in.'
    }
}

function Get-PreferredDomainController {
    $logonServer = [string]$env:LOGONSERVER
    $dnsDomain   = [string]$env:USERDNSDOMAIN
    $hostName    = $logonServer -replace '^\\\\', ''

    if ([string]::IsNullOrWhiteSpace($hostName)) {
        throw 'LOGONSERVER is empty; cannot determine a preferred DC.'
    }
    # LOGONSERVER is usually NetBIOS (\\DC01). Do not append the DNS suffix if
    # it is already an FQDN — that produced dc01.contoso.com.contoso.com.
    if ($hostName -match '\.' -or [string]::IsNullOrWhiteSpace($dnsDomain)) {
        return $hostName
    }
    return "$hostName.$dnsDomain"
}

function ConvertTo-AdminShareUnc {
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [Parameter(Mandatory)][string]$LocalPath
    )
    $normalized = $LocalPath.TrimEnd('\')
    if ($normalized -match '^[\\]{2}') {
        return $normalized
    }
    # C:\foo\bar -> \\server\C$\foo\bar  (any drive letter, any case)
    if ($normalized -match '^([A-Za-z]):(.*)$') {
        return '\\{0}\{1}${2}' -f $ComputerName, $Matches[1], $Matches[2]
    }
    throw "Cannot convert '$LocalPath' on $ComputerName to an admin UNC path."
}

function Get-DatabaseFileFolders {
    param([Parameter(Mandatory)][string]$DatabaseName)

    $dbObject = Get-MailboxDatabase -Identity $DatabaseName
    $edbFile  = [string]$dbObject.EdbFilePath
    $logDir   = [string]$dbObject.LogFolderPath
    if ([string]::IsNullOrWhiteSpace($edbFile) -or [string]::IsNullOrWhiteSpace($logDir)) {
        throw "Database $DatabaseName is missing EdbFilePath or LogFolderPath."
    }
    [pscustomobject]@{
        EdbFile     = $edbFile
        EdbFolder   = Split-Path -Parent $edbFile
        EdbFileName = Split-Path -Leaf $edbFile
        LogFolder   = $logDir
    }
}

function Test-DatabaseHasMailboxes {
    param([Parameter(Mandatory)][string]$DatabaseName)

    # Each hashtable is splatted onto Get-Mailbox. Empty = regular user mailboxes
    # (the original used a space character + Invoke-Expression for this case).
    $queries = @(
        @{ Name = 'User'         ; Params = @{} },
        @{ Name = 'Archive'      ; Params = @{ Archive      = $true } },
        @{ Name = 'Arbitration'  ; Params = @{ Arbitration  = $true } },
        @{ Name = 'AuditLog'     ; Params = @{ AuditLog     = $true } },
        @{ Name = 'AuxAuditLog'  ; Params = @{ AuxAuditLog  = $true } },
        @{ Name = 'Migration'    ; Params = @{ Migration    = $true } },
        @{ Name = 'PublicFolder' ; Params = @{ PublicFolder = $true } },
        @{ Name = 'Monitoring'   ; Params = @{ Monitoring   = $true } }
    )

    $found = @()
    foreach ($query in $queries) {
        $extra = $query.Params
        try {
            $mailboxes = @(Get-Mailbox -Database $DatabaseName -ResultSize 1 -ErrorAction Stop @extra)
        } catch {
            # -AuxAuditLog (and similar) do not exist on older Exchange builds.
            $missingParameter = $_.Exception -is [System.Management.Automation.ParameterBindingException] -or
                $_.Exception.Message -match 'A parameter cannot be found'
            if ($missingParameter) {
                Write-Log "Skipping Get-Mailbox -$($query.Name): parameter not supported on this Exchange version." 'WARN'
                continue
            }
            throw
        }
        if ($mailboxes.Count -gt 0) {
            $found += $query.Name
        }
    }
    return $found
}

function Assert-DatabaseHasNoMailboxes {
    param(
        [Parameter(Mandatory)][string]$DatabaseName,
        [Parameter(Mandatory)][string]$Stage
    )
    $occupiedBy = @(Test-DatabaseHasMailboxes -DatabaseName $DatabaseName)
    if ($occupiedBy.Count -gt 0) {
        throw "Database $DatabaseName is not empty ($Stage); found: $($occupiedBy -join ', ')."
    }
    Write-Log "Mailbox check passed ($Stage) for $DatabaseName"
}

function Wait-ForCondition {
    param(
        [Parameter(Mandatory)][scriptblock]$Condition,
        [Parameter(Mandatory)][string]$Description,
        [Parameter(Mandatory)][int]$TimeoutSeconds,
        [int]$PollSeconds = 5,
        [object[]]$ArgumentList = @()
    )
    if ($PollSeconds -lt 1) {
        $PollSeconds = 1
    }
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    while ($true) {
        $ready = [bool](& $Condition @ArgumentList)
        if ($ready) {
            Write-Log "Ready: $Description ($([int]$stopwatch.Elapsed.TotalSeconds)s)"
            return
        }
        if ($stopwatch.Elapsed.TotalSeconds -ge $TimeoutSeconds) {
            throw "Timed out after $TimeoutSeconds s waiting for: $Description"
        }
        Write-Log "Polling ($([int]$stopwatch.Elapsed.TotalSeconds)s/$TimeoutSeconds s): $Description"
        Start-Sleep -Seconds $PollSeconds
    }
}

function Test-IsNotFoundError {
    param($ErrorRecord)
    $message = $ErrorRecord.Exception.Message
    $fqid = [string]$ErrorRecord.FullyQualifiedErrorId
    return ($message -match 'couldn.?t find|cannot find|not found|doesn.?t exist|wasn.?t found') -or
        $fqid -match 'ManagementObjectNotFound'
}

function Test-CopyStatusAbsent {
    param(
        [Parameter(Mandatory)][string]$DatabaseName,
        [Parameter(Mandatory)][string]$CopyName
    )
    try {
        $status = @(Get-MailboxDatabaseCopyStatus -Identity $DatabaseName -ErrorAction Stop)
        return -not ($status | Where-Object { $_.Name -eq $CopyName })
    } catch {
        if (Test-IsNotFoundError $_) {
            return $true
        }
        throw
    }
}

function Test-DatabaseDismounted {
    param([Parameter(Mandatory)][string]$CopyName)
    try {
        $status = @(Get-MailboxDatabaseCopyStatus -Identity $CopyName -ErrorAction Stop)
        if ($status.Count -eq 0) {
            return $true
        }
        return $status[0].Status.ToString() -eq 'Dismounted'
    } catch {
        if (Test-IsNotFoundError $_) {
            return $true
        }
        throw
    }
}

function Test-MailboxDatabaseAbsent {
    param([Parameter(Mandatory)][string]$DatabaseName)
    try {
        $db = @(Get-MailboxDatabase -Identity $DatabaseName -ErrorAction Stop)
        return $db.Count -eq 0
    } catch {
        if (Test-IsNotFoundError $_) {
            return $true
        }
        throw
    }
}

function Assert-CopyFilesOnServer {
    param(
        [Parameter(Mandatory)][string]$MailboxServer,
        [Parameter(Mandatory)][string]$DatabaseName,
        [Parameter(Mandatory)][string]$EdbFolder,
        [Parameter(Mandatory)][string]$EdbFileName,
        [Parameter(Mandatory)][string]$LogFolder
    )

    $edbUnc     = ConvertTo-AdminShareUnc -ComputerName $MailboxServer -LocalPath $EdbFolder
    $logUnc     = ConvertTo-AdminShareUnc -ComputerName $MailboxServer -LocalPath $LogFolder
    $edbFileUnc = Join-Path $edbUnc $EdbFileName

    $missing = New-Object 'System.Collections.Generic.List[string]'
    if (-not (Test-Path -LiteralPath $edbUnc)) {
        [void]$missing.Add("UNC EDB folder $edbUnc")
    }
    if (-not (Test-Path -LiteralPath $edbFileUnc)) {
        [void]$missing.Add("UNC EDB file $edbFileUnc")
    }
    if (-not (Test-Path -LiteralPath $logUnc)) {
        [void]$missing.Add("UNC log folder $logUnc")
    }

    $remoteNote = 'skipped'
    try {
        $remote = Invoke-Command -ComputerName $MailboxServer -ErrorAction Stop -ScriptBlock {
            param($EdbFolder, $EdbFileName, $LogFolder)
            $edbFile = Join-Path $EdbFolder $EdbFileName
            [pscustomobject]@{
                EdbFolderExists = Test-Path -LiteralPath $EdbFolder
                EdbFileExists   = Test-Path -LiteralPath $edbFile
                LogFolderExists = Test-Path -LiteralPath $LogFolder
            }
        } -ArgumentList $EdbFolder, $EdbFileName, $LogFolder

        if (-not $remote.EdbFolderExists) {
            [void]$missing.Add("local EDB folder $EdbFolder on $MailboxServer")
        }
        if (-not $remote.EdbFileExists) {
            [void]$missing.Add("local EDB file $(Join-Path $EdbFolder $EdbFileName) on $MailboxServer")
        }
        if (-not $remote.LogFolderExists) {
            [void]$missing.Add("local log folder $LogFolder on $MailboxServer")
        }
        $remoteNote = 'ok'
    } catch {
        Write-Log "Remote path check on $MailboxServer failed ($($_.Exception.Message)); using UNC results." 'WARN'
        $remoteNote = 'failed'
    }

    if ($missing.Count -gt 0) {
        throw "Path verification failed for $DatabaseName on ${MailboxServer}: $($missing -join '; ')"
    }

    Write-Log "Verified paths on ${MailboxServer} (remote=$remoteNote): edb=$edbFileUnc log=$logUnc"
    [pscustomobject]@{
        EdbUnc     = $edbUnc
        LogUnc     = $logUnc
        EdbFileUnc = $edbFileUnc
    }
}

function Remove-FolderWithRetry {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][int]$MaxRetries,
        [Parameter(Mandatory)][int]$WaitSeconds
    )

    if ($WhatIfPreference) {
        Write-Log "WhatIf: would delete $Path"
        return
    }

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        if (-not (Test-Path -LiteralPath $Path)) {
            Write-Log "Path already gone: $Path"
            return
        }
        Write-Log "Deleting (attempt $attempt/$MaxRetries): $Path"
        try {
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
        } catch {
            Write-Log "Remove-Item failed: $($_.Exception.Message)" 'WARN'
        }
        if (-not (Test-Path -LiteralPath $Path)) {
            Write-Log "Deleted: $Path"
            return
        }
        Start-Sleep -Seconds $WaitSeconds
    }
    throw "Failed to delete '$Path' after $MaxRetries attempts."
}

function Remove-DatabaseFilesOnServer {
    param(
        [Parameter(Mandatory)][string]$MailboxServer,
        [Parameter(Mandatory)][string]$EdbFolder,
        [Parameter(Mandatory)][string]$EdbFileName,
        [Parameter(Mandatory)][string]$LogFolder,
        [Parameter(Mandatory)][string]$DatabaseName
    )

    $edbUnc = ConvertTo-AdminShareUnc -ComputerName $MailboxServer -LocalPath $EdbFolder
    $logUnc = ConvertTo-AdminShareUnc -ComputerName $MailboxServer -LocalPath $LogFolder
    $edbGone = -not (Test-Path -LiteralPath $edbUnc)
    $logGone = -not (Test-Path -LiteralPath $logUnc)
    if ($edbGone -and $logGone) {
        Write-Log "Files already absent on ${MailboxServer}: $edbUnc / $logUnc"
        return
    }

    $verified = Assert-CopyFilesOnServer -MailboxServer $MailboxServer -DatabaseName $DatabaseName `
        -EdbFolder $EdbFolder -EdbFileName $EdbFileName -LogFolder $LogFolder
    Remove-FolderWithRetry -Path $verified.EdbUnc -MaxRetries $FileDeleteMaxRetries -WaitSeconds $FileDeleteRetryWaitSeconds
    Remove-FolderWithRetry -Path $verified.LogUnc -MaxRetries $FileDeleteMaxRetries -WaitSeconds $FileDeleteRetryWaitSeconds
}

function Test-IsInteractiveHost {
    try {
        return [Environment]::UserInteractive -and
            $Host.Name -ne 'ServerRemoteHost' -and
            -not [Console]::IsInputRedirected
    } catch {
        return [Environment]::UserInteractive
    }
}

function Confirm-LastCopyRemoval {
    param(
        [Parameter(Mandatory)][string]$DatabaseName,
        [Parameter(Mandatory)][string]$CopyName,
        [Parameter(Mandatory)][string]$MailboxServer
    )

    if ($WhatIfPreference) {
        Write-Log "WhatIf: would require double confirmation to remove last copy $CopyName on $MailboxServer"
        return $true
    }
    if ($Force) {
        Write-Log "Force: skipping double confirmation for last copy $CopyName"
        return $true
    }
    if (-not (Test-IsInteractiveHost)) {
        throw "Last copy of '$DatabaseName' requires interactive double confirmation or -Force (non-interactive host)."
    }

    Write-Host ''
    Write-Host "WARNING: You are about to DISMOUNT and DELETE the LAST copy of '$DatabaseName'." -ForegroundColor Red
    Write-Host "Copy:   $CopyName"
    Write-Host "Server: $MailboxServer"
    Write-Host 'This removes the database from Active Directory and deletes its files.'
    Write-Host ''

    $first = Read-Host "Confirm 1/2: re-type the database name '$DatabaseName'"
    if ($first -cne $DatabaseName) {
        throw "First confirmation failed for '$DatabaseName' (typed '$first')."
    }

    $second = Read-Host "Confirm 2/2: type DELETE to permanently remove the last copy"
    if ($second -ne 'DELETE') {
        throw "Second confirmation failed for '$DatabaseName' (typed '$second', expected DELETE)."
    }

    Write-Log "Double confirmation accepted for last copy $CopyName"
    return $true
}

function Invoke-ConfirmedOperation {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory)][string]$Target,
        [Parameter(Mandatory)][string]$Action,
        [Parameter(Mandatory)][scriptblock]$Operation
    )
    if ($PSCmdlet.ShouldProcess($Target, $Action)) {
        & $Operation
        return $true
    }
    if ($WhatIfPreference) {
        Write-Log "WhatIf: $Action — $Target"
        return $false
    }
    throw "Operation cancelled: $Action — $Target"
}

# -----------------------------------------------------------------------------
Initialize-ExchangeShell

$dc = Get-PreferredDomainController
Write-Log "Preferred DC: $dc"
Set-AdServerSettings -PreferredServer $dc -ViewEntireForest $true

if (-not (Test-Path -LiteralPath $DatabaseListPath)) {
    throw "Database list file not found: $DatabaseListPath"
}

# @() is required: Get-Content of a one-line file returns a string, and
# foreach on a string iterates characters, not the database name.
$databases = @(
    Get-Content -LiteralPath $DatabaseListPath -ErrorAction Stop |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -and $_ -notmatch '^\s*#' }
)
if ($databases.Count -eq 0) {
    throw "No database names in $DatabaseListPath"
}
Write-Log "Loaded $($databases.Count) database(s) from $DatabaseListPath"
if ($ContinueOnError) {
    Write-Log 'ContinueOnError is enabled: a failure on one database will not stop the rest.'
}
if ($WhatIfPreference) {
    Write-Log 'WhatIf is enabled: no copies or files will be removed.'
}

foreach ($dbToRemove in $databases) {
    try {
        Write-Log "==== Processing database $dbToRemove ===="

        Assert-DatabaseHasNoMailboxes -DatabaseName $dbToRemove -Stage 'initial'

        $folders = Get-DatabaseFileFolders -DatabaseName $dbToRemove
        Write-Log "EdbFolder=$($folders.EdbFolder) EdbFile=$($folders.EdbFileName) LogFolder=$($folders.LogFolder)"

        $allCopies = @(Get-MailboxDatabaseCopyStatus -Identity $dbToRemove)
        $allCopies | Format-List * | Out-File -FilePath $HistoryLogPath -Encoding utf8 -Append

        $passiveCopies = @($allCopies | Where-Object { $_.Status.ToString() -ne 'Mounted' })
        Write-Log "Passive copies to remove: $(@($passiveCopies | ForEach-Object { $_.Name }) -join ', ')"

        Assert-DatabaseHasNoMailboxes -DatabaseName $dbToRemove -Stage 'before copy removal'

        foreach ($copy in $passiveCopies) {
            Write-Log "Preparing passive copy $($copy.Name) on $($copy.MailboxServer)"
            $null = Assert-CopyFilesOnServer -MailboxServer $copy.MailboxServer -DatabaseName $dbToRemove `
                -EdbFolder $folders.EdbFolder -EdbFileName $folders.EdbFileName -LogFolder $folders.LogFolder

            $copyAction = "Remove passive mailbox database copy and delete files on $($copy.MailboxServer)"
            $didCopy = Invoke-ConfirmedOperation -Target $copy.Name -Action $copyAction -Operation {
                Remove-MailboxDatabaseCopy -Identity $copy.Name -Confirm:$false
                Wait-ForCondition -Description "copy $($copy.Name) removed from copy status" `
                    -TimeoutSeconds $CopyRemovalTimeoutSeconds -PollSeconds $PollIntervalSeconds `
                    -Condition {
                        param($DatabaseName, $CopyName)
                        Test-CopyStatusAbsent -DatabaseName $DatabaseName -CopyName $CopyName
                    } -ArgumentList $dbToRemove, $copy.Name
                Remove-DatabaseFilesOnServer -MailboxServer $copy.MailboxServer -DatabaseName $dbToRemove `
                    -EdbFolder $folders.EdbFolder -EdbFileName $folders.EdbFileName -LogFolder $folders.LogFolder
            }
            if (-not $didCopy -and $WhatIfPreference) {
                Write-Log "WhatIf: would poll until $($copy.Name) is gone, then delete verified files on $($copy.MailboxServer)"
            }
        }

        $remaining = @(Get-MailboxDatabaseCopyStatus -Identity $dbToRemove -ErrorAction SilentlyContinue)
        $leftoverPassive = @($remaining | Where-Object { $_.Status.ToString() -ne 'Mounted' })
        $mountedCopies = @($remaining | Where-Object { $_.Status.ToString() -eq 'Mounted' })

        if ($WhatIfPreference) {
            if ($mountedCopies.Count -ne 1) {
                Write-Log "WhatIf: expected one Mounted copy of $dbToRemove after removing passives; found $($mountedCopies.Count)." 'WARN'
            }
        } else {
            if ($leftoverPassive.Count -gt 0) {
                $leftoverNames = @($leftoverPassive | ForEach-Object { $_.Name }) -join ', '
                throw "Not all database copies were successfully deleted: $leftoverNames"
            }
            if ($remaining.Count -ne 1) {
                throw "Expected exactly one mounted copy of $dbToRemove, found $($remaining.Count)."
            }
        }

        if ($mountedCopies.Count -eq 0) {
            throw "No mounted copy of $dbToRemove is available for last-copy removal."
        }

        $lastCopy = $mountedCopies[0]
        Assert-DatabaseHasNoMailboxes -DatabaseName $dbToRemove -Stage 'before last copy removal'
        $null = Assert-CopyFilesOnServer -MailboxServer $lastCopy.MailboxServer -DatabaseName $dbToRemove `
            -EdbFolder $folders.EdbFolder -EdbFileName $folders.EdbFileName -LogFolder $folders.LogFolder

        Write-Log "Preparing LASTCOPY $($lastCopy.Name) (status=$($lastCopy.Status))"
        $null = Confirm-LastCopyRemoval -DatabaseName $dbToRemove -CopyName $lastCopy.Name `
            -MailboxServer $lastCopy.MailboxServer

        $lastAction = "Dismount and remove LAST mailbox database copy, then delete files on $($lastCopy.MailboxServer)"
        $didLast = Invoke-ConfirmedOperation -Target $lastCopy.Name -Action $lastAction -Operation {
            Dismount-Database -Identity $lastCopy.DatabaseName -Confirm:$false
            Wait-ForCondition -Description "database $($lastCopy.DatabaseName) dismounted" `
                -TimeoutSeconds $DismountTimeoutSeconds -PollSeconds $PollIntervalSeconds `
                -Condition {
                    param($CopyName)
                    Test-DatabaseDismounted -CopyName $CopyName
                } -ArgumentList $lastCopy.Name
            Remove-MailboxDatabase -Identity $lastCopy.DatabaseName -Confirm:$false
            Wait-ForCondition -Description "database $($lastCopy.DatabaseName) removed from AD" `
                -TimeoutSeconds $DatabaseRemovalTimeoutSeconds -PollSeconds $PollIntervalSeconds `
                -Condition {
                    param($DatabaseName)
                    Test-MailboxDatabaseAbsent -DatabaseName $DatabaseName
                } -ArgumentList $lastCopy.DatabaseName
            Remove-DatabaseFilesOnServer -MailboxServer $lastCopy.MailboxServer -DatabaseName $dbToRemove `
                -EdbFolder $folders.EdbFolder -EdbFileName $folders.EdbFileName -LogFolder $folders.LogFolder
        }
        if (-not $didLast -and $WhatIfPreference) {
            Write-Log "WhatIf: would dismount/remove $($lastCopy.Name) and delete verified files on $($lastCopy.MailboxServer)"
        }

        [void]$script:SucceededDatabases.Add($dbToRemove)
        Write-Log "==== Finished database $dbToRemove ===="
    } catch {
        Add-OperationError -Database $dbToRemove -Stage 'Processing' -Message $_.Exception.Message
        if (-not $ContinueOnError) {
            Write-Log 'ContinueOnError is off; remaining databases will not be processed.' 'WARN'
            break
        }
    }
}

Write-Log "Succeeded: $($script:SucceededDatabases.Count); errors: $($script:OperationErrors.Count)"
if ($script:SucceededDatabases.Count -gt 0) {
    Write-Log ("OK: " + ($script:SucceededDatabases -join ', '))
}
if ($script:OperationErrors.Count -gt 0) {
    foreach ($err in $script:OperationErrors) {
        Write-Log ("FAIL {0} [{1}] {2}" -f $err.Database, $err.Stage, $err.Message) 'ERROR'
    }
    exit 1
}

Write-Log 'All listed databases processed.'
exit 0
