#requires -Version 5.1
<#
.SYNOPSIS
    Removes Exchange mailbox database copies (passive first), then the database
    itself, and deletes leftover .db / .log folders on each DAG node.

.DESCRIPTION
    Reads database names from a text file (one per line) and for each database:
      1. Aborts if any mailbox still lives on the database (user, archive,
         arbitration, audit, migration, public folder, monitoring).
      2. Removes every non-Mounted copy and deletes its files via admin UNC.
      3. Dismounts and removes the last (active) copy, then deletes its files.

    Original script bugs that this rewrite fixes:
      - Curly/smart quotes (“ ”) around -replace broke parsing.
      - Invoke-Expression + a space (' ') hack to list user mailboxes.
      - Database names with spaces were not quoted in the IE string.
      - Get-Content of a single-line file made foreach iterate characters.
      - Get-MailboxDatabaseCopyStatus has no DatabaseVolumeMountPoint /
        LogVolumeMountPoint — those properties are empty, so UNC paths were wrong.
        Real paths come from Get-MailboxDatabase (EdbFilePath / LogFolderPath)
        and are identical on every DAG copy.
      - Last-copy UNC path omitted the backslash before "<db>.db" / "<db>.log".
      - .Replace('C:','C$') is case-sensitive and only handles the C: drive.
      - while (Test-Path) + Remove-Item with no retry cap loops forever if a
        file is locked; Remove-Item also errored on the path that was already gone.
      - Set-MailboxDatabase -CircularLoggingEnabled $false does nothing useful
        here and on some versions requires a dismount to take effect.
      - Add-PSSnapin *exch* can load Setup/Support snapins and collide with EMS.
      - No -ResultSize / -Monitoring / missing-parameter handling on Get-Mailbox.
      - $lastcopy could be a collection; identities are taken from the DB name.

.NOTES
    Run from Exchange Management Shell (or a session with EMS snap-in loaded)
    as an Organization Management admin that can reach C$/D$/... on DAG nodes.
#>
[CmdletBinding()]
param(
    [string]$DatabaseListPath = 'C:\BEA\dag4databases.txt',
    [string]$HistoryLogPath   = 'C:\BEA\dbcopystatushistory.txt',
    [int]$CopyRemovalWaitSeconds      = 120,
    [int]$LastCopyDismountWaitSeconds = 30,
    [int]$FileDeleteRetryWaitSeconds  = 60,
    [int]$FileDeleteMaxRetries        = 10
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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
        EdbFolder = Split-Path -Parent $edbFile
        LogFolder = $logDir
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

function Remove-FolderWithRetry {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][int]$MaxRetries,
        [Parameter(Mandatory)][int]$WaitSeconds
    )

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
        [Parameter(Mandatory)][string]$LogFolder
    )

    $dbUnc  = ConvertTo-AdminShareUnc -ComputerName $MailboxServer -LocalPath $EdbFolder
    $logUnc = ConvertTo-AdminShareUnc -ComputerName $MailboxServer -LocalPath $LogFolder
    Write-Log "File paths on ${MailboxServer}: db=$dbUnc log=$logUnc"
    Remove-FolderWithRetry -Path $dbUnc  -MaxRetries $FileDeleteMaxRetries -WaitSeconds $FileDeleteRetryWaitSeconds
    Remove-FolderWithRetry -Path $logUnc -MaxRetries $FileDeleteMaxRetries -WaitSeconds $FileDeleteRetryWaitSeconds
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

foreach ($dbToRemove in $databases) {
    try {
        Write-Log "==== Processing database $dbToRemove ===="

        $occupiedBy = @(Test-DatabaseHasMailboxes -DatabaseName $dbToRemove)
        if ($occupiedBy.Count -gt 0) {
            Write-Log "Database $dbToRemove is not empty (found: $($occupiedBy -join ', ')). Stopping." 'ERROR'
            exit 1
        }

        $folders = Get-DatabaseFileFolders -DatabaseName $dbToRemove
        Write-Log "EdbFolder=$($folders.EdbFolder) LogFolder=$($folders.LogFolder)"

        $allCopies = @(Get-MailboxDatabaseCopyStatus -Identity $dbToRemove)
        $allCopies | Format-List * | Out-File -FilePath $HistoryLogPath -Encoding utf8 -Append

        $passiveCopies = @($allCopies | Where-Object { $_.Status.ToString() -ne 'Mounted' })
        Write-Log "Passive copies to remove: $(@($passiveCopies | ForEach-Object { $_.Name }) -join ', ')"

        foreach ($copy in $passiveCopies) {
            Write-Log "DELETING copy $($copy.Name) on $($copy.MailboxServer)"
            Remove-MailboxDatabaseCopy -Identity $copy.Name -Confirm:$false
            Start-Sleep -Seconds $CopyRemovalWaitSeconds
            Remove-DatabaseFilesOnServer -MailboxServer $copy.MailboxServer `
                -EdbFolder $folders.EdbFolder -LogFolder $folders.LogFolder
        }

        $remaining = @(Get-MailboxDatabaseCopyStatus -Identity $dbToRemove)
        $leftoverPassive = @($remaining | Where-Object { $_.Status.ToString() -ne 'Mounted' })
        if ($leftoverPassive.Count -gt 0) {
            $leftoverNames = @($leftoverPassive | ForEach-Object { $_.Name }) -join ', '
            Write-Log "Not all database copies were successfully deleted: $leftoverNames" 'ERROR'
            exit 1
        }
        if ($remaining.Count -ne 1) {
            Write-Log "Expected exactly one mounted copy of $dbToRemove, found $($remaining.Count)." 'ERROR'
            exit 1
        }

        $lastCopy = $remaining[0]
        Write-Log "DELETING LASTCOPY $($lastCopy.Name) (status=$($lastCopy.Status))"
        Dismount-Database -Identity $lastCopy.DatabaseName -Confirm:$false
        Start-Sleep -Seconds $LastCopyDismountWaitSeconds
        Remove-MailboxDatabase -Identity $lastCopy.DatabaseName -Confirm:$false
        Start-Sleep -Seconds $CopyRemovalWaitSeconds
        Remove-DatabaseFilesOnServer -MailboxServer $lastCopy.MailboxServer `
            -EdbFolder $folders.EdbFolder -LogFolder $folders.LogFolder

        Write-Log "==== Finished database $dbToRemove ===="
    } catch {
        Write-Log "Failed while processing '$dbToRemove': $($_.Exception.Message)" 'ERROR'
        throw
    }
}

Write-Log 'All listed databases processed.'
exit 0
