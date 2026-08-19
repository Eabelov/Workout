<#
.SYNOPSIS
    Копирует каталог Filebeat на серверы Exchange и при необходимости ставит службу.

.DESCRIPTION
    Copy-Item -Recurse с маской '*' и заранее созданной папкой назначения — известный
    баг Windows PowerShell 5.1: командлет возвращает успех, а на части серверов
    каталог остаётся пустым (особенно если папку уже создали New-Item).
    -LiteralPath к тому же не раскрывает '*', поэтому копироваться может 0 файлов
    без понятной ошибки.

    Этот скрипт копирует дерево через robocopy и сверяет число файлов с источником.
    Локальный сервер (откуда запускаете) пропускается — файлы там уже есть.

    Если служба уже запущена, filebeat.exe держит файлы: robocopy тогда
    пропускает занятые бинарники или падает. Перед копированием служба
    останавливается, после копирования запускается снова, если она была
    Running. Повторно install-service-*.ps1 не вызывается — New-Service
    упрётся в «служба уже существует».

.PARAMETER Source
    Локальный каталог Filebeat. По умолчанию C:\Program Files\filebeat-smtp.

.PARAMETER Servers
    Имена или FQDN. Если не заданы — все серверы из Get-ExchangeServer.

.PARAMETER ExchangeUri
    URI удалённого PowerShell Exchange, если скрипт запущен не из EMS.

.PARAMETER InstallService
    Если службы ещё нет — выполнить install-service-*.ps1 и запустить её.
    Если служба уже есть, установщик не трогается.

.PARAMETER NoServiceRestart
    Не останавливать и не запускать службу. Имеет смысл только если она
    точно не запущена: иначе filebeat.exe останется залочен.

.PARAMETER ServiceName
    Имя службы Windows. По умолчанию совпадает с именем каталога источника.

.PARAMETER Credential
    Учётные данные для Invoke-Command. Для C$ и Get-Service -ComputerName
    обычно достаточно текущего токена администратора.

.EXAMPLE
    .\Copy-FilebeatToExchange.ps1

.EXAMPLE
    .\Copy-FilebeatToExchange.ps1 -InstallService

.EXAMPLE
    # Служба уже запущена: скрипт остановит её, скопирует файлы и запустит снова.
    .\Copy-FilebeatToExchange.ps1

.EXAMPLE
    .\Copy-FilebeatToExchange.ps1 -Source 'C:\Program Files\filebeat-iis' -InstallService
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter()]
    [string]$Source = $(Join-Path $env:ProgramFiles 'filebeat-smtp'),

    [Parameter()]
    [string[]]$Servers,

    [Parameter()]
    [string]$ExchangeUri,

    [Parameter()]
    [switch]$InstallService,

    [Parameter()]
    [switch]$NoServiceRestart,

    [Parameter()]
    [string]$ServiceName,

    [Parameter()]
    [System.Management.Automation.PSCredential]$Credential
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Connect-ExchangeIfNeeded {
    param([string]$Uri)

    if (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue) {
        Write-Verbose 'Командлеты Exchange уже загружены.'
        return $null
    }

    foreach ($name in @(
            'Microsoft.Exchange.Management.PowerShell.SnapIn',
            'Microsoft.Exchange.Management.PowerShell.E2010'
        )) {
        if (Get-PSSnapin -Registered -Name $name -ErrorAction SilentlyContinue) {
            Add-PSSnapin $name -ErrorAction Stop
            return $null
        }
    }

    if (-not $Uri) {
        throw @'
Не найдены командлеты Exchange. Запустите скрипт из Exchange Management Shell
или передайте -ExchangeUri, например:
  -ExchangeUri http://exch01.contoso.local/PowerShell/
'@
    }

    Write-Host "Подключение к Exchange: $Uri"
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $Uri -Authentication Kerberos
    Import-PSSession $session -DisableNameChecking | Out-Null
    return $session
}

function Test-IsLocalServer {
    param([string]$Name)

    if (-not $Name) { return $false }

    $short = $Name.Split('.')[0]
    $candidates = @(
        $env:COMPUTERNAME
        [System.Net.Dns]::GetHostName()
    )
    if ($env:USERDNSDOMAIN) {
        $candidates += "$($env:COMPUTERNAME).$($env:USERDNSDOMAIN)"
    }

    foreach ($candidate in $candidates) {
        if (-not $candidate) { continue }
        if ($Name -eq $candidate) { return $true }
        if ($short -eq $candidate.Split('.')[0]) { return $true }
    }

    return $false
}

function Get-AdminShareRoot {
    param([string]$ComputerName)

    # C$ в одинарных кавычках: в двойных $ может стать началом переменной.
    $root = '\\{0}\C$\Program Files' -f $ComputerName
    if (Test-Path -LiteralPath $root) {
        return $root
    }
    return $null
}

function Resolve-RemoteComputer {
    param(
        [string]$Name,
        [string]$Fqdn
    )

    $candidates = @()
    if ($Fqdn) { $candidates += $Fqdn }
    if ($Name -and ($candidates -notcontains $Name)) { $candidates += $Name }

    foreach ($candidate in $candidates) {
        $root = Get-AdminShareRoot -ComputerName $candidate
        if ($root) {
            return [PSCustomObject]@{
                ComputerName = $candidate
                ShareRoot    = $root
            }
        }
    }

    $tried = $candidates -join ', '
    throw "Нет доступа к C`$ (File and Printer Sharing / права администратора). Пробовали: $tried"
}

function Get-FileCount {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return 0
    }

    return @(Get-ChildItem -LiteralPath $Path -Recurse -Force -File -ErrorAction SilentlyContinue).Count
}

function Test-ServiceNotFound {
    param($ErrorRecord)

    $msg = [string]$ErrorRecord.Exception.Message
    return (
        $msg -match 'Cannot find any service' -or
        $msg -match 'No service with' -or
        $msg -match 'не найден' -or
        $ErrorRecord.CategoryInfo.Category -eq 'ObjectNotFound'
    )
}

function Invoke-RemoteServiceAction {
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName,

        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [ValidateSet('Get', 'Stop', 'Start')]
        [string]$Action,

        [System.Management.Automation.PSCredential]$Cred
    )

    $block = {
        param([string]$SvcName, [string]$Act)

        $ErrorActionPreference = 'Stop'
        $svc = Get-Service -Name $SvcName -ErrorAction SilentlyContinue
        if ($Act -eq 'Get') {
            if (-not $svc) { return 'Absent' }
            return $svc.Status.ToString()
        }

        if (-not $svc) {
            if ($Act -eq 'Stop') { return 'Absent' }
            throw "Служба $SvcName не найдена"
        }

        if ($Act -eq 'Stop') {
            if ($svc.Status -ne 'Stopped') {
                Stop-Service -Name $SvcName -Force
                $svc.WaitForStatus('Stopped', '00:00:45')
            }
            return 'Stopped'
        }

        if ($svc.Status -ne 'Running') {
            Start-Service -Name $SvcName
            $svc.WaitForStatus('Running', '00:00:45')
        }
        return 'Running'
    }

    # RPC, те же права что и C$ — WinRM не нужен.
    try {
        $svc = Get-Service -ComputerName $ComputerName -Name $Name -ErrorAction Stop

        if ($Action -eq 'Get') {
            return $svc.Status.ToString()
        }

        if ($Action -eq 'Stop') {
            if ($svc.Status -ne 'Stopped') {
                Stop-Service -InputObject $svc -Force -ErrorAction Stop
                $svc.WaitForStatus(
                    [System.ServiceProcess.ServiceControllerStatus]::Stopped,
                    [timespan]::FromSeconds(45)
                )
            }
            return 'Stopped'
        }

        if ($svc.Status -ne 'Running') {
            Start-Service -InputObject $svc -ErrorAction Stop
            $svc.WaitForStatus(
                [System.ServiceProcess.ServiceControllerStatus]::Running,
                [timespan]::FromSeconds(45)
            )
        }
        return 'Running'
    }
    catch {
        if ($Action -eq 'Get' -and (Test-ServiceNotFound $_)) {
            return 'Absent'
        }
        if ($Action -eq 'Stop' -and (Test-ServiceNotFound $_)) {
            return 'Absent'
        }

        $icm = @{
            ComputerName = $ComputerName
            ScriptBlock  = $block
            ArgumentList = @($Name, $Action)
            ErrorAction  = 'Stop'
        }
        if ($Cred) {
            $icm.Credential = $Cred
        }
        return Invoke-Command @icm
    }
}

function Copy-DirectoryWithRobocopy {
    param(
        [string]$SourcePath,
        [string]$DestinationPath
    )

    $robocopy = Join-Path $env:SystemRoot 'System32\robocopy.exe'
    if (-not (Test-Path -LiteralPath $robocopy)) {
        throw "robocopy.exe не найден: $robocopy"
    }

    # /E  — подкаталоги, включая пустые
    # /COPY:DAT — данные, атрибуты, timestamps (без ACL, они на C$ часто мешают)
    # /R:2 /W:2 — не крутиться минутами на залоченных файлах
    # /XO не используем: нужна полная копия, а не «пропуск более новых»
    $args = @(
        $SourcePath
        $DestinationPath
        '/E'
        '/COPY:DAT'
        '/R:2'
        '/W:2'
        '/NP'
        '/NFL'
        '/NDL'
        '/NJH'
        '/NJS'
    )

    Write-Verbose ("robocopy {0}" -f ($args -join ' '))
    # Не направлять вывод в пайп: у native-команд тогда часто теряется LASTEXITCODE.
    $output = & $robocopy @args
    $code = $LASTEXITCODE

    # robocopy: 0–7 это успех (0 = нечего копировать, 1 = файлы скопированы, …)
    if ($code -ge 8) {
        $tail = @($output | Select-Object -Last 20) -join [Environment]::NewLine
        throw "robocopy завершился с кодом $code (8+ = ошибка) для '$DestinationPath'`n$tail"
    }

    # Чтобы сессия/CI не считали код 1 аварией.
    $global:LASTEXITCODE = 0
}

function Install-FilebeatServiceRemote {
    param(
        [string]$ComputerName,
        [string]$PackageName,
        [string]$WindowsServiceName,
        [System.Management.Automation.PSCredential]$Cred,
        [bool]$IsLocal
    )

    $installBlock = {
        param(
            [string]$Package,
            [string]$SvcName
        )

        $ErrorActionPreference = 'Stop'
        $root = Join-Path $env:ProgramFiles $Package
        $installScript = Join-Path $root "install-service-$Package.ps1"

        if (-not (Test-Path -LiteralPath $installScript)) {
            throw "Не найден скрипт установки: $installScript"
        }

        $existing = Get-Service -Name $SvcName -ErrorAction SilentlyContinue
        if (-not $existing) {
            $p = Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe') -ArgumentList @(
                '-NoProfile'
                '-ExecutionPolicy', 'Bypass'
                '-File', $installScript
            ) -Wait -PassThru -WindowStyle Hidden
            if ($p.ExitCode -ne 0) {
                throw "install-service вернул код $($p.ExitCode)"
            }
        }

        $svc = Get-Service -Name $SvcName -ErrorAction Stop
        if ($svc.Status -ne 'Running') {
            Start-Service -Name $SvcName
            $svc.Refresh()
        }

        [PSCustomObject]@{
            ComputerName = $env:COMPUTERNAME
            Service      = $SvcName
            Status       = $svc.Status.ToString()
        }
    }

    if ($IsLocal) {
        return & $installBlock $PackageName $WindowsServiceName
    }

    $icm = @{
        ComputerName = $ComputerName
        ScriptBlock  = $installBlock
        ArgumentList = @($PackageName, $WindowsServiceName)
        ErrorAction  = 'Stop'
    }
    if ($Cred) {
        $icm.Credential = $Cred
    }

    return Invoke-Command @icm
}

# --- main -----------------------------------------------------------------

if (-not (Test-Path -LiteralPath $Source)) {
    throw "Источник не найден: $Source"
}

$packageName = Split-Path -Path $Source -Leaf
if (-not $ServiceName) {
    $ServiceName = $packageName
}

$sourceFiles = Get-FileCount -Path $Source
if ($sourceFiles -eq 0) {
    throw "Источник пустой, копировать нечего: $Source"
}

$exchangeSession = Connect-ExchangeIfNeeded -Uri $ExchangeUri

try {
    if ($Servers -and $Servers.Count -gt 0) {
        $targets = @(
            $Servers | ForEach-Object {
                [PSCustomObject]@{ Name = $_; Fqdn = $_ }
            }
        )
    }
    else {
        $targets = @(Get-ExchangeServer | Select-Object Name, Fqdn)
        if ($targets.Count -eq 0) {
            throw 'Get-ExchangeServer не вернул серверов.'
        }
    }

    $results = New-Object System.Collections.Generic.List[object]

    foreach ($server in $targets) {
        $display = if ($server.Fqdn) { $server.Fqdn } else { $server.Name }

        if ((Test-IsLocalServer $server.Name) -or (Test-IsLocalServer $server.Fqdn)) {
            Write-Host "SKIP (локальный): $display"
            $results.Add([PSCustomObject]@{
                    Server  = $display
                    Status  = 'SkippedLocal'
                    Files   = $sourceFiles
                    Message = 'Источник уже на этой машине'
                })
            continue
        }

        $row = [PSCustomObject]@{
            Server  = $display
            Status  = 'Failed'
            Files   = 0
            Message = ''
        }

        $stoppedByScript = $false
        $remote = $null

        try {
            $remote = Resolve-RemoteComputer -Name $server.Name -Fqdn $server.Fqdn
            $dest = Join-Path $remote.ShareRoot $packageName

            $copyTarget = $dest
            if (-not $NoServiceRestart) {
                $copyTarget = "$dest (с остановкой $ServiceName при необходимости)"
            }

            if (-not $PSCmdlet.ShouldProcess($copyTarget, "Копировать $packageName ($sourceFiles файлов)")) {
                $row.Status = 'WhatIf'
                $row.Message = $dest
                $results.Add($row)
                continue
            }

            $previousState = 'Absent'
            if (-not $NoServiceRestart -or $InstallService) {
                $previousState = Invoke-RemoteServiceAction -ComputerName $remote.ComputerName -Name $ServiceName -Action Get -Cred $Credential
            }

            if (-not $NoServiceRestart -and ($previousState -eq 'Running' -or $previousState -eq 'StartPending' -or $previousState -eq 'StopPending')) {
                Write-Host "STOP: $($remote.ComputerName) $ServiceName ($previousState)"
                Invoke-RemoteServiceAction -ComputerName $remote.ComputerName -Name $ServiceName -Action Stop -Cred $Credential | Out-Null
                $stoppedByScript = $true
                Start-Sleep -Seconds 2
            }
            elseif ($previousState -ne 'Absent') {
                Write-Verbose "$($remote.ComputerName): $ServiceName = $previousState, установщик пропускаем"
            }

            Copy-DirectoryWithRobocopy -SourcePath $Source -DestinationPath $dest

            $copied = Get-FileCount -Path $dest
            $row.Files = $copied

            if ($copied -eq 0) {
                throw "Папка назначения пустая после копирования: $dest"
            }
            if ($copied -lt $sourceFiles) {
                throw "Скопировано файлов $copied из $sourceFiles в $dest"
            }

            $row.Status = 'Copied'
            $row.Message = $dest
            Write-Host "OK: $($remote.ComputerName) ($copied файлов)"

            if ($InstallService -and $previousState -eq 'Absent') {
                if ($PSCmdlet.ShouldProcess($remote.ComputerName, "Установить и запустить службу $ServiceName")) {
                    $svcInfo = Install-FilebeatServiceRemote -ComputerName $remote.ComputerName -PackageName $packageName -WindowsServiceName $ServiceName -Cred $Credential -IsLocal $false
                    $row.Status = 'Installed'
                    $row.Message = "$($row.Message); служба $($svcInfo.Status)"
                    Write-Host "SERVICE: $($remote.ComputerName) $ServiceName = $($svcInfo.Status)"
                }
            }
            elseif ($stoppedByScript) {
                $started = Invoke-RemoteServiceAction -ComputerName $remote.ComputerName -Name $ServiceName -Action Start -Cred $Credential
                $row.Status = 'Restarted'
                $row.Message = "$($row.Message); служба $started"
                Write-Host "START: $($remote.ComputerName) $ServiceName = $started"
            }
            elseif ($InstallService -and $previousState -ne 'Absent' -and $previousState -ne 'Running') {
                $started = Invoke-RemoteServiceAction -ComputerName $remote.ComputerName -Name $ServiceName -Action Start -Cred $Credential
                $row.Status = 'Started'
                $row.Message = "$($row.Message); служба $started"
                Write-Host "START: $($remote.ComputerName) $ServiceName = $started"
            }
        }
        catch {
            $row.Status = 'Failed'
            $row.Message = $_.Exception.Message
            Write-Warning "${display}: $($row.Message)"

            if ($stoppedByScript -and $remote) {
                try {
                    Invoke-RemoteServiceAction -ComputerName $remote.ComputerName -Name $ServiceName -Action Start -Cred $Credential | Out-Null
                    Write-Warning "${display}: служба $ServiceName запущена обратно после ошибки копирования"
                }
                catch {
                    Write-Warning "${display}: не удалось вернуть службу $ServiceName : $($_.Exception.Message)"
                }
            }
        }

        $results.Add($row)
    }

    Write-Host ''
    $results | Format-Table -AutoSize

    $failed = @($results | Where-Object { $_.Status -eq 'Failed' })
    if ($failed.Count -gt 0) {
        throw "Не удалось обработать серверов: $($failed.Count). См. таблицу выше."
    }
}
finally {
    if ($exchangeSession) {
        Remove-PSSession $exchangeSession -ErrorAction SilentlyContinue
    }
}
