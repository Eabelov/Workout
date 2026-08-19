#Requires -Version 5.1
<#
.SYNOPSIS
    Устанавливает и запускает службу filebeat-smtp на серверах Exchange.

.DESCRIPTION
    Получает список серверов Exchange (или принимает имена явно), на каждом
    удалённо выполняет install-service-filebeat-smtp.ps1 и гарантирует, что
    служба filebeat-smtp запущена и имеет тип запуска Automatic.

    Скрипт идемпотентен: если служба уже есть, установщик не вызывается повторно.
    Ошибка на одном сервере не прерывает обработку остальных.

    Запускайте из Exchange Management Shell (или сессии, где доступен
    Get-ExchangeServer), с правами, достаточными для WinRM и установки службы.

.PARAMETER ComputerName
    Имена или FQDN серверов. Если не задано — берётся Get-ExchangeServer.

.PARAMETER InstallPath
    Каталог Filebeat SMTP на целевом сервере.

.PARAMETER InstallScriptName
    Имя скрипта установки службы внутри InstallPath.

.PARAMETER ServiceName
    Имя службы Windows.

.PARAMETER ServerRole
    Необязательный фильтр ролей Get-ExchangeServer (например Mailbox).
    Если не задан, обрабатываются все серверы.

.PARAMETER Credential
    Учётные данные для Invoke-Command. По умолчанию — текущий пользователь.

.PARAMETER ThrottleLimit
    Сколько удалённых сессий запускать параллельно.

.PARAMETER TimeoutSec
    Таймаут открытия WinRM-сессии в секундах.

.PARAMETER RestartIfRunning
    Перезапустить службу, даже если она уже работает.

.PARAMETER SkipInstall
    Не вызывать скрипт установки, только проверить и запустить службу.

.EXAMPLE
    .\Install-FilebeatSmtp.ps1

    Установить и запустить службу на всех серверах Exchange.

.EXAMPLE
    .\Install-FilebeatSmtp.ps1 -ComputerName 'mbx01.contoso.com','mbx02.contoso.com' -Verbose

    Обработать только указанные серверы с подробным логом.

.EXAMPLE
    .\Install-FilebeatSmtp.ps1 -WhatIf

    Показать, какие серверы были бы обработаны, без изменений.

.OUTPUTS
    PSCustomObject с полями ComputerName, Success, ServiceStatus, Action, Message.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [Alias('Name', 'Fqdn', 'CN')]
    [string[]]$ComputerName,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$InstallPath = 'C:\Program Files\filebeat-smtp',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$InstallScriptName = 'install-service-filebeat-smtp.ps1',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$ServiceName = 'filebeat-smtp',

    [Parameter()]
    [string[]]$ServerRole,

    [Parameter()]
    [System.Management.Automation.PSCredential]$Credential,

    [Parameter()]
    [ValidateRange(1, 32)]
    [int]$ThrottleLimit = 5,

    [Parameter()]
    [ValidateRange(5, 600)]
    [int]$TimeoutSec = 30,

    [Parameter()]
    [switch]$RestartIfRunning,

    [Parameter()]
    [switch]$SkipInstall
)

begin {
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    function Get-TargetExchangeServer {
        param(
            [string[]]$ComputerName,
            [string[]]$ServerRole
        )

        if ($ComputerName -and $ComputerName.Count -gt 0) {
            return $ComputerName |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                ForEach-Object { $_.Trim() } |
                Select-Object -Unique
        }

        if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
            throw @'
Командлет Get-ExchangeServer недоступен. Запустите скрипт из Exchange Management Shell
либо передайте серверы явно: -ComputerName server1,server2
'@
        }

        $servers = @(Get-ExchangeServer)
        if ($ServerRole -and $ServerRole.Count -gt 0) {
            $servers = @(
                $servers | Where-Object {
                    $roleText = $_.ServerRole.ToString()
                    $matched = $false
                    foreach ($role in $ServerRole) {
                        if ($roleText -match [regex]::Escape($role)) {
                            $matched = $true
                            break
                        }
                    }
                    $matched
                }
            )
        }

        if ($servers.Count -eq 0) {
            throw 'Не найдено серверов Exchange, подходящих под заданный фильтр.'
        }

        foreach ($server in $servers) {
            if ($server.PSObject.Properties['Fqdn'] -and -not [string]::IsNullOrWhiteSpace($server.Fqdn)) {
                $server.Fqdn
            }
            else {
                $server.Name
            }
        }
    }

    function New-ResultObject {
        param(
            [string]$ComputerName,
            [bool]$Success,
            [string]$ServiceStatus,
            [string]$Action,
            [string]$Message
        )

        [pscustomobject]@{
            ComputerName  = $ComputerName
            Success       = $Success
            ServiceStatus = $ServiceStatus
            Action        = $Action
            Message       = $Message
            TimeStamp     = Get-Date
        }
    }

    $remoteScript = {
        param(
            [string]$InstallPath,
            [string]$InstallScriptName,
            [string]$ServiceName,
            [bool]$RestartIfRunning,
            [bool]$SkipInstall
        )

        $ErrorActionPreference = 'Stop'
        Set-StrictMode -Version Latest

        $result = [pscustomobject]@{
            ComputerName  = $env:COMPUTERNAME
            Success       = $false
            ServiceStatus = 'Unknown'
            Action        = 'None'
            Message       = ''
        }

        try {
            $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue

            if (-not $SkipInstall -and -not $service) {
                if (-not (Test-Path -LiteralPath $InstallPath -PathType Container)) {
                    throw "Каталог не найден: $InstallPath"
                }

                $scriptPath = Join-Path -Path $InstallPath -ChildPath $InstallScriptName
                if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) {
                    throw "Скрипт установки не найден: $scriptPath"
                }

                Unblock-File -LiteralPath $scriptPath -ErrorAction SilentlyContinue

                # Установщик Filebeat часто ожидает текущий каталог = папка с filebeat.exe.
                Push-Location -LiteralPath $InstallPath
                try {
                    $installOutput = & powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $scriptPath 2>&1
                    if ($LASTEXITCODE -and $LASTEXITCODE -ne 0) {
                        $detail = ($installOutput | Out-String).Trim()
                        throw "Скрипт установки завершился с кодом ${LASTEXITCODE}. $detail"
                    }
                }
                finally {
                    Pop-Location
                }

                $service = Get-Service -Name $ServiceName -ErrorAction Stop
                $result.Action = 'Installed'
            }
            elseif ($SkipInstall -and -not $service) {
                throw "Служба $ServiceName не найдена, а -SkipInstall запрещает установку."
            }
            else {
                $result.Action = 'AlreadyInstalled'
            }

            Set-Service -Name $ServiceName -StartupType Automatic -ErrorAction Stop
            $service = Get-Service -Name $ServiceName -ErrorAction Stop

            if ($RestartIfRunning -and $service.Status -eq 'Running') {
                Restart-Service -Name $ServiceName -Force -ErrorAction Stop
                $result.Action = if ($result.Action -eq 'Installed') { 'InstalledAndRestarted' } else { 'Restarted' }
            }
            elseif ($service.Status -ne 'Running') {
                Start-Service -Name $ServiceName -ErrorAction Stop
                $result.Action = if ($result.Action -eq 'Installed') { 'InstalledAndStarted' } else { 'Started' }
            }
            else {
                if ($result.Action -eq 'AlreadyInstalled') {
                    $result.Action = 'AlreadyRunning'
                }
            }

            $service = Get-Service -Name $ServiceName -ErrorAction Stop
            $service.WaitForStatus([System.ServiceProcess.ServiceControllerStatus]::Running, [timespan]::FromSeconds(45))
            $result.ServiceStatus = $service.Status.ToString()
            if ($result.ServiceStatus -ne 'Running') {
                throw "Служба $ServiceName в состоянии $($result.ServiceStatus), ожидалось Running."
            }

            $result.Success = $true
            $result.Message = "Служба $ServiceName запущена на $env:COMPUTERNAME"
        }
        catch {
            $status = 'Unknown'
            try {
                $failedService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
                if ($failedService) {
                    $status = $failedService.Status.ToString()
                }
            }
            catch {
                # Состояние службы недоступно — оставляем Unknown.
            }

            $result.Success = $false
            $result.ServiceStatus = $status
            $result.Message = $_.Exception.Message
        }

        $result
    }

    $targets = New-Object System.Collections.Generic.List[string]
    $results = New-Object System.Collections.Generic.List[object]
}

process {
    foreach ($name in @($ComputerName)) {
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            [void]$targets.Add($name.Trim())
        }
    }
}

end {
    try {
        $resolved = @(Get-TargetExchangeServer -ComputerName ($targets.ToArray()) -ServerRole $ServerRole)
    }
    catch {
        Write-Error -ErrorRecord $_
        return
    }

    Write-Verbose "К обработке: $($resolved -join ', ')"

    if (-not $PSCmdlet.ShouldProcess(($resolved -join ', '), "Установить и запустить службу $ServiceName")) {
        foreach ($computer in $resolved) {
            $results.Add((New-ResultObject -ComputerName $computer -Success $true -ServiceStatus 'Skipped' -Action 'WhatIf' -Message 'Операция пропущена из-за -WhatIf'))
        }
        $results
        return
    }

    $sessionOption = New-PSSessionOption `
        -OpenTimeout ([timespan]::FromSeconds($TimeoutSec)) `
        -CancelTimeout ([timespan]::FromSeconds($TimeoutSec)) `
        -OperationTimeout ([timespan]::FromMinutes(5))

    $invokeParams = @{
        ComputerName   = $resolved
        ScriptBlock    = $remoteScript
        ArgumentList   = @($InstallPath, $InstallScriptName, $ServiceName, [bool]$RestartIfRunning, [bool]$SkipInstall)
        ThrottleLimit  = $ThrottleLimit
        SessionOption  = $sessionOption
        ErrorAction    = 'SilentlyContinue'
        ErrorVariable  = 'remoteErrors'
    }
    if ($Credential) {
        $invokeParams.Credential = $Credential
    }

    $remoteResults = @()
    try {
        $remoteResults = @(Invoke-Command @invokeParams)
    }
    catch {
        Write-Warning "Пакетный Invoke-Command завершился с ошибкой: $($_.Exception.Message). Повтор по одному серверу."
    }

    $seen = @{}
    foreach ($item in $remoteResults) {
        if (-not $item) { continue }
        $computer = $item.ComputerName
        $seen[$computer] = $true
        $results.Add((New-ResultObject -ComputerName $computer -Success $item.Success -ServiceStatus $item.ServiceStatus -Action $item.Action -Message $item.Message))
    }

    $failedNames = New-Object System.Collections.Generic.List[string]
    if ($remoteErrors) {
        foreach ($err in $remoteErrors) {
            $computer = $null
            if ($err.PSObject.Properties['OriginInfo'] -and $err.OriginInfo.PSComputerName) {
                $computer = $err.OriginInfo.PSComputerName
            }
            elseif ($err.TargetObject) {
                $computer = [string]$err.TargetObject
            }

            if ($computer) {
                $seen[$computer] = $true
                $results.Add((New-ResultObject -ComputerName $computer -Success $false -ServiceStatus 'Unreachable' -Action 'Failed' -Message $err.Exception.Message))
            }
            else {
                Write-Warning $err.Exception.Message
            }
        }
    }

    foreach ($computer in $resolved) {
        $already = $false
        foreach ($key in $seen.Keys) {
            if ($key -eq $computer -or $computer.StartsWith("$key.", [System.StringComparison]::OrdinalIgnoreCase)) {
                $already = $true
                break
            }
            if ($key.StartsWith("$computer.", [System.StringComparison]::OrdinalIgnoreCase)) {
                $already = $true
                break
            }
        }
        if (-not $already) {
            [void]$failedNames.Add($computer)
        }
    }

    foreach ($computer in $failedNames) {
        Write-Verbose "Повторный запуск для $computer"
        $singleParams = @{
            ComputerName  = $computer
            ScriptBlock   = $remoteScript
            ArgumentList  = @($InstallPath, $InstallScriptName, $ServiceName, [bool]$RestartIfRunning, [bool]$SkipInstall)
            SessionOption = $sessionOption
            ErrorAction   = 'Stop'
        }
        if ($Credential) {
            $singleParams.Credential = $Credential
        }

        try {
            $item = Invoke-Command @singleParams
            $results.Add((New-ResultObject -ComputerName $computer -Success $item.Success -ServiceStatus $item.ServiceStatus -Action $item.Action -Message $item.Message))
        }
        catch {
            $results.Add((New-ResultObject -ComputerName $computer -Success $false -ServiceStatus 'Unreachable' -Action 'Failed' -Message $_.Exception.Message))
        }
    }

    $output = @($results)
    $succeeded = @($output | Where-Object Success).Count
    $failed = $output.Count - $succeeded

    Write-Host ''
    Write-Host "Готово. Успешно: $succeeded, ошибок: $failed (всего $($output.Count))"
    if ($failed -gt 0) {
        Write-Warning 'Есть серверы с ошибками. Смотрите столбец Message в результате.'
    }

    $output | Sort-Object ComputerName
}
