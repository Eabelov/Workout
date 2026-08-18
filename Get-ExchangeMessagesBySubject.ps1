<#
.SYNOPSIS
    Ищет в message tracking всех серверов Exchange письма, у которых тема содержит заданную строку, и выгружает результат в CSV.

.DESCRIPTION
    Обходит все транспортные / mailbox-серверы и отбирает записи, где тема СОДЕРЖИТ
    шаблон, а не равна ему целиком. По умолчанию ищется символ «&» в теме.

    По умолчанию (SearchMethod Auto) логи читаются напрямую с диска каждого сервера
    параллельно — это быстрее, чем Get-MessageTrackingLog по всем DELIVER.
    Если WinRM/файлы недоступны, сервер обрабатывается командлетом.

    CSV-колонки:
      Дата и время, Статус, Отправитель, Получатель, Тема

.PARAMETER Start
    Начало периода. По умолчанию — 7 дней назад. Чем уже окно, тем быстрее поиск.

.PARAMETER End
    Конец периода. По умолчанию — сейчас.

.PARAMETER SubjectContains
    Подстрока, которая должна встречаться в теме. По умолчанию «&».

.PARAMETER EventId
    Событие message tracking. По умолчанию DELIVER.

.PARAMETER Sender
    Необязательный фильтр отправителя для режима Cmdlet (ускоряет Get-MessageTrackingLog).

.PARAMETER Recipients
    Необязательный фильтр получателя для режима Cmdlet.

.PARAMETER OutputPath
    Путь к CSV. По умолчанию — файл рядом со скриптом с меткой времени.

.PARAMETER Servers
    Список серверов. Если не задан — берутся все серверы с ролью транспорта.

.PARAMETER ExchangeUri
    URI для удалённого подключения к Exchange, если скрипт запущен не из EMS.

.PARAMETER SearchMethod
    Auto (по умолчанию): сначала файлы логов параллельно, неудачные серверы — через командлет.
    LogFile: только файлы MSGTRK*.log.
    Cmdlet: только Get-MessageTrackingLog (как раньше, медленнее).

.PARAMETER ThrottleLimit
    Сколько серверов опрашивать одновременно. По умолчанию 4.

.EXAMPLE
    .\Get-ExchangeMessagesBySubject.ps1

.EXAMPLE
    .\Get-ExchangeMessagesBySubject.ps1 -Start (Get-Date).AddDays(-2) -ThrottleLimit 8

.EXAMPLE
    .\Get-ExchangeMessagesBySubject.ps1 -SearchMethod Cmdlet -Recipients 'user@contoso.com'
#>
[CmdletBinding()]
param(
    [datetime]$Start = (Get-Date).AddDays(-7),

    [datetime]$End = (Get-Date),

    [string]$SubjectContains = '&',

    [string]$EventId = 'DELIVER',

    [string]$Sender,

    [string]$Recipients,

    [string]$OutputPath,

    [string[]]$Servers,

    [string]$ExchangeUri,

    [ValidateSet('Auto', 'LogFile', 'Cmdlet')]
    [string]$SearchMethod = 'Auto',

    [ValidateRange(1, 32)]
    [int]$ThrottleLimit = 4
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Connect-ExchangeIfNeeded {
    param([string]$Uri)

    if (Get-Command Get-MessageTrackingLog -ErrorAction SilentlyContinue) {
        Write-Verbose 'Командлеты Exchange уже загружены.'
        return $null
    }

    $snapins = @(
        'Microsoft.Exchange.Management.PowerShell.SnapIn',
        'Microsoft.Exchange.Management.PowerShell.E2010'
    )
    foreach ($name in $snapins) {
        if (Get-PSSnapin -Registered -Name $name -ErrorAction SilentlyContinue) {
            Add-PSSnapin $name -ErrorAction Stop
            Write-Verbose "Загружен snap-in $name"
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

function Get-TrackingServers {
    param([string[]]$Requested)

    if ($Requested -and $Requested.Count -gt 0) {
        return @($Requested | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }

    if (Get-Command Get-TransportService -ErrorAction SilentlyContinue) {
        $names = @(Get-TransportService | Select-Object -ExpandProperty Name)
        if ($names.Count -gt 0) { return $names }
    }

    $exchangeServers = @(Get-ExchangeServer | Where-Object {
        $_.IsMailboxServer -or $_.IsHubTransportServer
    } | Select-Object -ExpandProperty Name)

    if ($exchangeServers.Count -eq 0) {
        throw 'Не найдены серверы Exchange с ролью Mailbox / Hub Transport.'
    }

    return $exchangeServers
}

function Get-TrackingLogPathMap {
    $map = @{}
    if (-not (Get-Command Get-TransportService -ErrorAction SilentlyContinue)) {
        return $map
    }

    foreach ($svc in @(Get-TransportService)) {
        $path = [string]$svc.MessageTrackingLogPath
        if (-not $path) { continue }
        $map[$svc.Name] = $path
        $short = $svc.Name.Split('.')[0]
        if ($short -and -not $map.ContainsKey($short)) {
            $map[$short] = $path
        }
    }

    return $map
}

function ConvertTo-RecipientText {
    param($Recipients)

    if ($null -eq $Recipients) { return '' }
    if ($Recipients -is [string]) { return $Recipients }

    $list = @($Recipients | ForEach-Object { "$_" } | Where-Object { $_ })
    return ($list -join '; ')
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

function Resolve-ServerLogPath {
    param(
        [string]$Server,
        [hashtable]$PathMap
    )

    if ($PathMap) {
        if ($PathMap.ContainsKey($Server)) { return $PathMap[$Server] }
        $short = $Server.Split('.')[0]
        if ($PathMap.ContainsKey($short)) { return $PathMap[$short] }
    }

    return $null
}

# Используется и локально, и через Invoke-Command на каждом сервере Exchange.
$script:SearchTrackingFilesSb = {
    param(
        [datetime]$Start,
        [datetime]$End,
        [string]$EventId,
        [string]$SubjectContains,
        [string]$LogPath,
        [string]$Sender,
        [string]$Recipients
    )

    $ErrorActionPreference = 'Stop'
    $wildcard = '*' + $SubjectContains + '*'
    $eventToken = ',' + $EventId + ','
    $rows = New-Object System.Collections.Generic.List[object]
    $raw = 0

    if (-not $LogPath -or -not (Test-Path -LiteralPath $LogPath)) {
        foreach ($ver in @('v15', 'v14')) {
            $setupKey = "HKLM:\SOFTWARE\Microsoft\ExchangeServer\$ver\Setup"
            if (-not (Test-Path $setupKey)) { continue }
            $root = [string](Get-ItemProperty -Path $setupKey).MsiInstallPath
            if (-not $root) { continue }
            $candidate = Join-Path $root 'TransportRoles\Logs\MessageTracking'
            if (Test-Path -LiteralPath $candidate) {
                $LogPath = $candidate
                break
            }
        }
    }

    if (-not $LogPath -or -not (Test-Path -LiteralPath $LogPath)) {
        throw "Каталог message tracking не найден: $LogPath"
    }

    $startDate = $Start.Date.AddDays(-1)
    $endDate = $End.Date.AddDays(1)
    $files = @(Get-ChildItem -LiteralPath $LogPath -Filter 'MSGTRK*.log' -File -ErrorAction Stop)

    foreach ($file in $files) {
        $fileDate = $null
        if ($file.Name -match '(\d{8})-\d+\.log$') {
            $fileDate = [datetime]::ParseExact($Matches[1], 'yyyyMMdd', $null)
            if ($fileDate -lt $startDate -or $fileDate -gt $endDate) { continue }
        }

        $header = $null
        $reader = New-Object System.IO.StreamReader($file.FullName, [System.Text.Encoding]::UTF8)
        try {
            while ($null -ne ($line = $reader.ReadLine())) {
                if ($line.StartsWith('#Fields:')) {
                    $header = ($line.Substring(8).Trim() -split ',')
                    continue
                }
                if ($line.StartsWith('#')) { continue }
                if ($line.IndexOf($eventToken, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
                $raw++
                if ($line.IndexOf($SubjectContains, [System.StringComparison]::Ordinal) -lt 0) { continue }
                if (-not $header) { continue }

                $row = $line | ConvertFrom-Csv -Header $header
                if ($row.'event-id' -ne $EventId) { continue }
                if ($row.'message-subject' -notlike $wildcard) { continue }
                if ($Sender -and $row.'sender-address' -ne $Sender) { continue }
                if ($Recipients -and $row.'recipient-address' -notlike ('*' + $Recipients + '*')) { continue }

                $timestamp = [datetime]$row.'date-time'
                if ($timestamp -lt $Start -or $timestamp -gt $End) { continue }

                $rows.Add([pscustomobject]@{
                    Timestamp      = $timestamp.ToString('yyyy-MM-dd HH:mm:ss')
                    EventId        = $row.'event-id'
                    Sender         = $row.'sender-address'
                    Recipients     = $row.'recipient-address'
                    MessageSubject = $row.'message-subject'
                }) | Out-Null
            }
        }
        finally {
            $reader.Close()
        }
    }

    [pscustomobject]@{
        Raw     = $raw
        Rows    = $rows.ToArray()
        LogPath = $LogPath
    }
}

function ConvertTo-ResultRow {
    param($Entry)

    [pscustomobject]@{
        'Дата и время' = $Entry.Timestamp
        'Статус'       = $Entry.EventId
        'Отправитель'  = $Entry.Sender
        'Получатель'   = $Entry.Recipients
        'Тема'         = $Entry.MessageSubject
    }
}

function Search-ServerByCmdlet {
    param(
        [string]$Server,
        [datetime]$Start,
        [datetime]$End,
        [string]$EventId,
        [string]$SubjectContains,
        [string]$Sender,
        [string]$Recipients
    )

    $wildcard = '*' + $SubjectContains + '*'
    $gmtParams = @{
        Server      = $Server
        Start       = $Start
        End         = $End
        EventId     = $EventId
        ErrorAction = 'Stop'
    }
    if ($Sender) { $gmtParams.Sender = $Sender }
    if ($Recipients) { $gmtParams.Recipients = $Recipients }

    $raw = 0
    $rows = New-Object System.Collections.Generic.List[object]

    Get-MessageTrackingLog @gmtParams -ResultSize Unlimited | ForEach-Object {
        $raw++
        if ($_.MessageSubject -notlike $wildcard) { return }
        $rows.Add([pscustomobject]@{
            Timestamp      = $_.Timestamp.ToString('yyyy-MM-dd HH:mm:ss')
            EventId        = $_.EventId
            Sender         = $_.Sender
            Recipients     = ConvertTo-RecipientText -Recipients $_.Recipients
            MessageSubject = $_.MessageSubject
        }) | Out-Null
    }

    [pscustomobject]@{
        Raw  = $raw
        Rows = $rows.ToArray()
    }
}

function Add-ServerRows {
    param(
        $Payload,
        [System.Collections.Generic.List[object]]$Results
    )

    if ($null -eq $Payload -or $null -eq $Payload.Rows) { return }

    foreach ($entry in @($Payload.Rows)) {
        if ($null -eq $entry) { continue }
        $Results.Add((ConvertTo-ResultRow -Entry $entry)) | Out-Null
    }
}

if (-not $OutputPath) {
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $scriptDir = $PSScriptRoot
    if (-not $scriptDir) { $scriptDir = (Get-Location).Path }
    $OutputPath = Join-Path $scriptDir "exchange-deliver-subject_$stamp.csv"
}

$outputDir = Split-Path -Parent $OutputPath
if ($outputDir -and -not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

$exchangeSession = Connect-ExchangeIfNeeded -Uri $ExchangeUri
$serverList = Get-TrackingServers -Requested $Servers
$pathMap = Get-TrackingLogPathMap

Write-Host "Период: $($Start.ToString('yyyy-MM-dd HH:mm:ss')) — $($End.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "Серверы ($($serverList.Count)): $($serverList -join ', ')"
Write-Host "Событие: $EventId"
Write-Host "Тема содержит: $SubjectContains"
Write-Host "Метод: $SearchMethod, параллельность: $ThrottleLimit"
Write-Host "Файл: $OutputPath"
Write-Host ''

$errors = New-Object System.Collections.Generic.List[string]
$results = New-Object System.Collections.Generic.List[object]
$counters = @{ Raw = 0 }
$cmdletFallback = New-Object System.Collections.Generic.List[string]

if ($SearchMethod -ne 'Cmdlet') {
    $queue = New-Object System.Collections.Generic.Queue[string]
    foreach ($server in $serverList) { $queue.Enqueue($server) }

    $jobs = New-Object System.Collections.Generic.List[object]

    while ($queue.Count -gt 0 -or $jobs.Count -gt 0) {
        while ($jobs.Count -lt $ThrottleLimit -and $queue.Count -gt 0) {
            $server = $queue.Dequeue()
            $logPath = Resolve-ServerLogPath -Server $server -PathMap $pathMap
            Write-Host "Сервер $server — чтение MSGTRK-файлов..."

            $searchArgs = @($Start, $End, $EventId, $SubjectContains, $logPath, $Sender, $Recipients)
            if (Test-IsLocalServer -Name $server) {
                $job = Start-Job -Name "trk-$server" -ScriptBlock $script:SearchTrackingFilesSb -ArgumentList $searchArgs
            }
            else {
                $job = Invoke-Command -ComputerName $server -ScriptBlock $script:SearchTrackingFilesSb -ArgumentList $searchArgs -AsJob
            }

            $jobs.Add([pscustomobject]@{
                Server = $server
                Job    = $job
            }) | Out-Null
        }

        $finished = @($jobs | Where-Object { $_.Job.State -ne 'Running' })
        if ($finished.Count -eq 0) {
            Start-Sleep -Milliseconds 300
            continue
        }

        foreach ($item in $finished) {
            $server = $item.Server
            $job = $item.Job
            $null = $jobs.Remove($item)
            try {
                $payload = @(Receive-Job -Job $job -ErrorAction Stop) |
                    Where-Object { $_ -and ($_.PSObject.Properties.Name -contains 'Raw') } |
                    Select-Object -Last 1
                if (-not $payload) { throw 'Пустой ответ от поиска по файлам.' }
                $counters.Raw += [int]$payload.Raw
                Add-ServerRows -Payload $payload -Results $results
                $matchedCount = 0
                if ($null -ne $payload.Rows) { $matchedCount = @($payload.Rows).Count }
                Write-Host ("  {0}: DELIVER={1}, совпадений={2}" -f $server, $payload.Raw, $matchedCount)
            }
            catch {
                $msg = "Сервер ${server} (файлы): $($_.Exception.Message)"
                Write-Warning $msg
                if ($SearchMethod -eq 'Auto') {
                    $cmdletFallback.Add($server) | Out-Null
                    Write-Host "  $server — повтор через Get-MessageTrackingLog"
                }
                else {
                    $errors.Add($msg) | Out-Null
                }
            }
            finally {
                Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

$cmdletServers = @()
if ($SearchMethod -eq 'Cmdlet') {
    $cmdletServers = @($serverList)
}
elseif ($cmdletFallback.Count -gt 0) {
    $cmdletServers = @($cmdletFallback)
}

foreach ($server in $cmdletServers) {
    Write-Host "Сервер $server — Get-MessageTrackingLog..."
    try {
        $payload = Search-ServerByCmdlet -Server $server -Start $Start -End $End -EventId $EventId -SubjectContains $SubjectContains -Sender $Sender -Recipients $Recipients
        $counters.Raw += [int]$payload.Raw
        Add-ServerRows -Payload $payload -Results $results
        $matchedCount = 0
        if ($null -ne $payload.Rows) { $matchedCount = @($payload.Rows).Count }
        Write-Host ("  {0}: DELIVER={1}, совпадений={2}" -f $server, $payload.Raw, $matchedCount)
    }
    catch {
        $msg = "Сервер ${server}: $($_.Exception.Message)"
        $errors.Add($msg) | Out-Null
        Write-Warning $msg
    }
}

$utf8Bom = New-Object System.Text.UTF8Encoding $true
$header = '"Дата и время","Статус","Отправитель","Получатель","Тема"'
if ($results.Count -eq 0) {
    $csvLines = @($header)
}
else {
    $csvLines = $results | ConvertTo-Csv -NoTypeInformation
}
[System.IO.File]::WriteAllLines($OutputPath, $csvLines, $utf8Bom)

Write-Host ''
Write-Host "Просмотрено событий ${EventId}: $($counters.Raw)"
Write-Host "Найдено писем (тема содержит '$SubjectContains'): $($results.Count)"
Write-Host "CSV сохранён: $OutputPath"

if ($errors.Count -gt 0) {
    Write-Host ''
    Write-Warning "Ошибки по серверам ($($errors.Count)):"
    $errors | ForEach-Object { Write-Warning $_ }
}

if ($exchangeSession) {
    Remove-PSSession $exchangeSession -ErrorAction SilentlyContinue
}
