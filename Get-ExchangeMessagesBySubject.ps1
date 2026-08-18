<#
.SYNOPSIS
    Ищет в message tracking всех серверов Exchange письма, у которых тема содержит заданную строку, и выгружает результат в CSV.

.DESCRIPTION
    Обходит все транспортные / mailbox-серверы, читает логи Get-MessageTrackingLog
    (событие DELIVER по умолчанию) и отбирает записи, где тема СОДЕРЖИТ шаблон,
    а не равна ему целиком.

    По умолчанию ищется символ «&» в теме.

    CSV-колонки:
      Дата и время, Статус, Отправитель, Получатель, Тема

.PARAMETER Start
    Начало периода. По умолчанию — 7 дней назад.

.PARAMETER End
    Конец периода. По умолчанию — сейчас.

.PARAMETER SubjectContains
    Подстрока, которая должна встречаться в теме. По умолчанию «&».

.PARAMETER EventId
    Событие message tracking. По умолчанию DELIVER.

.PARAMETER OutputPath
    Путь к CSV. По умолчанию — файл рядом со скриптом с меткой времени.

.PARAMETER Servers
    Список серверов. Если не задан — берутся все серверы с ролью транспорта.

.PARAMETER ExchangeUri
    URI для удалённого подключения к Exchange, если скрипт запущен не из EMS.
    Пример: http://exch01.contoso.local/PowerShell/

.EXAMPLE
    .\Get-ExchangeMessagesBySubject.ps1

.EXAMPLE
    .\Get-ExchangeMessagesBySubject.ps1 -Start (Get-Date).AddDays(-30) -OutputPath C:\Temp\deliver-amp.csv

.EXAMPLE
    .\Get-ExchangeMessagesBySubject.ps1 -SubjectContains '&' -EventId DELIVER -Start '2026-08-01' -End '2026-08-18'
#>
[CmdletBinding()]
param(
    [datetime]$Start = (Get-Date).AddDays(-7),

    [datetime]$End = (Get-Date),

    [string]$SubjectContains = '&',

    [string]$EventId = 'DELIVER',

    [string]$OutputPath,

    [string[]]$Servers,

    [string]$ExchangeUri
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

function ConvertTo-RecipientText {
    param($Recipients)

    if ($null -eq $Recipients) { return '' }
    if ($Recipients -is [string]) { return $Recipients }

    $list = @($Recipients | ForEach-Object { "$_" } | Where-Object { $_ })
    return ($list -join '; ')
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

Write-Host "Период: $($Start.ToString('yyyy-MM-dd HH:mm:ss')) — $($End.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "Серверы ($($serverList.Count)): $($serverList -join ', ')"
Write-Host "Событие: $EventId"
Write-Host "Тема содержит: $SubjectContains"
Write-Host "Файл: $OutputPath"
Write-Host ''

$wildcard = '*' + $SubjectContains + '*'
$errors = New-Object System.Collections.Generic.List[string]
$results = New-Object System.Collections.Generic.List[object]
$counters = @{ Raw = 0 }

foreach ($server in $serverList) {
    Write-Host "Сервер $server — чтение message tracking..."
    try {
        $counters.MatchedOnServer = 0
        Get-MessageTrackingLog -Server $server -Start $Start -End $End -EventId $EventId -ResultSize Unlimited -ErrorAction Stop | ForEach-Object {
            $counters.Raw++
            if ($_.MessageSubject -like $wildcard) {
                $counters.MatchedOnServer++
                $results.Add([pscustomobject]@{
                    'Дата и время' = $_.Timestamp.ToString('yyyy-MM-dd HH:mm:ss')
                    'Статус'       = $_.EventId
                    'Отправитель'  = $_.Sender
                    'Получатель'   = ConvertTo-RecipientText -Recipients $_.Recipients
                    'Тема'         = $_.MessageSubject
                }) | Out-Null
            }
        }

        Write-Host ("  совпадений по теме: {0}" -f $counters.MatchedOnServer)
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
Write-Host "Просмотрено событий $EventId: $($counters.Raw)"
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
