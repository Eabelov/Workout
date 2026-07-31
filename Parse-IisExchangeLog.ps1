<#
.SYNOPSIS
    Простой парсер IIS-лога Exchange (W3C Extended Log Format).

.DESCRIPTION
    Читает лог IIS (например u_exYYMMDD.log), разбирает строку #Fields
    и выводит записи как объекты PowerShell.

.PARAMETER Path
    Путь к файлу лога.

.PARAMETER Status
    Опциональный фильтр по HTTP-коду (sc-status), например 401 или 500.

.PARAMETER Top
    Показать только первые N записей.

.EXAMPLE
    .\Parse-IisExchangeLog.ps1

.EXAMPLE
    .\Parse-IisExchangeLog.ps1 -Path "C:\inetpub\logs\LogFiles\W3SVC1\u_ex260725.log" -Status 401

.EXAMPLE
    .\Parse-IisExchangeLog.ps1 | Group-Object cs-uri-stem | Sort-Object Count -Descending | Select-Object -First 10
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Path = "C:\inetpub\logs\LogFiles\W3SVC1\u_ex260725.log",

    [Parameter()]
    [int]$Status,

    [Parameter()]
    [int]$Top
)

if (-not (Test-Path -LiteralPath $Path)) {
    throw "Файл лога не найден: $Path"
}

$fieldNames = $null
$count = 0

Get-Content -LiteralPath $Path -ReadCount 1000 | ForEach-Object {
    foreach ($line in $_) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        if ($line.StartsWith("#")) {
            if ($line.StartsWith("#Fields:")) {
                $fieldNames = ($line -replace "^#Fields:\s*", "") -split "\s+"
            }
            continue
        }

        if (-not $fieldNames) {
            continue
        }

        $values = $line -split "\s+"
        if ($values.Count -lt $fieldNames.Count) {
            continue
        }

        $record = [ordered]@{}
        for ($i = 0; $i -lt $fieldNames.Count; $i++) {
            $record[$fieldNames[$i]] = $values[$i]
        }

        if ($PSBoundParameters.ContainsKey("Status")) {
            if ($record["sc-status"] -ne "$Status") {
                continue
            }
        }

        [pscustomobject]$record
        $count++

        if ($Top -gt 0 -and $count -ge $Top) {
            return
        }
    }
}
