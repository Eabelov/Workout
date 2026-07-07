<#
.SYNOPSIS
    Выгружает в CSV почтовые ящики с заданным значением extensionAttribute13.

.DESCRIPTION
    Ищет объекты в Active Directory по атрибуту extensionAttribute13 (в Exchange
    соответствует CustomAttribute13) и формирует CSV с отображаемым именем,
    основным SMTP-адресом и дополнительными адресами — каждый в отдельной колонке.

.PARAMETER ExtensionAttribute13Value
    Значение extensionAttribute13 для фильтрации. По умолчанию: 11.

.PARAMETER OutputPath
    Путь к выходному CSV-файлу.

.PARAMETER SearchBase
    DN контейнера для поиска (например, OU=Users,DC=contoso,DC=com).
    Если не указан, поиск выполняется по всему домену.

.PARAMETER MailEnabledOnly
    Включать только объекты с заполненным атрибутом Mail.

.EXAMPLE
    .\Export-MailboxesByExtensionAttribute13.ps1

.EXAMPLE
    .\Export-MailboxesByExtensionAttribute13.ps1 -ExtensionAttribute13Value 11 -OutputPath C:\temp\boxes.csv

.NOTES
    Требуется модуль ActiveDirectory (RSAT).
    Запускайте от имени пользователя с правами чтения AD.
#>
[CmdletBinding()]
param(
    [Parameter()]
    [string]$ExtensionAttribute13Value = '11',

    [Parameter()]
    [string]$OutputPath = (Join-Path -Path $PSScriptRoot -ChildPath "mailboxes_extattr13_$ExtensionAttribute13Value.csv"),

    [Parameter()]
    [string]$SearchBase,

    [Parameter()]
    [switch]$MailEnabledOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-EmailAddressesFromProxyAddresses {
    param(
        [AllowNull()]
        [string[]]$ProxyAddresses,

        [AllowNull()]
        [string]$Mail
    )

    $primary = $null
    $additional = [System.Collections.Generic.List[string]]::new()

    foreach ($address in ($ProxyAddresses | Where-Object { $_ })) {
        if ($address -clike 'SMTP:*') {
            $primary = $address.Substring(5)
            continue
        }

        if ($address -clike 'smtp:*') {
            $additional.Add($address.Substring(5))
        }
    }

    if (-not $primary -and $Mail) {
        $primary = $Mail
    }

    [PSCustomObject]@{
        Primary    = $primary
        Additional = $additional.ToArray()
    }
}

if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    throw 'Модуль ActiveDirectory не найден. Установите RSAT: Active Directory Domain Services Tools.'
}

Import-Module ActiveDirectory

$filterParts = @("extensionAttribute13 -eq '$ExtensionAttribute13Value'")

if ($MailEnabledOnly) {
    $filterParts += "Mail -like '*'"
}

$adParams = @{
    Filter     = ($filterParts -join ' -and ')
    Properties = @(
        'DisplayName',
        'Mail',
        'proxyAddresses',
        'UserPrincipalName',
        'Enabled'
    )
}

if ($SearchBase) {
    $adParams.SearchBase = $SearchBase
}

Write-Host "Поиск объектов с extensionAttribute13 = '$ExtensionAttribute13Value'..." -ForegroundColor Cyan

$users = @(Get-ADUser @adParams)

if ($users.Count -eq 0) {
    Write-Warning "Объекты с extensionAttribute13 = '$ExtensionAttribute13Value' не найдены."
    return
}

Write-Host "Найдено объектов: $($users.Count)" -ForegroundColor Green

$parsed = foreach ($user in $users) {
    $emails = Get-EmailAddressesFromProxyAddresses -ProxyAddresses $user.proxyAddresses -Mail $user.Mail

    [PSCustomObject]@{
        DisplayName      = $user.DisplayName
        PrimaryEmail     = $emails.Primary
        AdditionalEmails = $emails.Additional
        UserPrincipalName = $user.UserPrincipalName
        Enabled          = $user.Enabled
    }
}

$maxAdditionalCount = ($parsed | ForEach-Object { $_.AdditionalEmails.Count } | Measure-Object -Maximum).Maximum
if (-not $maxAdditionalCount) {
    $maxAdditionalCount = 0
}

$exportRows = foreach ($item in ($parsed | Sort-Object DisplayName)) {
    $row = [ordered]@{
        'Отображаемое имя' = $item.DisplayName
        'Основной адрес'  = $item.PrimaryEmail
    }

    for ($index = 0; $index -lt $maxAdditionalCount; $index++) {
        $columnName = "Дополнительный адрес $($index + 1)"
        $row[$columnName] = if ($index -lt $item.AdditionalEmails.Count) {
            $item.AdditionalEmails[$index]
        }
        else {
            ''
        }
    }

    [PSCustomObject]$row
}

$exportRows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host "CSV сохранён: $OutputPath" -ForegroundColor Green
Write-Host "Строк: $($exportRows.Count), дополнительных колонок адресов: $maxAdditionalCount"
