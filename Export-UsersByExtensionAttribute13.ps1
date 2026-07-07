<#
.SYNOPSIS
    Выгружает в CSV пользователей с заданным значением extensionAttribute13.

.DESCRIPTION
    Ищет объекты в Active Directory по атрибуту extensionAttribute13 и формирует CSV
    с полями: Name, SamAccountName, mail, PrimarySmtpAddress, enabled.

    PrimarySmtpAddress берётся из Exchange (Get-Mailbox).
    Запускайте скрипт из Exchange Management Shell (EMS) — отдельное подключение не требуется.

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
    .\Export-UsersByExtensionAttribute13.ps1

.EXAMPLE
    .\Export-UsersByExtensionAttribute13.ps1 -ExtensionAttribute13Value 11 -OutputPath C:\temp\users.csv

.NOTES
    Требуется модуль ActiveDirectory (RSAT) и Exchange Management Shell.
#>
[CmdletBinding()]
param(
    [Parameter()]
    [string]$ExtensionAttribute13Value = '11',

    [Parameter()]
    [string]$OutputPath = (Join-Path -Path $PSScriptRoot -ChildPath "users_extattr13_$ExtensionAttribute13Value.csv"),

    [Parameter()]
    [string]$SearchBase,

    [Parameter()]
    [switch]$MailEnabledOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    throw 'Модуль ActiveDirectory не найден. Установите RSAT: Active Directory Domain Services Tools.'
}

if (-not (Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
    throw 'Команда Get-Mailbox недоступна. Запустите скрипт из Exchange Management Shell.'
}

Import-Module ActiveDirectory

$filterParts = @("extensionAttribute13 -eq '$ExtensionAttribute13Value'")

if ($MailEnabledOnly) {
    $filterParts += "Mail -like '*'"
}

$adParams = @{
    Filter     = ($filterParts -join ' -and ')
    Properties = @(
        'Name',
        'SamAccountName',
        'Mail',
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

Write-Host "Получение PrimarySmtpAddress из Exchange..." -ForegroundColor Cyan

$primarySmtpBySamAccount = @{}
Get-Mailbox -Filter "CustomAttribute13 -eq '$ExtensionAttribute13Value'" -ResultSize Unlimited |
    ForEach-Object {
        $primarySmtpBySamAccount[$_.SamAccountName] = $_.PrimarySmtpAddress
    }

$exportRows = foreach ($user in ($users | Sort-Object Name)) {
    [PSCustomObject]@{
        Name               = $user.Name
        SamAccountName     = $user.SamAccountName
        mail               = $user.Mail
        PrimarySmtpAddress = $primarySmtpBySamAccount[$user.SamAccountName]
        enabled            = $user.Enabled
    }
}

$exportRows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host "CSV сохранён: $OutputPath" -ForegroundColor Green
Write-Host "Строк: $($exportRows.Count)"
