<#
.SYNOPSIS
    Возвращает общее количество room mailbox и количество с заполненным BookInPolicy.

.DESCRIPTION
    Скрипт подключается к Exchange Online (если сессия ещё не установлена),
    получает все room mailbox и считает, у скольких из них свойство BookInPolicy
    содержит хотя бы одного пользователя.

.PARAMETER UserPrincipalName
    UPN администратора для подключения к Exchange Online.
    Если не указан, используется Connect-ExchangeOnline без параметров
    (интерактивная или существующая аутентификация).

.PARAMETER AlreadyConnected
    Пропустить подключение, если сессия Exchange Online уже активна.

.EXAMPLE
    .\Get-RoomMailboxBookInPolicyStats.ps1

.EXAMPLE
    .\Get-RoomMailboxBookInPolicyStats.ps1 -UserPrincipalName admin@contoso.com

.EXAMPLE
    .\Get-RoomMailboxBookInPolicyStats.ps1 -AlreadyConnected
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string] $UserPrincipalName,

    [Parameter()]
    [switch] $AlreadyConnected
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-ExchangeOnlineSession {
    try {
        Get-OrganizationConfig -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

function Connect-ExchangeOnlineIfNeeded {
    if ($AlreadyConnected -or (Test-ExchangeOnlineSession)) {
        Write-Verbose 'Сессия Exchange Online уже активна.'
        return
    }

    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw 'Модуль ExchangeOnlineManagement не установлен. Установите: Install-Module ExchangeOnlineManagement'
    }

    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    if ($UserPrincipalName) {
        Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowBanner:$false
    }
    else {
        Connect-ExchangeOnline -ShowBanner:$false
    }
}

function Test-BookInPolicyPopulated {
    param(
        [Parameter(Mandatory)]
        $BookInPolicy
    )

    if ($null -eq $BookInPolicy) {
        return $false
    }

    if ($BookInPolicy -is [System.Collections.ICollection]) {
        return $BookInPolicy.Count -gt 0
    }

    $value = [string] $BookInPolicy
    return -not [string]::IsNullOrWhiteSpace($value)
}

Connect-ExchangeOnlineIfNeeded

Write-Host 'Получение room mailbox...' -ForegroundColor Cyan

$roomMailboxes = Get-EXOMailbox -RecipientTypeDetails RoomMailbox -Properties BookInPolicy -ResultSize Unlimited

$totalCount = @($roomMailboxes).Count
$withBookInPolicyCount = @(
    $roomMailboxes | Where-Object { Test-BookInPolicyPopulated -BookInPolicy $_.BookInPolicy }
).Count
$withoutBookInPolicyCount = $totalCount - $withBookInPolicyCount

$result = [PSCustomObject]@{
    TotalRoomMailboxes           = $totalCount
    WithBookInPolicy             = $withBookInPolicyCount
    WithoutBookInPolicy          = $withoutBookInPolicyCount
    PercentWithBookInPolicy      = if ($totalCount -gt 0) {
        [math]::Round(($withBookInPolicyCount / $totalCount) * 100, 2)
    }
    else {
        0
    }
    CollectedAt                  = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
}

Write-Host ''
Write-Host '=== Room Mailbox / BookInPolicy ===' -ForegroundColor Green
Write-Host "Всего room mailbox:              $($result.TotalRoomMailboxes)"
Write-Host "С заполненным BookInPolicy:      $($result.WithBookInPolicy)"
Write-Host "Без BookInPolicy:                $($result.WithoutBookInPolicy)"
Write-Host "Доля с BookInPolicy:             $($result.PercentWithBookInPolicy)%"
Write-Host ''

return $result
