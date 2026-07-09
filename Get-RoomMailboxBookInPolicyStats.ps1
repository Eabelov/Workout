<#
.SYNOPSIS
    Возвращает общее количество room mailbox и количество с заполненным BookInPolicy.

.DESCRIPTION
    Скрипт для on-premises Exchange Server.
    Получает все room mailbox и считает, у скольких из них свойство BookInPolicy
    содержит хотя бы одного пользователя.

    Запускайте из Exchange Management Shell или укажите -ExchangeServer
    для подключения через удалённую PowerShell-сессию.

.PARAMETER ExchangeServer
    FQDN сервера Exchange (например, exchange01.contoso.local).
    Не нужен, если скрипт уже запущен в Exchange Management Shell.

.PARAMETER Credential
    Учётные данные для удалённого подключения. Если не указаны, используется
    текущий контекст Windows (Kerberos).

.PARAMETER UseSsl
    Подключаться по HTTPS вместо HTTP.

.PARAMETER AlreadyConnected
    Пропустить подключение, если сессия Exchange Management Shell уже активна.

.EXAMPLE
    .\Get-RoomMailboxBookInPolicyStats.ps1 -AlreadyConnected

.EXAMPLE
    .\Get-RoomMailboxBookInPolicyStats.ps1 -ExchangeServer exchange01.contoso.local

.EXAMPLE
    .\Get-RoomMailboxBookInPolicyStats.ps1 -ExchangeServer exchange01.contoso.local -UseSsl
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string] $ExchangeServer,

    [Parameter()]
    [System.Management.Automation.PSCredential] $Credential,

    [Parameter()]
    [switch] $UseSsl,

    [Parameter()]
    [switch] $AlreadyConnected
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-ExchangeSession {
    try {
        Get-Command Get-Mailbox -ErrorAction Stop | Out-Null
        Get-ExchangeServer -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

function Connect-ExchangeOnPremIfNeeded {
    if ($AlreadyConnected -or (Test-ExchangeSession)) {
        Write-Verbose 'Сессия Exchange Management Shell уже активна.'
        return
    }

    if (-not $ExchangeServer) {
        throw @'
Сессия Exchange Management Shell не обнаружена.
Запустите скрипт из EMS или укажите параметр -ExchangeServer.
Пример: .\Get-RoomMailboxBookInPolicyStats.ps1 -ExchangeServer exchange01.contoso.local
'@
    }

    $scheme = if ($UseSsl) { 'https' } else { 'http' }
    $connectionUri = '{0}://{1}/PowerShell/' -f $scheme, $ExchangeServer

    $sessionParams = @{
        ConfigurationName = 'Microsoft.Exchange'
        ConnectionUri     = $connectionUri
        ErrorAction       = 'Stop'
    }

    if ($Credential) {
        $sessionParams.Credential = $Credential
        $sessionParams.Authentication = 'Negotiate'
    }
    else {
        $sessionParams.Authentication = 'Kerberos'
    }

    Write-Host "Подключение к $connectionUri ..." -ForegroundColor Cyan
    $session = New-PSSession @sessionParams
    Import-PSSession $session -DisableNameChecking -AllowClobber | Out-Null
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

Connect-ExchangeOnPremIfNeeded

Write-Host 'Получение room mailbox...' -ForegroundColor Cyan

$roomMailboxes = Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited |
    Select-Object DisplayName, PrimarySmtpAddress, BookInPolicy

$totalCount = @($roomMailboxes).Count
$withBookInPolicyCount = @(
    $roomMailboxes | Where-Object { Test-BookInPolicyPopulated -BookInPolicy $_.BookInPolicy }
).Count
$withoutBookInPolicyCount = $totalCount - $withBookInPolicyCount

$result = [PSCustomObject]@{
    TotalRoomMailboxes      = $totalCount
    WithBookInPolicy        = $withBookInPolicyCount
    WithoutBookInPolicy     = $withoutBookInPolicyCount
    PercentWithBookInPolicy = if ($totalCount -gt 0) {
        [math]::Round(($withBookInPolicyCount / $totalCount) * 100, 2)
    }
    else {
        0
    }
    CollectedAt             = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
}

Write-Host ''
Write-Host '=== Room Mailbox / BookInPolicy ===' -ForegroundColor Green
Write-Host "Всего room mailbox:              $($result.TotalRoomMailboxes)"
Write-Host "С заполненным BookInPolicy:      $($result.WithBookInPolicy)"
Write-Host "Без BookInPolicy:                $($result.WithoutBookInPolicy)"
Write-Host "Доля с BookInPolicy:             $($result.PercentWithBookInPolicy)%"
Write-Host ''

return $result
