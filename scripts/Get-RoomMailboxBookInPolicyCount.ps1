<#
.SYNOPSIS
Counts room mailboxes and room mailboxes with a populated BookInPolicy.

.DESCRIPTION
Gets all Exchange Online room mailboxes, reads their calendar processing
settings, and returns the total room mailbox count plus the count of room
mailboxes where BookInPolicy contains at least one value.

Requires the ExchangeOnlineManagement PowerShell module.

.PARAMETER UserPrincipalName
Optional account used to connect to Exchange Online. If omitted, the script
uses the current Exchange Online session.

.PARAMETER SkipConnect
Use the existing Exchange Online session and do not call Connect-ExchangeOnline.

.EXAMPLE
.\Get-RoomMailboxBookInPolicyCount.ps1 -UserPrincipalName admin@contoso.com

.EXAMPLE
Connect-ExchangeOnline
.\Get-RoomMailboxBookInPolicyCount.ps1 -SkipConnect
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$UserPrincipalName,

    [Parameter()]
    [switch]$SkipConnect
)

$ErrorActionPreference = 'Stop'

if (-not $SkipConnect) {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw 'ExchangeOnlineManagement module is required. Install it with: Install-Module ExchangeOnlineManagement'
    }

    Import-Module ExchangeOnlineManagement

    if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
        Connect-ExchangeOnline
    }
    else {
        Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName
    }
}

$roomMailboxes = @(Get-EXOMailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited)
$withBookInPolicy = 0

foreach ($roomMailbox in $roomMailboxes) {
    $calendarProcessing = Get-CalendarProcessing -Identity $roomMailbox.PrimarySmtpAddress

    if ($null -ne $calendarProcessing.BookInPolicy -and @($calendarProcessing.BookInPolicy).Count -gt 0) {
        $withBookInPolicy++
    }
}

[PSCustomObject]@{
    TotalRoomMailboxes              = $roomMailboxes.Count
    RoomMailboxesWithBookInPolicy   = $withBookInPolicy
    RoomMailboxesWithoutBookInPolicy = $roomMailboxes.Count - $withBookInPolicy
}
