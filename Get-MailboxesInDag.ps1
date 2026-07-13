<#
.SYNOPSIS
    Показывает все почтовые ящики, размещённые в указанном Exchange DAG.

.DESCRIPTION
    Скрипт для on-premises Exchange Server. Находит все mailbox database,
    входящие в Database Availability Group, а затем возвращает ящики из каждой
    найденной базы.

    Запускайте скрипт из Exchange Management Shell или укажите -ExchangeServer
    для подключения через удалённую PowerShell-сессию.

.PARAMETER DagName
    Имя Database Availability Group.

.PARAMETER ExchangeServer
    FQDN сервера Exchange (например, exchange01.contoso.local).
    Не требуется при запуске из Exchange Management Shell.

.PARAMETER Credential
    Учётные данные для удалённого подключения. Если параметр не указан,
    используется текущий контекст Windows (Kerberos).

.PARAMETER UseSsl
    Подключаться к удалённой PowerShell-сессии по HTTPS.

.PARAMETER AlreadyConnected
    Указывает, что активная сессия Exchange Management Shell уже существует.

.EXAMPLE
    .\Get-MailboxesInDag.ps1 -DagName DAG01 -AlreadyConnected

.EXAMPLE
    .\Get-MailboxesInDag.ps1 -DagName DAG01 -ExchangeServer exchange01.contoso.local

.EXAMPLE
    .\Get-MailboxesInDag.ps1 -DagName DAG01 -AlreadyConnected |
        Export-Csv C:\Temp\DAG01-mailboxes.csv -NoTypeInformation -Encoding UTF8
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string] $DagName,

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
        Get-Command Get-MailboxDatabase -ErrorAction Stop | Out-Null
        Get-Command Get-DatabaseAvailabilityGroup -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

function Connect-ExchangeOnPremIfNeeded {
    if (Test-ExchangeSession) {
        Write-Verbose 'Сессия Exchange Management Shell уже активна.'
        return $null
    }

    if ($AlreadyConnected) {
        throw 'Параметр -AlreadyConnected указан, но активная сессия Exchange Management Shell не обнаружена.'
    }

    if (-not $ExchangeServer) {
        throw @'
Сессия Exchange Management Shell не обнаружена.
Запустите скрипт из EMS или укажите параметр -ExchangeServer.
Пример: .\Get-MailboxesInDag.ps1 -DagName DAG01 -ExchangeServer exchange01.contoso.local
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

    try {
        Import-PSSession $session -DisableNameChecking -AllowClobber | Out-Null
    }
    catch {
        Remove-PSSession $session -ErrorAction SilentlyContinue
        throw
    }

    return $session
}

function Test-DatabaseBelongsToDag {
    param(
        [Parameter(Mandatory)]
        $Database,

        [Parameter(Mandatory)]
        $Dag
    )

    $group = $Database.MasterServerOrAvailabilityGroup
    if ($null -eq $group) {
        return $false
    }

    $groupName = if ($group.PSObject.Properties['Name']) {
        [string] $group.Name
    }
    else {
        [string] $group
    }

    return (
        $groupName -eq [string] $Dag.Name -or
        [string] $group -eq [string] $Dag.Identity
    )
}

$createdSession = $null

try {
    $createdSession = Connect-ExchangeOnPremIfNeeded

    Write-Verbose "Поиск DAG '$DagName'."
    $dag = Get-DatabaseAvailabilityGroup -Identity $DagName -ErrorAction Stop

    $databases = @(
        Get-MailboxDatabase -ResultSize Unlimited |
            Where-Object { Test-DatabaseBelongsToDag -Database $_ -Dag $dag } |
            Sort-Object Name
    )

    if ($databases.Count -eq 0) {
        Write-Warning "В DAG '$($dag.Name)' не найдено mailbox database."
        return
    }

    Write-Verbose ("Найдено баз в DAG '{0}': {1}." -f $dag.Name, $databases.Count)

    $mailboxes = @(
        foreach ($database in $databases) {
            Write-Verbose "Получение ящиков из базы '$($database.Name)'."

            Get-Mailbox -Database $database.Identity -ResultSize Unlimited |
                Select-Object @(
                    @{ Name = 'Dag'; Expression = { [string] $dag.Name } }
                    'DisplayName'
                    'Alias'
                    'PrimarySmtpAddress'
                    'RecipientTypeDetails'
                    @{ Name = 'Database'; Expression = { [string] $_.Database } }
                    'ServerName'
                )
        }
    )

    Write-Verbose ("Всего найдено ящиков: {0}." -f $mailboxes.Count)
    return $mailboxes
}
finally {
    if ($null -ne $createdSession) {
        Remove-PSSession $createdSession -ErrorAction SilentlyContinue
    }
}
