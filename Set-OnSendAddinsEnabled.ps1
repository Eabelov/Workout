<#
.SYNOPSIS
    Просмотр и изменение параметра OnSendAddinsEnabled в политиках OWA.

.DESCRIPTION
    Параметр OnSendAddinsEnabled указывает, можно ли редактировать почтовый
    элемент, пока надстройка при отправке обрабатывает его в Outlook в Интернете
    или в новом Outlook в Windows.

    Допустимые значения:
      $true  — надстройки при отправке включены.
      $false — надстройки при отправке отключены. Это значение по умолчанию.

    Применимо: Exchange Server 2016, Exchange Server 2019, Exchange Server SE,
    Exchange Online.

    По умолчанию скрипт только выводит текущее состояние политик OWA.
    Чтобы изменить флаг, укажите -Enabled $true или -Enabled $false.

    Запускайте из Exchange Management Shell (on-premises) или из сессии
    Exchange Online PowerShell. Для удалённого подключения к on-premises
    укажите -ExchangeServer.

    После изменения политики подождите до 60 минут или перезапустите IIS
    (on-premises), чтобы настройка вступила в силу.

.PARAMETER Identity
    Имя политики OWA (например, OwaMailboxPolicy-Default).
    Если не указано, действие применяется ко всем политикам.

.PARAMETER Enabled
    Новое значение OnSendAddinsEnabled. Если параметр не задан, скрипт
    только показывает текущее состояние.

.PARAMETER AssignToMailbox
    SMTP-адрес, alias или identity ящика, которому нужно назначить политику
    из -Identity. Требует указания -Identity.

.PARAMETER ShowMailboxAssignments
    Дополнительно вывести ящики и назначенные им OWA-политики.

.PARAMETER ExchangeServer
    FQDN сервера Exchange (например, exchange01.contoso.local).
    Не требуется при запуске из Exchange Management Shell.

.PARAMETER Credential
    Учётные данные для удалённого подключения. Если параметр не указан,
    используется текущий контекст Windows (Kerberos).

.PARAMETER UseSsl
    Подключаться к удалённой PowerShell-сессии по HTTPS.

.PARAMETER AlreadyConnected
    Указывает, что активная сессия Exchange уже существует.

.EXAMPLE
    .\Set-OnSendAddinsEnabled.ps1 -AlreadyConnected

    Показывает OnSendAddinsEnabled для всех политик OWA.

.EXAMPLE
    .\Set-OnSendAddinsEnabled.ps1 -AlreadyConnected -Enabled $true

    Включает надстройки при отправке во всех политиках OWA.

.EXAMPLE
    .\Set-OnSendAddinsEnabled.ps1 -AlreadyConnected -Identity 'OwaMailboxPolicy-Default' -Enabled $true

    Включает надстройки при отправке только в политике по умолчанию.

.EXAMPLE
    .\Set-OnSendAddinsEnabled.ps1 -AlreadyConnected -Identity 'OWAOnSendAddinAllUserPolicy' -AssignToMailbox 'user@contoso.com'

    Назначает политику указанному ящику.

.EXAMPLE
    .\Set-OnSendAddinsEnabled.ps1 -ExchangeServer exchange01.contoso.local -ShowMailboxAssignments

    Подключается к on-premises Exchange и выводит политики вместе с
    назначенными ящиками.

.LINK
    https://learn.microsoft.com/powershell/module/exchangepowershell/set-owamailboxpolicy
    https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-on-send-addins
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter()]
    [string] $Identity,

    [Parameter()]
    [bool] $Enabled,

    [Parameter()]
    [string] $AssignToMailbox,

    [Parameter()]
    [switch] $ShowMailboxAssignments,

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

if ($PSBoundParameters.ContainsKey('AssignToMailbox') -and -not $Identity) {
    throw 'Параметр -AssignToMailbox требует указания -Identity (имя политики OWA).'
}

function Test-ExchangeSession {
    try {
        Get-Command Get-OwaMailboxPolicy -ErrorAction Stop | Out-Null
        Get-Command Set-OwaMailboxPolicy -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

function Connect-ExchangeOnPremIfNeeded {
    if (Test-ExchangeSession) {
        Write-Verbose 'Сессия Exchange уже активна.'
        return $null
    }

    if ($AlreadyConnected) {
        throw 'Параметр -AlreadyConnected указан, но командлеты Get-OwaMailboxPolicy / Set-OwaMailboxPolicy не найдены.'
    }

    if (-not $ExchangeServer) {
        throw @'
Сессия Exchange не обнаружена.
Запустите скрипт из Exchange Management Shell, из сессии Exchange Online
PowerShell или укажите параметр -ExchangeServer.

Примеры:
  .\Set-OnSendAddinsEnabled.ps1 -AlreadyConnected
  .\Set-OnSendAddinsEnabled.ps1 -ExchangeServer exchange01.contoso.local
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

function Get-TargetOwaMailboxPolicies {
    param(
        [string] $PolicyIdentity
    )

    if ($PolicyIdentity) {
        return @(Get-OwaMailboxPolicy -Identity $PolicyIdentity -ErrorAction Stop)
    }

    return @(Get-OwaMailboxPolicy -ErrorAction Stop | Sort-Object Name)
}

function Test-OnSendAddinsProperty {
    param(
        $Policy
    )

    return [bool] $Policy.PSObject.Properties['OnSendAddinsEnabled']
}

function Get-OwaPolicyAssignmentRows {
    param(
        [Parameter(Mandatory)]
        [string[]] $PolicyNames
    )

    if (-not (Get-Command Get-CASMailbox -ErrorAction SilentlyContinue)) {
        Write-Warning 'Командлет Get-CASMailbox недоступен — назначения политик ящикам пропущены.'
        return @()
    }

    $nameSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($name in $PolicyNames) {
        [void] $nameSet.Add($name)
    }

    Write-Verbose 'Получение назначений OWA-политик через Get-CASMailbox.'
    Get-CASMailbox -ResultSize Unlimited |
        Where-Object {
            $_.OwaMailboxPolicy -and $nameSet.Contains([string] $_.OwaMailboxPolicy)
        } |
        Select-Object @(
            @{ Name = 'Mailbox'; Expression = { [string] $_.Identity } }
            'PrimarySmtpAddress'
            @{ Name = 'OwaMailboxPolicy'; Expression = { [string] $_.OwaMailboxPolicy } }
        )
}

$createdSession = $null
$changeRequested = $PSBoundParameters.ContainsKey('Enabled')

try {
    $createdSession = Connect-ExchangeOnPremIfNeeded

    $policies = Get-TargetOwaMailboxPolicies -PolicyIdentity $Identity
    if ($policies.Count -eq 0) {
        Write-Warning 'Политики OWA не найдены.'
        return
    }

    $unsupported = @(
        $policies | Where-Object { -not (Test-OnSendAddinsProperty -Policy $_) }
    )
    if ($unsupported.Count -gt 0) {
        $names = ($unsupported | ForEach-Object { $_.Name }) -join ', '
        throw ("Свойство OnSendAddinsEnabled отсутствует у политик: {0}. Требуется Exchange 2016 CU6+ / 2019 CU1+ / SE / Exchange Online (и PrepareAD после установки CU)." -f $names)
    }

    if ($changeRequested) {
        foreach ($policy in $policies) {
            $target = [string] $policy.Identity
            $current = [bool] $policy.OnSendAddinsEnabled
            if ($current -eq $Enabled) {
                Write-Host ("Политика '{0}': OnSendAddinsEnabled уже {1}." -f $policy.Name, $Enabled) -ForegroundColor Yellow
                continue
            }

            if ($PSCmdlet.ShouldProcess($target, "Set-OwaMailboxPolicy -OnSendAddinsEnabled `$$Enabled")) {
                Set-OwaMailboxPolicy -Identity $target -OnSendAddinsEnabled:$Enabled -ErrorAction Stop
                Write-Host ("Политика '{0}': OnSendAddinsEnabled {1} -> {2}." -f $policy.Name, $current, $Enabled) -ForegroundColor Green
            }
        }

        $policies = Get-TargetOwaMailboxPolicies -PolicyIdentity $Identity
    }

    if ($AssignToMailbox) {
        if ($policies.Count -ne 1) {
            throw 'Для -AssignToMailbox укажите одну политику через -Identity.'
        }

        if (-not (Get-Command Set-CASMailbox -ErrorAction SilentlyContinue)) {
            throw 'Командлет Set-CASMailbox недоступен — назначить политику ящику нельзя.'
        }

        $policyName = [string] $policies[0].Name
        if ($PSCmdlet.ShouldProcess($AssignToMailbox, "Set-CASMailbox -OwaMailboxPolicy '$policyName'")) {
            Set-CASMailbox -Identity $AssignToMailbox -OwaMailboxPolicy $policyName -ErrorAction Stop
            Write-Host ("Ящику '{0}' назначена политика OWA '{1}'." -f $AssignToMailbox, $policyName) -ForegroundColor Green
        }
    }

    $report = @(
        foreach ($policy in $policies) {
            $isEnabled = [bool] $policy.OnSendAddinsEnabled
            [PSCustomObject]@{
                Name                 = [string] $policy.Name
                Identity             = [string] $policy.Identity
                IsDefault            = [bool] $policy.IsDefault
                OnSendAddinsEnabled  = $isEnabled
                Status               = if ($isEnabled) { 'надстройки при отправке включены' } else { 'надстройки при отправке отключены (по умолчанию)' }
            }
        }
    )

    Write-Host ''
    Write-Host '=== OnSendAddinsEnabled / политики OWA ===' -ForegroundColor Cyan
    $report | Format-Table Name, IsDefault, OnSendAddinsEnabled, Status -AutoSize | Out-String | Write-Host

    if ($changeRequested) {
        Write-Host 'Изменение может применяться до 60 минут. На on-premises можно ускорить перезапуском IIS (iisreset).' -ForegroundColor DarkYellow
        Write-Host ''
    }

    if ($ShowMailboxAssignments) {
        $assignmentRows = @(
            Get-OwaPolicyAssignmentRows -PolicyNames @($report | ForEach-Object { $_.Name })
        )
        Write-Host '=== Назначения политик OWA ===' -ForegroundColor Cyan
        if ($assignmentRows.Count -eq 0) {
            Write-Host 'Ящики с указанными политиками не найдены (или Get-CASMailbox недоступен).'
        }
        else {
            $assignmentRows | Format-Table -AutoSize | Out-String | Write-Host
        }

        foreach ($row in $report) {
            $assigned = @(
                $assignmentRows | Where-Object { [string] $_.OwaMailboxPolicy -eq $row.Name }
            )
            $row | Add-Member -NotePropertyName AssignedMailboxCount -NotePropertyValue $assigned.Count
            $row | Add-Member -NotePropertyName AssignedMailboxes -NotePropertyValue $assigned
        }
    }

    return $report
}
finally {
    if ($null -ne $createdSession) {
        Remove-PSSession $createdSession -ErrorAction SilentlyContinue
    }
}
