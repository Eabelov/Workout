# Workout

## OnSendAddinsEnabled

Параметр `OnSendAddinsEnabled` указывает, можно ли редактировать почтовый элемент, пока надстройка при отправке обрабатывает его в Outlook в Интернете или в новом Outlook в Windows.

Применимо: Exchange Server 2016, Exchange Server 2019, Exchange Server SE, Exchange Online.

Допустимые значения:

- `$true` — надстройки при отправке включены.
- `$false` — надстройки при отправке отключены. Это значение является значением по умолчанию.

Скрипт `Set-OnSendAddinsEnabled.ps1` показывает текущее состояние политик OWA и при необходимости включает или отключает флаг.

### Примеры

Из Exchange Management Shell или уже открытой сессии Exchange Online:

```powershell
# Только просмотр
.\Set-OnSendAddinsEnabled.ps1 -AlreadyConnected

# Включить для политики по умолчанию
.\Set-OnSendAddinsEnabled.ps1 -AlreadyConnected -Identity 'OwaMailboxPolicy-Default' -Enabled $true

# Отключить
.\Set-OnSendAddinsEnabled.ps1 -AlreadyConnected -Identity 'OwaMailboxPolicy-Default' -Enabled $false

# Назначить политику ящику
.\Set-OnSendAddinsEnabled.ps1 -AlreadyConnected -Identity 'OwaMailboxPolicy-Default' -AssignToMailbox 'user@contoso.com'
```

Удалённое подключение к on-premises Exchange:

```powershell
.\Set-OnSendAddinsEnabled.ps1 -ExchangeServer exchange01.contoso.local -ShowMailboxAssignments
```

После изменения политики подождите до 60 минут или перезапустите IIS на серверах Exchange.
