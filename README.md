# Workout

## Room mailbox BookInPolicy count

PowerShell script:

```powershell
.\scripts\Get-RoomMailboxBookInPolicyCount.ps1 -UserPrincipalName admin@contoso.com
```

Or use an existing Exchange Online session:

```powershell
Connect-ExchangeOnline
.\scripts\Get-RoomMailboxBookInPolicyCount.ps1 -SkipConnect
```

The script prints:

- total room mailbox count
- count of room mailboxes with a populated `BookInPolicy`
- count of room mailboxes without a populated `BookInPolicy`
