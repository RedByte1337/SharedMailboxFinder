[![GitHub Sponsors](https://img.shields.io/github/sponsors/RedByte1337?style=flat&logo=githubsponsors)](https://github.com/sponsors/RedByte1337)
[![Twitter](https://img.shields.io/twitter/follow/RedByte1337?label=RedByte1337&style=social)](https://twitter.com/intent/follow?screen_name=RedByte1337)
[![LinkedIn](https://img.shields.io/badge/in-Keanu_Nys-white?style=flat&logoColor=blue&labelColor=blue)](https://www.linkedin.com/in/keanunys/)

# SharedMailboxFinder

Finds mailboxes you can access with a Microsoft Graph token. Handy for spotting shared / delegated mailboxes during M365 assessments.

Microsoft does not disclose any API endpoint which low-privilege users can use to list all mailboxes that are accessible to them. Therefore, this script will use a bruteforce approach where it will check if we can access each mailbox one by one.

Provided with an MsGraph access token with `Mail.Read` / `Mail.ReadWrite` (`.All` or `.Shared`) permissions, the script enumerates all users in the tenant (or alternatively takes a custom target list) and probes each account's `mailFolders` endpoint. Accessible mailboxes are reported with folder count and inbox item count.

After you find accessible mailboxes, you can open them in [GraphSpy](https://github.com/RedByte1337/GraphSpy) via the [Shared Mailboxes](https://github.com/RedByte1337/GraphSpy/wiki/Outlook-Graph#shared-mailboxes) feature of the Outlook Graph module.

## Requirements

- PowerShell 5.1+
- A valid Microsoft Graph access token with one of the following scopes:
  - `Mail.Read.All`
  - `Mail.ReadWrite.All`
  - `Mail.Read.Shared`
  - `Mail.ReadWrite.Shared`

> [!TIP]
> You can use the Office 365 Management (`00b41c95-dab0-4487-9791-b9d2c32c80f2`) FOCI Client ID to obtain a token with the `Mail.ReadWrite.All` scope, or Outlook Mobile (`27922004-5251-4030-b22d-91ecd9a37ea4`) for the `Mail.Read.Shared` scope.

## Usage

```powershell
# Enumerate all users and check mailbox access
$results = .\Invoke-SharedMailboxFinder.ps1 -AccessToken "eyJ..."

# Only users with the Mail attribute set
$results = .\Invoke-SharedMailboxFinder.ps1 -AccessToken "eyJ..." -MailSet

# Only users with at least one ProxyAddress (can combine with -MailSet)
$results = .\Invoke-SharedMailboxFinder.ps1 -AccessToken "eyJ..." -MailSet -ProxySet

# Check a custom list instead of enumerating the tenant
$results = .\Invoke-SharedMailboxFinder.ps1 -AccessToken "eyJ..." -EmailList .\targets.txt

# Don't prompt on token warnings
$results = .\Invoke-SharedMailboxFinder.ps1 -AccessToken "eyJ..." -Force
```

`targets.txt` is just one email/UPN per line:

```
shared@contoso.com
finance@contoso.com
user@contoso.com
```

## Parameters

| Parameter | Description |
|-----------|-------------|
| `-AccessToken` | MsGraph access token (required) |
| `-EmailList` | Optional text file of targets. This skips `/users` enumeration entirely |
| `-MailSet` | Only check users where the `mail` attribute is set |
| `-ProxySet` | Only check users with a proxyAddresses |
| `-Force` | Keep going even if the token looks wrong/expired |

`-MailSet` and `-ProxySet` are applied as Graph `$filter` queries (additive). They are ignored when `-EmailList` is used.

## Output

Returns a list of accessible mailboxes as PowerShell objects:

- `DisplayName`
- `UserPrincipalName`
- `Mail`
- `FolderCount` (Number of folders in the mailbox)
- `InboxMails` (Number of emails in the user's Inbox)

Console output marks accessible mailboxes with `[+]` and inaccessible ones with `[-]`.

### Example

```
[*] Enumerating users (filter: not(mail eq null))...
[*] Found 142 users
[*] Checking mailbox access for 142 accounts...

[+] alice.lovelace@e-corp.com | Folders: 8 | Inbox items: 214
[-] charles.babbage@e-corp.com
[-] alan.turing@e-corp.com
[+] finance@e-corp.com | Folders: 5 | Inbox items: 1832
[+] helpdesk@e-corp.com | Folders: 6 | Inbox items: 97
[-] grace.hopper@e-corp.com
...

[*] Done. Accessible mailboxes: 3 / 142

--- Accessible Mailboxes ---

DisplayName      UserPrincipalName           Mail                         FolderCount InboxMails
-----------      -----------------           ----                         ----------- ----------
Ada Lovelace     alice.lovelace@e-corp.com  alice.lovelace@e-corp.com             8        214
Finance          finance@e-corp.com         finance@e-corp.com                    5       1832
Helpdesk Shared  helpdesk@e-corp.com        helpdesk@e-corp.com                   6         97
```

## How it works

1. Decodes the JWT and warns on expiry, wrong audience, or missing Mail scopes
2. Resolves targets via MsGraph API `/users` endpoint (with optional `$filter`) or from `-EmailList` file
3. Calls `GET /users/{upn}/mailFolders` for each account
4. Collects successes and prints a summary table

## Disclaimer

For authorized security testing and legitimate administration only. Ensure you have permission to access the target tenant and mailboxes before running this tool.
