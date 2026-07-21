<#
.SYNOPSIS
    Enumerates all users in a tenant using the MsGraph API and checks mailbox access.

.DESCRIPTION
    Takes an access token and attempts to access each user's mailFolders endpoint.
    Useful for identifying shared mailboxes or delegated mailboxes accessible with
    the current token.

.PARAMETER AccessToken
    A valid MS Graph access token with a Mail.Read(Write).All or Mail.Read(Write).Shared permission.

.PARAMETER EmailList
    Use a custom list of email addresses instead of enumerating all users in the tenant. (Text file with one email address per line)

.PARAMETER MailSet
    Only check users that have the Mail attribute set.

.PARAMETER ProxySet
    Only check users that have at least one ProxyAddress set.
    
.PARAMETER Force
    Continue regardless of the access token pre-check results (e.g. expired tokens, incorrect audience/scope, etc...)

.EXAMPLE
    $results = .\Invoke-SharedMailboxFinder.ps1 -AccessToken "eyJ..."
    $results = .\Invoke-SharedMailboxFinder.ps1 -AccessToken "eyJ..." -MailSet
    $results = .\Invoke-SharedMailboxFinder.ps1 -AccessToken "eyJ..." -MailSet -ProxySet
    $results = .\Invoke-SharedMailboxFinder.ps1 -AccessToken "eyJ..." -EmailList .\targets.txt
    $results = .\Invoke-SharedMailboxFinder.ps1 -AccessToken "eyJ..." -Force
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$AccessToken,

    [string]$EmailList,
    [switch]$MailSet,
    [switch]$ProxySet,
    [switch]$Force
)

$headers = @{
    Authorization  = "Bearer $AccessToken"
    'Content-Type' = 'application/json'
}

$baseUrl = "https://graph.microsoft.com/v1.0"

# --- Token validation ---

$warnings = @()

try {
    $parts  = $AccessToken.Split('.')
    $padded = $parts[1].PadRight($parts[1].Length + (4 - $parts[1].Length % 4) % 4, '=')
    $claims = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($padded)) | ConvertFrom-Json
} catch {
    $warnings +=  "[!] Could not decode token, likely not a valid JWT."
}

if ($warnings.Count -eq 0){
    # Expiry
    $exp = [DateTimeOffset]::FromUnixTimeSeconds($claims.exp).UtcDateTime
    if ((Get-Date).ToUniversalTime() -gt $exp) {
        $warnings += "Token is expired (exp: $exp UTC)"
    }

    # Audience - must target MS Graph
    $aud = $claims.aud
    if ($aud -notmatch "graph\.microsoft\.com|00000003-0000-0000-c000-000000000000") {
        $warnings += "Token audience '$aud' does not appear to target MS Graph"
    }

    # Scopes - look for Mail.Read/Write variants
    $scp = "$($claims.scp) $($claims.roles)"
    if ($scp -notmatch "Mail\.(Read|ReadWrite)(\.All|\.Shared)") {
        $warnings += "No 'Mail.Read(Write).All' or 'Mail.Read(Write).Shared' scope found in token. We will likely not be able to access other user's mailboxes with this token."
    }
}

if ($warnings.Count -gt 0) {
    foreach ($w in $warnings) {
        Write-Host "[!] $w" -ForegroundColor Yellow
    }
    if (-not $Force) {
        $answer = Read-Host "`nContinue anyway? (y/N)"
        if ($answer -notmatch '^[Yy]$') { exit 0 }
    }
}

# --- Resolve user list ---

if ($EmailList) {
    if (-not (Test-Path $EmailList)) {
        Write-Host "[!] Email list not found: $EmailList" -ForegroundColor Red
        exit 1
    }
    Write-Host "[*] Loading accounts from $EmailList..." -ForegroundColor Cyan
    $users = Get-Content $EmailList | Where-Object { $_.Trim() } | ForEach-Object {
        $id = $_.Trim()
        [PSCustomObject]@{ displayName = $id; userPrincipalName = $id; mail = $id }
    }
    Write-Host "[*] Loaded $($users.Count) accounts from file" -ForegroundColor Cyan
}
else {
    $filters = @()
    if ($MailSet)  { $filters += 'not(mail eq null)' }
    if ($ProxySet) { $filters += 'not(proxyAddresses/$count eq 0)' }

    $url = "$baseUrl/users?`$select=displayName,userPrincipalName,mail,proxyAddresses&`$top=999"
    if ($filters) {
        $url += "&`$filter=$([uri]::EscapeDataString(($filters -join ' and ')))"
        $headers['ConsistencyLevel'] = 'eventual'
        Write-Host "[*] Enumerating users (filter: $($filters -join ' and '))..." -ForegroundColor Cyan
    }
    else {
        Write-Host "[*] Enumerating users..." -ForegroundColor Cyan
    }

    $users = @()
    do {
        try {
            $response = Invoke-RestMethod -Uri $url -Headers $headers -Method GET -ErrorAction Stop
            $users += $response.value
            $url     = $response.'@odata.nextLink'
        }
        catch {
            Write-Verbose "URL Failed: $url"
            Write-Host "[!] Failed to enumerate users: $_" -ForegroundColor Red
            exit 1
        }
    } while ($url)

    Write-Host "[*] Found $($users.Count) users" -ForegroundColor Cyan
}

Write-Host "[*] Checking mailbox access for $($users.Count) accounts...`n" -ForegroundColor Cyan

# --- Check mailbox access ---

$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$i       = 0

foreach ($user in $users) {
    $i++
    $upn = $user.userPrincipalName

    Write-Progress -Activity "Checking mailbox access" `
        -Status "[$i/$($users.Count)] $upn" `
        -PercentComplete (($i / $users.Count) * 100)

    $folderCount = 0
    $inboxMails  = 0

    try {
        $encodedUpn   = [uri]::EscapeDataString($upn)
        $foldersUrl   = "$baseUrl/users/$encodedUpn/mailFolders?`$top=100&`$select=id,displayName,totalItemCount"
        $foldersResp  = Invoke-RestMethod -Uri $foldersUrl -Headers $headers -Method GET -ErrorAction Stop

        $folders     = $foldersResp.value
        $folderCount = $folders.Count

        $inbox = $folders | Where-Object { $_.displayName -eq "Inbox" }
        if ($inbox) {
            $inboxMails = $inbox.totalItemCount
        }

        Write-Host "[+] $upn" -ForegroundColor Green -NoNewline
        Write-Host " | Folders: $folderCount | Inbox items: $inboxMails" -ForegroundColor White

        $results.Add([PSCustomObject]@{
            DisplayName       = $user.displayName
            UserPrincipalName = $upn
            Mail              = $user.mail
            FolderCount       = $folderCount
            InboxMails        = $inboxMails
        })
    }
    catch {
        Write-Host "[-] $upn" -ForegroundColor DarkGray
    }
}

Write-Progress -Activity "Checking mailbox access" -Completed

# --- Summary ---

Write-Host "`n[*] Done. Accessible mailboxes: $($results.Count) / $($users.Count)" -ForegroundColor Cyan

if ($results.Count -gt 0) {
    Write-Host "`n--- Accessible Mailboxes ---" -ForegroundColor Yellow
    $results | Format-Table -AutoSize
}

return $results
