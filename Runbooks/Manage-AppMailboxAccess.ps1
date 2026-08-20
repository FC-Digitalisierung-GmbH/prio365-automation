param
(
    [Parameter (Mandatory = $false)]
    [object] $WebhookData,
    [object] $WebhookSecret,
    [object] $CustomerDomain
)

$ScopeGroupIdentity = 'Prio365-MailboxScope'

function Set-MailboxScopeMembership {
    param(
        [Parameter(Mandatory)][ValidateSet('Add','Remove')][string] $Action,
        [Parameter(Mandatory)][string] $MailboxSmtp,
        [Parameter(Mandatory)][string] $GroupIdentity
    )
    $isMember = [bool](Get-DistributionGroupMember -Identity $GroupIdentity -ResultSize Unlimited -ErrorAction Stop |
                       Where-Object { $_.PrimarySmtpAddress -eq $MailboxSmtp })
    # -BypassSecurityGroupManagerCheck: Admin/MI darf Mitglieder ändern, auch ohne ManagedBy-Owner zu sein
    if ($Action -eq 'Add' -and -not $isMember) {
        Add-DistributionGroupMember -Identity $GroupIdentity -Member $MailboxSmtp -BypassSecurityGroupManagerCheck -ErrorAction Stop
        return 'Added'
    }
    if ($Action -eq 'Remove' -and $isMember) {
        Remove-DistributionGroupMember -Identity $GroupIdentity -Member $MailboxSmtp -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop
        return 'Removed'
    }
    return 'NoChange'
}

function Invoke-ManageBatch {
    param(
        [Parameter(Mandatory)][object[]] $Items,
        [Parameter(Mandatory)][string] $GroupIdentity
    )
    $changed = 0; $failed = 0; $noop = 0
    $log = New-Object System.Collections.Generic.List[string]
    foreach ($item in $Items) {
        try {
            $result = Set-MailboxScopeMembership -Action $item.action -MailboxSmtp $item.mailboxSmtp -GroupIdentity $GroupIdentity
            if ($result -eq 'NoChange') { $noop++ } else { $changed++ }
            $log.Add("OK  $($item.action) $($item.mailboxSmtp) -> $result")
        }
        catch {
            $failed++
            $log.Add("ERR $($item.action) $($item.mailboxSmtp) -> $($_.Exception.Message)")
        }
    }
    # Log ins Rückgabeobjekt (NICHT per Write-Output — sonst wird es von '$summary = ...' verschluckt)
    return [pscustomobject]@{ Changed = $changed; NoChange = $noop; Failed = $failed; Log = $log }
}

function Invoke-Main {
    if (-not $WebhookData) { Write-Error "Only webhooks allowed."; return }

    # PowerShell-7-Automation liefert WebhookData ggf. als JSON-String -> re-parsen
    if ($WebhookData -is [string]) { $WebhookData = $WebhookData | ConvertFrom-Json }

    if ($WebhookData.RequestHeader.message -ne $WebhookSecret) {
        Write-Output "Authentication failed - invalid secret"; return
    }
    Write-Output "Request authenticated successfully"

    $tenant      = $CustomerDomain
    $WebhookBody = if ($WebhookData.RequestBody) { ConvertFrom-Json -InputObject $WebhookData.RequestBody } else { $null }
    $items       = @($WebhookBody | Select-Object -Property * -ExcludeProperty WebhookSecret, CustomerDomain)
    if (-not $items -or $items.Count -eq 0) {
        Write-Output "No items in request body - nothing to do (Body korrekt per -Body gesendet?)."; return
    }

    Connect-AzAccount -Identity | Out-Null
    Write-Output "Connecting to tenant: $tenant ($($items.Count) item(s))"
    Connect-ExchangeOnline -ManagedIdentity -Organization $tenant

    try {
        $summary = Invoke-ManageBatch -Items $items -GroupIdentity $ScopeGroupIdentity
        $summary.Log | ForEach-Object { Write-Output $_ }
        Write-Output "Done. Changed=$($summary.Changed) NoChange=$($summary.NoChange) Failed=$($summary.Failed)"
    }
    finally {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}

# Guard: bei Pester-Dot-Source (InvocationName '.') NICHT ausführen
if ($MyInvocation.InvocationName -ne '.') { Invoke-Main }
