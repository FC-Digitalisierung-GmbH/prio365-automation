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
    if ($Action -eq 'Add' -and -not $isMember) {
        Add-DistributionGroupMember -Identity $GroupIdentity -Member $MailboxSmtp -ErrorAction Stop
        return 'Added'
    }
    if ($Action -eq 'Remove' -and $isMember) {
        Remove-DistributionGroupMember -Identity $GroupIdentity -Member $MailboxSmtp -Confirm:$false -ErrorAction Stop
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
    foreach ($item in $Items) {
        try {
            $result = Set-MailboxScopeMembership -Action $item.action -MailboxSmtp $item.mailboxSmtp -GroupIdentity $GroupIdentity
            if ($result -eq 'NoChange') { $noop++ } else { $changed++ }
            Write-Output "OK  $($item.action) $($item.mailboxSmtp) -> $result"
        }
        catch {
            $failed++
            Write-Output "ERR $($item.action) $($item.mailboxSmtp) -> $($_.Exception.Message)"
        }
    }
    return [pscustomobject]@{ Changed = $changed; NoChange = $noop; Failed = $failed }
}

function Invoke-Main {
    if (-not $WebhookData) { Write-Error "Only webhooks allowed."; return }

    $WebhookBody = ConvertFrom-Json -InputObject $WebhookData.RequestBody
    if ($WebhookData.RequestHeader.message -ne $WebhookSecret) {
        Write-Output "Authentication failed - invalid secret"; return
    }
    Write-Output "Request authenticated successfully"

    $tenant = $CustomerDomain
    $items  = $WebhookBody | Select-Object -Property * -ExcludeProperty WebhookSecret, CustomerDomain

    Connect-AzAccount -Identity | Out-Null
    Write-Output "Connecting to tenant: $tenant"
    Connect-ExchangeOnline -ManagedIdentity -Organization $tenant

    try {
        $summary = Invoke-ManageBatch -Items $items -GroupIdentity $ScopeGroupIdentity
        Write-Output "Done. Changed=$($summary.Changed) NoChange=$($summary.NoChange) Failed=$($summary.Failed)"
    }
    finally {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}

# Guard: bei Pester-Dot-Source (InvocationName '.') NICHT ausführen
if ($MyInvocation.InvocationName -ne '.') { Invoke-Main }
