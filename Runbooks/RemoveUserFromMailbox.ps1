param
(
    [Parameter (Mandatory = $false)]
    [object] $WebhookData
)
if ($WebhookData) {
    # Retrieve variables from webhook request body
    $WebhookBody = ConvertFrom-Json -InputObject $WebhookData.RequestBody
    
    # Extract the webhook secret for authentication
    $WebhookSecret = $WebhookBody.WebhookSecret
    
    # Check header for secret to validate request
    if ($WebhookData.RequestHeader.message -eq $WebhookSecret)
    {
        Write-Output "Request authenticated successfully"
    }
    else
    {
        Write-Output "Authentication failed - invalid secret"
        exit
    }
    
    # Extract the customer domain from parameters
    $tenant = $WebhookBody.CustomerDomain
    
    # Get the array of mailbox actions to perform
    # Remove WebhookSecret and CustomerDomain from the body to avoid processing them in the loop
    $mailboxUserActions = $WebhookBody | Select-Object -Property * -ExcludeProperty WebhookSecret, CustomerDomain
    
    # Connect to exchange online
    Connect-AzAccount -Identity
    
    Write-Output "Connecting to tenant: $tenant"
    
    # Connect to Exchange Online with managed identity
    Connect-ExchangeOnline -ManagedIdentity -Organization $tenant
    
    Write-Output "mailboxUserActions: $($mailboxUserActions.count)"
    
    foreach ($mailboxUserAction in $mailboxUserActions) {
        $sharedMailboxPrimarysmtp = $mailboxUserAction.sharedMailboxPrimarySmtp
        $mailboxUser = $mailboxUserAction.mailboxUser
        
        Write-Output "Removing user $($mailboxUser) from shared mailbox $($sharedMailboxPrimarysmtp)"
        
        # Remove user permissions from the shared mailbox
        Remove-MailboxPermission -Identity $sharedMailboxPrimarysmtp -AccessRights FullAccess -Confirm:$false -User $mailboxUser
        Remove-RecipientPermission -Identity $sharedMailboxPrimarysmtp -AccessRights SendAs -Confirm:$false -Trustee $mailboxUser
    }
    
    # Disconnect from Exchange Online
    Disconnect-ExchangeOnline -Confirm:$false
    
    Write-Output "Successfully removed user permissions from shared mailboxes"
}
else {    
    Write-Error "Only webhooks allowed."
}
