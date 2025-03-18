param
(
    [Parameter (Mandatory = $false)]
    [object] $WebhookData
)

if ($WebhookData) {
    # Retrieve variables from webhook request body
    $WebhookBody = ConvertFrom-Json -InputObject $WebhookData.RequestBody
    
    # Extract the webhook secret parameter
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
    
    # Extract the parameters from webhook body
    $CustomerDomain = $WebhookBody.CustomerDomain
    
    # Extract mailbox user actions - can handle both direct array or nested within data property
    if ($WebhookBody.data) {
        $mailboxUserActions = $WebhookBody.data
    } else {
        # Remove the parameters from the body to get just the mailbox actions
        $mailboxUserActions = $WebhookBody | Select-Object -Property * -ExcludeProperty CustomerDomain, WebhookSecret
    }
    
    # Secret has already been validated in the header check above
    
    # Connect to exchange online
    Connect-AzAccount -Identity
    
    # Use the customer domain from the webhook parameters instead of hardcoded value
    $tenant = $CustomerDomain
    Write-Output "Connecting to tenant: $tenant"
    
    # Connect to Exchange Online with managed identity
    Connect-ExchangeOnline -ManagedIdentity -Organization $tenant
    
    # Process each mailbox user action
    foreach ($mailboxUserAction in $mailboxUserActions) {
        $sharedMailboxPrimarysmtp = $mailboxUserAction.sharedMailboxPrimarySmtp
        $mailboxUser = $mailboxUserAction.mailboxUser
        
        Write-Output "Adding user $($mailboxUser) to shared mailbox $($sharedMailboxPrimarysmtp)"
        
        # Add user permissions to the shared mailbox
        Add-MailboxPermission -Identity $sharedMailboxPrimarysmtp -AccessRights FullAccess -User $mailboxUser -AutoMapping $false -InheritanceType All -Confirm:$false 
        Add-RecipientPermission -Identity $sharedMailboxPrimarysmtp -AccessRights SendAs -Confirm:$false -Trustee $mailboxUser
    }
    
    # Disconnect from Exchange Online
    Disconnect-ExchangeOnline -Confirm:$false
    
    Write-Output "Successfully added user permissions to shared mailboxes"
}
else {    
    Write-Error "Only webhooks allowed."
}
