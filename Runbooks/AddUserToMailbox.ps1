param
(
    [Parameter (Mandatory = $false)]
    [object] $WebhookData,
    [object] $WebhookSecret,
    [object] $CustomerDomain
)

if ($WebhookData) {
    # Retrieve variables from webhook request body
    $WebhookBody = ConvertFrom-Json -InputObject $WebhookData.RequestBody
    
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
    $tenant = $CustomerDomain
    
    # Get the array of mailbox actions to perform
    # Remove WebhookSecret and CustomerDomain from the body to avoid processing them in the loop
    $cleanBody = $WebhookBody | Select-Object -Property * -ExcludeProperty WebhookSecret, CustomerDomain
    
    # Connect to exchange online
    Connect-AzAccount -Identity
    
    Write-Output "Connecting to tenant: $tenant"
    
    # Connect to Exchange Online with managed identity
    Connect-ExchangeOnline -ManagedIdentity -Organization $tenant
    
    # Process each mailbox user action
    foreach ($mailboxUserAction in $cleanBody) {
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
