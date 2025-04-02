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
    
    # Get the array of mailbox creation actions to perform
    # Remove WebhookSecret and CustomerDomain from the body to avoid processing them in the loop
    $mailboxCreationActions = $WebhookBody | Select-Object -Property * -ExcludeProperty WebhookSecret, CustomerDomain
    
    # Connect to exchange online
    Connect-AzAccount -Identity
    
    Write-Output "Connecting to tenant: $tenant"
    
    # Connect to Exchange Online with managed identity
    Connect-ExchangeOnline -ManagedIdentity -Organization $tenant
    
    Write-Output "mailboxCreationActions: $($mailboxCreationActions.count)"
    
    foreach ($mailboxCreationAction in $mailboxCreationActions) {
        $sharedMailboxPrimarysmtp = $mailboxCreationAction.sharedMailboxPrimarysmtp
        $sharedMailboxDisplayname = $mailboxCreationAction.sharedMailboxDisplayname
        $sharedMailboxName = $mailboxCreationAction.sharedMailboxName
        $sharedMailboxAlias = $mailboxCreationAction.sharedMailboxAlias
        
        # Create shared mailbox
        New-Mailbox -Shared -DisplayName $sharedMailboxDisplayname -Name $sharedMailboxName -Alias $sharedMailboxAlias -PrimarySmtpAddress $sharedMailboxPrimarysmtp -ResetPasswordOnNextLogon $false
        
        # Wait for mailbox to be fully created before configuring
        Start-Sleep -Seconds 60
        
        # Set regional configuration
        Set-MailboxRegionalConfiguration -Identity $sharedMailboxPrimarysmtp -Language 1031 -TimeZone "W. Europe Standard Time" -DateFormat "dd.MM.yyyy" -TimeFormat "HH:mm"
    }
    
    # Disconnect from Exchange Online
    Disconnect-ExchangeOnline -Confirm:$false
    
    Write-Output "Successfully created shared mailboxes"
}
else {    
    Write-Error "Only webhooks allowed."
}
