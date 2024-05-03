param
(
    [Parameter (Mandatory = $false)]
    [object] $WebhookData
)

# If runbook was called from webhook, WebhookData will not be null.
if ($WebhookData) {

    # Check header for message to validate request
    if ($WebhookData.RequestHeader.message -eq 'StartedByPrio365')
    {
        Write-Output "Header has required information"}
    else
    {
        Write-Output "Header missing required information";
        exit;
    }

    Write-Output "Convert Header";
    # Retrieve variables from webhook request body
    $WebhookBody =  (ConvertFrom-Json -InputObject $WebhookData.RequestBody)    
    
    # Retrieve variables from webhook request body
    $mailboxUserActions = $WebhookBody
    
    # Connect to exchange online#connect to exchange

    #deprecated
    #$connection = Get-AutomationConnection –Name AzureRunAsConnection
    Connect-AzAccount -Identity
    $tenant = 'DOMAIN.onmicrosoft.com'

    #deprecated 
    #Connect-ExchangeOnline –CertificateThumbprint $connection.CertificateThumbprint –AppId $connection.ApplicationID –ShowBanner:$false –Organization $tenant
    Connect-ExchangeOnline -ManagedIdentity -Organization $tenant
    Write-Output "mailboxUserActions: $($mailboxUserActions.count)"
    
    foreach ($mailboxUserAction in $mailboxUserActions) {
        $sharedMailboxPrimarysmtp = $mailboxUserAction.sharedMailboxPrimarySmtp
        $mailboxUser = $mailboxUserAction.mailboxUser
        
        #add user from smb
        Remove-MailboxPermission -Identity $sharedMailboxPrimarysmtp -AccessRights FullAccess -Confirm:$false -User $mailboxUser
        Remove-RecipientPermission -Identity $sharedMailboxPrimarysmtp -AccessRights SendAs -Confirm:$false -Trustee $mailboxUser
   }

    Disconnect-ExchangeOnline -Confirm:$false   
}
else {    
    write-Error "Only webhooks allowed."
}


