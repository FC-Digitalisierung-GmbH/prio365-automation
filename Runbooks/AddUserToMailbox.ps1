param
(
    [Parameter (Mandatory = $false)]
    [object] $WebhookData
)

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
    foreach ($mailboxUserAction in $mailboxUserActions) {
        $sharedMailboxPrimarysmtp = $mailboxUserAction.sharedMailboxPrimarySmtp
        $mailboxUser = $mailboxUserAction.mailboxUser
        Write-Output "Add user $($mailboxUser) to shared mailbox $($sharedMailboxPrimarysmtp)"
        #add user from smb
        Add-MailboxPermission -Identity $sharedMailboxPrimarysmtp -AccessRights FullAccess -User $mailboxUser -AutoMapping $false -InheritanceType All -Confirm:$false 
        Add-RecipientPermission -Identity $sharedMailboxPrimarysmtp -AccessRights SendAs -Confirm:$false -Trustee $mailboxUser
   }

   Disconnect-ExchangeOnline -Confirm:$false
}
else {    
    write-Error "Only webhooks allowed."
}