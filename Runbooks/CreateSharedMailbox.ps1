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
    $mailboxCreationActions = $WebhookBody
    
    # Connect to exchange online#connect to exchange

    #deprecated
    #$connection = Get-AutomationConnection –Name AzureRunAsConnection
    Connect-AzAccount -Identity
    $tenant = 'DOMAIN.onmicrosoft.com'

    #deprecated 
    #Connect-ExchangeOnline –CertificateThumbprint $connection.CertificateThumbprint –AppId $connection.ApplicationID –ShowBanner:$false –Organization $tenant
    Connect-ExchangeOnline -ManagedIdentity -Organization $tenant
    Write-Output "mailboxCreationActions: $($mailboxCreationActions.count)"

    foreach ($mailboxCreationAction in $mailboxCreationActions) {
        $sharedMailboxPrimarysmtp = $mailboxCreationAction.sharedMailboxPrimarysmtp
        $sharedMailboxDisplayname = $mailboxCreationAction.sharedMailboxDisplayname
        $sharedMailboxName = $mailboxCreationAction.sharedMailboxName
        $sharedMailboxAlias = $mailboxCreationAction.sharedMailboxAlias
        #create smb
        New-Mailbox -Shared -DisplayName $sharedMailboxDisplayname -Name $sharedMailboxName -Alias $sharedMailboxAlias -PrimarySmtpAddress $sharedMailboxPrimarysmtp -ResetPasswordOnNextLogon $false
	    Set-MailboxRegionalConfiguration -Identity $sharedMailboxPrimarysmtp -Language 1031 -TimeZone "W. Europe Standard Time" -DateFormat "dd.MM.yyyy" -TimeFormat "HH:mm"
    }
}
else {    
    write-Error "Only webhooks allowed."
}
