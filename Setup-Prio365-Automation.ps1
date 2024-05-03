#Uninstall-Module AzureRM (if local power shell)
#Install-Module -Name Az -Repository PSGallery -Force -scope CurrentUser -AllowClobber
#Install-Module Microsoft.Graph -Scope CurrentUser
#Install-Module -Name Az.Automation

$accountName = 'Prio365-Automation'
$rgName = 'Prio365'
$location = "westeurope"

#Connect-AzAccount
$tenant = (Get-AzTenant).Id
Connect-MgGraph -TenantId $tenant -Scopes AppRoleAssignment.ReadWrite.All,Application.Read.All,Directory.AccessAsUser.All

# Function to generate a random secret
function Generate-Secret {
    param (
        [int]$Length = 10
    )

    # Define the characters allowed in the secret
    $Chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"

    # Initialize an empty string to store the secret
    $Secret = ""

    # Generate the secret
    for ($i = 1; $i -le $Length; $i++) {
        $RandomIndex = Get-Random -Minimum 0 -Maximum $Chars.Length
        $Secret += $Chars[$RandomIndex]
    }

    # Output the generated secret
    return $Secret
}

# Function to create a runbook
function Create-Runbook {
    param (
        [string]$RunbookName,
        [string]$ResourceGroupName,
        [string]$AutomationAccountName,
        [string]$GitHubUrl
    )

    # Check if runbook already exists
    $existingRunbook = Get-AzAutomationRunbook -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName -Name $RunbookName -ErrorAction SilentlyContinue

    if (-not $existingRunbook) {
        # Download the script from GitHub
        $localPath = ".\$RunbookName.ps1"
        Invoke-WebRequest -Uri $GitHubUrl -OutFile $localPath

        # Read the content of the script
        $scriptContent = Get-Content $localPath -Raw

        # Generate secret and replace placeholders
        $scriptContent = $scriptContent -replace 'StartedByPrio365', $secret
        $scriptContent = $scriptContent -replace "'DOMAIN\.onmicrosoft\.com'", "'$customerDomain'"
        
        # Write the modified script content back to the file
        Set-Content -Path $localPath -Value $scriptContent

        # Import the runbook into Azure Automation
        Import-AzAutomationRunbook -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName -Path $localPath -Type PowerShell -Name $RunbookName -Force
    } else {
        Write-Output "Runbook '$RunbookName' already exists."
    }
}
# Function to publish a runbook
function Create-Webhook {
    param (
        [string]$RunbookName,
        [string]$ResourceGroupName,
        [string]$AutomationAccountName
    )

    # Generate a webhook name
    $webhookName = "Webhook-$RunbookName"

    # Create a new webhook
    $webhook = New-AzAutomationWebhook -Name $webhookName -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName -RunbookName $RunbookName -ExpiryTime (Get-Date).AddYears(3) -Force -IsEnabled $True

    # Get the webhook URL
    $webhookUrl = $webhook.WebhookUri

    # Output the webhook URL
    Write-Output "Webhook URL for $RunbookName - $webhookUrl - Expires on $($webhook.ExpiryTime)"
}

function Get-AutomationAccount {
    param (
        [string]$ResourceGroupName,
        [string]$AutomationAccountName,
        [int]$RetryIntervalSeconds = 10,
        [int]$MaxRetries = 6
    )

    $retries = 0
    while ($retries -lt $MaxRetries) {
        $automationAccount = Get-AzAutomationAccount -ResourceGroupName $ResourceGroupName -Name $AutomationAccountName -ErrorAction SilentlyContinue
        if ($automationAccount) {
            return $automationAccount
        } else {
            Write-Output "Automation Account '$AutomationAccountName' not found. Retrying in $RetryIntervalSeconds seconds..."
            Start-Sleep -Seconds $RetryIntervalSeconds
            $retries++
        }
    }
    Write-Error "Automation Account '$AutomationAccountName' not found after $MaxRetries retries."
}

$customerDomain = (Get-AzTenant).Domains | Where-Object { $_ -like "*.onmicrosoft.com" } | Select-Object -First 1
$secret = Generate-Secret

# Check if the Resource Group exists
$existingRG = Get-AzResourceGroup -Name $rgName -ErrorAction SilentlyContinue

if (!$existingRG) {
    # Resource Group does not exist, so create a new one
    $newRG = New-AzResourceGroup -Name $rgName -Location $location
    
    Write-Output "Resource Group '$rgName' created successfully in $location region."
} else {
    Write-Output "Resource Group '$rgName' already exists."
}

# Check if the Automation Account already exists
$existingAA = Get-AzAutomationAccount -ResourceGroupName $rgName -Name $accountName -ErrorAction SilentlyContinue


if (!$existingAA) {
    # Automation Account does not exist, so create a new one
    $automationAccount = New-AzAutomationAccount -ResourceGroupName $rgName `
                                                   -Name $accountName `
                                                   -Location $location `
                                                   -AssignSystemIdentity
    
    Write-Output "Automation Account '$accountName' created successfully in $location region."
    Write-Output "Wait until the Automation Account is created."
    Start-Sleep -Seconds 30
    $automationAccount = Get-AutomationAccount -ResourceGroupName $rgName -AutomationAccountName $accountName

} else {
    # Automation Account already exists
    Write-Output "Automation Account '$accountName' already exists in $location region."
}

$ServicePrincipal = Get-AzADServicePrincipal -DisplayName $accountName
$SPID = $ServicePrincipal.ID

$params = @{
    ServicePrincipalId = $SPID  # managed identity object id
    PrincipalId = $SPID  # managed identity object id
    ResourceId = (Get-MgServicePrincipal -Filter "AppId eq '00000002-0000-0ff1-ce00-000000000000'").id # Exchange online
    AppRoleId = "dc50a0fb-09a3-484d-be87-e023b12c6440" # Exchange.ManageAsApp
}

New-MgServicePrincipalAppRoleAssignedTo @params
$roleId = (Get-MgRoleManagementDirectoryRoleDefinition -Filter "DisplayName eq 'Exchange Administrator'").id
New-MgRoleManagementDirectoryRoleAssignment -PrincipalId $SPID -RoleDefinitionId $roleId -DirectoryScopeId "/"


$githubUrls = @{
    "AddUserToMailbox" = "https://raw.githubusercontent.com/FC-Digitalisierung-GmbH/prio365-automation/main/Runbooks/AddUserToMailbox.ps1"
    "CreateSharedMailbox" = "https://raw.githubusercontent.com/FC-Digitalisierung-GmbH/prio365-automation/main/Runbooks/CreateSharedMailbox.ps1"
    "RemoveUserFromMailbox" = "https://raw.githubusercontent.com/FC-Digitalisierung-GmbH/prio365-automation/main/Runbooks/RemoveUserFromMailbox.ps1"
}

# Create and import each runbook
foreach ($runbookName in $githubUrls.Keys) {
    Create-Runbook -RunbookName $runbookName -ResourceGroupName $rgName -AutomationAccountName $accountName -GitHubUrl $githubUrls[$runbookName]
}

# Publish each runbook
foreach ($runbookName in $githubUrls.Keys) {
    Publish-AzAutomationRunbook -ResourceGroupName $rgName -AutomationAccountName $accountName -Name $runbookName
}

# Create a webhook for each runbook
foreach ($runbookName in $githubUrls.Keys) {
    Create-Webhook -RunbookName $runbookName -ResourceGroupName $rgName -AutomationAccountName $accountName
}

# Output the webhook URLs and secret
Write-Output "Webhooks created successfully - Secret $secret."