#Uninstall-Module AzureRM
#Install-Module -Name Az -Repository PSGallery -Force -scope CurrentUser -AllowClobber
#Install-Module Microsoft.Graph -Scope CurrentUser

#Connect-AzAccount
$tenant = (Get-AzTenant).Id
Connect-Graph -TenantId $tenant
$accountName = 'Prio365-Automation'
$rgName = 'Prio365'
$location = "westeurope"


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
$existingAA


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

if (!$existingAA) {
    # Automation Account does not exist, so create a new one
    $automationAccount = New-AzAutomationAccount -ResourceGroupName $rgName `
                                                   -Name $accountName `
                                                   -Location $location `
                                                   -AssignSystemIdentity
    
    Write-Output "Automation Account '$accountName' created successfully in $location region."
 
    Start-Sleep -Seconds 10
    $automationAccount = Get-AutomationAccount -ResourceGroupName $rgName -AutomationAccountName $accountName

} else {
    # Automation Account already exists
    Write-Output "Automation Account '$accountName' already exists in $location region."
}

$ServicePrincipal = Get-AzADServicePrincipal -DisplayName $accountName
$SPID = $ServicePrincipal.ID



Connect-MgGraph -Scopes AppRoleAssignment.ReadWrite.All,Application.Read.All,Directory.AccessAsUser.All
$params = @{
    ServicePrincipalId = $SPID  # managed identity object id
    PrincipalId = $SPID  # managed identity object id
    ResourceId = (Get-MgServicePrincipal -Filter "AppId eq '00000002-0000-0ff1-ce00-000000000000'").id # Exchange online
    AppRoleId = "dc50a0fb-09a3-484d-be87-e023b12c6440" # Exchange.ManageAsApp
}
New-MgServicePrincipalAppRoleAssignedTo @params

$roleId = (Get-MgRoleManagementDirectoryRoleDefinition -Filter "DisplayName eq 'Exchange Administrator'").id
New-MgRoleManagementDirectoryRoleAssignment -PrincipalId $SPID -RoleDefinitionId $roleId -DirectoryScopeId "/"

# Create AddUserToMailbox
$runbookName = "AddUserToMailbox"
$runbook = New-AzAutomationRunbook -ResourceGroupName $rgName -AutomationAccountName $accountName -Name $runbookName -Type PowerShell


# Create CreateSharedMailbox
$runbookName = "CreateSharedMailbox"
$runbook = New-AzAutomationRunbook -ResourceGroupName $rgName -AutomationAccountName $accountName -Name $runbookName -Type PowerShell


# Create RemoveUserFromMailbox
$runbookName = "RemoveUserFromMailbox"
$runbook = New-AzAutomationRunbook -ResourceGroupName $rgName -AutomationAccountName $accountName -Name $runbookName -Type PowerShell