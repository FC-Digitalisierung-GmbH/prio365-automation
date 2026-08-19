param
(
    [Parameter(Mandatory = $true)]  [string] $SpAppId,
    [Parameter(Mandatory = $true)]  [string] $SpObjectId,
    [Parameter(Mandatory = $true)]  [string] $OrganizationDomain,
    [Parameter(Mandatory = $false)] [string] $VerifyInScopeMailbox,
    [Parameter(Mandatory = $false)] [string] $VerifyOutOfScopeMailbox
)

$ScopeGroupName  = 'Prio365-MailboxScope'
$ScopeGroupAlias = 'prio365-mailboxscope'
$ScopeName       = 'Prio365-MailboxScope'
$Roles = @(
    'Application Mail.ReadWrite','Application Mail.Send',
    'Application MailboxSettings.Read','Application Calendars.ReadWrite'
)

function Confirm-ScopeGroup {
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][string]$Alias)
    $existing = Get-DistributionGroup -Identity $Name -ErrorAction SilentlyContinue
    if (-not $existing) {
        New-DistributionGroup -Name $Name -Alias $Alias -Type Security -ErrorAction Stop | Out-Null
        Start-Sleep -Seconds 5
        $existing = Get-DistributionGroup -Identity $Name -ErrorAction Stop
    }
    return $existing
}

function Confirm-ExoServicePrincipal {
    param([Parameter(Mandatory)][string]$AppId, [Parameter(Mandatory)][string]$ObjectId)
    if (-not (Get-ServicePrincipal -Identity $AppId -ErrorAction SilentlyContinue)) {
        New-ServicePrincipal -AppId $AppId -ObjectId $ObjectId -DisplayName 'prio365-sp' -ErrorAction Stop | Out-Null
    }
}

function Confirm-ManagementScope {
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][string]$GroupDn)
    if (-not (Get-ManagementScope -Identity $Name -ErrorAction SilentlyContinue)) {
        New-ManagementScope -Name $Name -RecipientRestrictionFilter "MemberOfGroup -eq '$GroupDn'" -ErrorAction Stop | Out-Null
    }
}

function Confirm-RoleAssignments {
    param(
        [Parameter(Mandatory)][string]$AppId,
        [Parameter(Mandatory)][string]$ScopeName,
        [Parameter(Mandatory)][string[]]$Roles
    )
    $spId = (Get-ServicePrincipal -Identity $AppId -ErrorAction Stop).Identity
    $existing = Get-ManagementRoleAssignment -RoleAssignee $AppId -ErrorAction SilentlyContinue |
                Where-Object { $_.CustomResourceScope -eq $ScopeName }
    foreach ($role in $Roles) {
        if (-not ($existing | Where-Object { $_.Role -eq $role })) {
            New-ManagementRoleAssignment -App $spId -Role $role -CustomResourceScope $ScopeName -ErrorAction Stop | Out-Null
            Write-Output "  + RoleAssignment: $role"
        }
    }
}

function Test-ScopeGate {
    param(
        [Parameter(Mandatory)][string]$AppId,
        [Parameter(Mandatory)][string]$InScopeMailbox,
        [Parameter(Mandatory)][string]$OutOfScopeMailbox
    )
    $inRes  = Test-ServicePrincipalAuthorization -Identity $AppId -Resource $InScopeMailbox  -ErrorAction SilentlyContinue
    $outRes = Test-ServicePrincipalAuthorization -Identity $AppId -Resource $OutOfScopeMailbox -ErrorAction SilentlyContinue
    $inGranted  = [bool]($inRes  | Where-Object { $_.InScope -eq $true })
    $outGranted = [bool]($outRes | Where-Object { $_.InScope -eq $true })
    return ($inGranted -and -not $outGranted)
}

function Invoke-Main {
    Connect-AzAccount -Identity | Out-Null
    Connect-ExchangeOnline -ManagedIdentity -Organization $OrganizationDomain

    try {
        $group = Confirm-ScopeGroup -Name $ScopeGroupName -Alias $ScopeGroupAlias
        Confirm-ExoServicePrincipal -AppId $SpAppId -ObjectId $SpObjectId
        Confirm-ManagementScope -Name $ScopeName -GroupDn $group.DistinguishedName
        Confirm-RoleAssignments -AppId $SpAppId -ScopeName $ScopeName -Roles $Roles

        $ready = $true
        if ($VerifyInScopeMailbox -and $VerifyOutOfScopeMailbox) {
            $ready = Test-ScopeGate -AppId $SpAppId -InScopeMailbox $VerifyInScopeMailbox -OutOfScopeMailbox $VerifyOutOfScopeMailbox
        }
        Write-Output "ScopeGroupEmail=$($group.PrimarySmtpAddress)"
        Write-Output "SCOPE_READY=$($ready.ToString().ToLower())"
    }
    finally {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}

# Guard: bei Pester-Dot-Source (InvocationName '.') NICHT ausführen
if ($MyInvocation.InvocationName -ne '.') { Invoke-Main }
