param
(
    [Parameter(Mandatory = $true)]  [string] $ServicePrincipalsJson,
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
    param([Parameter(Mandatory)][string]$AppId, [Parameter(Mandatory)][string]$ObjectId, [Parameter(Mandatory)][string]$DisplayName)
    $sp = Get-ServicePrincipal -Identity $AppId -ErrorAction SilentlyContinue
    if (-not $sp) {
        $sp = New-ServicePrincipal -AppId $AppId -ObjectId $ObjectId -DisplayName $DisplayName -ErrorAction Stop
    }
    return $sp.Identity
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
        [Parameter(Mandatory)][string]$SpIdentity,
        [Parameter(Mandatory)][string]$ScopeName,
        [Parameter(Mandatory)][string[]]$Roles
    )
    $existing = Get-ManagementRoleAssignment -RoleAssignee $AppId -ErrorAction SilentlyContinue |
                Where-Object { $_.CustomResourceScope -eq $ScopeName }
    foreach ($role in $Roles) {
        if (-not ($existing | Where-Object { $_.Role -eq $role })) {
            New-ManagementRoleAssignment -App $SpIdentity -Role $role -CustomResourceScope $ScopeName -ErrorAction Stop | Out-Null
            Write-Output "  + [$AppId] RoleAssignment: $role"
        }
    }
}

function Invoke-SetupForServicePrincipals {
    param([Parameter(Mandatory)][object[]]$ServicePrincipals, [Parameter(Mandatory)][object]$Group,
          [Parameter(Mandatory)][string]$ScopeName, [Parameter(Mandatory)][string[]]$Roles)
    $i = 0
    foreach ($sp in $ServicePrincipals) {
        $i++
        $identity = Confirm-ExoServicePrincipal -AppId $sp.AppId -ObjectId $sp.ObjectId -DisplayName "prio365-sp-$i"
        Confirm-RoleAssignments -AppId $sp.AppId -SpIdentity $identity -ScopeName $ScopeName -Roles $Roles
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
        $sps   = ConvertFrom-Json -InputObject $ServicePrincipalsJson
        $group = Confirm-ScopeGroup -Name $ScopeGroupName -Alias $ScopeGroupAlias
        Confirm-ManagementScope -Name $ScopeName -GroupDn $group.DistinguishedName
        Invoke-SetupForServicePrincipals -ServicePrincipals @($sps) -Group $group -ScopeName $ScopeName -Roles $Roles

        $ready = $true
        if ($VerifyInScopeMailbox -and $VerifyOutOfScopeMailbox) {
            foreach ($sp in @($sps)) {
                if (-not (Test-ScopeGate -AppId $sp.AppId -InScopeMailbox $VerifyInScopeMailbox -OutOfScopeMailbox $VerifyOutOfScopeMailbox)) { $ready = $false }
            }
        }
        Write-Output "ScopeGroupEmail=$($group.PrimarySmtpAddress)"
        Write-Output "SCOPE_READY=$($ready.ToString().ToLower())"
    }
    finally { Disconnect-ExchangeOnline -Confirm:$false }
}

# Guard: bei Pester-Dot-Source (InvocationName '.') NICHT ausführen
if ($MyInvocation.InvocationName -ne '.') { Invoke-Main }
