BeforeAll {
    $script:RunbookPath = Join-Path $PSScriptRoot '..' 'Runbooks' 'Setup-Prio365MailboxScope.ps1'
    . $script:RunbookPath
    $script:Roles = @(
        'Application Mail.ReadWrite','Application Mail.Send',
        'Application MailboxSettings.Read','Application Calendars.ReadWrite'
    )
    function Get-DistributionGroup {}
    function New-DistributionGroup {}
    function Get-ServicePrincipal {}
    function New-ServicePrincipal {}
    function Get-ManagementScope {}
    function New-ManagementScope {}
    function Get-ManagementRoleAssignment {}
    function New-ManagementRoleAssignment {}
    function Test-ServicePrincipalAuthorization {}
}

Describe 'Confirm-ScopeGroup' {
    It 'legt Gruppe an, wenn nicht vorhanden' {
        Mock Get-DistributionGroup { $null }
        Mock New-DistributionGroup { [pscustomobject]@{ PrimarySmtpAddress='g@x.de'; DistinguishedName='CN=g' } }
        $g = Confirm-ScopeGroup -Name 'Prio365-MailboxScope' -Alias 'prio365-mailboxscope'
        Should -Invoke New-DistributionGroup -Times 1 -Exactly
        $g.DistinguishedName | Should -Be 'CN=g'
    }
    It 'ist idempotent, wenn Gruppe existiert' {
        Mock Get-DistributionGroup { [pscustomobject]@{ PrimarySmtpAddress='g@x.de'; DistinguishedName='CN=g' } }
        Mock New-DistributionGroup { }
        Confirm-ScopeGroup -Name 'Prio365-MailboxScope' -Alias 'prio365-mailboxscope' | Out-Null
        Should -Invoke New-DistributionGroup -Times 0 -Exactly
    }
}

Describe 'Confirm-RoleAssignments' {
    It 'legt genau die fehlenden gescopten RoleAssignments an' {
        Mock Get-ManagementRoleAssignment { @() }   # nichts vorhanden
        Mock New-ManagementRoleAssignment { }
        Confirm-RoleAssignments -AppId 'app1' -SpIdentity 'spId-1' -ScopeName 'Prio365-MailboxScope' -Roles $script:Roles
        Should -Invoke New-ManagementRoleAssignment -Times 4 -Exactly
    }
    It 'ist idempotent, wenn alle RoleAssignments existieren' {
        Mock Get-ManagementRoleAssignment {
            $script:Roles | ForEach-Object { [pscustomobject]@{ Role = $_; CustomResourceScope = 'Prio365-MailboxScope' } }
        }
        Mock New-ManagementRoleAssignment { }
        Confirm-RoleAssignments -AppId 'app1' -SpIdentity 'spId-1' -ScopeName 'Prio365-MailboxScope' -Roles $script:Roles
        Should -Invoke New-ManagementRoleAssignment -Times 0 -Exactly
    }
}

Describe 'Invoke-SetupForServicePrincipals' {
    It 'legt SP + RoleAssignments für JEDEN SP an' {
        Mock Get-DistributionGroup { [pscustomobject]@{ PrimarySmtpAddress='g@x.de'; DistinguishedName='CN=g' } }
        Mock Get-ManagementScope { $true }
        Mock Get-ServicePrincipal { $null }           # keiner existiert
        Mock New-ServicePrincipal { }
        Mock Get-ManagementRoleAssignment { @() }
        Mock New-ManagementRoleAssignment { }
        $sps = @(
            [pscustomobject]@{ AppId='a1'; ObjectId='o1' },
            [pscustomobject]@{ AppId='a2'; ObjectId='o2' }
        )
        Invoke-SetupForServicePrincipals -ServicePrincipals $sps -Group ([pscustomobject]@{ DistinguishedName='CN=g' }) -ScopeName 'Prio365-MailboxScope' -Roles @('Application Mail.ReadWrite')
        Should -Invoke New-ServicePrincipal -Times 2 -Exactly            # 2 SPs
        Should -Invoke New-ManagementRoleAssignment -Times 2 -Exactly    # 1 Rolle × 2 SPs
    }
}

Describe 'Test-ScopeGate' {
    It 'true, wenn In-Scope granted und Out-of-Scope denied' {
        Mock Test-ServicePrincipalAuthorization {
            param($Identity,$Resource)
            if ($Resource -eq 'in@x.de')  { [pscustomobject]@{ InScope = $true } }
            else                          { [pscustomobject]@{ InScope = $false } }
        }
        (Test-ScopeGate -AppId 'app1' -InScopeMailbox 'in@x.de' -OutOfScopeMailbox 'out@x.de') | Should -BeTrue
    }
    It 'false, wenn Out-of-Scope fälschlich granted (Union-Fallstrick)' {
        Mock Test-ServicePrincipalAuthorization { [pscustomobject]@{ InScope = $true } }
        (Test-ScopeGate -AppId 'app1' -InScopeMailbox 'in@x.de' -OutOfScopeMailbox 'out@x.de') | Should -BeFalse
    }
}
