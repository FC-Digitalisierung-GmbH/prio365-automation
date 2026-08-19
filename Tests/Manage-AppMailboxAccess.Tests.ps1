BeforeAll {
    $script:RunbookPath = Join-Path $PSScriptRoot '..' 'Runbooks' 'Manage-AppMailboxAccess.ps1'
    # Dot-Source: durch InvocationName '.' wird der Main-Block NICHT ausgeführt (nur Funktionen geladen)
    . $script:RunbookPath
    function Get-DistributionGroupMember {}
    function Add-DistributionGroupMember {}
    function Remove-DistributionGroupMember {}
}

Describe 'Set-MailboxScopeMembership' {
    Context 'Add' {
        It 'fügt hinzu, wenn noch kein Mitglied' {
            Mock Get-DistributionGroupMember { @() }
            Mock Add-DistributionGroupMember { }
            $r = Set-MailboxScopeMembership -Action Add -MailboxSmtp 'a@x.de' -GroupIdentity 'Prio365-MailboxScope'
            $r | Should -Be 'Added'
            Should -Invoke Add-DistributionGroupMember -Times 1 -Exactly
        }
        It 'ist idempotent: kein Add, wenn bereits Mitglied' {
            Mock Get-DistributionGroupMember { [pscustomobject]@{ PrimarySmtpAddress = 'a@x.de' } }
            Mock Add-DistributionGroupMember { }
            $r = Set-MailboxScopeMembership -Action Add -MailboxSmtp 'a@x.de' -GroupIdentity 'Prio365-MailboxScope'
            $r | Should -Be 'NoChange'
            Should -Invoke Add-DistributionGroupMember -Times 0 -Exactly
        }
    }
    Context 'Remove' {
        It 'entfernt, wenn Mitglied' {
            Mock Get-DistributionGroupMember { [pscustomobject]@{ PrimarySmtpAddress = 'a@x.de' } }
            Mock Remove-DistributionGroupMember { }
            $r = Set-MailboxScopeMembership -Action Remove -MailboxSmtp 'a@x.de' -GroupIdentity 'Prio365-MailboxScope'
            $r | Should -Be 'Removed'
            Should -Invoke Remove-DistributionGroupMember -Times 1 -Exactly
        }
        It 'ist idempotent: kein Remove, wenn kein Mitglied' {
            Mock Get-DistributionGroupMember { @() }
            Mock Remove-DistributionGroupMember { }
            $r = Set-MailboxScopeMembership -Action Remove -MailboxSmtp 'a@x.de' -GroupIdentity 'Prio365-MailboxScope'
            $r | Should -Be 'NoChange'
            Should -Invoke Remove-DistributionGroupMember -Times 0 -Exactly
        }
    }
}

Describe 'Invoke-ManageBatch' {
    It 'verarbeitet weitere Items, wenn eines wirft (per-Item try/catch)' {
        $script:added = 0
        Mock Get-DistributionGroupMember { @() }
        Mock Add-DistributionGroupMember {
            if ($Member -eq 'boom@x.de') { throw 'simulierter Fehler' }
            $script:added++
        }
        $items = @(
            [pscustomobject]@{ action = 'Add'; mailboxSmtp = 'boom@x.de' },
            [pscustomobject]@{ action = 'Add'; mailboxSmtp = 'ok@x.de' }
        )
        $summary = Invoke-ManageBatch -Items $items -GroupIdentity 'Prio365-MailboxScope'
        $script:added | Should -Be 1
        $summary.Failed  | Should -Be 1
        $summary.Changed | Should -Be 1
    }
}
