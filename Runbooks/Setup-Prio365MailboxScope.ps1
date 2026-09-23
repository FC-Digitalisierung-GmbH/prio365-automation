param
(
    [Parameter(Mandatory = $true)]  [string] $ServicePrincipalsJson,
    [Parameter(Mandatory = $true)]  [string] $OrganizationDomain,
    [Parameter(Mandatory = $false)] [string] $VerifyInScopeMailbox,
    [Parameter(Mandatory = $false)] [string] $VerifyOutOfScopeMailbox
)

$ErrorActionPreference = 'Stop'

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
    if ($existing) { return $existing }

    # Neu anlegen: das ZURÜCKGEGEBENE Objekt direkt verwenden (hat DistinguishedName sofort) –
    # nicht auf ein erneutes Get verlassen, das wegen Exchange-Replikation kurz null liefern kann
    # (genau das führte in den Folgeschritten zu null-Referenzen).
    $created = New-DistributionGroup -Name $Name -Alias $Alias -Type Security -ErrorAction Stop
    if ($created) { return $created }

    # Fallback, falls New- ausnahmsweise nichts zurückgibt: kurz auf Replikation warten und erneut lesen.
    for ($i = 0; $i -lt 6 -and -not $existing; $i++) {
        Start-Sleep -Seconds 10
        $existing = Get-DistributionGroup -Identity $Name -ErrorAction SilentlyContinue
    }
    if (-not $existing) {
        throw "Scope-Gruppe '$Name' ist nach der Erstellung nicht auffindbar (Exchange-Replikationsverzögerung)."
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
        $identity = Confirm-ExoServicePrincipal -AppId $sp.AppId -ObjectId $sp.ObjectId -DisplayName "prio365-mailSp-$i"
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
    Write-Output "STEP: AzAccount connected"

    try {
        Write-Output "STEP: OrganizationDomain='$OrganizationDomain' (Länge=$($OrganizationDomain.Length))"
        $exoCmd = Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue
        Write-Output "STEP: ExchangeOnlineManagement module = $($exoCmd.Source) v$($exoCmd.Version)"

        # Diagnose des mutmaßlichen Az.Accounts-5.x-Breaking-Changes (Get-AzAccessToken -> SecureString),
        # der Connect-ExchangeOnline -ManagedIdentity mit einer NRE brechen lässt.
        $azAcc = Get-Module Az.Accounts -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
        Write-Output "STEP: Az.Accounts v$($azAcc.Version)"
        try {
            $diagTok = Get-AzAccessToken -ResourceUrl 'https://outlook.office365.com' -ErrorAction Stop
            Write-Output "STEP: Get-AzAccessToken ok, Token-Typ = $($diagTok.Token.GetType().Name)"
        }
        catch {
            Write-Output "STEP: Get-AzAccessToken failed: $($_.Exception.Message)"
        }

        # Fail-fast: ein Versuch. Der Fehler wird vom äußeren catch mit Typ/InnerException geloggt,
        # der Job endet sofort auf Failed (kein minutenlanges Retry, das den Backend-Timeout reißt).
        Connect-ExchangeOnline -ManagedIdentity -Organization $OrganizationDomain -ShowBanner:$false
        Write-Output "STEP: ExchangeOnline connected (org=$OrganizationDomain)"

        $sps = ConvertFrom-Json -InputObject $ServicePrincipalsJson
        $sps = @($sps)
        if (-not $sps -or $sps.Count -eq 0) {
            throw "ServicePrincipalsJson enthält keine Service Principals (nach ConvertFrom-Json leer)."
        }
        Write-Output "STEP: parsed $($sps.Count) service principal(s)"

        $group = Confirm-ScopeGroup -Name $ScopeGroupName -Alias $ScopeGroupAlias
        Write-Output "STEP: scope group ready (DN='$($group.DistinguishedName)')"

        Confirm-ManagementScope -Name $ScopeName -GroupDn $group.DistinguishedName
        Write-Output "STEP: management scope ensured"

        Invoke-SetupForServicePrincipals -ServicePrincipals $sps -Group $group -ScopeName $ScopeName -Roles $Roles
        Write-Output "STEP: role assignments processed"

        $ready = $true
        if ($VerifyInScopeMailbox -and $VerifyOutOfScopeMailbox) {
            foreach ($sp in $sps) {
                if (-not (Test-ScopeGate -AppId $sp.AppId -InScopeMailbox $VerifyInScopeMailbox -OutOfScopeMailbox $VerifyOutOfScopeMailbox)) { $ready = $false }
            }
        }
        Write-Output "ScopeGroupEmail=$($group.PrimarySmtpAddress)"
        Write-Output "SCOPE_READY=$($ready.ToString().ToLower())"
    }
    catch {
        # Genaue Fehlerstelle sichtbar machen (der Azure-"Ausnahme"-Tab zeigt sonst nur die Message).
        Write-Error "FAILED: $($_.Exception.Message)"
        Write-Error "TYPE: $($_.Exception.GetType().FullName)"
        Write-Error "INNER: $($_.Exception.InnerException.Message)"
        Write-Error "AT: $($_.InvocationInfo.PositionMessage)"
        Write-Error "STACK: $($_.ScriptStackTrace)"
        throw
    }
    finally { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue }
}

# Guard: bei Pester-Dot-Source (InvocationName '.') NICHT ausführen
if ($MyInvocation.InvocationName -ne '.') { Invoke-Main }
