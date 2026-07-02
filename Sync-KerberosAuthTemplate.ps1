#Requires -Version 5.1
<#
.SYNOPSIS
    Exports and imports the "Kerberos Authentication" certificate template between AD forests
    using certutil -dsTemplate / -dsAddTemplate, and (re)applies the standard AD CS enrollment
    permissions after import.

    Uses System.DirectoryServices directly - no ActiveDirectory PowerShell module required.

.DESCRIPTION
    Two modes of operation:

      -Mode Export
          Run in the SOURCE forest. Dumps the template attributes to a text file using
          "certutil -dsTemplate". Only requires read access to the template (Authenticated
          Users normally has this by default) - no CA role and no AD PowerShell module are
          needed on the machine, just certutil.

      -Mode Import
          Run in the TARGET forest. Imports the template from the file using
          "certutil -f -dsAddTemplate", and then (unless -SkipAcl) resets the ACEs for the
          following principals to the standard AD CS set:
              - Domain Controllers                              -> Read, Enroll, Autoenroll
              - Enterprise Read-only Domain Controllers (RODC)  -> Read, Enroll, Autoenroll
              - Authenticated Users                             -> Read
          These three principals are purged and re-added, so re-running is idempotent (no
          duplicate ACEs). Any OTHER ACEs that certutil seeds in the template's default
          security descriptor (e.g. Enterprise Admins / Domain Admins Full Control) are left
          intact. Requires Enterprise Admin (or delegated write access to
          CN=Certificate Templates,... in the Configuration partition).

    The principals above are resolved by well-known SID/RID - Domain Controllers = domainSID-516,
    Enterprise Read-only Domain Controllers = forest-root-domainSID-498, Authenticated Users =
    S-1-5-11 - so the script also works on non-English forests where those group names are
    localized.

    Note: the security descriptor (ACL) is not included in the certutil export, since it is
    forest-specific and doesn't make sense outside the source environment. -Mode Import
    therefore always reapplies the permissions afterwards (unless -SkipAcl is specified).

.PARAMETER Mode
    "Export" or "Import".

.PARAMETER Path
    File path for the template file (.inf/.txt). Used to write (Export) or read (Import).

.PARAMETER TemplateName
    The CN/internal name of the template as certutil knows it. Default: "KerberosAuthentication"
    (no space - this is the internal name, not the display name).

.PARAMETER TemplateDisplayName
    The display name used to locate the object in AD when setting permissions after import.
    Default: "Kerberos Authentication".

.PARAMETER Server
    Optional. Target a specific domain controller for BOTH the certutil operation and the AD
    reads/writes. Pinning every step to the same DC avoids replication races immediately after
    import (in a multi-DC / multi-site forest a serverless bind may hit a DC that has not yet
    received the just-imported object). Example: dc01.target.forest.local.

.PARAMETER SkipAcl
    Skips the permission setup after import (only relevant together with -Mode Import).

.EXAMPLE
    # In the source forest:
    .\Sync-KerberosAuthTemplate.ps1 -Mode Export -Path .\KerberosAuthentication.inf

.EXAMPLE
    # Copy the file to the target forest, then:
    .\Sync-KerberosAuthTemplate.ps1 -Mode Import -Path .\KerberosAuthentication.inf

.EXAMPLE
    # Preview an import (no changes are made). If the template already exists it previews the
    # ACL step; if it does not exist yet, it reports that the ACL would be applied after import.
    .\Sync-KerberosAuthTemplate.ps1 -Mode Import -Path .\KerberosAuthentication.inf -WhatIf

.EXAMPLE
    # Pin every step to one DC (recommended in multi-site forests):
    .\Sync-KerberosAuthTemplate.ps1 -Mode Import -Path .\KerberosAuthentication.inf -Server dc01.target.local

.NOTES
    - "certutil.exe" must be available (part of RSAT "AD CS Tools" or the AD CS role), but
      no CA role and no ActiveDirectory PowerShell module need to be installed on the machine
      the script is run from.
    - Not all certutil versions expose these switches identically. Run
      "certutil -dsTemplate -?" / "certutil -dsAddTemplate -?" locally to confirm the
      syntax on your version if something fails.
    - Import overwrites an existing template of the same name (certutil -f). The script warns
      before doing so; use -WhatIf first to preview.
    - Each run emits a PSCustomObject describing the result, and writes progress to the
      verbose stream (use -Verbose to see it).
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidateSet("Export", "Import")]
    [string]$Mode,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$Path,

    [ValidateNotNullOrEmpty()]
    [string]$TemplateName = "KerberosAuthentication",

    [ValidateNotNullOrEmpty()]
    [string]$TemplateDisplayName = "Kerberos Authentication",

    [ValidateNotNullOrEmpty()]
    [string]$Server,

    [switch]$SkipAcl
)

#region Helpers

function Assert-Certutil {
    if (-not (Get-Command certutil.exe -ErrorAction SilentlyContinue)) {
        throw "certutil.exe was not found. Install RSAT 'AD CS Tools' (e.g. Add-WindowsFeature RSAT-ADCS,RSAT-ADCS-Mgmt) and try again."
    }
}

function ConvertTo-LdapFilterValue {
    # Escapes RFC 4515 filter metacharacters for an EXACT-match assertion value.
    # Also escapes '*' so a literal asterisk in a display name is treated as a literal, not a
    # wildcard (this lookup is an exact match, unlike the wildcard search in the sibling script).
    param([string]$Value)
    $Value -replace '\\', '\5c' -replace '\(', '\28' -replace '\)', '\29' -replace '\*', '\2a' -replace "`0", '\00'
}

function Get-DsContext {
    # Binds RootDSE (optionally on a specific DC) and returns the naming contexts we need.
    param([string]$Server)

    $rootDSE = if ($Server) { [ADSI]"LDAP://$Server/RootDSE" } else { [ADSI]'LDAP://RootDSE' }
    $configNC = $rootDSE.configurationNamingContext.Value
    if (-not $configNC) {
        throw "Could not read configurationNamingContext from RootDSE (Server: '$Server'). Is the machine domain-joined / is -Server reachable?"
    }

    [PSCustomObject]@{
        ConfigNC     = $configNC
        DefaultNC    = $rootDSE.defaultNamingContext.Value
        RootDomainNC = $rootDSE.rootDomainNamingContext.Value
        TemplatesDN  = "CN=Certificate Templates,CN=Public Key Services,CN=Services,$configNC"
    }
}

function Get-NamingContextSid {
    # Reads objectSid off a domain naming context and returns it as a SecurityIdentifier.
    param([string]$Server, [string]$NamingContext)

    if (-not $NamingContext) { throw "Naming context not available for SID resolution." }

    $path = if ($Server) { "LDAP://$Server/$NamingContext" } else { "LDAP://$NamingContext" }
    $de = New-Object System.DirectoryServices.DirectoryEntry($path)
    try {
        $bytes = $de.Properties['objectSid'].Value
        if (-not $bytes) { throw "Could not read objectSid from '$NamingContext'." }
        return New-Object System.Security.Principal.SecurityIdentifier([byte[]]$bytes, 0)
    }
    finally {
        $de.Dispose()
    }
}

function Find-TemplateDN {
    # Returns the DistinguishedName of the pKICertificateTemplate with the given display name,
    # or $null if not found. Throws if more than one matches (a broadened/ambiguous lookup).
    param([string]$Server, [string]$TemplatesDN, [string]$DisplayName)

    $path = if ($Server) { "LDAP://$Server/$TemplatesDN" } else { "LDAP://$TemplatesDN" }
    $base = New-Object System.DirectoryServices.DirectoryEntry($path)
    $filter = "(&(objectClass=pKICertificateTemplate)(displayName=$(ConvertTo-LdapFilterValue $DisplayName)))"

    $searcher = New-Object System.DirectoryServices.DirectorySearcher($base, $filter)
    $searcher.SearchScope = [System.DirectoryServices.SearchScope]::OneLevel
    [void]$searcher.PropertiesToLoad.Add('distinguishedName')
    try {
        $results = $searcher.FindAll()
        try {
            $found = @($results)
            if ($found.Count -eq 0) { return $null }
            if ($found.Count -gt 1) {
                throw "Multiple templates matched displayName '$DisplayName' under $TemplatesDN; refusing to continue."
            }
            return [string]$found[0].Properties['distinguishedname'][0]
        }
        finally {
            $results.Dispose()
        }
    }
    finally {
        $searcher.Dispose()
        $base.Dispose()
    }
}

#endregion Helpers

function Export-Template {
    param([string]$TemplateName, [string]$Path, [string]$Server)

    Assert-Certutil

    Write-Verbose "Exporting template '$TemplateName' to '$Path'..."

    $certArgs = @()
    if ($Server) { $certArgs += @('-dc', $Server) }
    $certArgs += @('-dsTemplate', $TemplateName)

    # Capture stdout only into the file; keep stderr separate so warnings/errors are never baked
    # into the exported .inf. Force UTF-8 so non-ASCII template names/descriptions survive.
    $errFile = [System.IO.Path]::GetTempFileName()
    $prevEncoding = $null
    try {
        try { $prevEncoding = [Console]::OutputEncoding; [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

        $output = & certutil.exe @certArgs 2>$errFile
        if ($LASTEXITCODE -ne 0) {
            $errText = Get-Content -LiteralPath $errFile -Raw -ErrorAction SilentlyContinue
            throw "certutil -dsTemplate failed (exit code $LASTEXITCODE):`n$($output -join "`n")`n$errText"
        }

        $output | Out-File -FilePath $Path -Encoding utf8 -Force
    }
    finally {
        if ($null -ne $prevEncoding) { try { [Console]::OutputEncoding = $prevEncoding } catch { } }
        Remove-Item -LiteralPath $errFile -ErrorAction SilentlyContinue
    }

    Write-Verbose "Export completed: $Path"

    [PSCustomObject]@{
        Mode         = 'Export'
        TemplateName = $TemplateName
        Path         = (Resolve-Path -LiteralPath $Path).ProviderPath
        Status       = 'Exported'
    }
}

function Import-Template {
    # Returns $true if the template was actually imported, $false if the import was skipped
    # (e.g. -WhatIf or a declined -Confirm).
    param(
        [string]$Path,
        [string]$TemplateDisplayName,
        [string]$Server,
        [System.Management.Automation.PSCmdlet]$CallerCmdlet
    )

    Assert-Certutil

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "File '$Path' was not found."
    }

    # Warn (also under -WhatIf) if a template with this display name already exists: certutil -f
    # will overwrite it, discarding any target-forest customizations.
    try {
        $ctx = Get-DsContext -Server $Server
        $existingDN = Find-TemplateDN -Server $Server -TemplatesDN $ctx.TemplatesDN -DisplayName $TemplateDisplayName
        if ($existingDN) {
            Write-Warning "A template with display name '$TemplateDisplayName' already exists ($existingDN) and will be OVERWRITTEN (certutil -f); any target-forest customizations to it will be lost."
        }
    }
    catch {
        Write-Verbose "Pre-import existence check skipped: $_"
    }

    if ($CallerCmdlet.ShouldProcess($Path, "Import certificate template into AD (certutil -f -dsAddTemplate; OVERWRITES an existing template of the same name)")) {
        Write-Verbose "Importing template from '$Path'..."

        $certArgs = @()
        if ($Server) { $certArgs += @('-dc', $Server) }
        $certArgs += @('-f', '-dsAddTemplate', $Path)

        $output = & certutil.exe @certArgs 2>&1
        if ($LASTEXITCODE -ne 0) {
            throw "certutil -dsAddTemplate failed (exit code $LASTEXITCODE):`n$($output -join "`n")"
        }

        Write-Verbose ($output -join "`n")
        Write-Verbose "Import completed."
        return $true
    }

    return $false
}

function Set-TemplateDefaultAcl {
    # Returns $true if the ACL was applied, $false otherwise (e.g. -WhatIf, or the template was
    # not created because the import was skipped).
    param(
        [string]$TemplateDisplayName,
        [string]$Server,
        [bool]$TemplateWasImported,
        [System.Management.Automation.PSCmdlet]$CallerCmdlet
    )

    $ctx = Get-DsContext -Server $Server

    # Locate the template object. After a real import, tolerate brief replication lag when the
    # read is not pinned to the same DC that certutil wrote to.
    $templateDN = $null
    $maxAttempts = if ($TemplateWasImported -and -not $Server) { 6 } else { 1 }
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        $templateDN = Find-TemplateDN -Server $Server -TemplatesDN $ctx.TemplatesDN -DisplayName $TemplateDisplayName
        if ($templateDN) { break }
        if ($attempt -lt $maxAttempts) {
            Write-Verbose "Template not visible yet (attempt $attempt/$maxAttempts); waiting for replication..."
            Start-Sleep -Seconds 5
        }
    }

    if (-not $templateDN) {
        if (-not $TemplateWasImported) {
            # Import was skipped (e.g. -WhatIf on a template that does not exist yet): a clean
            # preview, not an error.
            Write-Warning "Template '$TemplateDisplayName' not found (import was skipped, e.g. -WhatIf); the standard ACL would be applied after a real import."
            return $false
        }
        throw "Could not find template '$TemplateDisplayName' under $($ctx.TemplatesDN) after import. Check the certutil output above - was the object actually created?"
    }

    Write-Verbose "Found template: $templateDN"

    if (-not $CallerCmdlet.ShouldProcess($templateDN, "Set standard AD CS permissions (DC/RODC: Read+Enroll+Autoenroll, Authenticated Users: Read)")) {
        return $false
    }

    # Resolve the well-known principals by SID/RID (locale-independent).
    $domainSid     = Get-NamingContextSid -Server $Server -NamingContext $ctx.DefaultNC
    $rootDomainSid = Get-NamingContextSid -Server $Server -NamingContext $ctx.RootDomainNC
    $dcSID         = New-Object System.Security.Principal.SecurityIdentifier("$($domainSid.Value)-516")      # Domain Controllers
    $rodcSID       = New-Object System.Security.Principal.SecurityIdentifier("$($rootDomainSid.Value)-498")  # Enterprise Read-only Domain Controllers
    $authUsersSID  = New-Object System.Security.Principal.SecurityIdentifier("S-1-5-11")                     # Authenticated Users

    $EnrollGUID     = [Guid]"0e10c968-78fb-11d2-90d4-00c04f79dc55"  # Certificate-Enrollment
    $AutoEnrollGUID = [Guid]"a05b8cc2-17bc-4802-a710-e7c15ab866a2"  # Certificate-AutoEnrollment

    $ldapPath = if ($Server) { "LDAP://$Server/$templateDN" } else { "LDAP://$templateDN" }
    $entry    = New-Object System.DirectoryServices.DirectoryEntry($ldapPath)
    # Read/write only the DACL so CommitChanges does not attempt to rewrite owner/group.
    $entry.Options.SecurityMasks = [System.DirectoryServices.SecurityMasks]::Dacl
    $sec = $entry.ObjectSecurity

    function Add-TemplateAce {
        param(
            [System.DirectoryServices.ActiveDirectorySecurity]$Security,
            [System.Security.Principal.IdentityReference]$Sid,
            [System.DirectoryServices.ActiveDirectoryRights]$Rights,
            [Guid]$ObjectType = [Guid]::Empty
        )
        if ($ObjectType -eq [Guid]::Empty) {
            $ace = New-Object System.DirectoryServices.ActiveDirectoryAccessRule(
                $Sid, $Rights, [System.Security.AccessControl.AccessControlType]::Allow
            )
        } else {
            $ace = New-Object System.DirectoryServices.ActiveDirectoryAccessRule(
                $Sid, $Rights, [System.Security.AccessControl.AccessControlType]::Allow, $ObjectType
            )
        }
        $Security.AddAccessRule($ace)
    }

    # Purge only the three principals we manage, then re-add their exact ACEs. This keeps the
    # result idempotent (no duplicate ACEs on re-run) while leaving certutil's default admin
    # ACEs (Enterprise/Domain Admins Full Control, etc.) untouched.
    $sec.PurgeAccessRules($dcSID)
    $sec.PurgeAccessRules($rodcSID)
    $sec.PurgeAccessRules($authUsersSID)

    Add-TemplateAce -Security $sec -Sid $dcSID -Rights ([System.DirectoryServices.ActiveDirectoryRights]::GenericRead)
    Add-TemplateAce -Security $sec -Sid $dcSID -Rights ([System.DirectoryServices.ActiveDirectoryRights]::ExtendedRight) -ObjectType $EnrollGUID
    Add-TemplateAce -Security $sec -Sid $dcSID -Rights ([System.DirectoryServices.ActiveDirectoryRights]::ExtendedRight) -ObjectType $AutoEnrollGUID

    Add-TemplateAce -Security $sec -Sid $rodcSID -Rights ([System.DirectoryServices.ActiveDirectoryRights]::GenericRead)
    Add-TemplateAce -Security $sec -Sid $rodcSID -Rights ([System.DirectoryServices.ActiveDirectoryRights]::ExtendedRight) -ObjectType $EnrollGUID
    Add-TemplateAce -Security $sec -Sid $rodcSID -Rights ([System.DirectoryServices.ActiveDirectoryRights]::ExtendedRight) -ObjectType $AutoEnrollGUID

    Add-TemplateAce -Security $sec -Sid $authUsersSID -Rights ([System.DirectoryServices.ActiveDirectoryRights]::GenericRead)

    $entry.ObjectSecurity = $sec
    $entry.CommitChanges()
    $entry.Dispose()

    Write-Host "Permissions set on '$templateDN':" -ForegroundColor Green
    Write-Host " - Domain Controllers: Read, Enroll, Autoenroll"
    Write-Host " - Enterprise Read-only Domain Controllers: Read, Enroll, Autoenroll"
    Write-Host " - Authenticated Users: Read"
    return $true
}

# --- Main logic ---
switch ($Mode) {
    "Export" {
        $result = Export-Template -TemplateName $TemplateName -Path $Path -Server $Server
        Write-Host "Export completed: $($result.Path)" -ForegroundColor Green
        Write-Host "Copy this file to the target forest and run the script there with -Mode Import." -ForegroundColor Yellow
        $result
    }
    "Import" {
        $imported = Import-Template -Path $Path -TemplateDisplayName $TemplateDisplayName -Server $Server -CallerCmdlet $PSCmdlet

        $aclApplied = $false
        if (-not $SkipAcl) {
            $aclApplied = Set-TemplateDefaultAcl -TemplateDisplayName $TemplateDisplayName -Server $Server -TemplateWasImported $imported -CallerCmdlet $PSCmdlet
        }
        elseif ($imported) {
            Write-Host "Skipped permission setup (-SkipAcl was specified)." -ForegroundColor Yellow
        }

        $status =
            if (-not $imported)  { 'Skipped (WhatIf/declined)' }
            elseif ($SkipAcl)    { 'Imported (ACL skipped)' }
            elseif ($aclApplied) { 'Imported + ACL set' }
            else                 { 'Imported (ACL not applied)' }

        [PSCustomObject]@{
            Mode                = 'Import'
            TemplateName        = $TemplateName
            TemplateDisplayName = $TemplateDisplayName
            Imported            = $imported
            AclApplied          = $aclApplied
            Status              = $status
        }
    }
}
