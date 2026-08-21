#Requires -Modules ActiveDirectory

<#
.SYNOPSIS
    Copies a certificate template from one AD forest to another using direct LDAP attribute copy -
    either through a JSON file (Export/Import) or forest-to-forest in a single run (Sync).
    Optionally renames the template, either preserves or regenerates the template OID, and applies
    standard AD CS permissions after import. Works even against a target forest that has never had
    AD CS (Certificate Services) installed.

.DESCRIPTION
    This is an LDAP-based template copy (no certutil). It uses the ActiveDirectory PowerShell
    module for ALL modes, so the module is required for Export as well as Import/Sync.

      -Mode Export
          Run in the SOURCE forest. Reads the template's functional attributes (flags, revision,
          and all msPKI-*/pKI* attributes) via LDAP and writes them to a JSON file. Forest-specific
          data is deliberately NOT exported: the security descriptor, distinguishedName, and
          objectCategory are left out because they are derived/reapplied in the target forest.
          The source template OID and identity fields (name/displayName) can additionally be
          stripped from the file (-StripOid / -StripIdentity).

      -Mode Import
          Run in the TARGET forest. Recreates the template from the JSON file via LDAP:
            * handles the template OID per -OidHandling (Preserve carries the source OID; Generate
              mints under the forest's real OID root; GenerateFromRoot mints under a base you supply
              in -OidRoot; GenerateRandom mints under a synthesized base). Every mode also registers an
              OID "display" object so Windows resolves the OID to the template name. Only "Generate"
              needs a pre-existing PKI OID root (i.e. AD CS deployed once);
            * derives the container DN from the TARGET forest's configuration NC and lets AD
              assign objectClass/objectCategory automatically;
            * lets you rename the template (internal cn and display name) via parameters;
            * sets template permissions (unless -SkipAcl), from a base chosen by -AclBase plus optional
              -EnrollPrincipals additions. -AclBase Standard (default) writes the standard Kerberos
              Authentication ACL, replacing the schema default so admins are not left with Full Control:
              Authenticated Users -> Read; Domain Admins and Enterprise Admins -> Read/Write/Enroll;
              Domain Controllers, Enterprise RODCs and Enterprise Domain Controllers -> Enroll/Autoenroll
              (Read comes via Authenticated Users); SYSTEM is not granted. Other -AclBase values keep or
              extend AD's schema-default ACL instead. This ACL is what a consumer such as EJBCA reads to
              decide who may enrol.
          Requires Enterprise Admin (or delegated write access to the Certificate Templates and
          OID containers in the Configuration partition).

      -Mode Sync
          Direct forest-to-forest copy in one run - no intermediate file. Reads the template from a
          DC in the SOURCE forest (-SourceServer, optionally -SourceCredential) and recreates it on
          the TARGET side exactly as -Mode Import would (-Server / discovered DC, optionally
          -Credential) - including -OidHandling, -NewTemplateName / -NewDisplayName, and the full
          ACL handling described under -Mode Import. Because it feeds the read attributes straight
          into the import pipeline, it also skips the JSON serialization step that -Mode Validate
          exists to check.

          Authentication: with a (two-way) trust between the forests, the identity running the
          script can typically read the source as-is (Authenticated Users has read access to
          templates) while holding Enterprise Admin rights in the target - then only -SourceServer
          is needed. Without a trust, or when running as neither identity, pass -SourceCredential
          and/or -Credential: explicit credentials against explicitly named servers need no trust
          at all.

      -Mode Validate
          Proves round-trip fidelity in a single forest, without touching a CA. Reads a source
          template, exports it to a (temp) JSON file, imports it under a throwaway name (with a
          unique throwaway OID that needs no OID root), reads the new object back, and diffs every
          attribute it copied - byte[] attributes included, since those are the JSON round-trip risk. By default
          the throwaway template, its companion OID object, and the temp file are removed afterwards
          (keep them for inspection with -KeepArtifacts). Requires the same write access as Import.

    No CA required:
      * Export, Import and Validate operate ONLY on the certificate TEMPLATE objects in AD's
        Configuration partition. No CA has to be installed, online, or reachable, and no AD CS
        role / RSAT "AD CS Tools" is needed (only the AD PowerShell module).
      * With the default -OidHandling Preserve, Import/Validate need NO PKI OID root, so the target
        forest can be one that never had AD CS. It only needs the (CA-independent) Certificate
        Templates container, part of every forest's Public Key Services structure; the script checks
        for it and fails clearly if the whole structure is somehow absent.
      * Preserve, GenerateFromRoot and GenerateRandom need no PKI OID root - only the (CA-independent)
        Certificate Templates and OID containers, present in every forest. Only -OidHandling Generate
        requires the forest's actual "CN=OID,..." base OID (present once AD CS has been deployed once);
        without it, Generate fails with a clear message pointing you to the other modes.
      * This moves a template DEFINITION - not an issued certificate, and no private keys. A CA is
        only involved later, when you publish the imported template on an issuing CA so it can
        enroll certificates from it.

    Identity / OID / placement handling:
      * OID: -OidHandling Preserve (default) reuses the source template's OID (so do NOT combine it
        with -StripOid on export); Generate / GenerateFromRoot / GenerateRandom instead mint a new OID
        (from the forest's real root, a supplied -OidRoot, or a synthesized base, respectively).
      * objectCategory and the DN suffix (everything after "CN=Certificate Templates,...") are
        derived from the TARGET forest automatically - you never edit them by hand.
      * The new internal name (cn) and display name come from -NewTemplateName / -NewDisplayName.
        If you did not strip identity on export, the source name/displayName in the file are used
        as fallbacks.

.PARAMETER Mode
    "Export", "Import", "Sync", or "Validate". The file-based and direct flows mix freely: a JSON
    exported earlier imports into a remote forest with -Mode Import -Server <target DC> (plus
    -Credential when needed), and -Mode Export equally accepts -Server/-Credential to read from a
    remote source forest.

.PARAMETER Path
    JSON file path. Written on Export, read on Import. Required for Export and Import; not used by
    Sync (no intermediate file is involved). On Validate it is optional: if given, the intermediate
    export is written there and kept for inspection; if omitted, a temporary file is used and
    deleted afterwards.

.PARAMETER TemplateName
    Export/Validate/Sync. The internal name (cn) of the source template to read. cn is unique
    within the templates container, so this matches exactly one template.
    Default: "KerberosAuthentication".

.PARAMETER StripIdentity
    Export only. Removes the source name and displayName from the JSON file, forcing you to supply
    -NewTemplateName / -NewDisplayName on import.

.PARAMETER StripOid
    Export only. Removes the source msPKI-Cert-Template-OID from the JSON file. Only safe with
    -OidHandling Generate on import; with the default -OidHandling Preserve the import needs that OID
    and will error if it was stripped.

.PARAMETER NewTemplateName
    Import only. New internal name (cn) for the template in the target forest (no spaces). Falls
    back to the file's stored name if omitted.

.PARAMETER NewDisplayName
    Import only. New display name for the template in the target forest. Falls back to the file's
    stored displayName if omitted.

.PARAMETER OidHandling
    Import only. How the template's OID is chosen. Every mode also registers a companion
    msPKI-Enterprise-Oid "display" object (when the OID container exists) so Windows resolves the OID
    to the template name.
      Preserve         (default) carry the source template's OID from the file. Needs no PKI OID root,
                       so it works in a forest that never had AD CS.
      Generate         mint a fresh OID under the target forest's REAL enterprise OID root. Requires
                       that AD CS was provisioned in the target forest at least once.
      GenerateFromRoot mint under the base OID you pass in -OidRoot (no AD CS needed). Use the same
                       root across imports to give those templates a shared, stable base.
      GenerateRandom   mint under a freshly synthesized, forest-independent base (no AD CS, no input).

.PARAMETER OidRoot
    Import only. Required with -OidHandling GenerateFromRoot: the base OID to generate the template OID
    under, e.g. "1.3.6.1.4.1.311.21.8.100000001.100000002.100000003.100000004.100000005". Ignored by
    the other modes.

.PARAMETER Server
    Optional domain controller to target for configuration-partition operations - on Sync this is
    the TARGET side (the source side is -SourceServer). If omitted on Import/Sync/Validate, a
    writable DC in the CURRENT forest is discovered and used consistently for the write and the
    follow-up ACL step; point it at a DC in another forest (with -Credential as needed) to operate
    there instead. Required whenever -Credential is given, so the credentials are guaranteed to be
    used against the forest you intend.

.PARAMETER Credential
    Optional credentials for the target-side operations (-Server): the config-partition reads and
    writes, the principal lookups, and the LDAP ACL bind. Works with every mode; combined with
    -Server it lets Export read from - or Import/Sync write to - a forest you are not logged on to,
    with no trust required. Requires -Server (see above).

.PARAMETER SourceServer
    Sync only (required there). A domain controller, or domain name, in the SOURCE forest to read
    the template from.

.PARAMETER SourceCredential
    Sync only. Optional credentials used against -SourceServer. Omit to read as the current
    identity (works across a trust, or when running inside the source forest itself).

.PARAMETER SkipAcl
    Import only. Skips the permission setup after import. Mutually exclusive with -EnrollPrincipals.

.PARAMETER AclBase
    Import only. The base the template ACL is built from; -EnrollPrincipals (if any) is always added on
    top. Default: Standard.
      Standard           the script's standard Kerberos Authentication set (see -Mode Import above),
                         REPLACING AD's schema-default ACL (so admins are not left with Full Control).
      Schema             leave AD's schema-default ACL as created (Domain/Enterprise Admins Full
                         Control, SYSTEM Full Control, Authenticated Users Read) and only add
                         -EnrollPrincipals to it.
      SchemaPlusStandard keep the schema-default ACL AND add the Standard set on top (nothing removed).
      PrincipalsOnly     no base - the ACL is exactly your -EnrollPrincipals (which is then required),
                         replacing the schema default.

.PARAMETER EnrollPrincipals
    Import only. Hashtable mapping each principal to the rights it should receive; these grants are
    ADDED on top of the -AclBase base (and are the sole content when -AclBase PrincipalsOnly). Keys: a
    SID (S-1-5-...), a sAMAccountName (optionally DOMAIN\-prefixed - the prefix must name the target
    domain), or a well-known token (DomainControllers, DomainComputers, DomainUsers, DomainAdmins,
    EnterpriseAdmins, EnterpriseRODCs, EnterpriseDomainControllers, AuthenticatedUsers, Everyone).
    Values: one or more of Read, Write, Enroll, Autoenroll, FullControl. Named principals are looked up
    in the target (-Server) domain; use a SID for a principal in another domain. Validated up front,
    before anything is created. Example:
        -AclBase PrincipalsOnly -EnrollPrincipals @{
            'DomainControllers'    = 'Enroll','Autoenroll'
            'AuthenticatedUsers'   = 'Read'
            'NOREFJELL\PKI-Admins' = 'FullControl'
        }

.PARAMETER KeepArtifacts
    Validate only. Leaves the throwaway template, its OID object, and the export file in place after
    the diff (default is to remove them).

.EXAMPLE
    # Source forest - export the built-in Kerberos Authentication template:
    .\Sync-KerberosAuthTemplate.ps1 -Mode Export -Path .\KerberosAuth.json

.EXAMPLE
    # Source forest - export a custom template as a clean, name-neutral copy:
    .\Sync-KerberosAuthTemplate.ps1 -Mode Export -TemplateName "XX-KerberosAuthentication" `
        -Path .\XX.json -StripIdentity -StripOid

.EXAMPLE
    # Target forest with NO AD CS (default) - import under a new name, carrying the source OID:
    .\Sync-KerberosAuthTemplate.ps1 -Mode Import -Path .\XX.json `
        -NewTemplateName "YY-KerberosAuthentication" -NewDisplayName "YY-Kerberos Authentication"

.EXAMPLE
    # Target forest that HAS its own PKI - mint a fresh target-forest OID instead of carrying it:
    .\Sync-KerberosAuthTemplate.ps1 -Mode Import -Path .\XX.json -OidHandling Generate `
        -NewTemplateName "YY-KerberosAuthentication" -NewDisplayName "YY-Kerberos Authentication"

.EXAMPLE
    # No AD CS in the target, but you want a fresh (synthetic) OID with a Windows-resolvable name:
    .\Sync-KerberosAuthTemplate.ps1 -Mode Import -Path .\XX.json -OidHandling GenerateRandom `
        -NewTemplateName "YY-KerberosAuthentication" -NewDisplayName "YY-Kerberos Authentication"

.EXAMPLE
    .\Sync-KerberosAuthTemplate.ps1 -Mode Import -Path .\XX.json -NewTemplateName "YY-KerberosAuthentication" -NewDisplayName "YY-Kerberos Authentication" -WhatIf

.EXAMPLE
    # Standard Kerberos Auth ACL (default) PLUS a template-admin group that EJBCA will read:
    .\Sync-KerberosAuthTemplate.ps1 -Mode Import -Path .\KerberosAuth.json -EnrollPrincipals @{
        'NOREFJELL\PKI-Admins' = 'FullControl'
    }

.EXAMPLE
    # Take exactly the ACL you specify (no standard set, no schema default):
    .\Sync-KerberosAuthTemplate.ps1 -Mode Import -Path .\KerberosAuth.json -AclBase PrincipalsOnly -EnrollPrincipals @{
        'DomainControllers'  = 'Enroll','Autoenroll'
        'AuthenticatedUsers' = 'Read'
    }

.EXAMPLE
    # Direct sync, no file - run in the TARGET forest and pull from the source forest over the trust:
    .\Sync-KerberosAuthTemplate.ps1 -Mode Sync -SourceServer dc01.source.example `
        -TemplateName "XX-KerberosAuthentication" `
        -NewTemplateName "YY-KerberosAuthentication" -NewDisplayName "YY-Kerberos Authentication"

.EXAMPLE
    # Direct sync from a third machine, explicit credentials on both sides (no trust needed):
    .\Sync-KerberosAuthTemplate.ps1 -Mode Sync `
        -SourceServer dc01.a.example -SourceCredential (Get-Credential A\template.reader) `
        -Server dc01.b.example -Credential (Get-Credential B\ent.admin)

.EXAMPLE
    # Mixed flow: import a previously exported JSON straight into another forest, no logon there:
    .\Sync-KerberosAuthTemplate.ps1 -Mode Import -Path .\KerberosAuth.json `
        -Server dc01.b.example -Credential (Get-Credential B\ent.admin)

.EXAMPLE
    # Prove export -> import preserves every functional attribute (creates and removes a throwaway copy):
    .\Sync-KerberosAuthTemplate.ps1 -Mode Validate -TemplateName "KerberosAuthentication"

.NOTES
    - Requires the ActiveDirectory PowerShell module (RSAT) for all modes. certutil and the
      AD CS role are no longer used or required, and no CA needs to be reachable.
    - With -OidHandling Preserve (default) the target forest need not have (or ever have had) AD CS;
      it only needs the Certificate Templates container. -OidHandling Generate needs the target
      forest's PKI OID root (present after AD CS has been deployed there once).
    - Run Export in the source forest and Import in the target forest (or point -Server, with
      -Credential as needed, at a DC in the relevant forest). -Mode Sync does both sides in one
      run and needs network reachability to a DC in EACH forest from where it runs.
    - Import/Sync do NOT publish the template to any CA; publish/issue it from the CA afterwards.
    - The security descriptor is intentionally not carried across forests; Import/Sync reapply
      standard permissions instead (unless -SkipAcl).
    - Authentication Mechanism Assurance (AMA) links are NOT carried over: an msDS-OIDToGroupLink
      on a source-forest issuance policy OID points at a group DN in THAT forest and cannot be
      copied. If you use AMA, recreate the link in the target forest manually (policy OID object ->
      a local universal group with no static members); until then, certificates from the synced
      template grant no AMA group membership there.
    - Only version 2+ templates can be meaningfully recreated (built-in v1 templates cannot).
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidateSet("Export", "Import", "Sync", "Validate")]
    [string]$Mode,

    [string]$Path,

    [string]$TemplateName = "KerberosAuthentication",

    [switch]$StripIdentity,

    [switch]$StripOid,

    [string]$NewTemplateName,

    [string]$NewDisplayName,

    [ValidateSet("Preserve", "Generate", "GenerateFromRoot", "GenerateRandom")]
    [string]$OidHandling = "Preserve",

    [string]$OidRoot,

    [string]$Server,

    [pscredential]$Credential,

    [string]$SourceServer,

    [pscredential]$SourceCredential,

    [switch]$SkipAcl,

    [ValidateSet("Standard", "Schema", "SchemaPlusStandard", "PrincipalsOnly")]
    [string]$AclBase = "Standard",

    [hashtable]$EnrollPrincipals,

    [switch]$KeepArtifacts
)

# Attribute -> type mapping used when rebuilding the template object in the target forest.
# (Mirrors the attribute set that MS PFE tooling copies for cross-forest template moves.)
$script:IntAttributes = @(
    'flags', 'revision',
    'msPKI-Certificate-Name-Flag', 'msPKI-Enrollment-Flag', 'msPKI-Minimal-Key-Size',
    'msPKI-Private-Key-Flag', 'msPKI-Template-Minor-Revision', 'msPKI-Template-Schema-Version',
    'msPKI-RA-Signature', 'pKIMaxIssuingDepth', 'pKIDefaultKeySpec'
)
$script:MultiValueAttributes = @(
    'msPKI-Certificate-Application-Policy', 'msPKI-RA-Application-Policies',
    'msPKI-Certificate-Policy', 'msPKI-RA-Policies',
    'msPKI-Supersede-Templates',
    'pKICriticalExtensions', 'pKIDefaultCSPs', 'pKIExtendedKeyUsage'
)
$script:ByteAttributes = @('pKIExpirationPeriod', 'pKIKeyUsage', 'pKIOverlapPeriod')

function Get-RandomHex {
    param([int]$Length)
    $hex = '0123456789ABCDEF'
    -join (1..$Length | ForEach-Object { $hex[(Get-Random -Minimum 0 -Maximum 16)] })
}

function Get-ConfigNC {
    param([hashtable]$ADParams)
    (Get-ADRootDSE @ADParams).configurationNamingContext
}

function Get-ADObjectIfPresent {
    # Existence check by DN. Get-ADObject -Identity throws ADIdentityNotFoundException when the object
    # is absent - even with -ErrorAction SilentlyContinue (the AD module treats an -Identity miss as
    # terminating) - so a plain -SilentlyContinue check surfaces/raises the error. This returns $null
    # on a miss and lets any other error (permissions, server unreachable, bad DN) propagate.
    param(
        [Parameter(Mandatory)][string]$Identity,
        [hashtable]$ADParams = @{},
        [string[]]$Properties
    )
    try {
        if ($Properties) {
            Get-ADObject @ADParams -Identity $Identity -Properties $Properties -ErrorAction Stop
        }
        else {
            Get-ADObject @ADParams -Identity $Identity -ErrorAction Stop
        }
    }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        $null
    }
}

function Test-UniqueTemplateOid {
    param([string]$OidObjectCn, [string]$TemplateOid, [string]$OidContainerDN, [hashtable]$ADParams)
    $match = Get-ADObject @ADParams -SearchBase $OidContainerDN `
        -LDAPFilter "(|(cn=$OidObjectCn)(msPKI-Cert-Template-OID=$TemplateOid))" -ErrorAction SilentlyContinue
    -not $match
}

function New-SyntheticOidBase {
    # A well-formed but forest-independent enterprise OID base: the Microsoft V2+ certificate-template
    # root (szOID_ENTERPRISE_OID_ROOT) plus 5 pseudo-forest arcs (mirrors the arc count AD CS derives
    # from the forest GUID). ~5x8 digits of entropy makes a collision with a real forest negligible.
    "1.3.6.1.4.1.311.21.8." + ((1..5 | ForEach-Object { Get-Random -Minimum 1000000 -Maximum 99999999 }) -join '.')
}

function New-TemplateOid {
    # Reserves a unique template OID (and the cn for its companion msPKI-Enterprise-Oid "display"
    # object) under a base OID. With -BaseOid the base is used as-is; otherwise it is read from the
    # TARGET forest's enterprise OID root (which requires AD CS to have been provisioned there).
    #   OID value : <base OID>.<8-digit>.<8-digit>
    #   OID cn    : <same 8-digit>.<32 hex chars>
    param([string]$ConfigNC, [hashtable]$ADParams, [string]$BaseOid)

    $oidContainerDN = "CN=OID,CN=Public Key Services,CN=Services,$ConfigNC"
    if (-not (Get-ADObjectIfPresent -Identity $oidContainerDN -ADParams $ADParams)) {
        throw "The OID container '$oidContainerDN' does not exist, so a generated OID cannot be registered. Use -OidHandling Preserve."
    }

    if ($BaseOid) {
        $forestBaseOid = $BaseOid
    }
    else {
        $forestBaseOid = (Get-ADObject @ADParams -Identity $oidContainerDN `
                -Properties 'msPKI-Cert-Template-OID').'msPKI-Cert-Template-OID'
        if (-not $forestBaseOid) {
            throw "Could not read the enterprise OID base from '$oidContainerDN' - this forest has no PKI OID root (AD CS was never provisioned here). Use -OidHandling Preserve, GenerateRandom, or GenerateFromRoot instead."
        }
    }

    do {
        $part1 = Get-Random -Minimum 10000000 -Maximum 99999999
        $part2 = Get-Random -Minimum 10000000 -Maximum 99999999
        $part3 = Get-RandomHex -Length 32
        $templateOid = "$forestBaseOid.$part1.$part2"
        $oidObjectCn = "$part2.$part3"
    } until (Test-UniqueTemplateOid -OidObjectCn $oidObjectCn -TemplateOid $templateOid -OidContainerDN $oidContainerDN -ADParams $ADParams)

    [pscustomobject]@{
        TemplateOid = $templateOid
        OidObjectCn = $oidObjectCn
        ContainerDN = $oidContainerDN
    }
}

function Resolve-OidDisplay {
    # For an already-decided OID (e.g. Preserve carrying the source OID), decide whether/where to
    # create a companion msPKI-Enterprise-Oid "display" object so Windows can resolve the OID to the
    # template name. Returns @{ CompanionCn; CompanionContainerDN } - both $null when none is needed
    # (no OID container in this forest, or a display object for that OID already exists).
    param([string]$TemplateOid, [string]$ConfigNC, [hashtable]$ADParams)

    $oidContainerDN = "CN=OID,CN=Public Key Services,CN=Services,$ConfigNC"
    if (-not (Get-ADObjectIfPresent -Identity $oidContainerDN -ADParams $ADParams)) {
        return @{ CompanionCn = $null; CompanionContainerDN = $null }
    }
    if (Get-ADObject @ADParams -SearchBase $oidContainerDN -LDAPFilter "(msPKI-Cert-Template-OID=$TemplateOid)" -ErrorAction SilentlyContinue) {
        return @{ CompanionCn = $null; CompanionContainerDN = $null }
    }
    do {
        $cn = "$(Get-Random -Minimum 10000000 -Maximum 99999999).$(Get-RandomHex -Length 32)"
    } until (-not (Get-ADObject @ADParams -SearchBase $oidContainerDN -LDAPFilter "(cn=$cn)" -ErrorAction SilentlyContinue))
    return @{ CompanionCn = $cn; CompanionContainerDN = $oidContainerDN }
}

function Resolve-TemplateOid {
    # Decides which OID the imported template carries, and (when applicable) the cn/container for its
    # companion msPKI-Enterprise-Oid "display" object. Returns @{ Oid; CompanionCn; CompanionContainerDN }.
    # A companion is created for every mode that produces a not-yet-registered OID, so Windows clients
    # can resolve the OID to the template name - except the transient -ExplicitOid (Validate) case.
    param(
        [string]$OidHandling,
        [string]$ExplicitOid,
        [string]$SourceOid,
        [string]$OidRoot,
        [string]$ConfigNC,
        [hashtable]$ADParams
    )

    if ($ExplicitOid) {
        return @{ Oid = $ExplicitOid; CompanionCn = $null; CompanionContainerDN = $null }
    }

    switch ($OidHandling) {
        'Generate' {
            # Mint a fresh OID under the TARGET forest's REAL enterprise OID root (needs AD CS to have
            # been provisioned there at least once) and register a companion display object for it.
            $oid = New-TemplateOid -ConfigNC $ConfigNC -ADParams $ADParams
            return @{ Oid = $oid.TemplateOid; CompanionCn = $oid.OidObjectCn; CompanionContainerDN = $oid.ContainerDN }
        }
        'GenerateFromRoot' {
            # Mint under a user-supplied base OID (the "root"). Needs no AD CS; use the same root for
            # every template you want to share a base. A companion display object is registered.
            if (-not $OidRoot) {
                throw "OidHandling 'GenerateFromRoot' requires -OidRoot (the base OID to generate under, e.g. 1.3.6.1.4.1.311.21.8.<5 arcs>)."
            }
            if ($OidRoot -notmatch '^(0|[1-9]\d*)(\.(0|[1-9]\d*))+$') {
                throw "-OidRoot '$OidRoot' is not a valid dotted OID (digits and dots only, no leading zeros in an arc)."
            }
            if ($OidRoot -notlike '1.3.6.1.4.1.311.21.8.*') {
                Write-Warning "-OidRoot '$OidRoot' is outside the conventional Microsoft template arc '1.3.6.1.4.1.311.21.8'; Windows matches template OIDs by exact value, so this generally works, but stays off the standard namespace."
            }
            $oid = New-TemplateOid -ConfigNC $ConfigNC -ADParams $ADParams -BaseOid $OidRoot
            return @{ Oid = $oid.TemplateOid; CompanionCn = $oid.OidObjectCn; CompanionContainerDN = $oid.ContainerDN }
        }
        'GenerateRandom' {
            # Mint under a freshly synthesized (forest-independent, "clearly fake") base OID. Needs no
            # AD CS and no user input. A companion display object is registered.
            $oid = New-TemplateOid -ConfigNC $ConfigNC -ADParams $ADParams -BaseOid (New-SyntheticOidBase)
            return @{ Oid = $oid.TemplateOid; CompanionCn = $oid.OidObjectCn; CompanionContainerDN = $oid.ContainerDN }
        }
        default {
            # Preserve: carry the source template's OID across. Needs no forest OID root, so this works
            # in a forest that never had AD CS (this is what certutil -dsAddTemplate did). A display
            # object is still registered (when possible) so Windows can resolve the carried OID.
            if (-not $SourceOid) {
                throw "OidHandling 'Preserve' needs the source template's OID, but msPKI-Cert-Template-OID is missing (a v1 template, or exported with -StripOid?). Re-export without -StripOid, or use -OidHandling GenerateFromRoot / GenerateRandom / Generate."
            }
            if ($SourceOid -notmatch '^(0|[1-9]\d*)(\.(0|[1-9]\d*))+$') {
                # Externally-supplied value: validate before it reaches any LDAP filter or gets written
                # as the template's identity (a tampered '*' would otherwise wildcard-match and be stored).
                throw "The source msPKI-Cert-Template-OID ('$SourceOid') is not a valid dotted OID - the export file (or source object) looks corrupted or tampered with."
            }
            $disp = Resolve-OidDisplay -TemplateOid $SourceOid -ConfigNC $ConfigNC -ADParams $ADParams
            return @{ Oid = $SourceOid; CompanionCn = $disp.CompanionCn; CompanionContainerDN = $disp.CompanionContainerDN }
        }
    }
}

function Get-SourceTemplate {
    # Reads the source template and returns the functional + identity attribute view that both the
    # file export (-Mode Export) and the direct forest-to-forest copy (-Mode Sync) work from.
    param(
        [string]$TemplateName,
        [hashtable]$ADParams
    )

    $configNC = Get-ConfigNC -ADParams $ADParams
    $templatesDN = "CN=Certificate Templates,CN=Public Key Services,CN=Services,$configNC"

    Write-Host "Reading template '$TemplateName' from $templatesDN ..." -ForegroundColor Cyan

    # Escape the name (a '*' would otherwise act as a wildcard and could silently export the wrong
    # template) and refuse a multi-match outright.
    $template = @(Get-ADObject @ADParams -SearchBase $templatesDN `
        -LDAPFilter "(&(objectClass=pKICertificateTemplate)(cn=$(ConvertTo-LdapFilterValue $TemplateName)))" -Properties *)
    if (-not $template.Count) {
        throw "Template with cn '$TemplateName' was not found under $templatesDN."
    }
    if ($template.Count -gt 1) {
        throw "The name '$TemplateName' matched $($template.Count) templates - refusing an ambiguous export."
    }
    $template = $template[0]

    # Select functional + identity attributes only. The DN, objectCategory and security descriptor
    # are intentionally excluded (they are forest-specific and reapplied/derived on import).
    $props = $template | Select-Object -Property name, displayName, objectClass, flags, revision, *pki*

    if ($props.'msPKI-Template-Schema-Version' -and [int]$props.'msPKI-Template-Schema-Version' -lt 2) {
        Write-Warning "Schema version 1 template: the object will round-trip, but v1 semantics are fixed in Windows (no editing, no autoenrollment) and v1 consumers match by NAME - import it under its original name, or duplicate it as v2+ in the source forest instead."
    }

    $props
}

function Export-Template {
    param(
        [string]$TemplateName,
        [string]$Path,
        [switch]$StripIdentity,
        [switch]$StripOid,
        [switch]$NoImportHint,   # internal (Validate): suppress the "copy to the target forest" hint
        [hashtable]$ADParams,
        [System.Management.Automation.PSCmdlet]$CallerCmdlet
    )

    $props = Get-SourceTemplate -TemplateName $TemplateName -ADParams $ADParams

    if ($StripOid) {
        $props.PSObject.Properties.Remove('msPKI-Cert-Template-OID')
    }
    if ($StripIdentity) {
        'name', 'displayName' | ForEach-Object { $props.PSObject.Properties.Remove($_) }
    }

    if (-not $CallerCmdlet.ShouldProcess($Path, "Write template export file")) {
        # -WhatIf, or the write was declined at the -Confirm prompt: no file is written, say so
        # honestly instead of printing a success message for a file that does not exist.
        Write-Host "Export file was NOT written (-WhatIf or declined): $Path" -ForegroundColor Yellow
        return
    }

    # Write via .NET with a UTF-8 BOM: PS7's Out-File -Encoding utf8 is BOM-less, which Windows
    # PowerShell 5.1 then misreads as ANSI (mojibake in non-ASCII names). A BOM is unambiguous for
    # both hosts. GetUnresolvedProviderPathFromPSPath resolves a PS-relative path for .NET.
    $fullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    $json = $props | ConvertTo-Json -Depth 10
    [System.IO.File]::WriteAllText($fullPath, $json, (New-Object System.Text.UTF8Encoding($true)))

    Write-Host "Export completed: $fullPath" -ForegroundColor Green
    if ($StripIdentity) {
        Write-Host "Identity stripped - you must supply -NewTemplateName and -NewDisplayName on import." -ForegroundColor Yellow
    }
    if (-not $NoImportHint) {
        Write-Host "Copy this file to the target forest and run the script there with -Mode Import." -ForegroundColor Yellow
    }
}

function Import-Template {
    param(
        [string]$Path,
        [psobject]$InputObject,   # in-memory alternative to -Path (used by -Mode Sync)
        [string]$NewTemplateName,
        [string]$NewDisplayName,
        [string]$OidHandling = 'Preserve',
        [string]$OidRoot,
        [string]$ExplicitOid,
        [hashtable]$ADParams,
        [System.Management.Automation.PSCmdlet]$CallerCmdlet
    )

    if ($InputObject) {
        # Direct sync: the attribute set was just read from the source forest, no file involved.
        $import = $InputObject
    }
    else {
        if (-not (Test-Path $Path)) {
            throw "File '$Path' was not found."
        }

        # Read as UTF-8 explicitly (BOM or BOM-less): Get-Content -Raw on Windows PowerShell 5.1 decodes a
        # BOM-less UTF-8 file (e.g. written by PS7) as ANSI, silently corrupting non-ASCII names.
        $fullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
        $import = [System.IO.File]::ReadAllText($fullPath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
    }

    # Resolve identity: explicit parameters win, else fall back to whatever the file carries.
    $cn = if ($NewTemplateName) { $NewTemplateName } elseif ($import.name) { $import.name } else { $null }
    $displayName = if ($NewDisplayName) { $NewDisplayName } elseif ($import.displayName) { $import.displayName } else { $null }

    if (-not $cn) {
        throw "No internal template name (cn) available. The file has no 'name' (stripped on export?) - supply -NewTemplateName."
    }
    if (-not $displayName) {
        throw "No display name available. The file has no 'displayName' (stripped on export?) - supply -NewDisplayName."
    }
    # Strict allowlist: the cn is interpolated into DNs, LDAP filters, and an ADSI ADsPath, where
    # characters like , + = " \ ; < > # / ( ) * or whitespace change meaning (DN injection, wildcard
    # matching, ADsPath separators). Real template cns are alphanumeric with . _ - so enforce that.
    if ($cn -notmatch '^[A-Za-z0-9._-]+$') {
        throw "The internal template name (cn) may only contain letters, digits, '.', '_' and '-': '$cn'. Put spaces and special characters in -NewDisplayName instead."
    }

    $configNC = Get-ConfigNC -ADParams $ADParams
    $templatesDN = "CN=Certificate Templates,CN=Public Key Services,CN=Services,$configNC"
    $newTemplateDN = "CN=$cn,$templatesDN"

    # Pre-flight: the Certificate Templates container must exist. It belongs to the forest's Public
    # Key Services structure and does NOT require a CA - but confirm it, so a forest missing that
    # structure fails with a clear message instead of a raw New-ADObject path error.
    if (-not (Get-ADObjectIfPresent -Identity $templatesDN -ADParams $ADParams)) {
        throw "The Certificate Templates container was not found at '$templatesDN'. This forest is missing the Public Key Services structure; this script does not provision it."
    }

    # Pre-flight: refuse to clobber an existing template of the same cn. (The allowlist above already
    # bars LDAP metacharacters; escaping is defense in depth.)
    $existing = Get-ADObject @ADParams -SearchBase $templatesDN -LDAPFilter "(cn=$(ConvertTo-LdapFilterValue $cn))" -ErrorAction SilentlyContinue
    if ($existing) {
        throw "A template with cn '$cn' already exists ($newTemplateDN). Choose a different -NewTemplateName or remove the existing template first."
    }

    # Build the functional attribute set (identity + OID handled separately below).
    $oa = @{}
    $unconsumed = @()
    foreach ($prop in ($import | Get-Member -MemberType NoteProperty)) {
        $name = $prop.Name
        if ($name -in $script:IntAttributes) {
            $oa[$name] = [System.Int32]$import.$name
        }
        elseif ($name -in $script:MultiValueAttributes) {
            $oa[$name] = [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection]$import.$name
        }
        elseif ($name -in $script:ByteAttributes) {
            $oa[$name] = [System.Byte[]]$import.$name
        }
        elseif ($name -match '^(msPKI-|pKI)' -and $name -ne 'msPKI-Cert-Template-OID' -and $null -ne $import.$name) {
            # A PKI attribute the type lists don't know. Never drop data silently - surface it.
            $unconsumed += $name
        }
        # identity/handled-elsewhere fields (name, displayName, objectClass, msPKI-Cert-Template-OID) are ignored
    }
    if ($unconsumed.Count) {
        Write-Warning "The source carries PKI attribute(s) this script does not know how to copy; they will NOT be written to the new template: $($unconsumed -join ', '). Add them to the attribute type lists if they matter."
    }

    # Resolve the OID (Preserve / Generate / explicit) BEFORE the ShouldProcess gate, so a missing
    # OID root or a missing source OID fails cleanly with no side effects. New-TemplateOid (Generate)
    # only reads/reserves here; the writes happen inside the gate below.
    $oidPlan  = Resolve-TemplateOid -OidHandling $OidHandling -ExplicitOid $ExplicitOid `
        -SourceOid $import.'msPKI-Cert-Template-OID' -OidRoot $OidRoot -ConfigNC $configNC -ADParams $ADParams
    $oidLabel = if ($ExplicitOid) { 'explicit' } else { $OidHandling }

    # Pre-flight: no existing template may already carry this OID. Windows and external CAs (EJBCA)
    # identify a v2+ template by its OID, so two templates sharing one OID make authorization lookups
    # ambiguous. Mainly bites Preserve (re-importing the same file under a new name); the Generate*
    # modes make a collision here statistically negligible but are checked all the same.
    $oidClash = Get-ADObject @ADParams -SearchBase $templatesDN `
        -LDAPFilter "(msPKI-Cert-Template-OID=$($oidPlan.Oid))" -ErrorAction SilentlyContinue
    if ($oidClash) {
        throw "A template already carries OID $($oidPlan.Oid): $($oidClash.DistinguishedName). Importing another template with the same OID would make OID-based template lookups (Windows, EJBCA) ambiguous. Use -OidHandling Generate, GenerateFromRoot, or GenerateRandom to mint a different OID."
    }

    $actionText = "Create certificate template via LDAP (OID handling: $oidLabel)"
    if ($oidPlan.CompanionCn) {
        $actionText += " and companion OID display object CN=$($oidPlan.CompanionCn),$($oidPlan.CompanionContainerDN)"
    }
    if (-not $CallerCmdlet.ShouldProcess($newTemplateDN, $actionText)) {
        return $newTemplateDN   # -WhatIf / declined: nothing created; DN returned for messaging only
    }

    # A companion msPKI-Enterprise-Oid "display" object is created whenever the resolved OID is not
    # yet registered in the OID container - all Generate* modes, and Preserve when the carried OID has
    # no display object yet. Only the transient -ExplicitOid (Validate) path never adds one.
    $companionDN = $null
    if ($oidPlan.CompanionCn) {
        $oidObjectAttrs = @{
            'DisplayName'             = $displayName
            'flags'                   = [System.Int32]1
            'msPKI-Cert-Template-OID' = $oidPlan.Oid
        }
        New-ADObject @ADParams -Path $oidPlan.CompanionContainerDN -Name $oidPlan.CompanionCn `
            -Type 'msPKI-Enterprise-Oid' -OtherAttributes $oidObjectAttrs -Confirm:$false
        $companionDN = "CN=$($oidPlan.CompanionCn),$($oidPlan.CompanionContainerDN)"
    }

    # Create the template object itself, referencing the resolved OID. -PassThru captures the DN that
    # AD actually assigned, so the ACL step and cleanup always bind the real object rather than a
    # string-built DN that could diverge under RDN escaping.
    $oa['msPKI-Cert-Template-OID'] = $oidPlan.Oid
    try {
        $createdObj = New-ADObject @ADParams -Path $templatesDN -Name $cn -DisplayName $displayName `
            -Type 'pKICertificateTemplate' -OtherAttributes $oa -Confirm:$false -PassThru
        $newTemplateDN = $createdObj.DistinguishedName
    }
    catch {
        # Best-effort rollback of a just-minted companion OID object (nothing to roll back otherwise).
        if ($companionDN) {
            try {
                Remove-ADObject @ADParams -Identity $companionDN -Confirm:$false -ErrorAction Stop
                Write-Warning "Template creation failed; rolled back the orphaned OID object."
            }
            catch {
                Write-Warning "Template creation failed AND the companion OID object could not be removed. Clean up manually: $companionDN"
            }
        }
        throw
    }

    Write-Host "Created template: $newTemplateDN" -ForegroundColor Green
    Write-Host " - Internal name (cn): $cn"
    Write-Host " - Display name:       $displayName"
    Write-Host " - Template OID:       $($oidPlan.Oid) ($oidLabel)"
    return $newTemplateDN
}

function ConvertTo-LdapFilterValue {
    # RFC 4515 escaping so a principal name cannot break the filter or inject a wildcard.
    param([string]$Value)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Value.ToCharArray()) {
        switch ($ch) {
            '\'  { [void]$sb.Append('\5c') }
            '*'  { [void]$sb.Append('\2a') }
            '('  { [void]$sb.Append('\28') }
            ')'  { [void]$sb.Append('\29') }
            "`0" { [void]$sb.Append('\00') }
            default { [void]$sb.Append($ch) }
        }
    }
    $sb.ToString()
}

function Resolve-PrincipalSid {
    # Resolves a principal (for -EnrollPrincipals) to a SID. Accepts a raw SID string, a language-
    # invariant well-known token, or a name (sAMAccountName / UPN / cn, optionally DOMAIN\-prefixed).
    # Named principals are looked up in the target (-Server) domain; use a SID for a principal that
    # lives in a different domain of a multi-domain forest.
    param(
        [string]$Identity,
        $TargetDomain,          # Get-ADDomain result (has .DomainSID)
        [string]$RootDomainSID, # SID string of the forest root domain
        [hashtable]$ADParams
    )

    $id = $Identity.Trim()

    if ($id -match '^S-1-\d+(-\d+)+$') {
        return New-Object System.Security.Principal.SecurityIdentifier($id)
    }

    $domSid = $TargetDomain.DomainSID.Value
    $norm   = ($id -replace '[\s_\-]', '').ToLowerInvariant()
    $wellKnown = @{
        'authenticatedusers'                  = 'S-1-5-11'
        'everyone'                            = 'S-1-1-0'
        'domaincontrollers'                   = "$domSid-516"
        'domaincomputers'                     = "$domSid-515"
        'domainusers'                         = "$domSid-513"
        'domainadmins'                        = "$domSid-512"
        'enterpriseadmins'                    = "$RootDomainSID-519"
        'enterprisereadonlydomaincontrollers' = "$RootDomainSID-498"
        'enterpriserodcs'                     = "$RootDomainSID-498"
        'enterprisedomaincontrollers'         = 'S-1-5-9'
    }
    if ($wellKnown.ContainsKey($norm)) {
        return New-Object System.Security.Principal.SecurityIdentifier($wellKnown[$norm])
    }

    # Directory lookup in the target (-Server) domain, matching sAMAccountName first (unique within a
    # domain), then UPN - both LDAP-escaped ('\' never goes into a filter). A DOMAIN\ prefix is
    # honoured only when it names the target domain itself; a mismatched prefix is rejected rather than
    # silently resolved against a same-named principal in the target domain (which would grant the
    # wrong object). For a principal in another domain, pass its SID.
    if ($id -match '\\') {
        $prefix = ($id -split '\\', 2)[0]
        $name   = ($id -split '\\', 2)[1]
        if ($prefix -and $TargetDomain.NetBIOSName -and $prefix -ne $TargetDomain.NetBIOSName) {
            throw "Principal '$Identity' names domain '$prefix', but named principals resolve only in the target domain '$($TargetDomain.NetBIOSName)'. Use a SID (S-1-5-...) for a principal in another domain."
        }
    }
    else {
        $name = $id
    }
    $escName = ConvertTo-LdapFilterValue $name

    $obj = @(Get-ADObject @ADParams -LDAPFilter "(sAMAccountName=$escName)" -Properties objectSid -ErrorAction SilentlyContinue |
            Where-Object { $_.objectSid }) | Select-Object -First 1

    if (-not $obj -and $id -match '@') {
        $escUpn = ConvertTo-LdapFilterValue $id
        $obj = @(Get-ADObject @ADParams -LDAPFilter "(userPrincipalName=$escUpn)" -Properties objectSid -ErrorAction SilentlyContinue |
                Where-Object { $_.objectSid }) | Select-Object -First 1
    }

    # Deliberately NO cn fallback: cn is not unique to security principals, and a domain-wide cn match
    # could resolve to an object anyone with create rights planted under that name - silently granting
    # template rights (which EJBCA enforces) to the wrong SID. Fail closed instead.
    if (-not $obj -or -not $obj.objectSid) {
        throw "Could not resolve principal '$Identity' by sAMAccountName or UPN in the target domain. Use its sAMAccountName, a SID (S-1-5-...), or a well-known token (DomainControllers, DomainComputers, DomainUsers, DomainAdmins, EnterpriseAdmins, EnterpriseRODCs, EnterpriseDomainControllers, AuthenticatedUsers, Everyone)."
    }
    return [System.Security.Principal.SecurityIdentifier]$obj.objectSid
}

function Resolve-TemplateGrants {
    # Resolves the grants to ADD to the DACL, UP FRONT (before anything is created). The 'Standard'
    # 6-entry Kerberos Authentication set is included when -AclBase is Standard or SchemaPlusStandard;
    # the user's -EnrollPrincipals (if any) are always added on top. Whether these replace or are added
    # to the schema-default ACL is decided later by Set-TemplateAcl (-ReplaceExisting). Each grant is
    # @{ Sid; Rights = @(lowercased keywords); Label }.
    param([string]$AclBase, [hashtable]$EnrollPrincipals, [hashtable]$ADParams)

    $targetDomain = Get-ADDomain @ADParams -ErrorAction Stop
    $targetForest = Get-ADForest @ADParams -ErrorAction Stop
    if ($targetDomain.DNSRoot -eq $targetForest.RootDomain) {
        $rootDomainSID = $targetDomain.DomainSID.Value
    }
    else {
        # Carry an explicit -Credential into the root-domain lookup too (this call replaces the
        # splat's -Server with the forest root, so it must re-add the credential itself).
        $rootParams = @{ Server = $targetForest.RootDomain; ErrorAction = 'Stop' }
        if ($ADParams.ContainsKey('Credential')) { $rootParams['Credential'] = $ADParams['Credential'] }
        $rootDomainSID = (Get-ADDomain @rootParams).DomainSID.Value
    }

    $validRights = @('read', 'write', 'enroll', 'autoenroll', 'fullcontrol')
    $grants = @()

    # The script's standard Kerberos Authentication set (matches the built-in template). Read is granted
    # via Authenticated Users (every DC computer account is a member), so the DC / RODC / Enterprise-DC
    # groups only need enroll + autoenroll; admins get Read/Write/Enroll (not Full Control / Autoenroll);
    # SYSTEM is deliberately not granted.
    if ($AclBase -in 'Standard', 'SchemaPlusStandard') {
        $domSid = $targetDomain.DomainSID.Value
        $grants += @{ Sid = (New-Object System.Security.Principal.SecurityIdentifier('S-1-5-11'));           Rights = @('read');                    Label = 'Authenticated Users' }
        $grants += @{ Sid = (New-Object System.Security.Principal.SecurityIdentifier("$rootDomainSID-498")); Rights = @('enroll', 'autoenroll');    Label = 'Enterprise Read-only Domain Controllers (RID 498)' }
        $grants += @{ Sid = (New-Object System.Security.Principal.SecurityIdentifier("$domSid-512"));        Rights = @('read', 'write', 'enroll'); Label = 'Domain Admins (RID 512)' }
        $grants += @{ Sid = (New-Object System.Security.Principal.SecurityIdentifier("$domSid-516"));        Rights = @('enroll', 'autoenroll');    Label = 'Domain Controllers (RID 516)' }
        $grants += @{ Sid = (New-Object System.Security.Principal.SecurityIdentifier("$rootDomainSID-519")); Rights = @('read', 'write', 'enroll'); Label = 'Enterprise Admins (RID 519)' }
        $grants += @{ Sid = (New-Object System.Security.Principal.SecurityIdentifier('S-1-5-9'));            Rights = @('enroll', 'autoenroll');    Label = 'Enterprise Domain Controllers (S-1-5-9)' }
    }

    if ($EnrollPrincipals -and $EnrollPrincipals.Count -gt 0) {
        foreach ($key in $EnrollPrincipals.Keys) {
            $rights = @(@($EnrollPrincipals[$key]) | ForEach-Object { "$_".Trim().ToLowerInvariant() } | Where-Object { $_ })
            if (-not $rights.Count) {
                throw "Principal '$key' was given no rights. Specify one or more of: Read, Write, Enroll, Autoenroll, FullControl."
            }
            $bad = @($rights | Where-Object { $_ -notin $validRights })
            if ($bad.Count) {
                throw "Unknown right(s) '$($bad -join ', ')' for principal '$key'. Valid rights: Read, Write, Enroll, Autoenroll, FullControl."
            }
            $sid = Resolve-PrincipalSid -Identity $key -TargetDomain $targetDomain -RootDomainSID $rootDomainSID -ADParams $ADParams
            $grants += @{ Sid = $sid; Rights = $rights; Label = $key }
        }
    }

    if ($AclBase -eq 'PrincipalsOnly' -and -not $grants.Count) {
        throw "-AclBase PrincipalsOnly requires -EnrollPrincipals; otherwise the template would be given an empty ACL."
    }

    return $grants
}

function Set-TemplateAcl {
    # Applies a pre-resolved grant list to the template's DACL. With -ReplaceExisting it first protects
    # the DACL from inheritance and removes the schema-default ACEs (so the result is exactly $Grants);
    # otherwise it ADDS $Grants on top of the schema-default ACL. The object owner (the account running
    # the import) can always re-permission the template, so no admin ACE is strictly required.
    param(
        [string]$TemplateDN,
        [array]$Grants,
        [bool]$ReplaceExisting,
        [hashtable]$ADParams,
        [System.Management.Automation.PSCmdlet]$CallerCmdlet
    )

    $exists = Get-ADObjectIfPresent -Identity $TemplateDN -ADParams $ADParams
    if (-not $exists) {
        # Reached only when the template was not created (-WhatIf, or the create was declined at the
        # -Confirm prompt). There is nothing to secure, so report and return rather than erroring.
        if ($WhatIfPreference) {
            if (-not $ReplaceExisting -and -not $Grants.Count) {
                Write-Host "What if: Would leave the schema-default ACL on $TemplateDN unchanged." -ForegroundColor Yellow
            }
            else {
                $verb = if ($ReplaceExisting) { 'set (replacing the schema default)' } else { 'add on top of the schema default' }
                Write-Host "What if: Would $verb these grants on ${TemplateDN}:" -ForegroundColor Yellow
                foreach ($g in $Grants) { Write-Host "         - $($g.Label): $($g.Rights -join ', ')" -ForegroundColor Yellow }
            }
        }
        else {
            Write-Host "Template '$TemplateDN' was not created; skipping permissions." -ForegroundColor Yellow
        }
        return
    }

    if (-not $ReplaceExisting -and -not $Grants.Count) {
        Write-Host "Left the schema-default ACL on '$TemplateDN' unchanged." -ForegroundColor Green
        return
    }

    $EnrollGUID     = [Guid]'0e10c968-78fb-11d2-90d4-00c04f79dc55'  # Certificate-Enrollment
    $AutoEnrollGUID = [Guid]'a05b8cc2-17bc-4802-a710-e7c15ab866a2'  # Certificate-AutoEnrollment
    $rightMap = @{
        'read'        = @{ Rights = [System.DirectoryServices.ActiveDirectoryRights]::GenericRead;    ObjectType = [Guid]::Empty; Label = 'Read' }
        'write'       = @{ Rights = [System.DirectoryServices.ActiveDirectoryRights]::GenericWrite;   ObjectType = [Guid]::Empty; Label = 'Write' }
        'enroll'      = @{ Rights = [System.DirectoryServices.ActiveDirectoryRights]::ExtendedRight;  ObjectType = $EnrollGUID;     Label = 'Enroll' }
        'autoenroll'  = @{ Rights = [System.DirectoryServices.ActiveDirectoryRights]::ExtendedRight;  ObjectType = $AutoEnrollGUID; Label = 'Autoenroll' }
        'fullcontrol' = @{ Rights = [System.DirectoryServices.ActiveDirectoryRights]::GenericAll;     ObjectType = [Guid]::Empty; Label = 'FullControl' }
    }

    $srv = if ($ADParams.ContainsKey('Server')) { $ADParams['Server'] } else { $null }
    $ldapPath = if ($srv) { "LDAP://$srv/$TemplateDN" } else { "LDAP://$TemplateDN" }

    $action = if ($ReplaceExisting) {
        "Set template ACL to $($Grants.Count) grant(s), replacing the schema default"
    }
    else {
        "Add $($Grants.Count) grant(s) to the template ACL, keeping the schema default"
    }
    if ($CallerCmdlet.ShouldProcess($TemplateDN, $action)) {
        # The DirectoryEntry work is the script's only LDAP-389 operation (everything else is ADWS
        # 9389), so a bind can fail where all prior calls succeeded. Statement-terminating errors do
        # NOT stop a function under the default ErrorActionPreference - without the try/catch a failed
        # bind would cascade through null $sec and still print the green success line.
        try {
            # An explicit -Credential must also reach this LDAP-389 bind (it does not flow through
            # the AD-cmdlet splat). The constructor defaults to secure (signed/sealed) binding.
            if ($ADParams.ContainsKey('Credential')) {
                $cred = $ADParams['Credential']
                $entry = New-Object System.DirectoryServices.DirectoryEntry(
                    $ldapPath, $cred.UserName, $cred.GetNetworkCredential().Password)
            }
            else {
                $entry = New-Object System.DirectoryServices.DirectoryEntry($ldapPath)
            }
            $sec = $entry.ObjectSecurity   # lazy bind fires here
            if (-not $sec) { throw "Bind returned no security descriptor." }

            if ($ReplaceExisting) {
                # Protect from inheritance and drop the inherited + schema-default ACEs (which otherwise
                # leave Domain/Enterprise Admins with Full Control and SYSTEM present) before adding grants.
                $sec.SetAccessRuleProtection($true, $false)
                foreach ($rule in @($sec.GetAccessRules($true, $false, [System.Security.Principal.SecurityIdentifier]))) {
                    $sec.RemoveAccessRuleSpecific($rule)
                }
            }

            foreach ($g in $Grants) {
                foreach ($r in $g.Rights) {
                    $m = $rightMap[$r]
                    if ($m.ObjectType -eq [Guid]::Empty) {
                        $ace = New-Object System.DirectoryServices.ActiveDirectoryAccessRule(
                            $g.Sid, $m.Rights, [System.Security.AccessControl.AccessControlType]::Allow)
                    }
                    else {
                        $ace = New-Object System.DirectoryServices.ActiveDirectoryAccessRule(
                            $g.Sid, $m.Rights, [System.Security.AccessControl.AccessControlType]::Allow, $m.ObjectType)
                    }
                    $sec.AddAccessRule($ace)
                }
            }

            $entry.ObjectSecurity = $sec
            $entry.CommitChanges()
        }
        catch {
            throw "Failed to apply the ACL on '$TemplateDN' (LDAP bind/commit via $ldapPath): $($_.Exception.Message). The template EXISTS but still carries the schema-default DACL - no Enroll rights are granted until this is fixed."
        }

        $how = if ($ReplaceExisting) { 'complete DACL, schema default replaced' } else { 'added on top of the schema-default ACL' }
        Write-Host "ACL set on '$TemplateDN' ($how):" -ForegroundColor Green
        foreach ($g in $Grants) {
            $pretty = ($g.Rights | ForEach-Object { $rightMap[$_].Label }) -join ', '
            Write-Host " - $($g.Label): $pretty"
        }
    }
    else {
        Write-Warning "ACL NOT applied (declined): '$TemplateDN' keeps the schema-default DACL (admins Full Control, NO Enroll rights). Apply permissions manually, or delete the template (and any companion OID object) and re-import."
    }
}

function Get-AttrCanonical {
    # Normalize an attribute value to a comparable string. Byte attributes are compared in order
    # (they encode fixed-layout values); everything else as an order-insensitive set of strings.
    param([string]$Name, $Value)
    if ($null -eq $Value) { return '<null>' }
    if ($Name -in $script:ByteAttributes) {
        return (([byte[]]$Value | ForEach-Object { $_.ToString('x2') }) -join '')
    }
    return ((@($Value) | ForEach-Object { "$_" } | Sort-Object) -join '|')
}

function Compare-TemplateAttributes {
    param($Source, $Target)
    # Diff the union of the import type lists AND every PKI attribute actually present on the SOURCE
    # object. Building the set only from the import lists would let an attribute the import cannot
    # write vanish from the diff entirely - a false PASS on a lossy round-trip. Identity and the OID
    # are excluded deliberately (renamed / regenerated by design).
    $sourcePkiAttrs = @($Source.PSObject.Properties.Name |
            Where-Object { $_ -match '^(msPKI-|pKI)' -and $_ -ne 'msPKI-Cert-Template-OID' })
    $attrs = @($script:IntAttributes + $script:MultiValueAttributes + $script:ByteAttributes + $sourcePkiAttrs) |
        Sort-Object -Unique
    foreach ($a in $attrs) {
        $sc = Get-AttrCanonical -Name $a -Value $Source.$a
        $tc = Get-AttrCanonical -Name $a -Value $Target.$a
        [pscustomobject]@{
            Attribute = $a
            Match     = ($sc -eq $tc)
            Source    = $sc
            Target    = $tc
        }
    }
}

function Invoke-RoundTripValidation {
    param(
        [string]$TemplateName,
        [string]$Path,
        [switch]$KeepArtifacts,
        [hashtable]$ADParams,
        [System.Management.Automation.PSCmdlet]$CallerCmdlet
    )

    $configNC       = Get-ConfigNC -ADParams $ADParams
    $templatesDN    = "CN=Certificate Templates,CN=Public Key Services,CN=Services,$configNC"
    $oidContainerDN = "CN=OID,CN=Public Key Services,CN=Services,$configNC"

    # Read the live source object (this is the left-hand side of the diff). Escaped + single-match
    # guarded, same as Export-Template.
    $source = @(Get-ADObject @ADParams -SearchBase $templatesDN `
        -LDAPFilter "(&(objectClass=pKICertificateTemplate)(cn=$(ConvertTo-LdapFilterValue $TemplateName)))" -Properties *)
    if (-not $source.Count) {
        throw "Source template cn '$TemplateName' was not found under $templatesDN."
    }
    if ($source.Count -gt 1) {
        throw "The name '$TemplateName' matched $($source.Count) templates - refusing an ambiguous validation."
    }
    $source = $source[0]

    if ($WhatIfPreference) {
        # Short-circuit: Validate creates AND removes a throwaway template. Under -WhatIf we make no
        # changes at all (and avoid the export->import round trip, whose Out-File would be suppressed).
        Write-Host "What if: would round-trip '$TemplateName' - export to a temp file, import a throwaway copy (with a throwaway OID), diff every functional attribute, then remove the throwaway template and file. Nothing is left changed." -ForegroundColor Yellow
        return
    }

    $usingTempFile = [string]::IsNullOrWhiteSpace($Path)
    if ($usingTempFile) {
        $Path = Join-Path $env:TEMP ("kerbtpl-roundtrip-" + (Get-RandomHex -Length 8) + ".json")
    }

    $suffix      = Get-RandomHex -Length 8
    $tempCn      = "RoundtripTest-$suffix"
    $tempDisplay = "Roundtrip Test $suffix"
    $tempDN      = $null

    try {
        # Exercise the real pipeline: export (serialize to JSON) -> import (deserialize + create).
        Export-Template -TemplateName $TemplateName -Path $Path -ADParams $ADParams -CallerCmdlet $CallerCmdlet -NoImportHint
        if (-not (Test-Path $Path)) {
            # Export write declined at a -Confirm prompt: abort gracefully instead of a raw
            # "file not found" from the import step.
            Write-Warning "Export file was not written (declined); validation aborted."
            return
        }

        # Give the throwaway a unique explicit OID so validation needs no forest OID root and cannot
        # collide with the source template's own OID. Derive it from the source OID when present; else
        # (e.g. a v1 template with no msPKI-Cert-Template-OID) synthesize a self-contained one so the
        # -ExplicitOid path is still taken (never $null, which would fall through to Preserve).
        $throwawayOid = if ($source.'msPKI-Cert-Template-OID') {
            "$($source.'msPKI-Cert-Template-OID').$(Get-Random -Minimum 1000000 -Maximum 99999999)"
        }
        else {
            "$(New-SyntheticOidBase).$(Get-Random -Minimum 10000000 -Maximum 99999999).$(Get-Random -Minimum 10000000 -Maximum 99999999)"
        }

        $tempDN = Import-Template -Path $Path -NewTemplateName $tempCn -NewDisplayName $tempDisplay `
            -ExplicitOid $throwawayOid -ADParams $ADParams -CallerCmdlet $CallerCmdlet

        $created = Get-ADObjectIfPresent -Identity $tempDN -ADParams $ADParams -Properties *
        if (-not $created) {
            # Not created - the create was declined at the -Confirm prompt (a genuine create failure
            # throws inside Import-Template; -WhatIf returned earlier). Nothing to diff; the finally
            # block still cleans up the temp file and any objects that exist.
            Write-Warning "Round-trip template was not created (declined at the -Confirm prompt?); nothing to compare."
            return
        }

        $diff       = Compare-TemplateAttributes -Source $source -Target $created
        $mismatches = @($diff | Where-Object { -not $_.Match })
        $byteRows   = @($diff | Where-Object { $_.Attribute -in $script:ByteAttributes })

        Write-Host ""
        Write-Host "Round-trip comparison: '$TemplateName' -> '$tempCn'" -ForegroundColor Cyan
        $diff | Format-Table -AutoSize Attribute, Match, Source, Target | Out-Host

        $byteOk = @($byteRows | Where-Object { -not $_.Match }).Count -eq 0
        Write-Host ("Byte[] attributes ({0}) - {1}" -f `
                (($byteRows | ForEach-Object { $_.Attribute }) -join ', '), `
                $(if ($byteOk) { 'all identical after JSON round-trip' } else { 'MISMATCH - see above' })) `
            -ForegroundColor $(if ($byteOk) { 'Green' } else { 'Red' })

        if ($mismatches.Count -eq 0) {
            Write-Host "PASS: all $($diff.Count) copied attributes are identical after round-trip." -ForegroundColor Green
        }
        else {
            Write-Host "FAIL: $($mismatches.Count) of $($diff.Count) attribute(s) differ after round-trip:" -ForegroundColor Red
            foreach ($m in $mismatches) {
                Write-Host ("   - {0}: source='{1}' target='{2}'" -f $m.Attribute, $m.Source, $m.Target) -ForegroundColor Red
            }
        }
    }
    finally {
        # Idempotent cleanup, safe on every exit path (early return, read-back failure after the
        # objects were created, or a mid-run error). Re-query by DN so it never relies on $created.
        $tpl = if ($tempDN) {
            Get-ADObjectIfPresent -Identity $tempDN -ADParams $ADParams -Properties 'msPKI-Cert-Template-OID'
        }
        else { $null }

        if ($KeepArtifacts) {
            if ($tpl) {
                Write-Host "-KeepArtifacts: left throwaway template '$tempDN' and file '$Path' in place." -ForegroundColor Yellow
            }
        }
        else {
            $cleanupOk = $true
            if ($tpl) {
                $oidVal = $tpl.'msPKI-Cert-Template-OID'
                try {
                    Remove-ADObject @ADParams -Identity $tempDN -Confirm:$false -ErrorAction Stop
                }
                catch {
                    $cleanupOk = $false
                    Write-Warning "Could not remove throwaway template '$tempDN': $($_.Exception.Message)"
                }
                if ($oidVal -and (Get-ADObjectIfPresent -Identity $oidContainerDN -ADParams $ADParams)) {
                    # Guard the -SearchBase: a missing CN=OID container would otherwise throw here
                    # (unsuppressed by SilentlyContinue) and abort the finally, leaking the temp file.
                    $oidObj = Get-ADObject @ADParams -SearchBase $oidContainerDN `
                        -LDAPFilter "(msPKI-Cert-Template-OID=$oidVal)" -ErrorAction SilentlyContinue
                    if ($oidObj) {
                        try {
                            Remove-ADObject @ADParams -Identity $oidObj.DistinguishedName -Confirm:$false -ErrorAction Stop
                        }
                        catch {
                            $cleanupOk = $false
                            Write-Warning "Could not remove throwaway OID object '$($oidObj.DistinguishedName)': $($_.Exception.Message)"
                        }
                    }
                }
            }
            if ($usingTempFile) {
                # -Confirm:$false: Remove-Item would otherwise raise its own prompt under a -Confirm run.
                Remove-Item -Path $Path -Force -Confirm:$false -ErrorAction SilentlyContinue
                if (Test-Path $Path) {
                    $cleanupOk = $false
                    Write-Warning "Could not remove the temp export file '$Path'."
                }
            }
            if ($cleanupOk) {
                Write-Host "Cleaned up round-trip artifacts." -ForegroundColor Green
            }
            else {
                Write-Warning "Round-trip cleanup was INCOMPLETE - see warnings above for what remains."
            }
        }
    }
}

# --- Main logic ---
Import-Module ActiveDirectory -ErrorAction Stop

if ($Mode -ne 'Sync' -and ($SourceServer -or $SourceCredential)) {
    throw "-SourceServer and -SourceCredential apply only to -Mode Sync. For the file-based flow, -Mode Export reads the source forest via -Server (and -Credential)."
}

$adParams = @{}
if ($Server) { $adParams['Server'] = $Server }
if ($Credential) { $adParams['Credential'] = $Credential }

# -Credential is meant for "operate on a forest I am not logged on to" - but DC discovery is
# DC-locator based and always finds the CURRENT forest. Requiring -Server alongside it guarantees
# the credentials are used against the forest the caller intends (a trust could otherwise let
# foreign credentials silently authenticate to - and write into - the local forest).
if ($Credential -and -not $Server) {
    throw "-Credential requires -Server: name the DC (in the forest those credentials belong to) explicitly, so the operation cannot land on a discovered DC in the current forest instead."
}

# Import, Sync and Validate write to the config partition and then read back / bind by DN. Pin a
# concrete DC so those operations hit the same server (avoids read-after-write races on serverless
# binding).
if ($Mode -in 'Import', 'Sync', 'Validate' -and -not $adParams.ContainsKey('Server')) {
    # -Writable: plain -Discover can return an RODC in an RODC-only site, and every config-partition
    # write would then fail after the reads succeeded.
    $dc = Get-ADDomainController -Discover -Writable -ErrorAction Stop
    $adParams['Server'] = ($dc.HostName | Select-Object -First 1)
}

switch ($Mode) {
    "Export" {
        if (-not $Path) { throw "-Path is required for -Mode Export." }
        Export-Template -TemplateName $TemplateName -Path $Path `
            -StripIdentity:$StripIdentity -StripOid:$StripOid -ADParams $adParams -CallerCmdlet $PSCmdlet
    }
    { $_ -in 'Import', 'Sync' } {
        if ($Mode -eq 'Import' -and -not $Path) { throw "-Path is required for -Mode Import." }
        if ($Mode -eq 'Sync') {
            if (-not $SourceServer) { throw "-SourceServer is required for -Mode Sync (a DC, or domain name, in the SOURCE forest to read the template from)." }
            if ($Path) { throw "-Path is not used by -Mode Sync (no intermediate file is involved). Use -Mode Export / -Mode Import for the file-based flow." }
            if ($StripIdentity -or $StripOid) { throw "-StripIdentity / -StripOid are Export-only. With -Mode Sync use -NewTemplateName / -NewDisplayName and -OidHandling directly." }
        }
        if ($SkipAcl -and $EnrollPrincipals -and $EnrollPrincipals.Count -gt 0) {
            throw "-SkipAcl and -EnrollPrincipals are mutually exclusive (-SkipAcl skips the very ACL that -EnrollPrincipals defines)."
        }

        # Validate keywords and resolve every principal to a SID BEFORE creating anything, so bad ACL
        # input aborts with nothing created (and -WhatIf still exercises the resolution).
        $grants = if (-not $SkipAcl) { @(Resolve-TemplateGrants -AclBase $AclBase -EnrollPrincipals $EnrollPrincipals -ADParams $adParams) } else { $null }

        $importArgs = @{
            NewTemplateName = $NewTemplateName
            NewDisplayName  = $NewDisplayName
            OidHandling     = $OidHandling
            OidRoot         = $OidRoot
            ADParams        = $adParams
            CallerCmdlet    = $PSCmdlet
        }
        if ($Mode -eq 'Sync') {
            # Direct forest-to-forest: read the template from the source forest and feed it straight
            # into the same import pipeline the file-based flow uses (no JSON round-trip at all).
            $sourceParams = @{ Server = $SourceServer }
            if ($SourceCredential) { $sourceParams['Credential'] = $SourceCredential }
            $importArgs['InputObject'] = Get-SourceTemplate -TemplateName $TemplateName -ADParams $sourceParams
        }
        else {
            $importArgs['Path'] = $Path
        }

        $templateDN = Import-Template @importArgs

        if (-not $SkipAcl) {
            # Standard/PrincipalsOnly replace the schema-default DACL; Schema/SchemaPlusStandard add to it.
            $replaceAcl = $AclBase -in 'Standard', 'PrincipalsOnly'
            try {
                Set-TemplateAcl -TemplateDN $templateDN -Grants $grants -ReplaceExisting $replaceAcl `
                    -ADParams $adParams -CallerCmdlet $PSCmdlet
            }
            catch {
                # The template was already created; make the half-state and the recovery path explicit
                # instead of dying with only the raw error (a re-run is blocked by the cn pre-flight).
                Write-Warning "The template was created at '$templateDN' but its ACL was NOT applied. Fix the cause and apply permissions manually, or delete the template (and any companion OID object created this run) and re-import."
                throw
            }
        }
        else {
            Write-Host "Skipped permission setup (-SkipAcl was specified)." -ForegroundColor Yellow
        }
    }
    "Validate" {
        Invoke-RoundTripValidation -TemplateName $TemplateName -Path $Path `
            -KeepArtifacts:$KeepArtifacts -ADParams $adParams -CallerCmdlet $PSCmdlet
    }
}
