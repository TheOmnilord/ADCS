#Requires -Modules ActiveDirectory

<#
.SYNOPSIS
    Copies a certificate template from one AD forest to another using direct attribute copy -
    either through a JSON file (Export/Import) or forest-to-forest in a single run (Sync).
    Optionally renames the template, either preserves or regenerates the template OID, and applies
    standard AD CS permissions after import. Works even against a target forest that has never had
    AD CS (Certificate Services) installed.

.DESCRIPTION
    This is a direct attribute-level template copy (no certutil), performed entirely over ADWS.
    It uses the ActiveDirectory PowerShell module for ALL modes, so the module is required for
    Export as well as Import/Sync.

      -Mode Export
          Run in the SOURCE forest. Reads the template's functional attributes (flags, revision,
          and all msPKI-*/pKI* attributes) via ADWS and writes them to a JSON file. Forest-specific
          data is deliberately NOT exported: the security descriptor, distinguishedName, and
          objectCategory are left out because they are derived/reapplied in the target forest.
          The source template OID and identity fields (name/displayName) can additionally be
          stripped from the file (-StripOid / -StripIdentity).

      -Mode Import
          Run in the TARGET forest. Recreates the template from the JSON file via ADWS:
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
          ACL handling described under -Mode Import. It feeds the read attributes straight into
          the import pipeline, so no JSON serialization happens at all; -Mode Validate proves the
          fidelity of this direct pipeline and of the file pipeline separately.

          Authentication: with a (two-way) trust between the forests, the identity running the
          script can typically read the source as-is (Authenticated Users has read access to
          templates) while holding Enterprise Admin rights in the target - then only -SourceServer
          is needed. Without a trust, or when running as neither identity, pass -SourceCredential
          and/or -Credential: explicit credentials against explicitly named servers need no trust
          at all.

      -Mode Validate
          Proves round-trip fidelity in a single forest, without touching a CA - for BOTH copy
          pipelines. It reads a source template and (1) exports it to a (temp) JSON file and
          imports that under a throwaway name, exercising the Export/Import file flow, then
          (2) feeds the live attribute view directly into the import under a second throwaway
          name, exercising exactly what -Mode Sync does (live AD values reach the import casts
          untouched by JSON, so the file check cannot stand in for it). Each throwaway gets a
          unique OID that needs no OID root; every copied attribute of each copy is diffed -
          byte[] attributes included. Any mismatch makes the run FAIL with a terminating error
          (non-zero exit code), so Validate can gate automation; cleanup still runs first. By
          default the throwaway templates and the temp file are removed afterwards (keep them for
          inspection with -KeepArtifacts). Requires the same write access as Import.

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
    Export only. Removes the source msPKI-Cert-Template-OID from the JSON file. Only safe when the
    import mints a new OID (-OidHandling Generate, GenerateFromRoot, or GenerateRandom); with the
    default -OidHandling Preserve the import needs that OID and will error if it was stripped.

.PARAMETER NewTemplateName
    Import/Sync. New internal name (cn) for the template in the target forest. Letters (non-ASCII
    included), digits, non-edge spaces and . _ - ( ) are allowed; characters that carry meaning in a
    DN (, + = " \ ; < >) and the LDAP wildcard (*) and / # are rejected. Parentheses are permitted
    and escaped where a name reaches an LDAP filter. Falls back to the source's name if omitted.

.PARAMETER NewDisplayName
    Import/Sync. New display name for the template in the target forest. Falls back to the
    source's displayName if omitted.

.PARAMETER OidHandling
    Import/Sync. How the template's OID is chosen. Every mode also registers a companion
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
    Import/Sync. Required with -OidHandling GenerateFromRoot: the base OID to generate the template OID
    under, e.g. "1.3.6.1.4.1.311.21.8.100000001.100000002.100000003.100000004.100000005". Supplying
    it with any other -OidHandling is rejected up front (it would otherwise be silently ignored).

.PARAMETER Server
    Optional domain controller to target for configuration-partition operations - on Sync this is
    the TARGET side (the source side is -SourceServer). If omitted on Import/Sync/Validate, a
    writable DC in the CURRENT forest is discovered and used consistently for the write and the
    follow-up ACL step; point it at a DC in another forest (with -Credential as needed) to operate
    there instead. Required whenever -Credential is given, so the credentials are guaranteed to be
    used against the forest you intend.

.PARAMETER Credential
    Optional credentials used against -Server - the TARGET side for Import/Sync/Validate, the
    SOURCE side for Export (which has no target side). Covers every operation the mode performs:
    config-partition reads and writes, principal lookups, and the ACL write (all over ADWS).
    Combined with -Server it lets Export read from - or Import/Sync write to - a forest you are
    not logged on to, with no trust required. Requires -Server (see above).

.PARAMETER SourceServer
    Sync only (required there). A domain controller, or domain name, in the SOURCE forest to read
    the template from.

.PARAMETER SourceCredential
    Sync only. Optional credentials used against -SourceServer. Omit to read as the current
    identity (works across a trust, or when running inside the source forest itself).

.PARAMETER SkipAcl
    Import/Sync. Skips the permission setup after import. Mutually exclusive with -EnrollPrincipals
    and with an explicit -AclBase.

.PARAMETER AclBase
    Import/Sync. The base the template ACL is built from; -EnrollPrincipals (if any) is always added on
    top. Default: Standard. NOTE: the default is the KERBEROS AUTHENTICATION set (DC-oriented) - when
    it applies only by default to a template whose name does not look like a Kerberos Authentication
    copy, the script warns so the DC-oriented grants are a conscious choice, not an accident.
      Standard           the script's standard Kerberos Authentication set (see -Mode Import above),
                         REPLACING AD's schema-default ACL (so admins are not left with Full Control).
      Schema             leave AD's schema-default ACL as created (Domain/Enterprise Admins Full
                         Control, SYSTEM Full Control, Authenticated Users Read) and only add
                         -EnrollPrincipals to it.
      SchemaPlusStandard keep the schema-default ACL AND add the Standard set on top (nothing removed).
      PrincipalsOnly     no base - the ACL is exactly your -EnrollPrincipals (which is then required),
                         replacing the schema default.

.PARAMETER EnrollPrincipals
    Import/Sync. Hashtable mapping each principal to the rights it should receive; these grants are
    ADDED on top of the -AclBase base (and are the sole content when -AclBase PrincipalsOnly). Keys: a
    SID (S-1-5-...), a sAMAccountName or UPN (user@domain), optionally DOMAIN\-prefixed - the prefix
    must name the target domain, or a well-known token (DomainControllers, DomainComputers,
    DomainUsers, DomainAdmins, EnterpriseAdmins, EnterpriseRODCs, EnterpriseDomainControllers,
    AuthenticatedUsers, Everyone). Resolution is fail-closed: a bare string that matches BOTH a
    well-known token AND a directory object with a DIFFERENT SID is refused (a planted account
    cannot hijack a token, and a token cannot shadow a distinct real group - disambiguate with a
    SID or a DOMAIN\ prefix); the built-in groups' own names (e.g. 'Domain Admins') resolve
    normally, since both readings yield the same SID. A name matching more than one object
    (duplicate UPNs) is refused rather than guessed. Values: one
    or more of Read, Write, Enroll, Autoenroll, FullControl. Named principals are looked up in the
    target (-Server) domain; use a SID for a principal in another domain. Validated up front, before
    anything is created. Example:
        -AclBase PrincipalsOnly -EnrollPrincipals @{
            'DomainControllers'    = 'Enroll','Autoenroll'
            'AuthenticatedUsers'   = 'Read'
            'NOREFJELL\PKI-Admins' = 'FullControl'
        }

.PARAMETER UpgradeCompatibility
    Import/Sync. Raises the imported template to the newest compatibility the Certificate Templates
    MMC offers - Certification Authority: Windows Server 2016, Certificate recipient:
    Windows 10 / Windows Server 2016 - as it is created in the target forest (schema version 4 plus
    the matching private-key-flag bits). Only schema v2/v3 templates can be upgraded in place; a
    schema v1 template is imported unchanged with a warning (v1 built-ins are read-only in the MMC),
    and a template already at v4 is left as-is. The source template/export is not modified; only the
    copy written to the target is upgraded.

.PARAMETER KeepArtifacts
    Validate only. Leaves the throwaway templates and the export file in place after the diff
    (default is to remove them). No companion OID object is ever created for the throwaways (they use
    a self-contained explicit OID), so there is none to keep.

.EXAMPLE
    # Source forest - export the built-in Kerberos Authentication template:
    .\Sync-ADCSTemplate.ps1 -Mode Export -Path .\KerberosAuth.json

.EXAMPLE
    # Source forest - export a custom template as a name-neutral copy (identity stripped; the OID
    # stays in the file so the default -OidHandling Preserve works on import - add -StripOid only
    # when the import will mint a new OID with one of the Generate modes):
    .\Sync-ADCSTemplate.ps1 -Mode Export -TemplateName "XX-KerberosAuthentication" `
        -Path .\XX.json -StripIdentity

.EXAMPLE
    # Target forest with NO AD CS (default) - import under a new name, carrying the source OID:
    .\Sync-ADCSTemplate.ps1 -Mode Import -Path .\XX.json `
        -NewTemplateName "YY-KerberosAuthentication" -NewDisplayName "YY-Kerberos Authentication"

.EXAMPLE
    # Target forest that HAS its own PKI - mint a fresh target-forest OID instead of carrying it:
    .\Sync-ADCSTemplate.ps1 -Mode Import -Path .\XX.json -OidHandling Generate `
        -NewTemplateName "YY-KerberosAuthentication" -NewDisplayName "YY-Kerberos Authentication"

.EXAMPLE
    # No AD CS in the target, but you want a fresh (synthetic) OID with a Windows-resolvable name:
    .\Sync-ADCSTemplate.ps1 -Mode Import -Path .\XX.json -OidHandling GenerateRandom `
        -NewTemplateName "YY-KerberosAuthentication" -NewDisplayName "YY-Kerberos Authentication"

.EXAMPLE
    # Import and raise the copy to the latest compatibility (Windows Server 2016 / Windows 10):
    .\Sync-ADCSTemplate.ps1 -Mode Import -Path .\Workstation.json -OidHandling GenerateRandom -UpgradeCompatibility

.EXAMPLE
    .\Sync-ADCSTemplate.ps1 -Mode Import -Path .\XX.json -NewTemplateName "YY-KerberosAuthentication" -NewDisplayName "YY-Kerberos Authentication" -WhatIf

.EXAMPLE
    # Standard Kerberos Auth ACL (default) PLUS a template-admin group that EJBCA will read:
    .\Sync-ADCSTemplate.ps1 -Mode Import -Path .\KerberosAuth.json -EnrollPrincipals @{
        'NOREFJELL\PKI-Admins' = 'FullControl'
    }

.EXAMPLE
    # Take exactly the ACL you specify (no standard set, no schema default):
    .\Sync-ADCSTemplate.ps1 -Mode Import -Path .\KerberosAuth.json -AclBase PrincipalsOnly -EnrollPrincipals @{
        'DomainControllers'  = 'Enroll','Autoenroll'
        'AuthenticatedUsers' = 'Read'
    }

.EXAMPLE
    # Direct sync, no file - run in the TARGET forest and pull from the source forest over the trust:
    .\Sync-ADCSTemplate.ps1 -Mode Sync -SourceServer dc01.source.example `
        -TemplateName "XX-KerberosAuthentication" `
        -NewTemplateName "YY-KerberosAuthentication" -NewDisplayName "YY-Kerberos Authentication"

.EXAMPLE
    # Direct sync from a third machine, explicit credentials on both sides (no trust needed):
    .\Sync-ADCSTemplate.ps1 -Mode Sync `
        -SourceServer dc01.a.example -SourceCredential (Get-Credential A\template.reader) `
        -Server dc01.b.example -Credential (Get-Credential B\ent.admin)

.EXAMPLE
    # Mixed flow: import a previously exported JSON straight into another forest, no logon there:
    .\Sync-ADCSTemplate.ps1 -Mode Import -Path .\KerberosAuth.json `
        -Server dc01.b.example -Credential (Get-Credential B\ent.admin)

.EXAMPLE
    # Prove both copy pipelines (file and direct/Sync) preserve every functional attribute
    # (creates and removes one throwaway copy per pipeline):
    .\Sync-ADCSTemplate.ps1 -Mode Validate -TemplateName "KerberosAuthentication"

.NOTES
    - Requires the ActiveDirectory PowerShell module (RSAT) for all modes. certutil and the
      AD CS role are no longer used or required, and no CA needs to be reachable.
    - ALL directory access - including the ACL write - runs over ADWS (TCP 9389); no LDAP (389)
      connectivity is needed. -Server (and -SourceServer) values are always used VERBATIM - the
      script never substitutes an endpoint the operator did not type. Name ONE DC: a DOMAIN name
      (DNS or NetBIOS) locates a different DC per connection, so the create, read-back and ACL
      steps could hit different replicas - the script detects that case and warns loudly (and a
      lagging read-back fails the run rather than leaving a template mis-secured).
    - Parameters a mode does not consume are rejected up front (e.g. -StripOid with -Mode Import,
      -OidRoot without -OidHandling GenerateFromRoot, -AclBase with -SkipAcl) rather than silently
      ignored.
    - PKI attributes the built-in type lists don't know (a genuine schema extension linked to the
      pKICertificateTemplate class) are typed automatically from the TARGET forest's schema; an
      attribute is dropped with a warning when the target schema cannot type it, does not permit it
      on the template class, or types it single-valued while the source value is empty or
      multi-valued (schema divergence between the forests). (Note: v3/v4 CNG algorithm settings are
      NOT separate directory attributes - they travel packed inside msPKI-RA-Application-Policies,
      which is copied as-is.)
    - v3/v4 templates can embed a private-key SDDL (msPKI-Key-Security-Descriptor) packed inside
      msPKI-RA-Application-Policies; it is copied verbatim and any domain SIDs in it are
      SOURCE-forest SIDs - the script warns so the key ACL can be reviewed in the target forest.
    - -Mode Validate exits non-zero when any attribute differs, so it can gate CI/automation.
    - With -OidHandling Preserve (default) the target forest need not have (or ever have had) AD CS;
      it only needs the Certificate Templates container. -OidHandling Generate needs the target
      forest's PKI OID root (present after AD CS has been deployed there once).
    - Run Export in the source forest and Import in the target forest (or point -Server, with
      -Credential as needed, at a DC in the relevant forest). -Mode Sync does both sides in one
      run and needs ADWS reachability to a DC in EACH forest from where it runs; it refuses to
      proceed when source and target resolve to the same forest unless -Server was given
      explicitly (guarding against a forgotten -Server silently targeting the source forest).
    - Import/Sync do NOT publish the template to any CA; publish/issue it from the CA afterwards.
    - The security descriptor is intentionally not carried across forests; Import/Sync reapply
      standard permissions instead (unless -SkipAcl).
    - Authentication Mechanism Assurance (AMA) links are NOT carried over: an msDS-OIDToGroupLink
      on a source-forest issuance policy OID points at a group DN in THAT forest and cannot be
      copied. If you use AMA, recreate the link in the target forest manually (policy OID object ->
      a local universal group with no static members); until then, certificates from the synced
      template grant no AMA group membership there.
    - v1 templates copy too (the export warns): the object round-trips faithfully, but Windows
      fixes v1 semantics in code - v1 consumers match by NAME, and the definition is not editable
      and never autoenrolls. Import a v1 template under its ORIGINAL name (a renamed copy is
      invisible to Windows v1 consumers); for non-Windows consumers such as EJBCA, which read the
      object and its ACL directly, a v1 copy works like any other. To get editable/autoenroll
      behavior, duplicate it as v2+ in the source forest first.
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

    [switch]$UpgradeCompatibility,

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

# Per-run caches: schema-derived types for attributes the static lists don't know, and the target
# forest-root domain SID (resolved lazily, only when a grant actually needs it).
$script:SchemaTypeCache     = @{}
$script:TemplateAllowedCache = $null
$script:RootDomainSidCache  = $null

function Convert-ToLatestCompatibility {
    # Raises a template's compatibility to the newest setting the Certificate Templates MMC offers -
    # Certification Authority: Windows Server 2016, Certificate recipient: Windows 10 / Windows
    # Server 2016 - by mutating the New-ADObject -OtherAttributes hashtable IN PLACE. Only schema
    # v2/v3 templates are upgraded: v1 built-ins are read-only in the MMC (not upgradable in place)
    # and templates already at v4 are the newest, so both are left untouched. Returns a report.
    #
    # Encoding (verified against real MMC-made v4 templates and a live-DC round-trip):
    #   msPKI-Template-Schema-Version -> 4
    #   msPKI-Private-Key-Flag        |= 0x06060000 (CA=Server2016 nibble | recipient=Win10/2016 nibble),
    #                                    plus 0x100 (CT_FLAG_USE_LEGACY_PROVIDER) for CSP-based
    #                                    templates only - NOT CNG/KSP templates (no pKIDefaultCSPs)
    #   flags                          IS_DEFAULT (0x10000) -> IS_MODIFIED (0x20000)
    #   msPKI-Template-Minor-Revision  += 1
    param([Parameter(Mandatory)][hashtable]$Attributes)

    $ver = if ($Attributes.ContainsKey('msPKI-Template-Schema-Version')) { [int]$Attributes['msPKI-Template-Schema-Version'] } else { 1 }
    if ($ver -lt 2) { return [pscustomobject]@{ Upgraded = $false; FromVersion = $ver; Reason = 'schema v1 template - not upgradable in place' } }
    if ($ver -ge 4) { return [pscustomobject]@{ Upgraded = $false; FromVersion = $ver; Reason = 'already at the latest compatibility (schema v4)' } }

    $pkf = if ($Attributes.ContainsKey('msPKI-Private-Key-Flag')) { [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$Attributes['msPKI-Private-Key-Flag']), 0) } else { [uint32]0 }
    $pkf = $pkf -bor 0x06060000
    # @($null).Count is 1, so test the value explicitly - do NOT rely on the count alone.
    $isCsp = $Attributes.ContainsKey('pKIDefaultCSPs') -and $null -ne $Attributes['pKIDefaultCSPs'] -and @($Attributes['pKIDefaultCSPs']).Count -gt 0
    if ($isCsp) { $pkf = $pkf -bor 0x100 }

    $flags = if ($Attributes.ContainsKey('flags')) { [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$Attributes['flags']), 0) } else { [uint32]0 }
    $flags = ($flags -band (-bnot 0x10000)) -bor 0x20000

    $Attributes['msPKI-Template-Schema-Version'] = [System.Int32]4
    $Attributes['msPKI-Private-Key-Flag']        = [BitConverter]::ToInt32([BitConverter]::GetBytes([uint32]$pkf), 0)
    $Attributes['flags']                         = [BitConverter]::ToInt32([BitConverter]::GetBytes([uint32]$flags), 0)
    if ($Attributes.ContainsKey('msPKI-Template-Minor-Revision')) {
        $Attributes['msPKI-Template-Minor-Revision'] = [System.Int32]([int]$Attributes['msPKI-Template-Minor-Revision'] + 1)
    }
    return [pscustomobject]@{ Upgraded = $true; FromVersion = $ver; PrivateKeyFlag = ('0x{0:X8}' -f $pkf); LegacyProvider = $isCsp }
}

function Get-RandomHex {
    param([int]$Length)
    $hex = '0123456789ABCDEF'
    -join (1..$Length | ForEach-Object { $hex[(Get-Random -Minimum 0 -Maximum 16)] })
}

function ConvertTo-LdapFilterValue {
    # RFC 4515 escaping so a value cannot break an LDAP filter or inject a wildcard. Defined before
    # every function that calls it.
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

function Get-ConfigNC {
    # An unreachable/typo'd server makes Get-ADRootDSE fail with a NON-terminating error under the
    # default ErrorActionPreference, returning $null and sending a malformed (empty-suffix) DN into
    # every later search - which then surfaces as a misleading "template not found" / "forest is
    # missing Public Key Services" error. Fail here, loudly, naming the server.
    param([hashtable]$ADParams)
    try {
        (Get-ADRootDSE @ADParams -ErrorAction Stop).configurationNamingContext
    }
    catch {
        $srv = if ($ADParams.ContainsKey('Server')) { "'$($ADParams['Server'])'" } else { 'the default domain controller' }
        throw "Could not read RootDSE from $srv - is the server name correct and reachable (ADWS, TCP 9389)? Underlying error: $($_.Exception.Message)"
    }
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
    # -ErrorAction Stop: "unique" may only mean "a successful search found nothing". A swallowed
    # query failure here would return $true unverified and let a colliding OID through.
    param([string]$OidObjectCn, [string]$TemplateOid, [string]$OidContainerDN, [hashtable]$ADParams)
    $match = Get-ADObject @ADParams -SearchBase $OidContainerDN `
        -LDAPFilter "(|(cn=$OidObjectCn)(msPKI-Cert-Template-OID=$TemplateOid))" -ErrorAction Stop
    -not $match
}

function New-SyntheticOidBase {
    # A well-formed but forest-independent enterprise OID base: the Microsoft V2+ certificate-template
    # root (szOID_ENTERPRISE_OID_ROOT) plus 5 pseudo-random arcs of 7-8 digits each. (Real AD CS bases
    # derive 6 GUID-based arcs of varying width, so this shape is recognizably synthetic; the ~35+
    # digits of entropy make a collision with a real forest base negligible.)
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
    # -ErrorAction Stop on both probes: a swallowed query failure would be indistinguishable from
    # "no display object yet" and produce a DUPLICATE companion (or an unverified cn).
    if (Get-ADObject @ADParams -SearchBase $oidContainerDN -LDAPFilter "(msPKI-Cert-Template-OID=$TemplateOid)" -ErrorAction Stop) {
        return @{ CompanionCn = $null; CompanionContainerDN = $null }
    }
    do {
        $cn = "$(Get-Random -Minimum 10000000 -Maximum 99999999).$(Get-RandomHex -Length 32)"
    } until (-not (Get-ADObject @ADParams -SearchBase $oidContainerDN -LDAPFilter "(cn=$cn)" -ErrorAction Stop))
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
        # Internal (Validate) path today, but validate all the same: this value reaches an LDAP
        # filter and becomes the new template's stored identity.
        if ($ExplicitOid -notmatch '^(0|[1-9]\d*)(\.(0|[1-9]\d*))+$') {
            throw "The explicit OID '$ExplicitOid' is not a valid dotted OID."
        }
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
            if ($OidRoot -eq '1.3.6.1.4.1.311.21.8') {
                Write-Warning "-OidRoot is the bare shared Microsoft template arc: template OIDs will be minted directly under it with only two random arcs, which risks prefix-colliding with real forests' GUID-derived bases. Prefer a deeper base of your own beneath this arc."
            }
            elseif ($OidRoot -notlike '1.3.6.1.4.1.311.21.8.*') {
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
    # -ConfigNC skips the RootDSE read when the caller already has it.
    param(
        [string]$TemplateName,
        [hashtable]$ADParams,
        [string]$ConfigNC
    )

    $configNC = if ($ConfigNC) { $ConfigNC } else { Get-ConfigNC -ADParams $ADParams }
    $templatesDN = "CN=Certificate Templates,CN=Public Key Services,CN=Services,$configNC"

    Write-Host "Reading template '$TemplateName' from $templatesDN ..." -ForegroundColor Cyan

    # Escape the name (a '*' would otherwise act as a wildcard and could silently export the wrong
    # template) and refuse a multi-match outright. -ErrorAction Stop: a connectivity failure here
    # must terminate as itself, not fall through to a misleading "not found".
    $template = @(Get-ADObject @ADParams -SearchBase $templatesDN -ErrorAction Stop `
        -LDAPFilter "(&(objectClass=pKICertificateTemplate)(cn=$(ConvertTo-LdapFilterValue $TemplateName)))" -Properties *)
    if (-not $template.Count) {
        throw "Template with cn '$TemplateName' was not found under $templatesDN."
    }
    if ($template.Count -gt 1) {
        throw "The name '$TemplateName' matched $($template.Count) templates - refusing an ambiguous export."
    }
    $template = $template[0]

    # Select functional + identity attributes only. The DN, objectCategory and security descriptors
    # are intentionally excluded (they are forest-specific and reapplied/derived on import) -
    # pKIEnrollmentAccess is an ACL-bearing attribute the *pki* wildcard would otherwise catch.
    $props = $template | Select-Object -Property name, displayName, objectClass, flags, revision, *pki* `
        -ExcludeProperty pKIEnrollmentAccess

    if ($props.'msPKI-Template-Schema-Version' -and [int]$props.'msPKI-Template-Schema-Version' -lt 2) {
        Write-Warning "Schema version 1 template: the object will round-trip, but v1 semantics are fixed in Windows (no editing, no autoenrollment) and v1 consumers match by NAME - import it under its original name, or duplicate it as v2+ in the source forest instead."
    }

    $props
}

function Export-Template {
    param(
        [string]$TemplateName,
        [psobject]$InputObject,  # internal (Validate): a pre-read source view; skips the directory read.
                                 # Callers passing it must not also use Strip* (which would mutate it).
        [string]$Path,
        [switch]$StripIdentity,
        [switch]$StripOid,
        [switch]$NoImportHint,   # internal (Validate): suppress the "copy to the target forest" hint
        [hashtable]$ADParams,
        [System.Management.Automation.PSCmdlet]$CallerCmdlet
    )

    $props = if ($InputObject) { $InputObject } else { Get-SourceTemplate -TemplateName $TemplateName -ADParams $ADParams }

    if ($StripOid) {
        $props.PSObject.Properties.Remove('msPKI-Cert-Template-OID')
    }
    if ($StripIdentity) {
        'name', 'displayName' | ForEach-Object { $props.PSObject.Properties.Remove($_) }
    }

    # An existing file at -Path is replaced: say so in the ShouldProcess action and warn, so a reused
    # path cannot silently destroy an earlier export (possibly the only copy in an air-gapped flow).
    $fileExists = Test-Path -LiteralPath $Path
    $action = if ($fileExists) { "OVERWRITE the existing template export file" } else { "Write template export file" }
    if (-not $CallerCmdlet.ShouldProcess($Path, $action)) {
        # -WhatIf, or the write was declined at the -Confirm prompt: no file is written, say so
        # honestly instead of printing a success message for a file that does not exist. Return
        # $false so a caller (Validate) can tell "not written" from "written" without inspecting the
        # filesystem (a declined OVERWRITE leaves the STALE file in place - Test-Path would lie).
        Write-Host "Export file was NOT written (-WhatIf or declined): $Path" -ForegroundColor Yellow
        return $false
    }
    if ($fileExists) {
        Write-Warning "Overwriting existing file: $Path"
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
    return $true
}

function Read-TemplateExport {
    # Reads and parses a -Mode Export JSON file. Split out of Import-Template so the import pipeline
    # has a single in-memory input contract regardless of where the attributes came from (file or
    # direct read from the source forest).
    param([Parameter(Mandatory)][string]$Path)

    # -LiteralPath: a filename containing [ ] must not be wildcard-expanded; -PathType Leaf: a
    # directory must fail here with a clear message, not inside ReadAllText with a raw exception.
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "File '$Path' was not found (or is not a file)."
    }

    # Read as UTF-8 explicitly (BOM or BOM-less): Get-Content -Raw on Windows PowerShell 5.1 decodes a
    # BOM-less UTF-8 file (e.g. written by PS7) as ANSI, silently corrupting non-ASCII names.
    $fullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    [System.IO.File]::ReadAllText($fullPath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
}

function Get-TemplateAllowedAttributes {
    # The effective set of attributes a pKICertificateTemplate may hold in the TARGET forest,
    # lowercased and cached for the run. Read as the constructed 'allowedAttributes' of an existing
    # template instance - authoritative because AD computes it from the whole class hierarchy AND any
    # auxiliary classes, so a vendor schema extension attached via an auxiliary class is included
    # (reading only the class's own mayContain would wrongly drop those). Falls back to the
    # classSchema's mayContain/systemMayContain if the container holds no template to sample.
    # Gates schema typing: an attribute not in this set is dropped with a warning rather than admitted
    # into the create (which would otherwise fail server-side with a raw objectClassViolation).
    param([string]$ConfigNC, [hashtable]$ADParams)

    if ($null -ne $script:TemplateAllowedCache) { return $script:TemplateAllowedCache }

    $set = @{}
    $templatesDN = "CN=Certificate Templates,CN=Public Key Services,CN=Services,$ConfigNC"
    $sample = Get-ADObject @ADParams -SearchBase $templatesDN -LDAPFilter '(objectClass=pKICertificateTemplate)' `
        -ResultSetSize 1 -Properties allowedAttributes -ErrorAction Stop
    if ($sample -and $sample.allowedAttributes) {
        foreach ($a in @($sample.allowedAttributes)) { if ($a) { $set[$a.ToLowerInvariant()] = $true } }
    }
    else {
        $schemaNC = "CN=Schema,$ConfigNC"
        $cls = Get-ADObject @ADParams -SearchBase $schemaNC -ErrorAction Stop `
            -LDAPFilter "(&(objectClass=classSchema)(lDAPDisplayName=pKICertificateTemplate))" `
            -Properties mayContain, systemMayContain
        foreach ($a in @(@($cls.mayContain) + @($cls.systemMayContain))) {
            if ($a) { $set[$a.ToLowerInvariant()] = $true }
        }
    }
    $script:TemplateAllowedCache = $set
    $set
}

function Get-SchemaAttributeType {
    # Types an attribute the static lists don't know by asking the TARGET forest's schema
    # (attributeSyntax + isSingleValued), so newer or vendor msPKI-* attributes copy correctly
    # without a script edit. Returns 'Int', 'String', 'MultiString', 'Bytes', or $null (attribute
    # unknown in the target schema, not permitted on the template class, or a syntax this script
    # cannot round-trip). Cached per attribute name for the run.
    param([string]$AttributeName, [string]$ConfigNC, [hashtable]$ADParams)

    if ($script:SchemaTypeCache.ContainsKey($AttributeName)) { return $script:SchemaTypeCache[$AttributeName] }

    $type = $null
    # Only type attributes the template class actually permits: an attributeSchema can exist while
    # the class link was never applied, and admitting it would make New-ADObject fail with a raw
    # objectClassViolation instead of a clean warn-and-drop.
    if ((Get-TemplateAllowedAttributes -ConfigNC $ConfigNC -ADParams $ADParams).ContainsKey($AttributeName.ToLowerInvariant())) {
        $schemaNC = "CN=Schema,$ConfigNC"
        $attr = Get-ADObject @ADParams -SearchBase $schemaNC -ErrorAction Stop `
            -LDAPFilter "(&(objectClass=attributeSchema)(lDAPDisplayName=$(ConvertTo-LdapFilterValue $AttributeName)))" `
            -Properties attributeSyntax, isSingleValued
        if ($attr) {
            $single = [bool]$attr.isSingleValued
            $type = switch ($attr.attributeSyntax) {
                '2.5.5.9'  { if ($single) { 'Int' } else { $null } }             # integer / enumeration
                '2.5.5.12' { if ($single) { 'String' } else { 'MultiString' } }  # unicode string
                '2.5.5.10' { if ($single) { 'Bytes' } else { $null } }           # octet string (multi-valued octet does not round-trip JSON unambiguously)
                default    { $null }                                             # incl. 2.5.5.15 NT-Sec-Desc: never copied
            }
        }
    }
    $script:SchemaTypeCache[$AttributeName] = $type
    $type
}

function Import-Template {
    param(
        [Parameter(Mandatory)]
        [psobject]$InputObject,   # from Read-TemplateExport (file flow) or Get-SourceTemplate (Sync)
        [string]$NewTemplateName,
        [string]$NewDisplayName,
        [string]$OidHandling = 'Preserve',
        [string]$OidRoot,
        [string]$ExplicitOid,
        [string]$ConfigNC,        # skips the RootDSE read when the caller already has it
        [hashtable]$ADParams,
        [switch]$UpgradeCompatibility,
        [System.Management.Automation.PSCmdlet]$CallerCmdlet
    )

    $import = $InputObject

    # Resolve identity: explicit parameters win, else fall back to whatever the source carries.
    $cn = if ($NewTemplateName) { $NewTemplateName } elseif ($import.name) { $import.name } else { $null }
    $displayName = if ($NewDisplayName) { $NewDisplayName } elseif ($import.displayName) { $import.displayName } else { $null }

    if (-not $cn) {
        throw "No internal template name (cn) available. The source carries no 'name' (stripped on export?) - supply -NewTemplateName."
    }
    if (-not $displayName) {
        throw "No display name available. The source carries no 'displayName' (a template created without one, or stripped on export?) - supply -NewDisplayName."
    }
    # Allowlist: the cn lands in DN strings and LDAP filters, so DN/filter metacharacters
    # (, + = " \ ; < > # * /) stay banned - but legal real-world names are allowed: letters
    # (non-ASCII included), digits, spaces (not leading/trailing) and . _ - ( ). Filter values are
    # escaped at every use site, the RDN is escaped by New-ADObject itself, and the authoritative
    # DN comes from -PassThru.
    if ($cn -notmatch '^[\p{L}\p{Nd}._()\- ]+$' -or $cn -match '^\s|\s$') {
        throw "The internal template name (cn) may only contain letters, digits, spaces and . _ - ( ), with no leading or trailing space: '$cn'. Put other special characters in -NewDisplayName instead."
    }

    $configNC = if ($ConfigNC) { $ConfigNC } else { Get-ConfigNC -ADParams $ADParams }
    $templatesDN = "CN=Certificate Templates,CN=Public Key Services,CN=Services,$configNC"
    $newTemplateDN = "CN=$cn,$templatesDN"

    # Pre-flight: the Certificate Templates container must exist. It belongs to the forest's Public
    # Key Services structure and does NOT require a CA - but confirm it, so a forest missing that
    # structure fails with a clear message instead of a raw New-ADObject path error.
    if (-not (Get-ADObjectIfPresent -Identity $templatesDN -ADParams $ADParams)) {
        throw "The Certificate Templates container was not found at '$templatesDN'. This forest is missing the Public Key Services structure; this script does not provision it."
    }

    # Pre-flight: refuse to clobber an existing template of the same cn. -ErrorAction Stop: the
    # guard may only pass on a SUCCESSFUL empty search - a swallowed query failure would skip it.
    $existing = Get-ADObject @ADParams -SearchBase $templatesDN -LDAPFilter "(cn=$(ConvertTo-LdapFilterValue $cn))" -ErrorAction Stop
    if ($existing) {
        throw "A template with cn '$cn' already exists ($newTemplateDN). Choose a different -NewTemplateName or remove the existing template first."
    }

    # Build the functional attribute set (identity + OID handled separately below).
    $oa = @{}
    $unconsumed = @()
    foreach ($prop in ($import | Get-Member -MemberType NoteProperty)) {
        $name = $prop.Name
        # Identity/handled-elsewhere fields are skipped; a $null value means "attribute not set on
        # the source" and is skipped too - coercing it would fabricate a value ([int]$null -> 0)
        # or hand New-ADObject a $null it rejects mid-create.
        if ($name -in 'name', 'displayName', 'objectClass', 'msPKI-Cert-Template-OID') { continue }
        if ($null -eq $import.$name) { continue }
        if ($name -eq 'pKIEnrollmentAccess') {
            # An ACL-bearing attribute (present only in exports made before it was excluded):
            # security descriptors are deliberately never copied - the ACL is rebuilt on import.
            Write-Verbose "Skipping pKIEnrollmentAccess (ACLs are rebuilt on import, never copied)."
            continue
        }
        if ($name -in $script:IntAttributes) {
            $oa[$name] = [System.Int32]$import.$name
        }
        elseif ($name -in $script:MultiValueAttributes) {
            $oa[$name] = [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection]$import.$name
        }
        elseif ($name -in $script:ByteAttributes) {
            $oa[$name] = [System.Byte[]]$import.$name
        }
        elseif ($name -match '^(msPKI-|pKI)') {
            # A PKI attribute the static lists don't know (a schema extension linked to the template
            # class): type it from the TARGET forest's schema instead of dropping it. Returns $null
            # when the attribute is not permitted on the template class or has a type this script
            # cannot round-trip - then it is dropped (surfaced below), never silently.
            # The Int/String branches are SINGLE-valued in the target schema: exactly one element is
            # required - a multi-element value would be corrupted (space-joined) or crash the cast,
            # and an EMPTY array would fabricate a value ([int]$null -> 0), so both shapes drop to
            # $unconsumed. @(...)[0] also unwraps a one-element array, which [System.Int32] alone
            # would not. The try/catch turns any remaining cast failure (schema-divergent shapes,
            # e.g. a string where the target expects octets) into the same drop-with-warning instead
            # of a raw terminating cast error.
            $schemaType = Get-SchemaAttributeType -AttributeName $name -ConfigNC $configNC -ADParams $ADParams
            $oneValue = @($import.$name).Count -eq 1
            try {
                switch ($schemaType) {
                    'Int'         { if ($oneValue) { $oa[$name] = [System.Int32]@($import.$name)[0] } else { $unconsumed += $name } }
                    'String'      { if ($oneValue) { $oa[$name] = [string]@($import.$name)[0] } else { $unconsumed += $name } }
                    'MultiString' { $oa[$name] = [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection]$import.$name }
                    'Bytes'       { $oa[$name] = [System.Byte[]]$import.$name }
                    default       { $unconsumed += $name }
                }
            }
            catch {
                $oa.Remove($name)
                $unconsumed += $name
            }
        }
    }
    if ($unconsumed.Count) {
        Write-Warning "The source carries PKI attribute(s) the target forest's schema does not permit on the template class (or whose type/shape this script cannot copy there); they will NOT be written to the new template: $($unconsumed -join ', ')."
    }

    # v3/v4 templates can pack a private-key security descriptor (msPKI-Key-Security-Descriptor,
    # an SDDL string) inside msPKI-RA-Application-Policies. It copies verbatim, but any domain SIDs
    # inside it are SOURCE-forest SIDs that will not resolve in this forest - surface that.
    if ($oa.ContainsKey('msPKI-RA-Application-Policies') -and
        (@($oa['msPKI-RA-Application-Policies']) -match 'msPKI-Key-Security-Descriptor')) {
        Write-Warning "msPKI-RA-Application-Policies embeds an msPKI-Key-Security-Descriptor (private-key SDDL). Domain SIDs inside it are from the SOURCE forest and will not resolve here - review and adjust the key security descriptor on the copied template if it carries custom key permissions."
    }

    # Optional: raise the copy to the newest compatibility as it is created (schema v2/v3 -> v4 plus
    # the matching private-key-flag bits). Mutates $oa in place; the source is never touched.
    $compatNote = ''
    if ($UpgradeCompatibility) {
        $compat = Convert-ToLatestCompatibility -Attributes $oa
        if ($compat.Upgraded) {
            Write-Verbose "Compatibility raised to latest: schema v$($compat.FromVersion) -> v4, msPKI-Private-Key-Flag $($compat.PrivateKeyFlag)$(if ($compat.LegacyProvider) { ' (legacy provider)' } else { ' (CNG/KSP)' })."
            $compatNote = ' [compatibility upgraded to latest: CA Windows Server 2016 / recipient Windows 10]'
        }
        else {
            Write-Warning "-UpgradeCompatibility: $($compat.Reason); the template is imported at its existing compatibility."
        }
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
        -LDAPFilter "(msPKI-Cert-Template-OID=$($oidPlan.Oid))" -ErrorAction Stop
    if ($oidClash) {
        throw "A template already carries OID $($oidPlan.Oid): $($oidClash.DistinguishedName). Importing another template with the same OID would make OID-based template lookups (Windows, EJBCA) ambiguous. Use -OidHandling Generate, GenerateFromRoot, or GenerateRandom to mint a different OID."
    }

    $actionText = "Create certificate template (OID handling: $oidLabel)$compatNote"
    if ($oidPlan.CompanionCn) {
        $actionText += " and companion OID display object CN=$($oidPlan.CompanionCn),$($oidPlan.CompanionContainerDN)"
    }
    if (-not $CallerCmdlet.ShouldProcess($newTemplateDN, $actionText)) {
        if ($WhatIfPreference) {
            return $newTemplateDN   # -WhatIf: DN returned so the ACL preview can name its target
        }
        return $null                # declined at the -Confirm prompt: nothing was created
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
        # -ErrorAction Stop: without it a failure here (e.g. write access to CN=Certificate Templates
        # but not CN=OID) is only statement-terminating - the import would carry on, create the
        # template anyway, and report green success for an action that half-happened.
        New-ADObject @ADParams -Path $oidPlan.CompanionContainerDN -Name $oidPlan.CompanionCn `
            -Type 'msPKI-Enterprise-Oid' -OtherAttributes $oidObjectAttrs -Confirm:$false -ErrorAction Stop
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

function Get-RootDomainSid {
    # Resolves the forest-root domain SID (for the RID-498/519 grants) - called ONLY when a grant
    # actually needs it, and cached for the run. When the named -Server DC is in a child domain,
    # the root domain lives somewhere the caller never named: prefer reading the root domain head's
    # objectSid from the named DC's own Global Catalog (port 3268), which keeps the explicitly-
    # named-servers contract intact (no DNS/firewall dependency on an unnamed root-domain DC). Only
    # if the named DC is not a GC fall back to locating a root-domain DC by domain name.
    param($TargetDomain, $TargetForest, [hashtable]$ADParams)

    if ($script:RootDomainSidCache) { return $script:RootDomainSidCache }

    if ($TargetDomain.DNSRoot -eq $TargetForest.RootDomain) {
        $script:RootDomainSidCache = $TargetDomain.DomainSID.Value
        return $script:RootDomainSidCache
    }

    $rootDomainSID = $null
    $rootNC = 'DC=' + (($TargetForest.RootDomain -split '\.') -join ',DC=')
    if ($ADParams.ContainsKey('Server')) {
        # Build the Global Catalog endpoint (port 3268) from -Server, replacing any explicit port.
        # Bracket a bare IPv6 literal so the port is unambiguous and the trailing hextet is not
        # mistaken for a port (a plain ':\d+$' strip would turn '2001:db8::1' into '2001:db8:').
        $srv = $ADParams['Server']
        $gcServer = if ($srv -match '^\[(.+)\](:\d+)?$') { "[$($Matches[1])]:3268" }        # [ipv6] or [ipv6]:port
                    elseif (($srv -split ':').Count -gt 2)   { "[$srv]:3268" }               # bare IPv6 literal
                    else                                     { ($srv -replace ':\d+$', '') + ':3268' }  # host / ipv4 (:port)
        $gcParams = @{} + $ADParams
        $gcParams['Server'] = $gcServer
        try {
            $rootSid = (Get-ADObject @gcParams -Identity $rootNC -Properties objectSid -ErrorAction Stop).objectSid
            if ($rootSid -is [byte[]]) { $rootSid = New-Object System.Security.Principal.SecurityIdentifier($rootSid, 0) }
            if ($rootSid) { $rootDomainSID = $rootSid.Value }
        }
        catch {
            Write-Verbose "Global Catalog read of '$rootNC' via $($gcParams['Server']) failed ($($_.Exception.Message)); falling back to a root-domain DC lookup."
        }
    }
    if (-not $rootDomainSID) {
        $rootParams = @{ Server = $TargetForest.RootDomain; ErrorAction = 'Stop' }
        if ($ADParams.ContainsKey('Credential')) { $rootParams['Credential'] = $ADParams['Credential'] }
        try {
            $rootDomainSID = (Get-ADDomain @rootParams).DomainSID.Value
        }
        catch {
            throw "Could not determine the forest-root domain SID: the named DC's Global Catalog (port 3268) was not readable and no DC of root domain '$($TargetForest.RootDomain)' could be reached by name. Point -Server at a DC in the forest root domain, or allow GC (3268) access to the named DC. Underlying error: $($_.Exception.Message)"
        }
    }

    $script:RootDomainSidCache = $rootDomainSID
    $rootDomainSID
}

function Get-WellKnownTokenSid {
    # SID for a normalized well-known token name; $null when the name is not a token. The
    # enterprise tokens resolve the forest-root domain SID on first use (cached).
    param([string]$Norm, $TargetDomain, $TargetForest, [hashtable]$ADParams)

    $domSid = $TargetDomain.DomainSID.Value
    switch ($Norm) {
        'authenticatedusers'          { return New-Object System.Security.Principal.SecurityIdentifier('S-1-5-11') }
        'everyone'                    { return New-Object System.Security.Principal.SecurityIdentifier('S-1-1-0') }
        'enterprisedomaincontrollers' { return New-Object System.Security.Principal.SecurityIdentifier('S-1-5-9') }
        'domaincontrollers'           { return New-Object System.Security.Principal.SecurityIdentifier("$domSid-516") }
        'domaincomputers'             { return New-Object System.Security.Principal.SecurityIdentifier("$domSid-515") }
        'domainusers'                 { return New-Object System.Security.Principal.SecurityIdentifier("$domSid-513") }
        'domainadmins'                { return New-Object System.Security.Principal.SecurityIdentifier("$domSid-512") }
        'enterpriseadmins' {
            $rootSid = Get-RootDomainSid -TargetDomain $TargetDomain -TargetForest $TargetForest -ADParams $ADParams
            return New-Object System.Security.Principal.SecurityIdentifier("$rootSid-519")
        }
        { $_ -in 'enterprisereadonlydomaincontrollers', 'enterpriserodcs' } {
            $rootSid = Get-RootDomainSid -TargetDomain $TargetDomain -TargetForest $TargetForest -ADParams $ADParams
            return New-Object System.Security.Principal.SecurityIdentifier("$rootSid-498")
        }
    }
    return $null
}

function Resolve-PrincipalSid {
    # Resolves a principal (for -EnrollPrincipals) to a SID. Accepts a raw SID string, a name
    # (sAMAccountName / UPN, optionally DOMAIN\-prefixed), or a language-invariant well-known token.
    # Collisions are resolved fail-closed, never guessed:
    #   * a bare string that matches a well-known token (DomainAdmins, EnterpriseRODCs, ...) AND also
    #     matches a directory object is accepted only when both resolve to the SAME SID (e.g.
    #     'Domain Admins' on an English forest, where the built-in group's own sAMAccountName equals
    #     the token); a DIFFERENT SID means a planted or shadowing object and is refused -
    #     disambiguate with a SID or DOMAIN\ prefix;
    #   * a name (sAMAccountName/UPN) that matches more than one object is refused (AD does not enforce
    #     UPN uniqueness) - pick with a SID.
    # A DOMAIN\-prefixed input is an explicit directory reference: it never token-matches (the
    # backslash survives normalization) and is looked up only in the target (-Server) domain.
    param(
        [string]$Identity,
        $TargetDomain,          # Get-ADDomain result (has .DomainSID)
        $TargetForest,          # Get-ADForest result (for the lazily-resolved root-domain SID)
        [hashtable]$ADParams
    )

    $id = $Identity.Trim()

    if ($id -match '^S-1-\d+(-\d+)+$') {
        return New-Object System.Security.Principal.SecurityIdentifier($id)
    }

    # A DOMAIN\ prefix is honoured only when it names the target domain itself; a mismatched prefix
    # is rejected rather than silently resolved against a same-named principal in the target domain
    # (which would grant the wrong object). For a principal in another domain, pass its SID.
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

    # Is this a well-known token name? (Cheap string check; the SID - possibly a root-domain read for
    # the enterprise tokens - is resolved only if the token is actually used, below.) A DOMAIN\-prefixed
    # input never matches (the backslash survives normalization).
    $knownTokens = 'authenticatedusers', 'everyone', 'enterprisedomaincontrollers', 'domaincontrollers',
                   'domaincomputers', 'domainusers', 'domainadmins', 'enterpriseadmins',
                   'enterprisereadonlydomaincontrollers', 'enterpriserodcs'
    $norm = ($id -replace '[\s_\-]', '').ToLowerInvariant()
    $isToken = $norm -in $knownTokens

    # Directory lookup - sAMAccountName (unique within a domain), then UPN - both LDAP-escaped ('\'
    # never goes into a filter). -ErrorAction Stop: a failed search must fail the run, not silently
    # degrade into a token/not-found path that could resolve a different SID.
    $escName = ConvertTo-LdapFilterValue $name
    $obj = @(Get-ADObject @ADParams -LDAPFilter "(sAMAccountName=$escName)" -Properties objectSid -ErrorAction Stop |
            Where-Object { $_.objectSid })

    if (-not $obj.Count -and $id -match '@') {
        $escUpn = ConvertTo-LdapFilterValue $id
        $obj = @(Get-ADObject @ADParams -LDAPFilter "(userPrincipalName=$escUpn)" -Properties objectSid -ErrorAction Stop |
                Where-Object { $_.objectSid })
    }
    if ($obj.Count -gt 1) {
        throw "Principal '$Identity' is ambiguous: $($obj.Count) objects match ($(@($obj | ForEach-Object { $_.DistinguishedName }) -join '; ')). Use the intended principal's SID (S-1-5-...) instead."
    }
    if ($obj.Count -eq 1) {
        $objSid = [System.Security.Principal.SecurityIdentifier]$obj[0].objectSid
        if ($isToken) {
            # The input is both a well-known token AND a real directory object. Same SID both ways
            # (the built-in group's own sAMAccountName, e.g. 'Domain Admins' on an English forest):
            # no ambiguity, accept. A DIFFERENT SID means a planted object trying to hijack the
            # token, or the token shadowing a distinct real group - refuse rather than guess.
            $tokenSid = Get-WellKnownTokenSid -Norm $norm -TargetDomain $TargetDomain -TargetForest $TargetForest -ADParams $ADParams
            if (-not $tokenSid -or $tokenSid.Value -ne $objSid.Value) {
                throw "Principal '$Identity' matches BOTH the well-known token and a DIFFERENT directory object ($($obj[0].DistinguishedName)). Disambiguate with a SID (S-1-5-...) for the exact principal, or a DOMAIN\ prefix to force the directory object."
            }
        }
        return $objSid
    }

    # No directory object matched - resolve the well-known token if the name is one.
    if ($isToken) {
        $tokenSid = Get-WellKnownTokenSid -Norm $norm -TargetDomain $TargetDomain -TargetForest $TargetForest -ADParams $ADParams
        if ($tokenSid) { return $tokenSid }
    }

    # Deliberately NO cn fallback: cn is not unique to security principals, and a domain-wide cn
    # match could resolve to an object anyone with create rights planted under that name. Fail closed.
    throw "Could not resolve principal '$Identity' by sAMAccountName or UPN in the target domain. Use its sAMAccountName, a UPN (user@domain), a SID (S-1-5-...), or a well-known token (DomainControllers, DomainComputers, DomainUsers, DomainAdmins, EnterpriseAdmins, EnterpriseRODCs, EnterpriseDomainControllers, AuthenticatedUsers, Everyone)."
}

function Resolve-TemplateGrants {
    # Resolves the grants to ADD to the DACL, UP FRONT (before anything is created). The 'Standard'
    # 6-entry Kerberos Authentication set is included when -AclBase is Standard or SchemaPlusStandard;
    # the user's -EnrollPrincipals (if any) are always added on top. Whether these replace or are added
    # to the schema-default ACL is decided later by Set-TemplateAcl (-ReplaceExisting). Each grant is
    # @{ Sid; Rights = @(lowercased keywords); Label }.
    param([string]$AclBase, [hashtable]$EnrollPrincipals, [hashtable]$ADParams)

    $havePrincipals = $EnrollPrincipals -and $EnrollPrincipals.Count -gt 0
    if ($AclBase -eq 'PrincipalsOnly' -and -not $havePrincipals) {
        throw "-AclBase PrincipalsOnly requires -EnrollPrincipals; otherwise the template would be given an empty ACL."
    }
    $includeStandard = $AclBase -in 'Standard', 'SchemaPlusStandard'
    if (-not $includeStandard -and -not $havePrincipals) {
        # -AclBase Schema with nothing to add: the grant list is empty and Set-TemplateAcl will
        # leave the schema default untouched - no directory reads (and no root-SID resolution,
        # which could even fail) are needed at all.
        return @()
    }

    $targetDomain = Get-ADDomain @ADParams -ErrorAction Stop
    $targetForest = Get-ADForest @ADParams -ErrorAction Stop

    $validRights = @('read', 'write', 'enroll', 'autoenroll', 'fullcontrol')
    $grants = @()

    # The script's standard Kerberos Authentication set (matches the built-in template). Read is granted
    # via Authenticated Users (every DC computer account is a member), so the DC / RODC / Enterprise-DC
    # groups only need enroll + autoenroll; admins get Read/Write/Enroll (not Full Control / Autoenroll);
    # SYSTEM is deliberately not granted.
    if ($includeStandard) {
        $domSid        = $targetDomain.DomainSID.Value
        $rootDomainSID = Get-RootDomainSid -TargetDomain $targetDomain -TargetForest $targetForest -ADParams $ADParams
        $grants += @{ Sid = (New-Object System.Security.Principal.SecurityIdentifier('S-1-5-11'));           Rights = @('read');                    Label = 'Authenticated Users' }
        $grants += @{ Sid = (New-Object System.Security.Principal.SecurityIdentifier("$rootDomainSID-498")); Rights = @('enroll', 'autoenroll');    Label = 'Enterprise Read-only Domain Controllers (RID 498)' }
        $grants += @{ Sid = (New-Object System.Security.Principal.SecurityIdentifier("$domSid-512"));        Rights = @('read', 'write', 'enroll'); Label = 'Domain Admins (RID 512)' }
        $grants += @{ Sid = (New-Object System.Security.Principal.SecurityIdentifier("$domSid-516"));        Rights = @('enroll', 'autoenroll');    Label = 'Domain Controllers (RID 516)' }
        $grants += @{ Sid = (New-Object System.Security.Principal.SecurityIdentifier("$rootDomainSID-519")); Rights = @('read', 'write', 'enroll'); Label = 'Enterprise Admins (RID 519)' }
        $grants += @{ Sid = (New-Object System.Security.Principal.SecurityIdentifier('S-1-5-9'));            Rights = @('enroll', 'autoenroll');    Label = 'Enterprise Domain Controllers (S-1-5-9)' }
    }

    if ($havePrincipals) {
        foreach ($key in $EnrollPrincipals.Keys) {
            $rights = @(@($EnrollPrincipals[$key]) | ForEach-Object { "$_".Trim().ToLowerInvariant() } | Where-Object { $_ })
            if (-not $rights.Count) {
                throw "Principal '$key' was given no rights. Specify one or more of: Read, Write, Enroll, Autoenroll, FullControl."
            }
            $bad = @($rights | Where-Object { $_ -notin $validRights })
            if ($bad.Count) {
                throw "Unknown right(s) '$($bad -join ', ')' for principal '$key'. Valid rights: Read, Write, Enroll, Autoenroll, FullControl."
            }
            $sid = Resolve-PrincipalSid -Identity $key -TargetDomain $targetDomain -TargetForest $targetForest -ADParams $ADParams
            $grants += @{ Sid = $sid; Rights = $rights; Label = $key }
        }
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
        if ($WhatIfPreference) {
            # -WhatIf: nothing was created; preview the planned grants and return.
            if (-not $ReplaceExisting -and -not $Grants.Count) {
                Write-Host "What if: Would leave the schema-default ACL on $TemplateDN unchanged." -ForegroundColor Yellow
            }
            else {
                $verb = if ($ReplaceExisting) { 'set (replacing the schema default)' } else { 'add on top of the schema default' }
                Write-Host "What if: Would $verb these grants on ${TemplateDN}:" -ForegroundColor Yellow
                foreach ($g in $Grants) { Write-Host "         - $($g.Label): $($g.Rights -join ', ')" -ForegroundColor Yellow }
            }
            return
        }
        # A real run reached the ACL step for a template that is not visible although its creation
        # was reported (a declined -Confirm returns $null from Import-Template and never gets here).
        # That is replication lag (a domain-name -Server hitting another replica) or external deletion: fail
        # loudly instead of 'skipping permissions' with exit 0 - the template would silently keep
        # the schema-default DACL (no Enroll rights) that a consumer like EJBCA then enforces.
        throw "Template '$TemplateDN' is not visible on the targeted server although its creation was reported - replication lag (a domain-name -Server hitting another replica?) or external deletion. The ACL was NOT applied."
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

    $action = if ($ReplaceExisting) {
        "Set template ACL to $($Grants.Count) grant(s), replacing the schema default"
    }
    else {
        "Add $($Grants.Count) grant(s) to the template ACL, keeping the schema default"
    }
    if ($CallerCmdlet.ShouldProcess($TemplateDN, $action)) {
        # The security descriptor is read and written over ADWS (Get/Set-ADObject on
        # nTSecurityDescriptor), like every other operation in the script - one protocol (TCP 9389),
        # the same -Server, and the splat's -Credential applies as-is. Statement-terminating errors
        # do NOT stop a function under the default ErrorActionPreference, hence -ErrorAction Stop +
        # try/catch so a failure cannot cascade into the green success line.
        try {
            $sec = (Get-ADObject @ADParams -Identity $TemplateDN -Properties nTSecurityDescriptor -ErrorAction Stop).nTSecurityDescriptor
            if (-not $sec) { throw "The security descriptor could not be read." }

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

            Set-ADObject @ADParams -Identity $TemplateDN -Replace @{ nTSecurityDescriptor = $sec } -Confirm:$false -ErrorAction Stop
        }
        catch {
            throw "Failed to apply the ACL on '$TemplateDN': $($_.Exception.Message). The template EXISTS but still carries the schema-default DACL - no Enroll rights are granted until this is fixed."
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
    # Elements are escaped before joining so a literal '|' inside a value cannot make @('a|b') and
    # @('a','b') collide, and the sort is case-sensitive so the -ceq comparison downstream stays
    # deterministic for case-differing sets.
    param([string]$Name, $Value)
    # $null and an EMPTY collection both mean "attribute not set" and must canonicalize identically:
    # a raw ADObject materializes unset multi-valued attributes as empty collections, while a
    # Select-Object view of the same object yields $null for them (enumeration collapse).
    if ($null -eq $Value) { return '<null>' }
    $items = @($Value)
    if (-not $items.Count) { return '<null>' }
    # Ordered byte-wise for the known byte attributes AND any value that is actually a byte[] (a
    # schema-typed octet-string attribute the static list doesn't name): both diff operands come
    # from live -Properties * reads, so a real octet attribute is byte[] on both sides. Byte order
    # is significant, so it must not fall through to the order-insensitive set path below.
    if ($Name -in $script:ByteAttributes -or $Value -is [byte[]]) {
        return (([byte[]]$Value | ForEach-Object { $_.ToString('x2') }) -join '')
    }
    return (($items | ForEach-Object { "$_" -replace '\\', '\\' -replace '\|', '\|' } | Sort-Object -CaseSensitive) -join '|')
}

function Compare-TemplateAttributes {
    param($Source, $Target)
    # Diff the union of the import type lists AND every PKI attribute present on EITHER object.
    # Source-only attributes catch a lossy copy; target-only attributes catch the inverse (the copy
    # carrying values the source never had). Identity and the OID are excluded deliberately
    # (renamed / regenerated by design). The comparison is case-SENSITIVE - a case-only change is
    # still a change.
    $eitherPkiAttrs = @(@($Source.PSObject.Properties.Name) + @($Target.PSObject.Properties.Name) |
            Where-Object { $_ -match '^(msPKI-|pKI)' -and $_ -ne 'msPKI-Cert-Template-OID' })
    $attrs = @($script:IntAttributes + $script:MultiValueAttributes + $script:ByteAttributes + $eitherPkiAttrs) |
        Sort-Object -Unique
    foreach ($a in $attrs) {
        $sc = Get-AttrCanonical -Name $a -Value $Source.$a
        $tc = Get-AttrCanonical -Name $a -Value $Target.$a
        [pscustomobject]@{
            Attribute = $a
            Match     = ($sc -ceq $tc)
            Source    = $sc
            Target    = $tc
        }
    }
}

function Show-RoundTripDiff {
    # Prints the attribute diff for one round-trip pipeline and returns the number of mismatches.
    param([string]$Label, $Source, $Target)

    $diff       = Compare-TemplateAttributes -Source $Source -Target $Target
    $mismatches = @($diff | Where-Object { -not $_.Match })
    $byteRows   = @($diff | Where-Object { $_.Attribute -in $script:ByteAttributes })

    Write-Host ""
    Write-Host $Label -ForegroundColor Cyan
    $diff | Format-Table -AutoSize Attribute, Match, Source, Target | Out-Host

    $byteOk = @($byteRows | Where-Object { -not $_.Match }).Count -eq 0
    Write-Host ("Byte[] attributes ({0}) - {1}" -f `
            (($byteRows | ForEach-Object { $_.Attribute }) -join ', '), `
            $(if ($byteOk) { 'all identical after round-trip' } else { 'MISMATCH - see above' })) `
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
    $mismatches.Count
}

function Invoke-OneRoundTrip {
    # Runs ONE pipeline of the validation: import the given input under a throwaway identity, read
    # the copy back, and diff it against the source view. Returns the mismatch count, or $null when
    # the create was declined at a -Confirm prompt (the caller then aborts; its finally still
    # cleans up). $CreatedDN is [ref] so the caller can clean up even if this function throws.
    param(
        [string]$PipelineName,
        [string]$SourceName,
        [psobject]$ImportInput,
        [string]$Cn,
        [string]$Display,
        [string]$ExplicitOid,
        [string]$ConfigNC,
        [psobject]$Source,
        [hashtable]$ADParams,
        [System.Management.Automation.PSCmdlet]$CallerCmdlet,
        [ref]$CreatedDN
    )

    $CreatedDN.Value = Import-Template -InputObject $ImportInput -NewTemplateName $Cn -NewDisplayName $Display `
        -ExplicitOid $ExplicitOid -ConfigNC $ConfigNC -ADParams $ADParams -CallerCmdlet $CallerCmdlet
    if (-not $CreatedDN.Value) {
        # Declined at the -Confirm prompt (a genuine create failure throws inside Import-Template;
        # -WhatIf never reaches this function). Nothing was created, nothing to diff.
        Write-Warning "$PipelineName throwaway template was not created (declined at the -Confirm prompt); nothing to compare."
        return $null
    }

    $created = Get-ADObjectIfPresent -Identity $CreatedDN.Value -ADParams $ADParams -Properties *
    if (-not $created) {
        Write-Warning "$PipelineName throwaway template could not be read back after creation (replication lag on a domain-name -Server?); nothing to compare."
        return $null
    }

    Show-RoundTripDiff -Label "$PipelineName comparison: '$SourceName' -> '$Cn'" -Source $Source -Target $created
}

function Invoke-RoundTripValidation {
    param(
        [string]$TemplateName,
        [string]$Path,
        [switch]$KeepArtifacts,
        [hashtable]$ADParams,
        [System.Management.Automation.PSCmdlet]$CallerCmdlet
    )

    $configNC = Get-ConfigNC -ADParams $ADParams

    # ONE source read serves everything: the diff left-hand side, the file export's input, and the
    # direct-path import input (the latter being exactly what -Mode Sync feeds the import).
    $source = Get-SourceTemplate -TemplateName $TemplateName -ADParams $ADParams -ConfigNC $configNC

    if ($WhatIfPreference) {
        # Short-circuit: Validate creates AND removes throwaway templates. Under -WhatIf we make no
        # changes at all (and avoid the export->import round trip, whose file write would be suppressed).
        Write-Host "What if: would round-trip '$TemplateName' through BOTH pipelines - (1) export to a temp file and import a throwaway copy, (2) feed the live attribute view directly into the import as -Mode Sync does - diff every functional attribute of each copy, then remove the throwaway templates and file. Nothing is left changed." -ForegroundColor Yellow
        return
    }

    $usingTempFile = [string]::IsNullOrWhiteSpace($Path)
    if ($usingTempFile) {
        $Path = Join-Path $env:TEMP ("kerbtpl-roundtrip-" + (Get-RandomHex -Length 8) + ".json")
    }

    $suffix        = Get-RandomHex -Length 8
    $tempCn        = "RoundtripTest-$suffix"
    $tempDisplay   = "Roundtrip Test $suffix"
    $tempDN        = $null
    $directCn      = "DirectPathTest-$suffix"
    $directDisplay = "Direct Path Test $suffix"
    $directDN      = $null

    # Unique explicit OIDs so validation needs no forest OID root and cannot collide with the source
    # template's own OID. Derived from the source OID when present; else (e.g. a v1 template with no
    # msPKI-Cert-Template-OID) synthesized self-contained, so the -ExplicitOid path is always taken
    # (never $null, which would fall through to Preserve).
    $newThrowawayOid = {
        if ($source.'msPKI-Cert-Template-OID') {
            "$($source.'msPKI-Cert-Template-OID').$(Get-Random -Minimum 1000000 -Maximum 99999999)"
        }
        else {
            "$(New-SyntheticOidBase).$(Get-Random -Minimum 10000000 -Maximum 99999999).$(Get-Random -Minimum 10000000 -Maximum 99999999)"
        }
    }

    try {
        # Pipeline 1 - the FILE flow: export (serialize to JSON) -> read back -> import. Use the
        # written/not-written return, NOT Test-Path: a declined OVERWRITE of an explicit -Path leaves
        # the STALE file, which Test-Path would accept and then diff against, producing a false FAIL.
        $exported = Export-Template -InputObject $source -Path $Path -ADParams $ADParams -CallerCmdlet $CallerCmdlet -NoImportHint
        if (-not $exported) {
            Write-Warning "Export file was not written (declined); validation aborted."
            return
        }

        $fileMismatches = Invoke-OneRoundTrip -PipelineName 'File pipeline (Export/Import)' -SourceName $TemplateName `
            -ImportInput (Read-TemplateExport -Path $Path) -Cn $tempCn -Display $tempDisplay `
            -ExplicitOid (& $newThrowawayOid) -ConfigNC $configNC -Source $source `
            -ADParams $ADParams -CallerCmdlet $CallerCmdlet -CreatedDN ([ref]$tempDN)
        if ($null -eq $fileMismatches) { return }   # file-pipeline create declined; nothing validated

        # Pipeline 2 - the DIRECT flow: feed the live attribute view straight into the import,
        # exactly as -Mode Sync does. The file pipeline above cannot stand in for this one: live AD
        # values (ADPropertyValueCollection, byte[]) reach Import-Template's casts untouched by
        # JSON, so each pipeline needs its own round-trip proof.
        $directMismatches = Invoke-OneRoundTrip -PipelineName 'Direct pipeline (Sync)' -SourceName $TemplateName `
            -ImportInput $source -Cn $directCn -Display $directDisplay `
            -ExplicitOid (& $newThrowawayOid) -ConfigNC $configNC -Source $source `
            -ADParams $ADParams -CallerCmdlet $CallerCmdlet -CreatedDN ([ref]$directDN)
        if ($null -eq $directMismatches) {
            # Direct-pipeline create was declined. Still fail if the file pipeline already found
            # mismatches - a mismatch must never exit 0 (automation gates on that).
            if ($fileMismatches -gt 0) {
                throw "Round-trip validation FAILED: file pipeline $fileMismatches mismatch(es); the direct (Sync) pipeline was declined so could not be checked. See the diff output above."
            }
            Write-Warning "Direct (Sync) pipeline was declined; only the file pipeline was validated (it passed)."
            return
        }

        Write-Host ""
        if (($fileMismatches + $directMismatches) -eq 0) {
            Write-Host "OVERALL PASS: both the file (Export/Import) and the direct (Sync) pipeline reproduce every copied attribute." -ForegroundColor Green
        }
        else {
            Write-Host "OVERALL FAIL: file pipeline $fileMismatches mismatch(es), direct pipeline $directMismatches mismatch(es) - see above." -ForegroundColor Red
            # Fail as a real failure: automation gating on Validate must see a non-zero exit code,
            # not have to scrape console colors. Cleanup still runs (finally below).
            throw "Round-trip validation FAILED: file pipeline $fileMismatches mismatch(es), direct pipeline $directMismatches mismatch(es). The copy pipelines did not reproduce the source template faithfully - see the diff output above."
        }
    }
    finally {
        # Idempotent cleanup, safe on every exit path (early return, read-back failure after the
        # objects were created, a FAIL throw, or a mid-run error). Re-query by DN so it never relies
        # on in-try state. Throwaways are created with -ExplicitOid, which never registers a
        # companion OID display object, so only the templates themselves (and the temp file) need
        # removing.
        $throwawayDNs = @($tempDN, $directDN) | Where-Object { $_ }

        if ($KeepArtifacts) {
            foreach ($dn in $throwawayDNs) {
                Write-Host "-KeepArtifacts: left throwaway template '$dn' in place (if created)." -ForegroundColor Yellow
            }
            if ($throwawayDNs.Count) {
                Write-Host "-KeepArtifacts: left file '$Path' in place." -ForegroundColor Yellow
            }
        }
        else {
            $cleanupOk = $true
            foreach ($dn in $throwawayDNs) {
                # The whole cleanup runs in a finally that may be entered because the DC/network
                # dropped mid-run: the existence probe (Get-ADObjectIfPresent) would then re-throw
                # that connectivity error, aborting cleanup and MASKING the original failure. Wrap
                # each DN so one unreachable object cannot leak the rest (or the temp file).
                try {
                    if (-not (Get-ADObjectIfPresent -Identity $dn -ADParams $ADParams)) { continue }
                    Remove-ADObject @ADParams -Identity $dn -Confirm:$false -ErrorAction Stop
                }
                catch {
                    $cleanupOk = $false
                    Write-Warning "Could not remove throwaway template '$dn': $($_.Exception.Message)"
                }
            }
            if ($usingTempFile) {
                # -Confirm:$false: Remove-Item would otherwise raise its own prompt under a -Confirm run.
                Remove-Item -LiteralPath $Path -Force -Confirm:$false -ErrorAction SilentlyContinue
                if (Test-Path -LiteralPath $Path) {
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

# One table drives mode/parameter compatibility: any supplied parameter the mode does not consume
# is rejected up front instead of being silently ignored, so a mistyped combination fails loudly
# and a future mode or parameter needs exactly one list updated.
$modeParams = @{
    Export   = 'Path', 'TemplateName', 'StripIdentity', 'StripOid', 'Server', 'Credential'
    Import   = 'Path', 'NewTemplateName', 'NewDisplayName', 'OidHandling', 'OidRoot', 'SkipAcl', 'AclBase', 'EnrollPrincipals', 'UpgradeCompatibility', 'Server', 'Credential'
    Sync     = 'TemplateName', 'NewTemplateName', 'NewDisplayName', 'OidHandling', 'OidRoot', 'SkipAcl', 'AclBase', 'EnrollPrincipals', 'UpgradeCompatibility', 'Server', 'Credential', 'SourceServer', 'SourceCredential'
    Validate = 'TemplateName', 'Path', 'KeepArtifacts', 'Server', 'Credential'
}
$commonParams = @([System.Management.Automation.PSCmdlet]::CommonParameters) + @([System.Management.Automation.PSCmdlet]::OptionalCommonParameters) + 'Mode'
$notApplicable = @($PSBoundParameters.Keys | Where-Object { $_ -notin $commonParams -and $_ -notin $modeParams[$Mode] })
if ($notApplicable.Count) {
    throw "Parameter(s) not applicable to -Mode ${Mode}: $(($notApplicable | ForEach-Object { "-$_" }) -join ', ') - they would otherwise be silently ignored. -Mode $Mode consumes: $(($modeParams[$Mode] | ForEach-Object { "-$_" }) -join ', ')."
}

# Requirements and cross-parameter rules, in the same one place:
if ($Mode -in 'Export', 'Import' -and -not $Path) {
    throw "-Path is required for -Mode $Mode."
}
if ($Mode -eq 'Sync' -and -not $SourceServer) {
    throw "-SourceServer is required for -Mode Sync (a DC, or domain name, in the SOURCE forest to read the template from)."
}
# -Credential is meant for "operate on a forest I am not logged on to" - but DC discovery is
# DC-locator based and always finds the CURRENT forest. Requiring -Server alongside it guarantees
# the credentials are used against the forest the caller intends (a trust could otherwise let
# foreign credentials silently authenticate to - and write into - the local forest).
if ($Credential -and -not $Server) {
    throw "-Credential requires -Server: name the DC (in the forest those credentials belong to) explicitly, so the operation cannot land on a discovered DC in the current forest instead."
}
if ($SkipAcl -and $EnrollPrincipals -and $EnrollPrincipals.Count -gt 0) {
    throw "-SkipAcl and -EnrollPrincipals are mutually exclusive (-SkipAcl skips the very ACL that -EnrollPrincipals defines)."
}
if ($SkipAcl -and $PSBoundParameters.ContainsKey('AclBase')) {
    throw "-SkipAcl and -AclBase are mutually exclusive (-SkipAcl skips the very ACL step that -AclBase configures)."
}
if ($OidRoot -and $OidHandling -ne 'GenerateFromRoot') {
    throw "-OidRoot is only consumed by -OidHandling GenerateFromRoot; with '$OidHandling' it would be silently ignored. Add -OidHandling GenerateFromRoot, or drop -OidRoot."
}

$adParams = @{}
if ($Server) { $adParams['Server'] = $Server }
if ($Credential) { $adParams['Credential'] = $Credential }

# Import, Sync and Validate write to the config partition and then read back / bind by DN, so every
# step must hit ONE server. With no -Server, discover a writable DC that runs ADWS.
if ($Mode -in 'Import', 'Sync', 'Validate') {
    if ($adParams.ContainsKey('Server')) {
        # Whatever -Server names is used VERBATIM for every operation - this script never rewrites
        # the operator's endpoint. (Auto-"pinning" a domain name to a member DC was implemented and
        # deliberately abandoned: a client cannot safely pick a substitute server, because a
        # bit-identical clone, split DNS, or a NetBIOS-layer redirect can make a DIFFERENT directory
        # pass every health check a substitute could be given - so any rewrite risks silently
        # sending writes somewhere the operator did not name.)
        # What CAN be done safely is detection: a DOMAIN name (DNS - with or without the absolute
        # trailing dot - or NetBIOS) locates a different DC per connection, so the create,
        # read-back and ACL steps of one run may hit different replicas. Warn loudly and let the
        # operator choose; a lagging read-back then FAILS the run (see Set-TemplateAcl) rather than
        # mis-securing the template.
        try {
            $rootDSE = Get-ADRootDSE @adParams -ErrorAction Stop
        }
        catch {
            throw "Could not read RootDSE from '$($adParams['Server'])' - is the server name correct and reachable (ADWS, TCP 9389)? Underlying error: $($_.Exception.Message)"
        }
        $suppliedHost = $adParams['Server']
        if ($suppliedHost -notmatch ':.*:') { $suppliedHost = $suppliedHost -replace ':\d+$', '' }  # strip :port (never an IPv6 literal)
        $suppliedHost = $suppliedHost.TrimEnd('.')   # absolute DNS form 'domain.' is the same domain name
        $ncToDns = { param($nc) if ($nc) { (($nc -split ',') | ForEach-Object { $_ -replace '^DC=', '' }) -join '.' } }
        $domainNames = @((& $ncToDns $rootDSE.defaultNamingContext), (& $ncToDns $rootDSE.rootDomainNamingContext))
        if ($suppliedHost -notmatch '[.:]') {
            # Dotless and not an IPv6 literal: could be the domain's NetBIOS name (a NetBIOS HOST
            # name simply won't match it). Classification only - on any failure treat the value as
            # a host name; never abort the run for this.
            try { $domainNames += (Get-ADDomain @adParams -ErrorAction Stop).NetBIOSName }
            catch { Write-Verbose "Could not read the domain's NetBIOS name for -Server classification ($($_.Exception.Message))." }
        }
        if ($suppliedHost -in $domainNames) {
            Write-Warning "-Server '$($adParams['Server'])' is a DOMAIN name, which locates a different DC per connection: the create, read-back and ACL steps of this run may hit DIFFERENT replicas. Name one DC directly (e.g. '$($rootDSE.dnsHostName)') for a fully consistent run."
        }
    }
    else {
        # -Writable: plain -Discover can return an RODC in an RODC-only site, and every config-
        # partition write would then fail after the reads succeeded. -Service ADWS: every call this
        # script makes needs ADWS running on the selected DC.
        $dc = Get-ADDomainController -Discover -Writable -Service ADWS -ErrorAction Stop
        $adParams['Server'] = ($dc.HostName | Select-Object -First 1)
    }
}

switch ($Mode) {
    "Export" {
        # [void]: swallow the written/not-written boolean Export-Template returns (only Validate reads it).
        [void](Export-Template -TemplateName $TemplateName -Path $Path `
            -StripIdentity:$StripIdentity -StripOid:$StripOid -ADParams $adParams -CallerCmdlet $PSCmdlet)
    }
    { $_ -in 'Import', 'Sync' } {
        $importArgs = @{
            NewTemplateName      = $NewTemplateName
            NewDisplayName       = $NewDisplayName
            OidHandling          = $OidHandling
            OidRoot              = $OidRoot
            ADParams             = $adParams
            UpgradeCompatibility = $UpgradeCompatibility
            CallerCmdlet         = $PSCmdlet
        }
        if ($Mode -eq 'Sync') {
            # Direct forest-to-forest: read the template from the source forest and feed it straight
            # into the same import pipeline the file-based flow uses (no JSON round-trip at all).
            # After the target-side -Server validation above, the SOURCE is contacted before any other
            # target-side work, so the most error-prone inputs (a typo'd -SourceServer, a wrong
            # -TemplateName) fail before grants are resolved or anything is created.
            $sourceParams = @{ Server = $SourceServer }
            if ($SourceCredential) { $sourceParams['Credential'] = $SourceCredential }

            # Both RootDSEs up front: a clear connectivity error on either side, and a guard against
            # the same-forest accident (-Server omitted on a machine joined to the SOURCE forest
            # would otherwise silently make the source forest the write target). Same-forest is
            # detected by the Configuration NC head's objectGUID - forest-unique even when two
            # distinct forests share a DNS name (prod vs. its isolated clone) - with the DN string
            # as fallback if a GUID is unreadable.
            $sourceConfigNC = Get-ConfigNC -ADParams $sourceParams
            $targetConfigNC = Get-ConfigNC -ADParams $adParams
            $sourceConfigGuid = (Get-ADObjectIfPresent -Identity $sourceConfigNC -ADParams $sourceParams).ObjectGUID
            $targetConfigGuid = (Get-ADObjectIfPresent -Identity $targetConfigNC -ADParams $adParams).ObjectGUID
            $sameForest = if ($sourceConfigGuid -and $targetConfigGuid) { $sourceConfigGuid -eq $targetConfigGuid }
                          else { $sourceConfigNC -eq $targetConfigNC }
            if ($sameForest) {
                if (-not $PSBoundParameters.ContainsKey('Server')) {
                    throw "The discovered target DC '$($adParams['Server'])' is in the SAME forest as -SourceServer '$SourceServer' ($sourceConfigNC). Pass -Server naming a DC in the intended target forest - or, if a same-forest copy is intended, pass -Server explicitly to confirm."
                }
                Write-Warning "Source and target are the same forest ($sourceConfigNC) - proceeding with a same-forest copy."
            }

            $importArgs['InputObject'] = Get-SourceTemplate -TemplateName $TemplateName -ADParams $sourceParams -ConfigNC $sourceConfigNC
            $importArgs['ConfigNC']    = $targetConfigNC
        }
        else {
            $importArgs['InputObject'] = Read-TemplateExport -Path $Path
        }

        # The default -AclBase Standard writes the stock KERBEROS AUTHENTICATION ACL (DC-oriented:
        # Domain Controllers / Enterprise DCs / ERODCs get Enroll+Autoenroll, replacing the schema
        # default). Applied to an arbitrary template that is rarely what is wanted - and an external
        # CA (EJBCA) reads exactly this ACL as its authorization source - so say so whenever the
        # default was not an explicit choice and the template does not look like a Kerberos copy.
        if (-not $SkipAcl -and $AclBase -eq 'Standard' -and -not $PSBoundParameters.ContainsKey('AclBase')) {
            $intendedCn = if ($NewTemplateName) { $NewTemplateName } elseif ($importArgs['InputObject'].name) { "$($importArgs['InputObject'].name)" } else { '' }
            # Normalize away the spaces/hyphens the widened cn allowlist permits, so 'Kerberos
            # Authentication' / 'Kerberos-Authentication' are recognized as Kerberos copies. Skip the
            # warning entirely when no name is known yet (Import-Template raises the real error).
            if ($intendedCn -and (($intendedCn -replace '[\s_\-]', '') -notmatch 'KerberosAuthentication')) {
                Write-Warning "-AclBase defaulted to 'Standard', which writes the stock Kerberos Authentication ACL (domain controllers get Enroll+Autoenroll; schema default replaced) - but template '$intendedCn' does not look like a Kerberos Authentication copy. Pass -AclBase (and/or -EnrollPrincipals) explicitly if different permissions are intended."
            }
        }

        # Validate keywords and resolve every principal to a SID BEFORE creating anything, so bad ACL
        # input aborts with nothing created (and -WhatIf still exercises the resolution).
        $grants = if (-not $SkipAcl) { @(Resolve-TemplateGrants -AclBase $AclBase -EnrollPrincipals $EnrollPrincipals -ADParams $adParams) } else { $null }

        $templateDN = Import-Template @importArgs

        if (-not $templateDN) {
            # Declined at the -Confirm prompt: nothing was created, so there is nothing to secure.
            Write-Warning "Template creation was declined at the -Confirm prompt; nothing was created."
        }
        elseif (-not $SkipAcl) {
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
