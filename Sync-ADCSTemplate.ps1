<#PSScriptInfo
.VERSION 1.0.8
.GUID 689db74d-e668-410a-9a62-0b208179a369
.AUTHOR Sveinung Svea
.PROJECTURI https://github.com/TheOmnilord/ADCS
.LICENSEURI https://github.com/TheOmnilord/ADCS/blob/main/LICENSE
.TAGS ADCS PKI CertificateServices
.RELEASENOTES
1.0.8 - -OidHandling GenerateFromRoot without -OidRoot is now refused by the up-front parameter guards, before any domain controller is contacted and before any grant is resolved (the late check inside Resolve-TemplateOid is kept as a second line); the schema-typed conversion of a PKI attribute the static lists do not know is moved out of an inline switch in Import-Template into the new pure helper ConvertTo-SchemaTypedValue with the same rules (Int and String need exactly one element, MultiString and Bytes are cast, an unknown type or a failed cast drops the attribute with the existing warning) so that every arm can be unit-tested; the "Created template:" line now ends with " (objectGUID <guid>)" taken from the object New-ADObject -PassThru returned, and the companion OID display object is created with -PassThru and reported the same way on a new "Created OID object: <DN> (objectGUID <guid>)" line, so that a caller can identify exactly the objects this run created without a later lookup by name or OID; no other behaviour change
1.0.7 - Help text only: the comment-based help is rewritten to the repository writing style (STYLE.md, derived from ASD-STE100 Simplified Technical English) - short sentences, active voice, no figurative language, acronyms defined, a CAUTION line on -SkipAcl and -AllowLinkedIssuancePolicy; every fact, condition and default is kept; no code change
1.0.6 - ConvertTo-ImportAttributeValue checks integrality and range in the value's OWN numeric type before any [decimal] cast: a tiny double (1e-30) cast to decimal underflowed to 0 and was silently coerced to 0 (integer and byte-element branches both); Convert-ToLatestCompatibility computes and validates every replacement value - including the minor-revision increment, which now throws on Int32.MaxValue overflow - BEFORE mutating $Attributes, so a failure no longer leaves a template half-upgraded (v4 schema/flags with an un-bumped revision) while still reporting Upgraded; the Authentication Mechanism Assurance import guard scans only msPKI-Certificate-Policy (the issuance policies stamped into the ISSUED certificate), no longer msPKI-RA-Policies (which constrains the enrollment-agent SIGNING certificate and is not stamped into the issued cert, so an AMA link on it never grants the enrollee) - it was falsely refusing templates that merely require a signing-cert application policy
1.0.5 - The DOMAIN\user@domain principal form now takes the UPN-only resolution and sAMAccountName shadow check (matching on the raw key let a prefixed key skip to the sAMAccountName lookup, so a planted sAMAccountName could still capture the grant); the dotted-OID validation regexes are anchored with \z instead of $ (a trailing newline in a tampered msPKI-Cert-Template-OID passed validation and bypassed the template-OID uniqueness search); -UpgradeCompatibility refuses a schema-2 source carrying msPKI-RA-Application-Policies (its encoding differs at v3/v4, so upgrading in place would silently drop the RA-signature application-policy requirement)
1.0.4 - Every known attribute of an import is validated for type, shape and range and the import is refused when one is malformed (a failed cast previously dropped the attribute - msPKI-RA-Signature included - and the template was created without it; 0.4 was coerced to 0, a three-element period array passed as a period); an issuance policy OID that the TARGET forest links to a group via Authentication Mechanism Assurance refuses the import unless -AllowLinkedIssuancePolicy is given (new switch)
1.0.3 - Help text only: -EnrollPrincipals documents the UPN-only resolution of user@domain keys, -UpgradeCompatibility the legacy-provider rule, -Mode Validate the read-back failure; no code change
1.0.2 - A user@domain principal in -EnrollPrincipals resolves ONLY as a UPN, and a different object carrying that string as its sAMAccountName is refused (sAMAccountName may contain '@', so a planted account could previously capture a grant meant for the UPN); -UpgradeCompatibility sets CT_FLAG_USE_LEGACY_PROVIDER only for a schema-2 source with a provider list and preserves a schema-3 source's own bit (a v3 KSP template was switched to legacy provider handling); -Mode Validate fails with a terminating error when the throwaway template cannot be read back after creation (previously a warning and exit 0 with nothing validated)
1.0.1 - The template create is now -ErrorAction Stop with a returned-object check (a non-terminating New-ADObject failure previously printed a green "Created template" line with an empty DN, orphaned the companion OID object and exited 0); after the create the OID is re-queried and, if another template claimed it concurrently, the new template is rolled back and the run fails
1.0.0 - Initial release
#>

#Requires -Modules ActiveDirectory

<#
.SYNOPSIS
    Copies a certificate template from one AD forest to another, through a JSON file
    (Export/Import) or directly forest-to-forest in one run (Sync).

.DESCRIPTION
    The script copies the template attribute by attribute, without certutil, entirely over Active
    Directory Web Services (ADWS). It uses the ActiveDirectory PowerShell module for ALL modes, so
    the module is required for Export as well as for Import and Sync. The script can rename the
    template. It can keep the OID (object identifier) of the template, or generate a new OID. It
    applies standard AD CS (Active Directory Certificate Services) permissions after the import.
    The target forest can be one that has never had AD CS installed.

      -Mode Export
          Run this mode in the SOURCE forest. The script reads the functional attributes of the
          template through ADWS: flags, revision, and all msPKI-* and pKI* attributes. It writes
          them to a JSON file. The script deliberately does NOT export forest-specific data: the
          security descriptor, distinguishedName, and objectCategory. The target forest derives
          these values, or the script reapplies them there. You can also remove the source template
          OID and the identity fields (name and displayName) from the file with -StripOid and
          -StripIdentity.

      -Mode Import
          Run this mode in the TARGET forest. The script recreates the template from the JSON file
          through ADWS. The script:
            * handles the template OID as -OidHandling says. Preserve keeps the source OID. Generate
              generates a new OID under the real OID root of the forest. GenerateFromRoot generates a
              new OID under a base you supply in -OidRoot. GenerateRandom generates a new OID under a
              synthesized base.
            * registers an OID "display" object for every -OidHandling value, so Windows resolves the
              OID to the template name. Only Generate needs a pre-existing PKI OID root, that is, a
              forest where AD CS was deployed once.
            * derives the container DN from the Configuration partition of the TARGET forest. It lets
              Active Directory assign objectClass and objectCategory automatically.
            * lets you rename the template (internal cn and display name) with -NewTemplateName and
              -NewDisplayName.
            * prints one line for each object it creates. The line for the display object is
              "Created OID object: <DN> (objectGUID <guid>)". The line for the template is
              "Created template: <DN> (objectGUID <guid>)". The DN and the objectGUID come from
              the object that Active Directory returned at creation. A cleanup tool can identify
              the created object by that objectGUID, without a later lookup by name or OID.
            * sets the template permissions, unless you pass -SkipAcl. The base comes from -AclBase;
              -EnrollPrincipals adds optional grants on top. The default, -AclBase Standard, writes the
              standard Kerberos Authentication ACL. It replaces the schema default, so admins are not
              left with Full Control. Other -AclBase values keep or extend the schema-default ACL of
              Active Directory instead. A consumer such as EJBCA reads this ACL to decide who may
              enroll.
            * writes these entries in the standard Kerberos Authentication ACL: Authenticated Users
              get Read. Domain Admins and Enterprise Admins get Read, Write, and Enroll. Domain
              Controllers, Enterprise RODCs, and Enterprise Domain Controllers get Enroll and
              Autoenroll. Their Read comes through Authenticated Users. The script grants nothing to
              SYSTEM.

          Import requires Enterprise Admin rights, or delegated write access to the Certificate
          Templates and OID containers in the Configuration partition.

      -Mode Sync
          A direct forest-to-forest copy in one run, with no intermediate file. The script reads
          the template from a domain controller (DC) in the SOURCE forest, named by -SourceServer,
          with -SourceCredential when needed. It recreates the template on the TARGET side exactly
          as -Mode Import does. The target is -Server, or a discovered DC, with -Credential when
          needed. Sync supports -OidHandling, -NewTemplateName, -NewDisplayName, and the full ACL
          handling described under -Mode Import.

          The script sends the attributes it read straight into the import pipeline, so no JSON
          serialization happens at all. -Mode Validate proves the fidelity of this direct pipeline
          and of the file pipeline separately.

          Authentication: with a two-way trust between the forests, the identity that runs the
          script can usually read the source as-is. Authenticated Users has read access to
          templates. When that identity also holds Enterprise Admin rights in the target, only
          -SourceServer is needed. Without a trust, or when you run as neither identity, pass
          -SourceCredential, -Credential, or both. Explicit credentials against explicitly named
          servers need no trust at all.

      -Mode Validate
          Proves round-trip fidelity in a single forest, for BOTH copy pipelines, without touching
          a CA. The script reads a source template. First, it exports the template to a JSON file,
          temporary unless you give -Path, and imports that file under a throwaway name. This
          exercises the Export/Import
          file flow. Then it sends the live attribute view directly into the import under a second
          throwaway name. This exercises exactly what -Mode Sync does.

          Live Active Directory values reach the import casts untouched by JSON, so the file check
          cannot stand in for the direct check. Each throwaway template gets a unique OID that
          needs no OID root. The script compares every copied attribute of each copy with the
          source, byte[] attributes included. Any mismatch makes the run FAIL with a terminating
          error and a non-zero exit code, so automation can use the exit code to decide. Cleanup
          still runs first. A throwaway template that the script cannot read back after its
          confirmed creation fails the run the same way, because nothing was validated.

          The script never reports that read-back failure as a warning with exit 0. By default the
          script removes the throwaway templates and the temporary file afterwards. Pass
          -KeepArtifacts to keep them for inspection. Validate requires the same write access as
          Import.

    No CA required:
      * Export, Import, and Validate operate ONLY on the certificate TEMPLATE objects in the
        Configuration partition of Active Directory. No CA has to be installed, online, or
        reachable. The AD CS role and the RSAT "AD CS Tools" are not needed. Only the
        ActiveDirectory PowerShell module is needed.
      * Import and Validate need NO PKI OID root with the default -OidHandling Preserve. The target
        forest can therefore be one that never had AD CS. The target forest needs only the
        Certificate Templates container, which does not depend on a CA. This container is part of
        the Public Key Services structure of every forest. The script checks for it and fails with
        a clear message when the whole structure is absent.
      * Preserve, GenerateFromRoot, and GenerateRandom need no PKI OID root. They need only the
        Certificate Templates and OID containers, which do not depend on a CA and exist in every
        forest. Only -OidHandling Generate requires the actual "CN=OID,..." base OID of the forest,
        which exists once AD CS has been deployed there once. Without it, Generate fails with a
        clear message that names the other -OidHandling values.
      * The script moves a template DEFINITION. It moves no issued certificate and no private key.
        A CA takes part only later, when you publish the imported template on an issuing CA. The
        CA can then enroll certificates from it.

    Identity / OID / placement handling:
      * OID: the default, -OidHandling Preserve, reuses the OID of the source template. Do NOT
        combine Preserve with -StripOid on export. Generate, GenerateFromRoot, and GenerateRandom
        instead generate a new OID. Generate uses the real root of the forest. GenerateFromRoot
        uses the -OidRoot you supply. GenerateRandom uses a synthesized base.
      * The script derives objectCategory and the DN suffix from the TARGET forest automatically.
        The DN suffix is everything after "CN=Certificate Templates,...". You never edit them by
        hand.
      * The new internal name (cn) and the new display name come from -NewTemplateName and
        -NewDisplayName. If you did not strip the identity on export, the script uses the source
        name and displayName from the file as fallbacks.

.PARAMETER Mode
    "Export", "Import", "Sync", or "Validate". The file-based flow and the direct flow mix freely.
    A JSON file exported earlier imports into a remote forest with -Mode Import -Server <target DC>,
    plus -Credential when needed. In the same way, -Mode Export accepts -Server and -Credential to
    read from a remote source forest.

.PARAMETER Path
    The path of the JSON file. Export writes it, and Import reads it. It is required for Export and
    Import; Sync does not use it, because no intermediate file is involved. On Validate it is
    optional. If you give it, the script writes the intermediate export there and keeps the file
    for inspection. If you omit it, the script uses a temporary file and deletes it afterwards.

.PARAMETER TemplateName
    Export, Validate, and Sync. The internal name (cn) of the source template to read. The cn is
    unique within the templates container, so this name matches exactly one template.
    Default: "KerberosAuthentication".

.PARAMETER StripIdentity
    Export only. Removes the source name and displayName from the JSON file. You must then supply
    -NewTemplateName and -NewDisplayName on import.

.PARAMETER StripOid
    Export only. Removes the source msPKI-Cert-Template-OID from the JSON file. This is safe only
    when the import generates a new OID, that is, with -OidHandling Generate, GenerateFromRoot, or
    GenerateRandom. With the default -OidHandling Preserve, the import needs that OID and fails
    with an error when the OID was stripped.

.PARAMETER NewTemplateName
    Import and Sync. The new internal name (cn) for the template in the target forest. The script
    allows letters (non-ASCII included), digits, spaces that are not at an edge, and the
    characters period (.), underscore (_), hyphen (-), and parentheses. The script rejects the
    characters that have a meaning in a DN (, + = " \ ; < >), the LDAP wildcard (*), and the
    characters / and #. Parentheses are permitted; the script escapes them where a name reaches an
    LDAP filter. If you omit it, the script uses the name of the source.

.PARAMETER NewDisplayName
    Import and Sync. The new display name for the template in the target forest. If you omit it,
    the script uses the displayName of the source.

.PARAMETER OidHandling
    Import and Sync. How the script chooses the OID of the template. Every value also registers a
    companion msPKI-Enterprise-Oid "display" object when the OID container exists, so Windows
    resolves the OID to the template name.
      * Preserve (default): keeps the OID of the source template from the file. Needs no PKI OID
        root, so it works in a forest that never had AD CS.
      * Generate: generates a new OID under the REAL enterprise OID root of the target forest.
        Requires that AD CS was provisioned in the target forest at least once.
      * GenerateFromRoot: generates a new OID under the base OID you pass in -OidRoot. No AD CS is
        needed. Use the same root across imports to give those templates a shared, stable base.
      * GenerateRandom: generates a new OID under a freshly synthesized, forest-independent base.
        No AD CS is needed, and no input is needed.

.PARAMETER OidRoot
    Import and Sync. Required with -OidHandling GenerateFromRoot: the base OID under which the
    script generates the template OID, for example
    "1.3.6.1.4.1.311.21.8.100000001.100000002.100000003.100000004.100000005". The script rejects
    -OidRoot with any other -OidHandling value before it does anything else, because it would
    otherwise ignore the value without a message. The script also refuses -OidHandling
    GenerateFromRoot without -OidRoot at that point, before it contacts a domain controller.

.PARAMETER Server
    The optional domain controller (DC) to target for Configuration partition operations. On Sync
    this is the TARGET side; the source side is -SourceServer. If you omit it on Import, Sync, or
    Validate, the script discovers a writable DC in the CURRENT forest. The script then uses that
    DC consistently for the write and for the follow-up ACL step. Point -Server at a DC in another
    forest, with -Credential as needed, to operate there instead. -Server is required whenever you
    give -Credential, so the script is guaranteed to use the credentials against the forest you
    intend.

.PARAMETER Credential
    The optional credentials the script uses against -Server. That is the TARGET side for Import,
    Sync, and Validate, and the SOURCE side for Export, which has no target side. The credentials
    cover every operation the mode performs: Configuration partition reads and writes, principal
    lookups, and the ACL write, all over ADWS. Combined with -Server, -Credential lets Export read
    from a forest you are not logged on to, with no trust required. In the same way it lets Import
    and Sync write to such a forest. Requires -Server (see -Server above).

.PARAMETER SourceServer
    Sync only, and required there. A domain controller, or a domain name, in the SOURCE forest from
    which the script reads the template.

.PARAMETER SourceCredential
    Sync only. The optional credentials the script uses against -SourceServer. Omit it to read as
    the current identity. That works across a trust, or when you run inside the source forest
    itself.

.PARAMETER SkipAcl
    Import and Sync. Skips the permission setup after the import. It is mutually exclusive with
    -EnrollPrincipals and with an explicit -AclBase.

    CAUTION: the template keeps the schema-default DACL. Admins and SYSTEM get Full Control, and
    no principal gets Enroll until you apply permissions yourself.

.PARAMETER AclBase
    Import and Sync. The base from which the script builds the template ACL. The script always
    adds -EnrollPrincipals, when given, on top of the base. Default: Standard, the KERBEROS
    AUTHENTICATION set, which is oriented to domain controllers (DCs). NOTE: when that set applies
    only by default to a template whose name does not look like a Kerberos Authentication copy,
    the script warns. The warning makes the DC-oriented grants a conscious choice, not an accident.
      * Standard: the standard Kerberos Authentication set of the script (see -Mode Import above).
        It REPLACES the schema-default ACL of Active Directory, so admins are not left with Full
        Control.
      * Schema: leaves the schema-default ACL of Active Directory as created, and only adds
        -EnrollPrincipals to it. The schema default is: Domain Admins and Enterprise Admins Full
        Control, SYSTEM Full Control, Authenticated Users Read.
      * SchemaPlusStandard: keeps the schema-default ACL AND adds the Standard set on top. The
        script removes nothing.
      * PrincipalsOnly: no base. The ACL is exactly your -EnrollPrincipals, which is then required,
        and it replaces the schema default.

.PARAMETER EnrollPrincipals
    Import and Sync. A hashtable that maps each principal to the rights the script grants to it.
    The script ADDS these grants on top of the -AclBase base. With -AclBase PrincipalsOnly they
    are the sole content of the ACL. The script validates the hashtable and resolves every
    principal first, before it creates anything.

    Keys: a SID (S-1-5-...), a sAMAccountName, a UPN (user@domain), or a well-known token. A
    sAMAccountName or a UPN may have a DOMAIN\ prefix; the prefix must name the target domain. The
    well-known tokens are DomainControllers, DomainComputers, DomainUsers, DomainAdmins,
    EnterpriseAdmins, EnterpriseRODCs, EnterpriseDomainControllers, AuthenticatedUsers, and
    Everyone. The script looks up named principals in the target (-Server) domain. Use a SID for a
    principal in another domain.

    Resolution fails closed. The script refuses a bare string that matches BOTH a well-known token
    AND a directory object with a DIFFERENT SID. This way an account placed by an attacker cannot
    hijack a token, and a token cannot shadow a distinct real group. Disambiguate with a SID or
    with a DOMAIN\ prefix. The names of the built-in groups themselves, for example
    'Domain Admins', resolve normally, because both readings yield the same SID. The script
    refuses a name that matches more than one object (duplicate UPNs) rather than guess.

    A key that contains '@' (user@domain) resolves ONLY as a UPN. A sAMAccountName may legally
    contain '@', so the script refuses a different object that has that string as its
    sAMAccountName. It treats that object as an account placed to capture the grant, or as a
    colliding account.

    Values: one or more of Read, Write, Enroll, Autoenroll, FullControl. Example:
        -AclBase PrincipalsOnly -EnrollPrincipals @{
            'DomainControllers'    = 'Enroll','Autoenroll'
            'AuthenticatedUsers'   = 'Read'
            'NOREFJELL\PKI-Admins' = 'FullControl'
        }

.PARAMETER UpgradeCompatibility
    Import and Sync. Raises the imported template to the newest compatibility that the Certificate
    Templates MMC (Microsoft Management Console) snap-in offers. The script applies the upgrade as
    it creates the template in the target forest. That compatibility is Certification Authority:
    Windows Server 2016, and Certificate recipient: Windows 10 / Windows Server 2016. It means
    schema version 4 plus the matching private-key-flag bits.

    Only schema v2 and v3 templates can be upgraded in place. The script imports a schema v1
    template unchanged, with a warning, because v1 built-in templates are read-only in the MMC. It
    leaves a template already at v4 as-is. The script does not modify the source template or the
    export; it upgrades only the copy it writes to the target.

    The script also imports a schema v2 source that has msPKI-RA-Application-Policies at its
    existing compatibility, with a warning. The encoding of that attribute differs at v3 and v4. An
    upgrade in place would drop the registration authority (RA) signature application-policy
    requirement.

    The script sets the legacy-provider bit (CT_FLAG_USE_LEGACY_PROVIDER, 0x100) only for a
    schema-2 source with a provider list. The reason: v2 knows only CryptoAPI cryptographic service
    providers (CSPs). A schema-3 source keeps its own bit, so a v3 template that lists a key
    storage provider (KSP) stays on Cryptography Next Generation (CNG).

.PARAMETER AllowLinkedIssuancePolicy
    Import and Sync. By default the script REFUSES the import when the template has an issuance
    policy OID that the TARGET forest already links to a group. Such a link is an Authentication
    Mechanism Assurance (AMA) link, stored in msDS-OIDToGroupLink. The script scans
    msPKI-Certificate-Policy, which holds the issuance policies stamped into the ISSUED
    certificate.

    Certificates issued from the copy would grant the membership of that group at logon to every
    principal the enrollment ACL of the copy admits. No link is ever copied. Pass this switch to
    accept such a mapping deliberately; the script then lists the linked OIDs and groups in a
    warning.

    CAUTION: certificates from the copy would give the membership of the linked group at logon to
    everyone the enrollment ACL of the copy admits.

    The script also validates every known attribute of the import for type, shape, and range
    before it creates anything. It refuses a malformed or tampered export, for example a
    non-integer msPKI-RA-Signature, a three-byte validity period, or a non-OID application policy.
    It never drops or coerces such an attribute without an error.

.PARAMETER KeepArtifacts
    Validate only. Leaves the throwaway templates and the export file in place after the
    comparison. By default the script removes them. The script never creates a companion OID
    object for the throwaway templates, because they use a self-contained explicit OID. So there
    is no companion object to keep.

.EXAMPLE
    # Source forest - export the built-in Kerberos Authentication template:
    .\Sync-ADCSTemplate.ps1 -Mode Export -Path .\KerberosAuth.json

.EXAMPLE
    # Source forest - export a custom template as a name-neutral copy. The identity is stripped.
    # The OID stays in the file, so the default -OidHandling Preserve works on import. Add
    # -StripOid only when the import will generate a new OID with one of the Generate values:
    .\Sync-ADCSTemplate.ps1 -Mode Export -TemplateName "XX-KerberosAuthentication" `
        -Path .\XX.json -StripIdentity

.EXAMPLE
    # Target forest with NO AD CS (the default case) - import under a new name and keep the
    # source OID:
    .\Sync-ADCSTemplate.ps1 -Mode Import -Path .\XX.json `
        -NewTemplateName "YY-KerberosAuthentication" -NewDisplayName "YY-Kerberos Authentication"

.EXAMPLE
    # Target forest that HAS its own PKI - generate a new target-forest OID instead of keeping
    # the source OID:
    .\Sync-ADCSTemplate.ps1 -Mode Import -Path .\XX.json -OidHandling Generate `
        -NewTemplateName "YY-KerberosAuthentication" -NewDisplayName "YY-Kerberos Authentication"

.EXAMPLE
    # No AD CS in the target, but you want a new synthetic OID that Windows resolves to the name:
    .\Sync-ADCSTemplate.ps1 -Mode Import -Path .\XX.json -OidHandling GenerateRandom `
        -NewTemplateName "YY-KerberosAuthentication" -NewDisplayName "YY-Kerberos Authentication"

.EXAMPLE
    # Import and raise the copy to the latest compatibility (Windows Server 2016 / Windows 10):
    .\Sync-ADCSTemplate.ps1 -Mode Import -Path .\Workstation.json -OidHandling GenerateRandom -UpgradeCompatibility

.EXAMPLE
    .\Sync-ADCSTemplate.ps1 -Mode Import -Path .\XX.json -NewTemplateName "YY-KerberosAuthentication" -NewDisplayName "YY-Kerberos Authentication" -WhatIf

.EXAMPLE
    # The standard Kerberos Authentication ACL (default) PLUS a template-admin group that EJBCA
    # will read:
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
    # Direct sync with no file - run in the TARGET forest and read from the source forest over
    # the trust:
    .\Sync-ADCSTemplate.ps1 -Mode Sync -SourceServer dc01.source.example `
        -TemplateName "XX-KerberosAuthentication" `
        -NewTemplateName "YY-KerberosAuthentication" -NewDisplayName "YY-Kerberos Authentication"

.EXAMPLE
    # Direct sync from a third machine, explicit credentials on both sides (no trust needed):
    .\Sync-ADCSTemplate.ps1 -Mode Sync `
        -SourceServer dc01.a.example -SourceCredential (Get-Credential A\template.reader) `
        -Server dc01.b.example -Credential (Get-Credential B\ent.admin)

.EXAMPLE
    # Mixed flow: import a JSON file exported earlier straight into another forest, with no logon
    # there:
    .\Sync-ADCSTemplate.ps1 -Mode Import -Path .\KerberosAuth.json `
        -Server dc01.b.example -Credential (Get-Credential B\ent.admin)

.EXAMPLE
    # Prove that both copy pipelines (file and direct/Sync) preserve every functional attribute.
    # This creates and removes one throwaway copy per pipeline:
    .\Sync-ADCSTemplate.ps1 -Mode Validate -TemplateName "KerberosAuthentication"

.NOTES
    - The script requires the ActiveDirectory PowerShell module, from the Remote Server
      Administration Tools (RSAT), for all modes. The script no longer uses or requires certutil or
      the AD CS role. No CA needs to be reachable.
    - ALL directory access, including the ACL write, runs over Active Directory Web Services
      (ADWS, TCP 9389). No LDAP (389) connectivity is needed. The script always uses the -Server
      and -SourceServer values VERBATIM; it never substitutes an endpoint the operator did not
      type. Name ONE DC. A DOMAIN name (DNS or NetBIOS) locates a different DC per connection, so
      the create, read-back, and ACL steps could reach different replicas. The script detects that
      case and warns; a lagging read-back fails the run rather than leave the template mis-secured.
    - The script rejects a parameter that the mode does not consume before it does anything else,
      rather than ignore it without a message. Examples: -StripOid with -Mode Import, -OidRoot
      without -OidHandling GenerateFromRoot, and -AclBase with -SkipAcl. The script also refuses
      -OidHandling GenerateFromRoot without -OidRoot at that point.
    - The script types a PKI attribute that its built-in type lists do not know from the schema of
      the TARGET forest automatically. Such an attribute is a genuine schema extension linked to
      the pKICertificateTemplate class. The script drops the attribute with a warning when the
      target schema cannot type it. It also drops the attribute when the target schema does not
      permit it on the template class. It also drops the attribute when the target schema types it
      single-valued while the source value is empty or multi-valued. Each of these cases is a
      schema divergence between the forests.
    - Note: v3 and v4 Cryptography Next Generation (CNG) algorithm settings are NOT separate
      directory attributes. They are packed inside msPKI-RA-Application-Policies, which the script
      copies as-is.
    - A v3 or v4 template can embed a private-key Security Descriptor Definition Language (SDDL)
      string (msPKI-Key-Security-Descriptor) packed inside msPKI-RA-Application-Policies. The
      script copies it verbatim. Any domain SIDs in it are SOURCE-forest SIDs. The script warns, so
      that you can review the key ACL in the target forest.
    - -Mode Validate exits non-zero when any attribute differs, so continuous integration (CI) and
      automation can use its exit code to decide.
    - With the default -OidHandling Preserve, the target forest does not need AD CS now, and it
      never needed AD CS. It needs only the Certificate Templates container. -OidHandling Generate
      needs the PKI OID root of the target forest, which exists after AD CS has been deployed
      there once.
    - Run Export in the source forest and Import in the target forest. Or point -Server, with
      -Credential as needed, at a DC in the relevant forest. -Mode Sync does both sides in one run.
      It needs ADWS reachability to a DC in EACH forest from the machine where it runs. The script
      refuses to proceed when the source and the target resolve to the same forest, unless you
      gave -Server explicitly. This guards against a forgotten -Server that targets the source
      forest without a message.
    - Import and Sync do NOT publish the template to any CA. Publish and issue it from the CA
      afterwards.
    - The script intentionally does not copy the security descriptor across forests. Import and
      Sync reapply standard permissions instead, unless you pass -SkipAcl.
    - The script does NOT copy Authentication Mechanism Assurance (AMA) links. An
      msDS-OIDToGroupLink on an issuance policy OID of the source forest points at a group DN in
      THAT forest, and cannot be copied. If you use AMA, recreate the link in the target forest
      manually: policy OID object -> a local universal group with no static members.
    - The script DOES check the reverse case. When the copy has an issuance policy OID that the
      TARGET forest already links to a group, the script refuses the import. Pass
      -AllowLinkedIssuancePolicy to accept it. The reason for the refusal: certificates from the
      copy would grant the membership of that group at logon.
    - A v1 template copies too, and the export warns. The object round-trips faithfully, but
      Windows fixes the v1 semantics in code. The v1 consumers match by NAME, the definition is not
      editable, and it never autoenrolls. Import a v1 template under its ORIGINAL name, because a
      renamed copy is invisible to Windows v1 consumers. For a non-Windows consumer such as EJBCA,
      which reads the object and its ACL directly, a v1 copy works like any other template. To get
      an editable template that autoenrolls, duplicate it as v2 or later in the source forest first.
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

    [switch]$AllowLinkedIssuancePolicy,

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
# Multi-value attributes whose every element must be a dotted OID (the others - pKIDefaultCSPs,
# msPKI-RA-Application-Policies, msPKI-Supersede-Templates - carry provider strings, packed
# name/value text and template names).
$script:OidListAttributes = @('msPKI-Certificate-Application-Policy', 'msPKI-Certificate-Policy', 'msPKI-RA-Policies', 'pKIExtendedKeyUsage', 'pKICriticalExtensions')

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
    #                                    plus 0x100 (CT_FLAG_USE_LEGACY_PROVIDER) for a schema-2
    #                                    source with a provider list (v2 knows only CryptoAPI CSPs);
    #                                    a schema-3 source keeps its own 0x100 bit as-is (its list
    #                                    may legitimately name KSPs)
    #   flags                          IS_DEFAULT (0x10000) -> IS_MODIFIED (0x20000)
    #   msPKI-Template-Minor-Revision  += 1
    param([Parameter(Mandatory)][hashtable]$Attributes)

    $ver = if ($Attributes.ContainsKey('msPKI-Template-Schema-Version')) { [int]$Attributes['msPKI-Template-Schema-Version'] } else { 1 }
    if ($ver -lt 2) { return [pscustomobject]@{ Upgraded = $false; FromVersion = $ver; Reason = 'schema v1 template - not upgradable in place' } }
    if ($ver -ge 4) { return [pscustomobject]@{ Upgraded = $false; FromVersion = $ver; Reason = 'already at the latest compatibility (schema v4)' } }

    # A schema-2 source carrying msPKI-RA-Application-Policies cannot be upgraded in place: that
    # attribute's ENCODING is schema-version dependent - a bare list of required application-policy
    # OIDs at v1/v2, but a packed `name`type`value` string at v3/v4 (this script's own .NOTES states
    # it, and the msPKI-Key-Security-Descriptor handling depends on the same packed form). Stamping
    # v4 while leaving the v2 encoding makes a v4-aware CA parse the OID list as packed triples, find
    # no entry, and stop enforcing the application-policy constraint on the required RA signature -
    # msPKI-RA-Signature still demands a co-signature, but now from any enrollment agent rather than
    # one holding the named policy. Refuse (the caller warns and imports at stock compatibility)
    # rather than silently weaken the co-signing requirement; re-encode the attribute by hand if a
    # v4 copy is genuinely needed.
    if ($ver -eq 2 -and $Attributes.ContainsKey('msPKI-RA-Application-Policies') -and
        $null -ne $Attributes['msPKI-RA-Application-Policies'] -and @($Attributes['msPKI-RA-Application-Policies']).Count -gt 0) {
        return [pscustomobject]@{ Upgraded = $false; FromVersion = $ver; Reason = 'schema v2 source carries msPKI-RA-Application-Policies, whose encoding differs at v3/v4; upgrading in place would drop the RA-signature application-policy requirement - imported at its stock compatibility instead' }
    }

    $pkf = if ($Attributes.ContainsKey('msPKI-Private-Key-Flag')) { [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$Attributes['msPKI-Private-Key-Flag']), 0) } else { [uint32]0 }
    # The two compatibility levels are VERSION NIBBLES, not independent bits: CA min-version at
    # bits 16-19, recipient min-version at bits 24-27. REPLACE them (clear both nibbles, then set
    # 6/6 = Windows Server 2016 / Windows 10) - a plain -bor would ACCUMULATE, so a source already
    # at e.g. 2008R2/Win7 (nibble 3) would become 3|6 = 7, an invalid level. 0xF0F0FFFF clears only
    # the two version nibbles and preserves every other flag (0x100 legacy provider, 0x200000, ...).
    $pkf = ($pkf -band 0xF0F0FFFF) -bor 0x06060000
    # @($null).Count is 1, so test the value explicitly - do NOT rely on the count alone.
    $isCsp = $Attributes.ContainsKey('pKIDefaultCSPs') -and $null -ne $Attributes['pKIDefaultCSPs'] -and @($Attributes['pKIDefaultCSPs']).Count -gt 0
    # CT_FLAG_USE_LEGACY_PROVIDER is derived from the SOURCE's own semantics, never from the mere
    # presence of a provider list: a schema-2 (Server 2003) template can only name CryptoAPI CSPs,
    # so a populated list there means legacy; a schema-3 template's list may name KSPs ("Microsoft
    # Software Key Storage Provider") and its own 0x100 bit already says which - that bit survived
    # the mask above and is left exactly as the source had it. Setting 0x100 on a KSP template
    # would switch the v4 copy to legacy (CryptoAPI) key handling.
    if ($ver -eq 2 -and $isCsp) { $pkf = $pkf -bor 0x100 }
    $legacy = [bool]($pkf -band 0x100)

    $flags = if ($Attributes.ContainsKey('flags')) { [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$Attributes['flags']), 0) } else { [uint32]0 }
    $flags = ($flags -band (-bnot 0x10000)) -bor 0x20000

    # Compute EVERY replacement value (including the minor-revision increment) and validate it
    # BEFORE mutating $Attributes, so a failure leaves the hashtable untouched. The old order set
    # schema/flags first and incremented the revision last: a source revision of Int32.MaxValue
    # overflowed the `[System.Int32](... + 1)` cast, which under the default error preference left
    # the template half-upgraded (v4 schema/flags, un-bumped revision) and still returned Upgraded.
    $newSchema = [System.Int32]4
    $newPkf    = [BitConverter]::ToInt32([BitConverter]::GetBytes([uint32]$pkf), 0)
    $newFlags  = [BitConverter]::ToInt32([BitConverter]::GetBytes([uint32]$flags), 0)
    $newRev    = $null
    if ($Attributes.ContainsKey('msPKI-Template-Minor-Revision')) {
        $curRev = [int]$Attributes['msPKI-Template-Minor-Revision']
        if ($curRev -eq [int]::MaxValue) {
            throw "Import refused: msPKI-Template-Minor-Revision is already Int32.MaxValue ($curRev) and cannot be incremented for the compatibility upgrade - the source object or export is corrupt or tampered with."
        }
        $newRev = [System.Int32]($curRev + 1)
    }
    $Attributes['msPKI-Template-Schema-Version'] = $newSchema
    $Attributes['msPKI-Private-Key-Flag']        = $newPkf
    $Attributes['flags']                         = $newFlags
    if ($null -ne $newRev) { $Attributes['msPKI-Template-Minor-Revision'] = $newRev }
    return [pscustomobject]@{ Upgraded = $true; FromVersion = $ver; PrivateKeyFlag = ('0x{0:X8}' -f $pkf); LegacyProvider = $legacy }
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
                -Properties 'msPKI-Cert-Template-OID' -ErrorAction Stop).'msPKI-Cert-Template-OID'
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
        if ($ExplicitOid -notmatch '^(0|[1-9]\d*)(\.(0|[1-9]\d*))+\z') {
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
            if ($OidRoot -notmatch '^(0|[1-9]\d*)(\.(0|[1-9]\d*))+\z') {
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
            if ($SourceOid -notmatch '^(0|[1-9]\d*)(\.(0|[1-9]\d*))+\z') {
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

function ConvertTo-SchemaTypedValue {
    # Converts the value of a PKI attribute the static lists do not know into the .NET value
    # New-ADObject must receive, from the type Get-SchemaAttributeType read in the TARGET schema.
    # Returns $null when the value cannot be copied: an unknown type, a single-valued type (Int,
    # String) given zero or several elements, or a failed cast (a schema-divergent shape, e.g. a
    # string where the target expects octets). The caller then drops the attribute with a warning,
    # never silently. Int and String take exactly ONE element: a multi-element value would be
    # corrupted (space-joined) or crash the cast, and an EMPTY array would fabricate a value
    # ([int]$null -> 0). @(...)[0] also unwraps a one-element array, which [System.Int32] alone
    # would not. MultiString comes back as an object[] of strings and the caller casts it to the AD
    # collection type, so this function needs no AD module and can be unit-tested.
    param([string]$SchemaType, [AllowNull()]$Value)
    $arr = @($Value)
    try {
        switch ($SchemaType) {
            'Int'         { if ($arr.Count -eq 1) { return [System.Int32]$arr[0] }; return $null }
            'String'      { if ($arr.Count -eq 1) { return [string]$arr[0] }; return $null }
            'MultiString' { return , [object[]]@($arr | ForEach-Object { [string]$_ }) }
            'Bytes'       { return , [System.Byte[]]$Value }
            default       { return $null }
        }
    }
    catch {
        return $null
    }
}

function ConvertTo-ImportAttributeValue {
    # Converts ONE known template attribute from the import view (JSON export or live AD read) to
    # the exact value New-ADObject must receive, and REFUSES anything malformed with a terminating
    # error that names the attribute. The casts used to run bare: under the default error
    # preference a failed [int] cast is only statement-terminating, so a corrupted (or tampered)
    # export silently DROPPED the attribute - msPKI-RA-Signature among them, a CA-enforced control -
    # and the template was created without it. JSON also arrives in shapes a bare cast quietly
    # coerces: 0.4 -> 0, a three-element period array -> a three-byte "period", "5" -> 5. Every
    # known attribute is therefore checked for exact type, shape and range, and the whole input
    # is converted before anything is created. Returns: [int] for the integer attributes, [byte[]]
    # for the period/key-usage attributes, [object[]] of strings for the multi-value attributes
    # (the caller casts that to the AD collection type; this function needs no AD module).
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][AllowNull()][AllowEmptyCollection()][AllowEmptyString()]$Value)
    $fail = { param($why) throw "Import refused: attribute '$Name' is malformed ($why). The export is corrupt or was tampered with - re-export it from the source, or remove the attribute deliberately." }
    if ($null -eq $Value) { & $fail 'no value' }
    if ($Name -in $script:IntAttributes) {
        $v = $Value
        if ($v -is [System.Collections.IEnumerable] -and $v -isnot [string]) {
            $arr = @($v)
            if ($arr.Count -ne 1) { & $fail "expected one integer, got $($arr.Count) values" }
            $v = $arr[0]
        }
        if ($null -eq $v)   { & $fail 'no value' }
        if ($v -is [string]) { & $fail 'expected an integer, got a string' }
        if ($v -is [bool])   { & $fail 'expected an integer, got a boolean' }
        if ($v -is [double] -or $v -is [single]) {
            # Finiteness, integrality AND range are checked in DOUBLE space BEFORE any [decimal] cast:
            # a tiny double (1e-30) cast to decimal UNDERFLOWS to 0, which would pass a post-cast
            # integer check and SILENTLY COERCE the malformed value to 0 - exactly the coercion this
            # validation exists to prevent. A giant/non-finite double (1e40, Infinity, NaN) is refused
            # here too (it is not integral or is out of Int32 range), with the attribute named.
            $dbl = [double]$v
            if ([double]::IsNaN($dbl) -or [double]::IsInfinity($dbl)) { & $fail "expected an integer, got $v" }
            if ([math]::Truncate($dbl) -ne $dbl) { & $fail "expected an integer, got $v" }
            if ($dbl -lt [int]::MinValue -or $dbl -gt [int]::MaxValue) { & $fail "value $v is outside the Int32 range" }
            return [System.Int32]$dbl
        }
        if ($v -is [decimal]) {
            # decimal is exact within its range, so no underflow: check integrality and range directly.
            if ([math]::Truncate($v) -ne $v) { & $fail "expected an integer, got $v" }
            if ($v -lt [int]::MinValue -or $v -gt [int]::MaxValue) { & $fail "value $v is outside the Int32 range" }
            return [System.Int32]$v
        }
        if ($v -isnot [byte] -and $v -isnot [sbyte] -and $v -isnot [int16] -and $v -isnot [uint16] -and
                $v -isnot [int] -and $v -isnot [uint32] -and $v -isnot [long] -and $v -isnot [uint64]) {
            & $fail "expected an integer, got $($v.GetType().Name)"
        }
        # An integer type: [decimal] is exact (no underflow); only the Int32 range remains to check.
        $d = [decimal]$v
        if ($d -lt [int]::MinValue -or $d -gt [int]::MaxValue) { & $fail "value $v is outside the Int32 range" }
        return [System.Int32]$d
    }
    if ($Name -in $script:ByteAttributes) {
        if ($Value -is [string]) { & $fail 'expected a byte array, got a string' }
        $arr = @($Value)
        $expect = if ($Name -eq 'pKIKeyUsage') { @(1, 2) } else { @(8) }
        if ($arr.Count -notin $expect) { & $fail "expected $($expect -join ' or ') byte(s), got $($arr.Count)" }
        $bytes = New-Object byte[] $arr.Count
        for ($i = 0; $i -lt $arr.Count; $i++) {
            $b = $arr[$i]
            if ($null -eq $b -or $b -is [string] -or $b -is [bool]) { & $fail "element $i is not a byte" }
            if ($b -is [double] -or $b -is [single]) {
                # As in the integer branch: check integrality in DOUBLE space, since [decimal]1e-30
                # underflows to 0 and would silently coerce a malformed element to a valid byte.
                $db = [double]$b
                if ([double]::IsNaN($db) -or [double]::IsInfinity($db) -or [math]::Truncate($db) -ne $db -or $db -lt 0 -or $db -gt 255) { & $fail "element $i ($b) is not a byte (0-255)" }
                $bytes[$i] = [byte]$db
            }
            else {
                $d = try { [decimal]$b } catch { $null }
                if ($null -eq $d -or [math]::Truncate($d) -ne $d -or $d -lt 0 -or $d -gt 255) { & $fail "element $i ($b) is not a byte (0-255)" }
                $bytes[$i] = [byte]$d
            }
        }
        return , $bytes
    }
    if ($Name -in $script:MultiValueAttributes) {
        $arr = if ($Value -is [string]) { @($Value) } else { @($Value) }
        $out = New-Object System.Collections.Generic.List[string]
        foreach ($e in $arr) {
            if ($null -eq $e) { & $fail 'a $null element' }
            if ($e -isnot [string]) { & $fail "element '$e' is $($e.GetType().Name), expected a string" }
            if ($e.Length -eq 0) { & $fail 'an empty element' }
            if ($e -match '[\x00-\x1F\x7F]') { & $fail 'an element contains control characters' }
            if ($Name -in $script:OidListAttributes -and $e -notmatch '^\d+(\.\d+)+\z') { & $fail "'$e' is not a dotted OID" }
            $out.Add($e)
        }
        if (-not $out.Count) { & $fail 'no values' }
        return , [object[]]$out.ToArray()
    }
    $Value
}

function Get-LinkedIssuancePolicy {
    # ISSUANCE-policy OIDs stamped into the issued certificate (msPKI-Certificate-Policy) that the
    # TARGET forest already binds to a group through Authentication Mechanism Assurance
    # (msDS-OIDToGroupLink on the OID object). A copy carrying such an OID issues certificates that
    # grant that group's membership at logon - to everyone its enrollment ACL admits - with no link
    # ever being copied: the link already exists here. An export the operator did not author, or
    # one an attacker edited, can carry exactly such an OID. Returns one record per linked OID:
    # @{ Oid; OidObjectDN; GroupDN }.
    # ONLY msPKI-Certificate-Policy is scanned: those OIDs go into the Certificate Policies extension
    # of the ISSUED certificate, which is what AMA maps to a group at logon. msPKI-RA-Policies is
    # deliberately EXCLUDED - it constrains the enrollment-agent SIGNING certificate (the co-signer
    # must hold those application policies), it is not stamped into the issued certificate, so an AMA
    # link on such an OID does not grant the enrollee anything.
    param([hashtable]$Attributes, [string]$ConfigNC, [hashtable]$ADParams)
    $oids = @()
    foreach ($a in 'msPKI-Certificate-Policy') {
        if ($Attributes.ContainsKey($a) -and $null -ne $Attributes[$a]) { $oids += @($Attributes[$a] | ForEach-Object { "$_" }) }
    }
    $oids = @($oids | Where-Object { $_ } | Sort-Object -Unique)
    if (-not $oids.Count) { return @() }
    $oidContainerDN = "CN=OID,CN=Public Key Services,CN=Services,$ConfigNC"
    if (-not (Get-ADObjectIfPresent -Identity $oidContainerDN -ADParams $ADParams)) { return @() }   # no OID container: no links can exist
    $clauses = ($oids | ForEach-Object { "(msPKI-Cert-Template-OID=$(ConvertTo-LdapFilterValue $_))" }) -join ''
    # -ErrorAction Stop: "no links" may only mean that a SUCCESSFUL search found none.
    @(Get-ADObject @ADParams -SearchBase $oidContainerDN -ErrorAction Stop `
        -LDAPFilter "(&(objectClass=msPKI-Enterprise-Oid)(msDS-OIDToGroupLink=*)(|$clauses))" `
        -Properties 'msPKI-Cert-Template-OID', 'msDS-OIDToGroupLink' |
        ForEach-Object { [pscustomobject]@{ Oid = "$($_.'msPKI-Cert-Template-OID')"; OidObjectDN = $_.DistinguishedName; GroupDN = "$($_.'msDS-OIDToGroupLink')" } })
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
        [switch]$AllowLinkedIssuancePolicy,
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
        if ($name -in $script:IntAttributes -or $name -in $script:ByteAttributes) {
            # Validated and converted, or the import is REFUSED - never dropped or coerced (see
            # ConvertTo-ImportAttributeValue).
            $oa[$name] = ConvertTo-ImportAttributeValue -Name $name -Value $import.$name
        }
        elseif ($name -in $script:MultiValueAttributes) {
            $oa[$name] = [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection](ConvertTo-ImportAttributeValue -Name $name -Value $import.$name)
        }
        elseif ($name -match '^(msPKI-|pKI)') {
            # A PKI attribute the static lists don't know (a schema extension linked to the template
            # class): type it from the TARGET forest's schema instead of dropping it. Get-SchemaAttributeType
            # returns $null when the attribute is not permitted on the template class or has a type
            # this script cannot round-trip; ConvertTo-SchemaTypedValue returns $null for that, for a
            # single-valued type given zero or several elements, and for a failed cast - then the
            # attribute is dropped (surfaced below), never silently. MultiString arrives as an
            # object[] of strings; the try/catch turns a failure of the collection cast into the
            # same drop-with-warning instead of a raw terminating cast error.
            $schemaType = Get-SchemaAttributeType -AttributeName $name -ConfigNC $configNC -ADParams $ADParams
            $typed = ConvertTo-SchemaTypedValue -SchemaType $schemaType -Value $import.$name
            if ($null -eq $typed) {
                $unconsumed += $name
            }
            else {
                # Direct assignments: an `if` used as an expression would enumerate a byte[] (or the
                # AD collection) into an object[] on its way into the hashtable.
                try {
                    if ($schemaType -eq 'MultiString') { $oa[$name] = [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection]$typed }
                    else { $oa[$name] = $typed }
                }
                catch {
                    $oa.Remove($name)
                    $unconsumed += $name
                }
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

    # Pre-flight: issuance policies the copy carries must not be ones THIS forest already binds to
    # a group through Authentication Mechanism Assurance - only the template OID was collision-checked
    # so far, while a policy OID rides along verbatim. Skipped for Validate's throwaway copies
    # (-ExplicitOid): they live in the SOURCE forest, where the source template carries the very same
    # policy legitimately. -AllowLinkedIssuancePolicy accepts the mapping deliberately.
    if (-not $ExplicitOid) {
        $linked = @(Get-LinkedIssuancePolicy -Attributes $oa -ConfigNC $configNC -ADParams $ADParams)
        if ($linked.Count) {
            $list = ($linked | ForEach-Object { "$($_.Oid) -> $($_.GroupDN)" }) -join '; '
            if ($AllowLinkedIssuancePolicy) {
                Write-Warning "The template carries issuance policy OID(s) that THIS forest links to a group via Authentication Mechanism Assurance ($list). Certificates issued from the copy will grant that group's membership at logon to every principal its enrollment ACL admits (-AllowLinkedIssuancePolicy given; proceeding)."
            }
            else {
                throw "Import refused: the template carries issuance policy OID(s) that THIS forest already links to a group via Authentication Mechanism Assurance: $list. Certificates issued from the copy would grant that group's membership at logon to every principal its enrollment ACL admits - a security decision, not a copy detail. Strip the OID from the source's msPKI-Certificate-Policy, or rerun with -AllowLinkedIssuancePolicy to accept the mapping deliberately."
            }
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
        # -PassThru: the returned object carries the DN AD assigned and the objectGUID. The line
        # printed below reports both, so a caller can identify the very object this run created
        # without a later lookup by name or OID (a later lookup could find a replacement).
        $companionObj = New-ADObject @ADParams -Path $oidPlan.CompanionContainerDN -Name $oidPlan.CompanionCn `
            -Type 'msPKI-Enterprise-Oid' -OtherAttributes $oidObjectAttrs -Confirm:$false -PassThru -ErrorAction Stop
        $companionDN = "CN=$($oidPlan.CompanionCn),$($oidPlan.CompanionContainerDN)"
        if ($companionObj -and $companionObj.DistinguishedName) { $companionDN = $companionObj.DistinguishedName }
        $companionGuidText = if ($companionObj -and $companionObj.ObjectGUID) { " (objectGUID $($companionObj.ObjectGUID))" } else { '' }
        Write-Host "Created OID object: $companionDN$companionGuidText" -ForegroundColor Green
    }

    # Create the template object itself, referencing the resolved OID. -PassThru captures the DN that
    # AD actually assigned, so the ACL step and cleanup always bind the real object rather than a
    # string-built DN that could diverge under RDN escaping.
    $oa['msPKI-Cert-Template-OID'] = $oidPlan.Oid
    $createdObj = $null
    try {
        # -ErrorAction Stop: AD-module failures are frequently NON-terminating. Without it a failed
        # create (constraint violation, permissions, ADWS hiccup) would leave $createdObj $null,
        # skip this catch and the companion rollback, print a green "Created template:" line with
        # an empty DN and return $null - which the caller reads as "declined at the prompt": exit
        # code 0 and an orphaned companion OID object. The returned object is checked as well.
        $createdObj = New-ADObject @ADParams -Path $templatesDN -Name $cn -DisplayName $displayName `
            -Type 'pKICertificateTemplate' -OtherAttributes $oa -Confirm:$false -PassThru -ErrorAction Stop
        if (-not $createdObj -or -not $createdObj.DistinguishedName) {
            throw "New-ADObject returned no object for template '$cn' - the create did not complete."
        }
        $newTemplateDN = $createdObj.DistinguishedName

        # Post-create uniqueness. The pre-flight OID search above is check-then-create: a
        # concurrent import (or any other writer) could have created another template with this
        # OID in between, and AD does not enforce msPKI-Cert-Template-OID uniqueness. Re-query on
        # the same pinned server and fail if anything but this object carries the OID, rather
        # than leave two templates sharing one OID (ambiguous authorization lookups in Windows and
        # EJBCA). The catch below rolls the new template back for this - and for a failed re-query.
        $carriers = @(Get-ADObject @ADParams -SearchBase $templatesDN -ErrorAction Stop `
                -LDAPFilter "(msPKI-Cert-Template-OID=$($oidPlan.Oid))" |
            Where-Object { $_.ObjectGUID -ne $createdObj.ObjectGUID })
        if ($carriers.Count) {
            throw "OID $($oidPlan.Oid) was claimed concurrently by another template ($(@($carriers | ForEach-Object { $_.DistinguishedName }) -join '; ')) - refusing to leave two templates sharing one OID."
        }
    }
    catch {
        $failure = $_
        # Rollback, in dependency order. Whatever failed AFTER the create (the uniqueness re-query,
        # a collision) must not leave the new template behind - without its ACL, and blocking a
        # re-run through the cn pre-flight - so it is removed by the GUID captured at creation.
        # The companion OID object is removed only once the template is gone (or was never
        # created): a template that could not be removed keeps its companion, and both survivors
        # are reported for manual cleanup instead of half of a pair being deleted.
        $templateGone = $true
        if ($createdObj) {
            try {
                Remove-ADObject @ADParams -Identity $createdObj.ObjectGUID -Confirm:$false -ErrorAction Stop
                Write-Warning "Import failed after the template was created; rolled back the new template '$cn'."
            }
            catch {
                $templateGone = $false
                Write-Warning "Import failed after the template was created AND the new template could not be removed ($($_.Exception.Message)). It remains WITHOUT its ACL - clean up manually: $newTemplateDN"
            }
        }
        if ($companionDN) {
            if ($templateGone) {
                try {
                    Remove-ADObject @ADParams -Identity $companionDN -Confirm:$false -ErrorAction Stop
                    Write-Warning "Rolled back the companion OID object."
                }
                catch {
                    Write-Warning "The companion OID object could not be removed. Clean up manually: $companionDN"
                }
            }
            else {
                Write-Warning "The companion OID object was KEPT because the template it belongs to still exists: $companionDN"
            }
        }
        throw $failure
    }

    # The objectGUID comes from the object New-ADObject -PassThru returned, never from a re-read.
    Write-Host "Created template: $newTemplateDN (objectGUID $($createdObj.ObjectGUID))" -ForegroundColor Green
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
    if ($name -match '@') {
        # ANY value whose NAME PART (after an optional DOMAIN\ prefix, already validated to name the
        # target domain) contains '@' is a UPN and is resolved ONLY as one - classifying on $name, not
        # the raw $id, so the documented DOMAIN\user@domain form takes this branch too (matching on $id
        # let a 'DOMAIN\'-prefixed key skip straight to the sAMAccountName lookup). A stricter shape
        # test would route an odd-but-real UPN such as 'ann lee@x.test' back to sAMAccountName, and
        # with it around the shadow check. sAMAccountName may legally contain '@', so an attacker with
        # account-creation rights could plant a principal whose sAMAccountName equals the victim's UPN;
        # consulting sAMAccountName first (as before) would have handed that principal the grant. The
        # UPN lookup is the authority, and a sAMAccountName match for the same string that is a
        # DIFFERENT object is refused rather than guessed.
        $escUpn = ConvertTo-LdapFilterValue $name
        $obj = @(Get-ADObject @ADParams -LDAPFilter "(userPrincipalName=$escUpn)" -Properties objectSid -ErrorAction Stop |
                Where-Object { $_.objectSid })
        $shadow = @(Get-ADObject @ADParams -LDAPFilter "(sAMAccountName=$escName)" -Properties objectSid -ErrorAction Stop |
                Where-Object { $_.objectSid })
        if ($shadow.Count) {
            $upnSids = @($obj | ForEach-Object { ([System.Security.Principal.SecurityIdentifier]$_.objectSid).Value })
            $other = @($shadow | Where-Object { ([System.Security.Principal.SecurityIdentifier]$_.objectSid).Value -notin $upnSids })
            if ($other.Count) {
                throw "Principal '$Identity' is a UPN, but a different directory object carries it as its sAMAccountName ($(@($other | ForEach-Object { $_.DistinguishedName }) -join '; ')) - a planted or colliding account. Use the intended principal's SID (S-1-5-...) instead."
            }
        }
    }
    else {
        $obj = @(Get-ADObject @ADParams -LDAPFilter "(sAMAccountName=$escName)" -Properties objectSid -ErrorAction Stop |
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
        # A throw, not a warning + $null: $null means "declined at the prompt" to the caller, which
        # then returns normally - and a deployment gate on the exit code would accept a copy that
        # was never validated. The throwaway object (if it exists) is removed by the caller's finally.
        throw "$PipelineName throwaway template '$($CreatedDN.Value)' was created but could not be read back (replication lag on a domain-name -Server? pass a single DC). Validation did NOT run."
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
    Import   = 'Path', 'NewTemplateName', 'NewDisplayName', 'OidHandling', 'OidRoot', 'SkipAcl', 'AclBase', 'EnrollPrincipals', 'UpgradeCompatibility', 'AllowLinkedIssuancePolicy', 'Server', 'Credential'
    Sync     = 'TemplateName', 'NewTemplateName', 'NewDisplayName', 'OidHandling', 'OidRoot', 'SkipAcl', 'AclBase', 'EnrollPrincipals', 'UpgradeCompatibility', 'AllowLinkedIssuancePolicy', 'Server', 'Credential', 'SourceServer', 'SourceCredential'
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
# The mirror case, refused here as well so it fails before a DC is contacted or a grant resolved.
# Resolve-TemplateOid keeps its own check as a second line for the internal callers.
if ($OidHandling -eq 'GenerateFromRoot' -and -not $OidRoot) {
    throw "OidHandling 'GenerateFromRoot' requires -OidRoot (the base OID to generate under, e.g. 1.3.6.1.4.1.311.21.8.<5 arcs>)."
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
            AllowLinkedIssuancePolicy = $AllowLinkedIssuancePolicy
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
