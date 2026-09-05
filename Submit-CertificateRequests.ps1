<#PSScriptInfo
.VERSION 1.0.3
.GUID 6f98f16e-0c56-4a72-ba31-443938175c06
.AUTHOR Sveinung Svea
.PROJECTURI https://github.com/TheOmnilord/ADCS
.LICENSEURI https://github.com/TheOmnilord/ADCS/blob/main/LICENSE
.TAGS ADCS PKI CertificateServices
.RELEASENOTES
1.0.3 - The delivery folder is also judged on what its ACL hands to the FILES created inside it: an inheritable "files" entry (inherit-only or not) granting an untrusted principal write, append, delete, write-attributes or re-permission rights is refused (a warning under -AllowUnprotectedOutputFolder; -TrustedOutputPrincipal applies), because certreq's staging file and the delivered certificate inherit it while the folder-swap checks rightly ignore inherit-only entries - previously such a user could alter the certificate's bytes before or after delivery with every folder check passing. CREATOR OWNER / OWNER RIGHTS placeholders resolve to the running account and are trusted; CREATOR GROUP is not
1.0.2 - A retrieval that reports Issued without producing the .cer file is recorded as Error (retried next Retrieve) instead of a final Issued; certreq exit code + .cer presence now decide Issued independently of certreq's (localized) wording; certificates are written to a private staging file inside the destination folder and delivered with a no-overwrite rename (an existing file is moved aside as <name>.superseded-<stamp>.cer, never deleted); every folder from the destination to the volume root is checked for reparse points and untrusted owners/writers before delivery (new -AllowUnprotectedOutputFolder and -TrustedOutputPrincipal); a tracking row's OutputCertFile must resolve beneath the tracking folder or -OutputFolder; an issued certificate that cannot be delivered gets the new Undelivered status (never resubmitted); values that reach the certreq command line are validated; concurrent runs on one tracking file are refused (exclusive <TrackingFile>.lock) and the tracking file is replaced via a unique temp file; cmdlets inside an approved action no longer raise their own -Confirm prompts
1.0.1 - Retrieve mode honours an explicit -OutputFolder (previously the path recorded at submit time was always used)
1.0.0 - Initial release
#>

<#
.SYNOPSIS
    Batch submission and retrieval of certificates via ADCS (certreq.exe).

.DESCRIPTION
    Submits all .req/.csr/.txt files from a folder to an ADCS CA,
    tracks request IDs in a CSV file, and can retrieve issued certificates
    later based on stored request IDs.

.PARAMETER InputPath
    Folder containing .req/.csr/.txt request files for submission.
    Required when -Mode is Submit or Both. Not used, and not required, for -Mode Retrieve.

.PARAMETER CAConfig
    CA configuration string for certreq, e.g. "CA01.domain.com\Contoso Issuing CA 1".
    Always required.

.PARAMETER CertificateTemplate
    Certificate template name used for submission.
    Required when -Mode is Submit or Both. Not used, and not required, for -Mode Retrieve.

.PARAMETER TrackingFile
    Path to the CSV file that tracks request IDs and statuses.
    Relative paths are resolved against the current directory and stored as absolute paths.

.PARAMETER OutputFolder
    Folder where issued certificates (.cer) are saved.
    Relative paths are resolved against the current directory and stored as absolute paths.
    In Retrieve mode each row is written to the path recorded at submit time; pass -OutputFolder
    explicitly to redirect retrieved .cer files there instead (the tracking row is updated).

.PARAMETER Mode
    Submit   = Submit new certificate requests.
    Retrieve = Retrieve issued certificates for unresolved requests.
    Both     = Run Submit, then Retrieve.

.PARAMETER KeepRspFile
    By default, the .rsp file created next to each .cer (at both submit and retrieve) is deleted.
    Specify -KeepRspFile to leave it in place.

.PARAMETER Force
    Resubmit request files that already have a tracked RequestID without prompting.
    Without -Force, the script asks y/n for each already-submitted file (default = No / skip).

.PARAMETER AllowUnprotectedOutputFolder
    By default the script refuses to deliver certificates into (or through) a folder that an
    untrusted principal owns, or can delete, rename or write to - trusted being SYSTEM,
    Administrators, TrustedInstaller, the running account, its Domain/Enterprise Admins and
    -TrustedOutputPrincipal - because such a user could swap the folder for a junction (or, with
    mere write-data / write-attributes rights, turn an empty folder into one in place) while a
    certificate is being delivered and redirect the privileged write. The delivery folder is also
    refused when its ACL would hand an untrusted principal write, append, delete, write-attributes
    or re-permission rights on the FILES created inside it (an inheritable "files" entry,
    inherit-only or not): the staging file certreq writes and the delivered certificate inherit
    such a grant. CREATOR OWNER entries resolve to the running account and are trusted; a CREATOR
    GROUP entry resolves to its primary group and is not (name S-1-3-1 in -TrustedOutputPrincipal
    to accept it). Pass this
    switch to accept those risks (e.g. a shared drop folder whose ACL cannot be tightened); the
    conditions are then only warned about.

.PARAMETER TrustedOutputPrincipal
    Additional principals (SIDs or account names, e.g. 'CONTOSO\PKI-Operators') that may own, or
    hold delete/rename/write rights on, the folders certificates are delivered through (including
    file-inheritable write rights in the delivery folder itself), on top of
    the built-in trusted set (see -AllowUnprotectedOutputFolder). Use it when the output folders
    are managed by a dedicated operator group rather than by Administrators.

.EXAMPLE
    .\Submit-CertificateRequests.ps1 -InputPath "C:\CSRs" `
        -CAConfig "CA01.domain.com\Contoso Issuing CA 1" `
        -CertificateTemplate "WebServer" -Mode Submit

.EXAMPLE
    .\Submit-CertificateRequests.ps1 -CAConfig "CA01.domain.com\Contoso Issuing CA 1" -Mode Retrieve

.EXAMPLE
    .\Submit-CertificateRequests.ps1 -InputPath "C:\CSRs" `
        -CAConfig "CA01.domain.com\Contoso Issuing CA 1" `
        -CertificateTemplate "WebServer" -Mode Submit -WhatIf

.NOTES
    - Status detection: whether a request was issued is decided by certreq's exit code plus the
      presence of the certificate it wrote (into a fresh temp file, so the file is unambiguously
      this run's output), which is language-independent. On success the certificate is delivered
      to the destination; a file already there (a -Force resubmit; a retry of an unresolved row)
      is moved aside as <name>.superseded-<UTC stamp>.cer - never deleted - and that copy is
      removed again only when the fresh certificate is byte-identical. The RequestID and the finer
      Pending/Denied dispositions are parsed from certreq's console text, which Windows localizes;
      on a non-English CA/client those may parse as Unknown (a retrieval is then retried on the
      next -Mode Retrieve, never silently finalized) and a missing RequestID is reported so the
      row can be completed by hand from the CA database.
    - One run per tracking file: the script holds an exclusive lock file (<TrackingFile>.lock,
      removed when the run ends) for its whole run and refuses to start while another run - on
      this or any other machine, via any alias of the path - holds it. The tracking file's name is
      canonicalized first (an 8.3 short name resolves to the long name, so both spellings share
      one lock) and a hard-linked or symlinked tracking file is refused (a second name elsewhere
      could not share the lock). A -WhatIf run takes no lock (it never writes the CSV).
    - certreq always writes into a private temp file; the destination is touched only after a
      successful write (a pending, denied or failed request never disturbs it). A row's
      OutputCertFile is an identifier inside an operator-chosen boundary, not an authority: it
      must be a rooted .cer path in an existing folder beneath the tracking file's folder or the
      run's -OutputFolder (canonically, with no junction/symlink in between, re-checked right
      before delivery; the destination itself must not be a folder or a link), and RequestID must
      be numeric; rows failing this are skipped with an error. If an issued certificate cannot be
      delivered (locked destination, denied rename), the row is recorded as Status 'Undelivered'
      with its RequestID kept and the certificate left in its staging file beside the destination (path in ErrorMessage): it counts
      as submitted on later runs (never resubmitted automatically) and -Mode Retrieve re-fetches
      it; an Undelivered row with no RequestID is reported for manual reconciliation. Values that
      reach the certreq command line (-CAConfig, -CertificateTemplate, RequestID, paths) must not
      contain double quotes or control characters.
    - Output folders must not be swappable by untrusted users: every folder from the destination up
      to the volume/share root is checked, and the script REFUSES to deliver when any of them is a
      reparse point (junction/symlink/mount point), is owned by an untrusted principal, can be
      deleted, renamed or written to by one (delete/rename lets such a user replace it with a
      junction between the path check and the privileged delivery; write-data or write-attributes
      access on a directory handle is all FSCTL_SET_REPARSE_POINT needs, so "create files" rights
      let them turn an EMPTY folder - e.g. a freshly created output folder - into a junction in
      place), or has a security descriptor that cannot be read. Trusted principals are SYSTEM,
      BUILTIN\Administrators, TrustedInstaller, the running account, its Domain Admins /
      Enterprise Admins, and anything named in -TrustedOutputPrincipal;
      -AllowUnprotectedOutputFolder downgrades the refusals to warnings. Only the create-SUBFOLDER
      right by itself (what the C:\ root grants Users on itself) and inherit-only ACEs are
      tolerated: neither can set a reparse point on the folder, a pre-planted junction is refused by
      the reparse-point checks, an attacker-created subfolder is refused by the owner check, and a
      folder or junction planted under the exact destination name in the moment between check and
      delivery receives no file, because every
      delivery is a no-overwrite rename (File.Move) that fails when anything occupies the name
      instead of moving the file into it. The .rsp companion written with -KeepRspFile and the
      tracking-file replacement follow the same rule.
#>

# NOTE: '#Requires' deliberately sits AFTER the help comment - placed before it, Get-Help
# fails to bind the comment-based help and shows only auto-generated syntax (verified).
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [string]$InputPath,

    [Parameter(Mandatory)]
    [string]$CAConfig,

    [string]$CertificateTemplate,

    [string]$TrackingFile = ".\CertTracking.csv",

    [string]$OutputFolder = ".\Certificates",

    [ValidateSet("Submit", "Retrieve", "Both")]
    [string]$Mode = "Submit",

    [switch]$KeepRspFile,

    [switch]$Force,

    [switch]$AllowUnprotectedOutputFolder,

    [string[]]$TrustedOutputPrincipal
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Functions

function Resolve-FullPath {
    param([Parameter(Mandatory)][string]$Path)

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }
    return [System.IO.Path]::GetFullPath((Join-Path (Get-Location).ProviderPath $Path))
}

function Resolve-TrackingFilePath {
    # The lock that serializes runs is a sibling file named after the tracking file, so every name
    # for the same CSV must produce the same lock name. A path that already resolves absolutely is
    # not enough: an 8.3 alias (CERTIF~1.CSV) names the same file under a different string. When
    # the file exists its canonical long name is taken from the directory entry (Get-Item), and a
    # hard link or symlink is REFUSED - those give the same data a second name in another place,
    # which no sibling lock can cover.
    param([Parameter(Mandatory)][string]$Path)
    $full = Resolve-FullPath -Path $Path
    if (-not (Test-Path -LiteralPath $full -PathType Leaf)) { return $full }
    $item = Get-Item -LiteralPath $full -Force
    if ($item.LinkType) {
        throw "-TrackingFile '$full' is a $($item.LinkType): the same tracking data would be reachable under another name that a concurrent run could lock separately. Point -TrackingFile at a regular file."
    }
    $item.FullName
}

function Assert-SafeNativeArgument {
    # Start-Process -ArgumentList joins its elements into ONE command-line string, so a double
    # quote inside a value would close the quoting this script wraps around it and let the rest
    # of the value be parsed as further certreq switches (argument injection); control characters
    # never belong in these values either. Checked up front for every value that reaches the
    # certreq command line - including fields read back from the tracking CSV, which may have
    # been edited or imported from elsewhere.
    param(
        [Parameter(Mandatory)][string]$Name,
        [AllowEmptyString()][string]$Value
    )
    if ($Value -match '["\x00-\x1F]') {
        throw "$Name must not contain double quotes or control characters (it is passed to certreq.exe on its command line): '$Value'"
    }
}

function Assert-CertificateOutputPath {
    # The .cer path of a tracking row is where a privileged process delivers a certificate and
    # moves a predecessor aside - and the row comes from an editable CSV. The path is therefore
    # treated as an identifier that must stay INSIDE an operator-chosen boundary, not as an
    # authority of its own: it must be rooted, end in .cer, and resolve (canonically, after
    # '..' normalization) beneath one of -AllowedRoots - the tracking file's own folder or the
    # run's -OutputFolder - with no reparse point (junction/symlink) between that root and the
    # file, so the boundary cannot be escaped through a redirecting folder. The folder must
    # already exist, and the value must also be safe on the certreq command line.
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string[]]$AllowedRoots
    )
    Assert-SafeNativeArgument -Name $Name -Value $Path
    if (-not [System.IO.Path]::IsPathRooted($Path)) {
        throw "$Name must be a rooted (absolute) path: '$Path'"
    }
    $full = [System.IO.Path]::GetFullPath($Path)
    if ([System.IO.Path]::GetExtension($full) -ne '.cer') {
        throw "$Name must be a .cer file: '$Path'"
    }
    $dir = [System.IO.Path]::GetDirectoryName($full)
    if (-not $dir -or -not (Test-Path -LiteralPath $dir -PathType Container)) {
        throw "$Name points into a folder that does not exist: '$Path' (pass -OutputFolder to redirect retrieved files)"
    }
    # The destination itself must be a plain file or absent: Move-Item onto an existing FOLDER
    # named x.cer would move the certificate INTO it (a junction there would carry it anywhere),
    # and a file symlink would deliver through to wherever it points.
    if (Test-Path -LiteralPath $full) {
        $existing = Get-Item -LiteralPath $full -Force
        if ($existing.PSIsContainer) {
            throw "$Name is an existing folder, not a file: '$Path'"
        }
        if ($existing.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
            throw "$Name is a reparse point (symlink); refusing to deliver through it: '$Path'"
        }
    }
    # The DEEPEST matching root wins, so the walk below covers as few components as possible (the
    # default layout .\Certificates beside .\CertTracking.csv matches -OutputFolder itself, not
    # the tracking folder above it).
    $root = $null
    foreach ($r in $AllowedRoots) {
        $canon = [System.IO.Path]::GetFullPath($r).TrimEnd('\', '/')
        if (($dir.Equals($canon, [System.StringComparison]::OrdinalIgnoreCase) -or
             $dir.StartsWith($canon + '\', [System.StringComparison]::OrdinalIgnoreCase)) -and
            (-not $root -or $canon.Length -gt $root.Length)) { $root = $canon }
    }
    if (-not $root) {
        throw "$Name is outside the allowed output locations ($($AllowedRoots -join '; ')): '$Path'. Pass -OutputFolder to redirect retrieved files into a folder you choose."
    }
    # Walk from the destination folder all the way up to the volume (or share) root - NOT just to
    # the allowed root: a junction or a swappable folder ABOVE the allowed root redirects everything
    # beneath it just the same. Every component must be a real folder (a junction/symlink/mount
    # point anywhere on the way would redirect the write while the textual path still looks
    # contained), and none may be swappable by untrusted users: a broad principal holding
    # delete/rename-class rights on a component could replace it with a junction between this
    # check and the privileged move. That window cannot be closed from PowerShell, so such a path
    # is REFUSED (fail closed) - as is one whose ACL cannot be read, since protection cannot be
    # established then - unless -AllowUnprotectedOutputFolder was passed, which turns these
    # refusals into warnings. Only the create-SUBFOLDER right (FILE_APPEND_DATA, what the C:\ root
    # grants Users on itself) is tolerated: it cannot set a reparse point; what it could pre-plant
    # is caught by the owner, reparse-point and destination-shape checks. The delivery folder is
    # additionally judged on what its ACL hands to the FILES created inside it (inheritable
    # entries, inherit-only or not): certreq's staging file and the delivered certificate inherit
    # them, so an untrusted write/append/delete grant there is refused as well.
    Assert-ProtectedDirectoryChain -Name $Name -Directory $dir -Path $Path
}

function Assert-ProtectedDirectoryChain {
    # See Assert-CertificateOutputPath: every folder from -Directory up to the volume/share root
    # must be a real, non-swappable folder with a readable ACL. Refusals become warnings under
    # -AllowUnprotectedOutputFolder ($script:AllowUnprotectedOutput).
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Directory,
        [Parameter(Mandatory)][string]$Path
    )
    $probe = $Directory
    while ($probe) {
        $item = Get-Item -LiteralPath $probe -Force
        if ($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
            $msg = "$Name passes through a reparse point (junction/symlink/mount point) at '$probe', which could redirect the write outside the allowed location: '$Path'."
            if ($script:AllowUnprotectedOutput) { Write-BatchLog "$msg (-AllowUnprotectedOutputFolder: proceeding anyway)" -Level Warning }
            else { throw "$msg Use a real folder, or pass -AllowUnprotectedOutputFolder to accept the risk." }
        }
        $swappers = $null; $badOwner = $null; $fileWriters = @()
        try {
            $swappers = @(Get-UntrustedGrant -Path $probe -Kind Swap)
            $badOwner = Get-UntrustedOwner -Path $probe
            # The delivery folder itself is also judged on what its ACL hands to the FILES created
            # inside it: the staging file and the delivered certificate inherit those entries
            # (inherit-only or not), which the folder-swap check above rightly ignores.
            if ($probe -eq $Directory) { $fileWriters = @(Get-UntrustedGrant -Path $probe -Kind File) }
        }
        catch {
            $msg = "$Name lies under '$probe', whose security descriptor could not be read ($($_.Exception.Message)), so it cannot be established that untrusted users are unable to swap it during a delivery."
            if ($script:AllowUnprotectedOutput) { Write-BatchLog "$msg (-AllowUnprotectedOutputFolder: proceeding anyway)" -Level Warning }
            else { throw "$msg Fix the folder's ACL/permissions, or pass -AllowUnprotectedOutputFolder to accept the risk." }
        }
        if ($badOwner) {
            $msg = "$Name lies under '$probe', which is OWNED by untrusted principal $badOwner (an owner can always re-permission and replace a folder) - it could be swapped for a junction while a certificate is being delivered."
            if ($script:AllowUnprotectedOutput) { Write-BatchLog "$msg (-AllowUnprotectedOutputFolder: proceeding anyway)" -Level Warning }
            else { throw "$msg Deliver only through folders owned by trusted principals (SYSTEM, Administrators, the running account, its Domain/Enterprise Admins, or -TrustedOutputPrincipal), or pass -AllowUnprotectedOutputFolder to accept the risk." }
        }
        if ($swappers -and $swappers.Count) {
            $msg = "$Name lies under '$probe', which untrusted principal(s) $($swappers -join ', ') can delete, rename or write to (write-data/write-attributes rights are enough to turn an empty folder into a junction in place) - such a user could redirect it while a certificate is being delivered."
            if ($script:AllowUnprotectedOutput) { Write-BatchLog "$msg (-AllowUnprotectedOutputFolder: proceeding anyway)" -Level Warning }
            else { throw "$msg Restrict that folder's ACL to trusted principals (name additional ones with -TrustedOutputPrincipal), or pass -AllowUnprotectedOutputFolder to accept the risk." }
        }
        if ($fileWriters -and $fileWriters.Count) {
            $msg = "$Name would be written in '$probe', whose ACL grants untrusted principal(s) $($fileWriters -join ', ') write, append, delete, write-attributes or re-permission rights on the FILES created inside it (an inheritable 'files' entry) - the staging file certreq writes and the delivered certificate inherit that grant, so such a user could alter the certificate's bytes, or turn the delivered file into a reparse point, before or after delivery."
            if ($script:AllowUnprotectedOutput) { Write-BatchLog "$msg (-AllowUnprotectedOutputFolder: proceeding anyway)" -Level Warning }
            else { throw "$msg Remove that entry from the folder's ACL (or name the principal with -TrustedOutputPrincipal), or pass -AllowUnprotectedOutputFolder to accept the risk." }
        }
        $probe = [System.IO.Path]::GetDirectoryName($probe)   # $null at the volume root (C:\) or the share root (\\server\share)
    }
}

function New-TempCertificatePath {
    # certreq writes into a private STAGING file first (see Move-RetrievedCertificate): the
    # destination named by the tracking row is touched only after the write is known to have
    # succeeded. The staging file lives INSIDE the destination's own folder - which the caller has
    # already verified as protected - never in %TEMP%: a TEMP whose inherited file ACL lets other
    # accounts modify files would let them tamper with the certificate before delivery, and a
    # same-volume rename keeps the file's own DACL, so that loose ACL would travel onto the
    # delivered .cer inside the protected folder. Staged beside the destination, the file inherits
    # the protected folder's ACL from the start and the delivery is a same-folder rename. The name
    # is random (cannot be pre-planted) and ends in .cer so certreq's .rsp lands beside it.
    param([Parameter(Mandatory)][string]$Destination)
    $dir  = [System.IO.Path]::GetDirectoryName($Destination)
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($Destination)
    Join-Path $dir ('{0}.{1}.staging.cer' -f $stem, [guid]::NewGuid().ToString('N'))
}

function Get-TrustedPrincipalSet {
    # The principals allowed to OWN - or to hold delete/rename rights on - the folders certificates
    # are delivered through: SYSTEM, BUILTIN\Administrators, TrustedInstaller (owner of the volume
    # root and of Windows' own folders), the account running this script and, when that is a domain
    # account, its domain's Domain Admins (RID 512) and Enterprise Admins (RID 519), plus whatever
    # -TrustedOutputPrincipal names (SIDs or account names). EVERYONE ELSE is untrusted - a fixed
    # list of "broad" groups would miss a Modify grant to Domain Users, or a folder an attacker
    # created (and therefore owns) beneath a root that tolerates create rights. Returns SID -> label.
    param([string[]]$Extra)
    $set = @{
        'S-1-5-18'     = 'SYSTEM'
        'S-1-5-32-544' = 'BUILTIN\Administrators'
        'S-1-5-80-956008885-3418522649-1831038044-1853292631-2271478464' = 'TrustedInstaller'
    }
    $me = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $set[$me.User.Value] = $me.Name
    if ($me.User.AccountDomainSid) {
        $dom = $me.User.AccountDomainSid.Value
        $set["$dom-512"] = 'Domain Admins'
        $set["$dom-519"] = 'Enterprise Admins'
    }
    foreach ($p in @($Extra | Where-Object { $_ })) {
        $sid = if ($p -match '^S-1-\d+(-\d+)+$') { New-Object System.Security.Principal.SecurityIdentifier($p) }
               else { (New-Object System.Security.Principal.NTAccount($p)).Translate([System.Security.Principal.SecurityIdentifier]) }
        $set[$sid.Value] = $p
    }
    $set
}

function ConvertTo-PrincipalLabel {
    param([Parameter(Mandatory)][string]$Sid)
    try { "$((New-Object System.Security.Principal.SecurityIdentifier($Sid)).Translate([System.Security.Principal.NTAccount]).Value) ($Sid)" }
    catch { $Sid }
}

function Get-UntrustedGrant {
    # Labels of every principal NOT in $script:TrustedSids (see Get-TrustedPrincipalSet) that holds
    # rights of the given kind on a folder:
    #   Swap    rights that let a folder on the delivery path be REPLACED BY or TURNED INTO a
    #           junction between the path check and the privileged move: delete/rename-class
    #           (Delete, DeleteSubdirectoriesAndFiles, ChangePermissions, TakeOwnership,
    #           GENERIC_ALL) AND write-data / write-attributes class (CreateFiles = FILE_WRITE_DATA,
    #           WriteAttributes, GENERIC_WRITE). The latter matter because FSCTL_SET_REPARSE_POINT
    #           needs only a directory handle opened with FILE_WRITE_DATA or FILE_WRITE_ATTRIBUTES,
    #           and a mount-point junction needs no privilege - so a user with mere "create files"
    #           rights can convert an EMPTY existing folder (a freshly created output folder) into
    #           a junction in place, with nothing deleted or renamed. Output folders must not grant
    #           any of these to untrusted principals; Assert-ProtectedDirectoryChain refuses otherwise.
    #   Create  CreateDirectories (FILE_APPEND_DATA) alone - the right the C:\ root grants Users
    #           on itself. It cannot set a reparse point or remove anything; what it can plant
    #           (a new subfolder) is caught by the owner and reparse-point checks and by the
    #           no-overwrite delivery rename, so it is only reported.
    #   File    what the folder's ACL hands to the FILES created inside it - judged on the DELIVERY
    #           folder only. Every ACE carrying ObjectInherit is considered, inherit-only or not
    #           (an inherit-only "files only" entry is exactly what the Swap check must ignore),
    #           against the file write-class: write-data, append-data, delete, change permissions,
    #           take ownership, write-attributes, GENERIC_WRITE/ALL. certreq's staging file and the
    #           delivered certificate inherit such an entry, so an untrusted principal holding one
    #           could alter or replace the certificate's bytes once certreq has closed the file -
    #           before delivery or after it - with every folder on the chain fully protected.
    #           Write-attributes is included because FSCTL_SET_REPARSE_POINT accepts a FILE handle
    #           opened with FILE_WRITE_ATTRIBUTES alone, which would let such a user turn the
    #           delivered certificate into a reparse point after the plain-file post-condition.
    #           The inheritance placeholders are resolved the way the file system resolves them:
    #           CREATOR OWNER (S-1-3-0) and OWNER RIGHTS (S-1-3-4) become the creating account -
    #           the trusted running account, since certreq creates the file - so they are skipped
    #           (the C:\ root hands every unprotected folder an inherit-only CREATOR OWNER Full
    #           Control entry); CREATOR GROUP (S-1-3-1) becomes the running account's primary
    #           group, which can be as broad as Domain Users, so it stays untrusted unless named
    #           in -TrustedOutputPrincipal.
    # Generic access bits are included because generic ACEs carry them. For Swap and Create,
    # inherit-only ACEs (the "subfolders and files only" entries the C:\ root carries) do not apply
    # to the folder itself and are skipped - the folders they propagate to are judged on their own
    # ACLs; File looks at precisely those entries. An ACL that cannot be read is a terminating
    # error: the caller must treat "unknown" as unprotected.
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][ValidateSet('Swap', 'Create', 'File')][string]$Kind
    )
    $mask = if ($Kind -eq 'Swap') {
        ([int64][System.Security.AccessControl.FileSystemRights]'Delete, DeleteSubdirectoriesAndFiles, ChangePermissions, TakeOwnership, CreateFiles, WriteAttributes') -bor 0x10000000 -bor 0x40000000   # + GENERIC_ALL, GENERIC_WRITE
    } elseif ($Kind -eq 'File') {
        # CreateFiles = FILE_WRITE_DATA and CreateDirectories = FILE_APPEND_DATA on a file
        ([int64][System.Security.AccessControl.FileSystemRights]'CreateFiles, CreateDirectories, Delete, ChangePermissions, TakeOwnership, WriteAttributes') -bor 0x10000000 -bor 0x40000000   # + GENERIC_ALL, GENERIC_WRITE
    } else {
        [int64][System.Security.AccessControl.FileSystemRights]'CreateDirectories'
    }
    $found = @()
    # Get-Acl, not DirectoryInfo.GetAccessControl(): on PowerShell 7 the latter is an extension
    # method (System.IO.FileSystem.AccessControl) that a PowerShell method call cannot reach.
    # -ErrorAction Stop: a failed read must reach the caller, never read as "no grants".
    $acl = Get-Acl -LiteralPath $Path -ErrorAction Stop
    if (-not $acl) { throw "Get-Acl returned no security descriptor for '$Path'." }
    foreach ($rule in $acl.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier])) {
        if ($rule.AccessControlType -ne [System.Security.AccessControl.AccessControlType]::Allow) { continue }
        if ($Kind -eq 'File') {
            # Only entries that propagate to files matter; a folders-only (ContainerInherit) entry
            # is never inherited by the staging file or the delivered certificate.
            if (-not ($rule.InheritanceFlags -band [System.Security.AccessControl.InheritanceFlags]::ObjectInherit)) { continue }
        }
        elseif ($rule.PropagationFlags -band [System.Security.AccessControl.PropagationFlags]::InheritOnly) { continue }
        $sid = $rule.IdentityReference.Value
        if ($script:TrustedSids.ContainsKey($sid)) { continue }
        # Inheritance placeholders (only meaningful in inheritable entries): CREATOR OWNER and
        # OWNER RIGHTS resolve to the account that creates the file - the running account, which
        # is trusted by construction. CREATOR GROUP resolves to that account's primary group and
        # is deliberately left untrusted (fail closed); its label says what it becomes.
        if ($sid -eq 'S-1-3-0' -or $sid -eq 'S-1-3-4') { continue }
        $rights = ([int64][int]$rule.FileSystemRights) -band 0xFFFFFFFF
        if ($rights -band $mask) {
            $found += if ($sid -eq 'S-1-3-1') { "CREATOR GROUP ($sid, resolves to the running account's primary group)" } else { ConvertTo-PrincipalLabel -Sid $sid }
        }
    }
    @($found | Sort-Object -Unique)
}

function Get-UntrustedOwner {
    # The label of a folder's owner when that owner is NOT a trusted principal, else $null. An
    # owner can always re-permission the object (WRITE_DAC is implicit for owners) and so rename
    # or replace it - an attacker-created subfolder is owned by the attacker whatever its ACEs say.
    param([Parameter(Mandatory)][string]$Path)
    $acl = Get-Acl -LiteralPath $Path -ErrorAction Stop
    $owner = if ($acl) { $acl.GetOwner([System.Security.Principal.SecurityIdentifier]) } else { $null }
    if (-not $owner) { throw "The owner of '$Path' could not be read." }
    if ($script:TrustedSids.ContainsKey($owner.Value)) { return $null }
    ConvertTo-PrincipalLabel -Sid $owner.Value
}

function Move-RetrievedCertificate {
    # Deliver a certificate certreq wrote into the temp file to its destination - only now, with
    # the write known to have succeeded, is anything at the destination touched: an existing file
    # there is moved aside (never deleted; dropped again only if identical), then the fresh file
    # takes its place. With -KeepRspFile the .rsp is placed beside the destination as before;
    # otherwise it is removed with the temp file. A pending, denied or failed request never
    # reaches this function, so it never disturbs the destination. With -AllowedRoots the
    # destination is re-validated immediately before it is touched, so the window between the
    # earlier check and this privileged write is as narrow as a pathname check allows (the
    # residual window is why Assert-CertificateOutputPath refuses folders that untrusted users
    # could swap). The .rsp companion gets the same destination-shape checks as the .cer.
    param(
        [Parameter(Mandatory)][string]$TempCer,
        [Parameter(Mandatory)][string]$Destination,
        [switch]$KeepRspFile,
        [string[]]$AllowedRoots
    )
    if ($AllowedRoots) { Assert-CertificateOutputPath -Name 'the delivery destination' -Path $Destination -AllowedRoots $AllowedRoots }
    $aside = Move-StaleCertificateAside -Path $Destination
    # [System.IO.File]::Move (MoveFileEx WITHOUT MOVEFILE_REPLACE_EXISTING), not Move-Item: the
    # destination was just cleared, so the rename must find NOTHING there. Move-Item -Force would
    # happily move the certificate INTO a folder - or a junction - that an untrusted user with
    # create rights planted under this exact name after the check above; File.Move fails instead
    # (the certificate stays in its staging file, the row becomes Undelivered, and the planted
    # junction receives no file). This is what closes the create-only race the ACL checks
    # deliberately tolerate.
    [System.IO.File]::Move($TempCer, $Destination)
    # Post-condition: the certificate must now be a plain file at exactly this path.
    if (-not (Test-Path -LiteralPath $Destination -PathType Leaf) -or
        ((Get-Item -LiteralPath $Destination -Force).Attributes -band [System.IO.FileAttributes]::ReparsePoint)) {
        throw "Delivery did not produce a plain file at '$Destination'."
    }
    # The delivered file must be governed by the folder just verified as protected and by nothing
    # else: a rename keeps a file's own DACL, so drop any explicit ACE the staging file carries and
    # re-enable inheritance from the destination folder.
    try {
        $fileAcl = Get-Acl -LiteralPath $Destination
        $fileAcl.SetAccessRuleProtection($false, $false)
        foreach ($rule in @($fileAcl.Access | Where-Object { -not $_.IsInherited })) { [void]$fileAcl.RemoveAccessRule($rule) }
        Set-Acl -LiteralPath $Destination -AclObject $fileAcl -Confirm:$false
    }
    catch { Write-BatchLog "  Could not reset the delivered certificate's ACL to inherit from its folder ($($_.Exception.Message)); review the file's permissions." -Level Warning }
    Remove-AsideIfIdentical -Path $Destination -Aside $aside
    $tmpRsp = [System.IO.Path]::ChangeExtension($TempCer, '.rsp')
    if ($KeepRspFile -and (Test-Path -LiteralPath $tmpRsp -PathType Leaf)) {
        # Same folder as the (validated) certificate; the .rsp entry itself must be a plain file
        # or absent - a folder or link named x.rsp would carry the response file elsewhere.
        $rspDest = [System.IO.Path]::ChangeExtension($Destination, '.rsp')
        $rspExisting = if (Test-Path -LiteralPath $rspDest) { Get-Item -LiteralPath $rspDest -Force } else { $null }
        if ($rspExisting -and ($rspExisting.PSIsContainer -or ($rspExisting.Attributes -band [System.IO.FileAttributes]::ReparsePoint))) {
            Write-BatchLog "  .rsp NOT placed beside the certificate: '$rspDest' is a folder or a link. The response file was discarded." -Level Warning
        }
        else {
            try {
                # Same no-overwrite rename as the certificate (a stale plain .rsp is removed first;
                # anything else appearing under that name makes the rename fail, never redirect).
                if ($rspExisting) { Remove-Item -LiteralPath $rspDest -Force -Confirm:$false }
                [System.IO.File]::Move($tmpRsp, $rspDest)
                if (-not (Test-Path -LiteralPath $rspDest -PathType Leaf) -or
                    ((Get-Item -LiteralPath $rspDest -Force).Attributes -band [System.IO.FileAttributes]::ReparsePoint)) {
                    Write-BatchLog "  .rsp placement did not produce a plain file at '$rspDest'." -Level Warning
                }
            }
            catch { Write-BatchLog "  .rsp NOT placed beside the certificate ($($_.Exception.Message)); the response file was discarded." -Level Warning }
        }
    }
}

function Move-StaleCertificateAside {
    # "Did certreq write the certificate?" must be answerable without ambiguity, so a file that
    # already sits at the destination (a -Force resubmit of a request issued earlier; an
    # unresolved row whose .cer exists from some earlier attempt) is moved out of the way BEFORE
    # certreq runs. Afterwards, a file at the destination can only have come from this
    # invocation. (Comparing timestamps instead would be unreliable: FAT and some SMB shares
    # keep coarse last-write times, so a rapid rewrite can leave the stamp unchanged.)
    # Nothing is deleted: the file is renamed <name>.superseded-<UTC stamp>.cer next to the
    # original and the path is returned, so the caller can drop that copy again if the fresh
    # certificate turns out to be identical (see Remove-AsideIfIdentical). $null when there was
    # nothing to move.
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
    $dir   = [System.IO.Path]::GetDirectoryName($Path)
    $stem  = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $ext   = [System.IO.Path]::GetExtension($Path)
    $aside = Join-Path $dir ('{0}.superseded-{1:yyyyMMddHHmmssfff}{2}' -f $stem, [DateTime]::UtcNow, $ext)
    # No-overwrite rename (see Move-RetrievedCertificate): fails rather than moving the old
    # certificate INTO something planted under the aside name.
    [System.IO.File]::Move($Path, $aside)
    Write-BatchLog "  A file already existed at the destination; moved aside as: $aside"
    return $aside
}

function Remove-AsideIfIdentical {
    # After a successful write: if the freshly written certificate is byte-identical to the copy
    # Move-StaleCertificateAside set aside (the same certificate retrieved again), the aside copy
    # is redundant and is removed. A different certificate (a resubmission) keeps its
    # predecessor next to it.
    param([Parameter(Mandatory)][string]$Path, [string]$Aside)
    if (-not $Aside -or -not (Test-Path -LiteralPath $Aside -PathType Leaf) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) { return }
    if ((Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash -eq (Get-FileHash -LiteralPath $Aside -Algorithm SHA256).Hash) {
        Remove-Item -LiteralPath $Aside -Force -Confirm:$false
        Write-BatchLog "  The retrieved certificate is identical to the copy moved aside; removed the duplicate: $Aside"
    }
}

function Write-BatchLog {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$timestamp] [$Level] $Message"

    switch ($Level) {
        'Warning' { Write-Warning $Message }
        'Error'   { Write-Error -Message $entry -ErrorAction Continue }
        default   { Write-Host $entry }
    }

    if (-not $script:SuppressLogFile) {
        # -Confirm:$false on every cmdlet that supports ShouldProcess inside an already-approved
        # action: with the script's -Confirm the preference is Low, and Out-File / Export-Csv /
        # Remove-Item / New-Item would each raise their OWN prompt. Declining one of those after
        # certreq has submitted would lose the checkpoint (RequestID) of a request the CA already
        # has - the one thing that must never be optional.
        $entry | Out-File -FilePath $script:LogFile -Append -Encoding utf8 -Confirm:$false
    }
}

function Test-CAConnectivity {
    param([string]$CAConfig)

    Write-BatchLog "Testing connectivity to CA: $CAConfig"
    try {
        # Localized EAP: with $ErrorActionPreference = 'Stop', 2>&1 turns native
        # stderr lines into terminating errors in Windows PowerShell 5.1
        $output = & {
            $ErrorActionPreference = 'Continue'
            certutil.exe -ping -config $CAConfig 2>&1 | ForEach-Object { "$_" }
        }
        $exitCode = $LASTEXITCODE
        if ($exitCode -ne 0) {
            Write-BatchLog "certutil -ping failed (exit $exitCode): $($output -join ' ')" -Level Error
            return $false
        }
        Write-BatchLog "CA connectivity OK."
        return $true
    }
    catch {
        Write-BatchLog "Could not reach CA: $_" -Level Error
        return $false
    }
}

function Get-RequestFiles {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        Write-BatchLog "InputPath does not exist: $Path" -Level Error
        return @()
    }

    if (-not (Get-Item $Path).PSIsContainer) {
        Write-BatchLog "InputPath must be a folder, not a file: $Path" -Level Error
        return @()
    }

    $files = @(
        Get-ChildItem -Path $Path -Filter '*.req' -File
        Get-ChildItem -Path $Path -Filter '*.csr' -File
        Get-ChildItem -Path $Path -Filter '*.txt' -File
    )

    if ($files.Count -eq 0) {
        Write-BatchLog "No .req/.csr/.txt files found in: $Path" -Level Warning
        $other = @(Get-ChildItem -Path $Path -File -ErrorAction SilentlyContinue)
        if ($other.Count -gt 0) {
            $sample = ($other | Select-Object -First 5 -ExpandProperty Name) -join ', '
            $suffix = if ($other.Count -gt 5) { ", ..." } else { '' }
            Write-BatchLog "  Folder contains $($other.Count) other file(s): $sample$suffix" -Level Warning
            Write-BatchLog "  Rename them to .req/.csr/.txt or move CSRs into this folder." -Level Warning
        }
        else {
            Write-BatchLog "  Folder is empty." -Level Warning
        }
    }

    return $files
}

function Get-RequestIdFromOutput {
    param([string[]]$Output)

    foreach ($line in $Output) {
        if ($line -match 'RequestId:\s*"?(\d+)"?') {
            return [int]$Matches[1]
        }
    }
    return $null
}

function Get-DispositionFromOutput {
    param([string[]]$Output)

    $joined = $Output -join "`n"
    if ($joined -match 'retrieved\(Issued\)')            { return 'Issued' }
    if ($joined -match 'pending|Taken Under Submission') { return 'Pending' }
    if ($joined -match 'denied')                         { return 'Denied' }
    return 'Unknown'
}

function Get-FriendlyErrorHint {
    param(
        [string[]]$Output,
        [string]$CertificateTemplate,
        [string]$CAConfig
    )

    $joined = $Output -join "`n"
    $hints = @()

    if ($joined -match '0x80094800|CERTSRV_E_UNSUPPORTED_CERT_TYPE|template that is not supported|certificate template is not supported') {
        $hints += "Template '$CertificateTemplate' is not supported by CA '$CAConfig'. Likely causes:"
        $hints += "  - Template name is misspelled (use the template's internal name, e.g. 'WebServer', not the display name)."
        $hints += "  - Template is not published on this CA (Certification Authority MMC -> Certificate Templates -> New -> Certificate Template to Issue)."
        $hints += "  - You accidentally included the 'CertificateTemplate:' prefix in -CertificateTemplate. Pass only the template name."
        $hints += "  Verify with:  certutil -config `"$CAConfig`" -CATemplates"
    }
    elseif ($joined -match '0x80094012|CERTSRV_E_TEMPLATE_DENIED|Denied by Policy Module.*[Aa]ccess|not have permission|not allowed') {
        $hints += "Enrollment was denied by policy. Likely causes:"
        $hints += "  - The calling account lacks Enroll permission on template '$CertificateTemplate'."
        $hints += "  - Template requires Autoenroll/manager approval/signature you are not providing."
        $hints += "  - Try running the script as an account that has Enroll rights on the template."
    }
    elseif ($joined -match '0x80094004|CERTSRV_E_BAD_REQUESTSUBJECT|subject') {
        $hints += "Request subject was rejected. The CSR subject/SAN may not match what the template requires."
    }
    elseif ($joined -match '0x80070005|Access is denied|access denied') {
        $hints += "Access denied by the CA. The calling account likely lacks 'Request Certificates' rights on CA '$CAConfig'."
    }

    return $hints
}

function Import-TrackingData {
    param([string]$Path)

    if (Test-Path $Path) {
        return @(Import-Csv -Path $Path -Encoding utf8)
    }
    return @()
}

function Export-TrackingData {
    param(
        [object[]]$Data,
        [string]$Path
    )

    $filtered = @($Data | Where-Object { $_ -ne $null })
    if ($filtered.Count -eq 0) {
        Write-BatchLog "No tracking data to export." -Level Warning
        return
    }
    # Unique temp name (not a fixed "$Path.tmp"): a second process writing the same fixed name
    # could otherwise clobber or move a half-written file. Written next to the destination so
    # the final Move-Item is a same-volume rename.
    $tempFile = "$Path.$([System.IO.Path]::GetRandomFileName()).tmp"
    try {
        # -Confirm:$false: the checkpoint is a mandatory part of an already-approved submission,
        # never a separately declinable prompt (see Write-BatchLog).
        $filtered | Export-Csv -Path $tempFile -NoTypeInformation -Encoding utf8 -Confirm:$false
        # File.Replace / File.Move rather than Move-Item -Force: both fail if the tracking-file
        # name is occupied by a folder or a link (Move-Item would move the CSV INTO a folder).
        if (Test-Path -LiteralPath $Path -PathType Leaf) { [System.IO.File]::Replace($tempFile, $Path, $null) }
        else { [System.IO.File]::Move($tempFile, $Path) }
    }
    finally {
        if (Test-Path -LiteralPath $tempFile) { Remove-Item -LiteralPath $tempFile -Force -Confirm:$false -ErrorAction SilentlyContinue }
    }
}

function Remove-RspFile {
    param([string]$CerPath)

    $rspPath = [System.IO.Path]::ChangeExtension($CerPath, '.rsp')
    if (Test-Path $rspPath) {
        Remove-Item -Path $rspPath -Force -Confirm:$false -ErrorAction SilentlyContinue
        Write-BatchLog "  Deleted .rsp file: $rspPath"
    }
}

function Submit-SingleRequest {
    param(
        [System.IO.FileInfo]$RequestFile,
        [string]$CAConfig,
        [string]$CertificateTemplate,
        [string]$CerPath,
        [switch]$KeepRspFile,
        [string[]]$AllowedRoots
    )

    Write-BatchLog "Submitting: $($RequestFile.Name)"

    $stdoutFile = [System.IO.Path]::GetTempFileName()
    $stderrFile = [System.IO.Path]::GetTempFileName()

    $tmpCer = New-TempCertificatePath -Destination $CerPath
    $keepTemp = $false
    try {
        # certreq writes into a private staging file beside the destination; the destination (which
        # may hold the .cer of a request this -Force run resubmits) is touched only after a
        # successful write.
        # -q is load-bearing: without it certreq may pop a MODAL GUI dialog on some error paths
        # (observed live: an unsupported-template denial), hanging a batch/headless run forever.
        $proc = Start-Process -FilePath 'certreq.exe' -ArgumentList @(
            '-submit', '-q', '-f',
            '-config', "`"$CAConfig`"",
            '-attrib', "`"CertificateTemplate:$CertificateTemplate`"",
            "`"$($RequestFile.FullName)`"",
            "`"$tmpCer`""
        ) -NoNewWindow -Wait -PassThru -RedirectStandardOutput $stdoutFile -RedirectStandardError $stderrFile

        $stdout = @(Get-Content -Path $stdoutFile -ErrorAction SilentlyContinue)
        $stderr = @(Get-Content -Path $stderrFile -ErrorAction SilentlyContinue)

        $requestId = Get-RequestIdFromOutput $stdout
        $disposition = Get-DispositionFromOutput $stdout
        $errorMsg = ''

        if ($proc.ExitCode -ne 0 -and $disposition -ne 'Pending') {
            $errorMsg = ($stderr + $stdout) -join ' '
            if (-not $requestId) {
                $disposition = 'Error'
            }
            Write-BatchLog "certreq failed for $($RequestFile.Name): $errorMsg" -Level Warning
        }

        # Language-independent success signal: certreq returned 0 AND wrote the temp file (which
        # did not exist before, so it is necessarily this invocation's output). It overrides the
        # text parse in both directions - localized output that never says "Issued" is still
        # Issued, and "Issued" wording with no .cer on disk is NOT (recorded as Error so the row
        # stays eligible for -Mode Retrieve instead of being finalized without a file). Only a
        # success delivers to the destination.
        $cerWritten = ($proc.ExitCode -eq 0) -and (Test-Path -LiteralPath $tmpCer -PathType Leaf)
        if ($cerWritten) {
            $disposition = 'Issued'
            # Delivery is a separate step from issuance: the CA HAS issued, and the RequestID (if
            # parsed) is final. A failure here - locked destination, denied rename, a destination
            # that no longer passes validation - must keep that RequestID and become a RETRYABLE
            # Error row, never propagate as an exception that the caller turns into an
            # "unsubmitted" row (which a later run would resubmit: a duplicate request at the CA).
            # The undelivered certificate is left in its staging file for recovery.
            try {
                Move-RetrievedCertificate -TempCer $tmpCer -Destination $CerPath -KeepRspFile:$KeepRspFile -AllowedRoots $AllowedRoots
            }
            catch {
                # 'Undelivered' (not 'Error') records the fact that the CA issued: the row counts as
                # SUBMITTED on later runs even without a RequestID, so it is never resubmitted
                # automatically, and -Mode Retrieve re-fetches it when a RequestID is present.
                $disposition = 'Undelivered'
                $keepTemp = $true
                $errorMsg = "Certificate ISSUED but could not be delivered to $CerPath ($($_.Exception.Message)). The undelivered certificate is at $tmpCer; -Mode Retrieve fetches it again by RequestID."
                Write-BatchLog "  $errorMsg" -Level Warning
            }
        }
        elseif ($disposition -eq 'Issued') {
            $disposition = 'Error'
            $errorMsg = "certreq reported Issued (exit $($proc.ExitCode)) but no certificate file was written - the row stays eligible for -Mode Retrieve. Output: $(($stderr + $stdout) -join ' ')"
            Write-BatchLog "  $errorMsg" -Level Warning
        }

        if ($requestId) {
            Write-BatchLog "  RequestID: $requestId - Status: $disposition"
        }
        elseif ($cerWritten) {
            # Certificate written, but no RequestId line parsed (localized certreq output?). A
            # delivered certificate keeps Issued - resubmitting would duplicate the request at the
            # CA - and the gap is recorded so the ID can be filled in from the CA database.
            $note = "RequestID could not be parsed from certreq output (localized wording?); the certificate was issued. Output: $($stdout -join ' ')"
            $errorMsg = if ($errorMsg) { "$errorMsg $note" } else { $note }
            Write-BatchLog "  $note" -Level Warning
        }
        else {
            Write-BatchLog "  Could not parse RequestID from output" -Level Warning
            $disposition = 'Error'
            $errorMsg = "No RequestID in output: $($stdout -join ' ')"
        }

        if ($disposition -in 'Denied', 'Error') {
            $hints = Get-FriendlyErrorHint -Output ($stdout + $stderr) `
                -CertificateTemplate $CertificateTemplate -CAConfig $CAConfig
            foreach ($h in $hints) {
                Write-BatchLog $h -Level Error
            }
        }

        return [PSCustomObject]@{
            RequestFile    = $RequestFile.FullName
            RequestID      = $requestId
            SubmitTime     = (Get-Date -Format 'o')
            Status         = $disposition
            OutputCertFile = $CerPath
            LastCheckTime  = (Get-Date -Format 'o')
            ErrorMessage   = $errorMsg
        }
    }
    finally {
        Remove-Item -Path $stdoutFile, $stderrFile -Force -Confirm:$false -ErrorAction SilentlyContinue
        # Whatever certreq left beside the destination (the .rsp of a pending or denied request) is this
        # invocation's own and is cleaned up here - except an issued certificate that could not be
        # delivered, which stays for recovery (its path is in the row's ErrorMessage).
        if (-not $keepTemp) {
            Remove-Item -LiteralPath $tmpCer -Force -Confirm:$false -ErrorAction SilentlyContinue
            Remove-RspFile -CerPath $tmpCer
        }
    }
}

function Get-IssuedCertificate {
    param(
        [PSCustomObject]$Record,
        [string]$CAConfig,
        [switch]$KeepRspFile,
        [string[]]$AllowedRoots
    )

    Write-BatchLog "Retrieving certificate for RequestID: $($Record.RequestID)"

    $stdoutFile = [System.IO.Path]::GetTempFileName()
    $stderrFile = [System.IO.Path]::GetTempFileName()

    $tmpCer = New-TempCertificatePath -Destination $Record.OutputCertFile
    $keepTemp = $false
    $deliveryError = $null
    try {
        # certreq writes into a private staging file beside the destination (see
        # Submit-SingleRequest); the destination named by the tracking row is touched only after a
        # successful retrieval.
        # -q for the same reason as in Submit-SingleRequest: never let certreq raise UI.
        $proc = Start-Process -FilePath 'certreq.exe' -ArgumentList @(
            '-retrieve', '-q', '-f',
            '-config', "`"$CAConfig`"",
            "$($Record.RequestID)",
            "`"$tmpCer`""
        ) -NoNewWindow -Wait -PassThru -RedirectStandardOutput $stdoutFile -RedirectStandardError $stderrFile

        $stdout = @(Get-Content -Path $stdoutFile -ErrorAction SilentlyContinue)
        $stderr = @(Get-Content -Path $stderrFile -ErrorAction SilentlyContinue)

        $disposition = Get-DispositionFromOutput $stdout
        # Issued means "certreq returned 0 AND wrote the temp file" - language-independent,
        # immune to a stale file at the destination, and it never finalizes a row whose file was
        # not written (Retrieve treats Issued as final and would never retry it; as Error it is
        # picked up again on the next Retrieve). Only a success delivers to the destination.
        $cerWritten = ($proc.ExitCode -eq 0) -and (Test-Path -LiteralPath $tmpCer -PathType Leaf)
        if ($cerWritten) {
            $disposition = 'Issued'
            # Delivery failure (locked destination, denied rename, failed re-validation) keeps the
            # row retryable with its RequestID intact; the certificate stays in its staging file for recovery.
            try {
                Move-RetrievedCertificate -TempCer $tmpCer -Destination $Record.OutputCertFile -KeepRspFile:$KeepRspFile -AllowedRoots $AllowedRoots
            }
            catch {
                $disposition = 'Undelivered'
                $keepTemp = $true
                $deliveryError = "Certificate retrieved but could not be delivered to $($Record.OutputCertFile) ($($_.Exception.Message)). The undelivered certificate is at $tmpCer; the row stays eligible for the next Retrieve."
            }
        }
        elseif ($disposition -eq 'Issued') {
            $disposition = 'Error'
        }
        $Record.Status = $disposition
        $Record.LastCheckTime = (Get-Date -Format 'o')

        switch ($disposition) {
            'Issued' {
                Write-BatchLog "  Certificate retrieved: $($Record.OutputCertFile)"
            }
            'Undelivered' {
                $Record.ErrorMessage = $deliveryError
                Write-BatchLog "  $deliveryError" -Level Warning
            }
            'Error' {
                $Record.ErrorMessage = "certreq reported Issued (exit $($proc.ExitCode)) but no certificate file was written; the request stays eligible for the next Retrieve. Output: $(($stderr + $stdout) -join ' ')"
                Write-BatchLog "  Status Issued, but the .cer file was not created - recorded as Error so the next Retrieve retries it" -Level Warning
            }
            'Pending' {
                Write-BatchLog "  Still pending" -Level Warning
            }
            'Denied' {
                $Record.ErrorMessage = ($stderr + $stdout) -join ' '
                Write-BatchLog "  Request DENIED" -Level Error
                $hints = Get-FriendlyErrorHint -Output ($stdout + $stderr) `
                    -CertificateTemplate '(see original submit)' -CAConfig $CAConfig
                foreach ($h in $hints) {
                    Write-BatchLog $h -Level Error
                }
            }
            default {
                $Record.ErrorMessage = ($stderr + $stdout) -join ' '
                Write-BatchLog "  Unknown status: $disposition" -Level Warning
            }
        }
    }
    finally {
        Remove-Item -Path $stdoutFile, $stderrFile -Force -Confirm:$false -ErrorAction SilentlyContinue
        if (-not $keepTemp) {
            Remove-Item -LiteralPath $tmpCer -Force -Confirm:$false -ErrorAction SilentlyContinue
            Remove-RspFile -CerPath $tmpCer
        }
    }
}

function Write-Summary {
    param(
        [PSCustomObject[]]$RunData,
        [PSCustomObject[]]$AllData,
        [string]$RunLabel = 'This run'
    )

    $run = @($RunData | Where-Object { $_ -ne $null })
    $all = @($AllData | Where-Object { $_ -ne $null })

    Write-BatchLog "--- Summary: $RunLabel ---"
    if ($run.Count -eq 0) {
        Write-BatchLog "  (no records processed in this run)"
    }
    else {
        $run | Group-Object Status | Sort-Object Name | ForEach-Object {
            Write-BatchLog "  $($_.Name): $($_.Count)"
        }
        Write-BatchLog "  Total processed this run: $($run.Count)"
    }

    Write-BatchLog "--- Tracking file total (all history) ---"
    if ($all.Count -eq 0) {
        Write-BatchLog "  (tracking file is empty)"
    }
    else {
        $all | Group-Object Status | Sort-Object Name | ForEach-Object {
            Write-BatchLog "  $($_.Name): $($_.Count)"
        }
    }
    Write-BatchLog "Tracking file: $TrackingFile"
}

#endregion

#region Main

# Resolve to absolute paths so tracking records stay valid when later runs
# use a different working directory
$TrackingFile = Resolve-TrackingFilePath -Path $TrackingFile   # canonical long name; hard links / symlinks refused (lock identity)
$OutputFolder = Resolve-FullPath -Path $OutputFolder
$script:LogFile = Resolve-FullPath -Path (".\CertBatch_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))
$script:SuppressLogFile = [bool]$WhatIfPreference

# Validation
if ($Mode -in 'Submit', 'Both') {
    if (-not $InputPath) {
        throw "-InputPath is required when -Mode is '$Mode'."
    }
    if (-not $CertificateTemplate) {
        throw "-CertificateTemplate is required when -Mode is '$Mode'."
    }
}
Assert-SafeNativeArgument -Name '-CAConfig' -Value $CAConfig
if ($CertificateTemplate) { Assert-SafeNativeArgument -Name '-CertificateTemplate' -Value $CertificateTemplate }

# One run per tracking file. Two concurrent runs would each read the CSV, append their own rows
# and replace the file - the last writer silently discarding the other's request IDs, which a
# later run would then resubmit (duplicate requests at the CA). The guard is a lock file beside
# the CSV, opened with FileShare.None and held for the whole run: the exclusion is enforced by
# the file system - the SMB server included - so it covers runs from OTHER machines and every
# alias of the DIRECTORY path (drive letter vs. UNC, mapped drive, junction), none of which a
# named mutex would; aliases of the FILE name are canonicalized or refused up front by
# Resolve-TrackingFilePath (8.3 short name -> long name; hard link / symlink -> error), so the
# lock file's own name is unique per CSV. DeleteOnClose removes the file when the handle closes (the finally at the
# end of the run, or the OS if the process dies). A -WhatIf run never writes the CSV, so it takes
# no lock (and leaves no transient file behind).
$script:RunLock = $null
if (-not $WhatIfPreference) {
    $trackingDir = [System.IO.Path]::GetDirectoryName($TrackingFile)
    if (-not (Test-Path -LiteralPath $trackingDir -PathType Container)) {
        throw "The folder for -TrackingFile does not exist: '$trackingDir'. Create it first."
    }
    $lockPath = "$TrackingFile.lock"
    try {
        $script:RunLock = New-Object System.IO.FileStream($lockPath, [System.IO.FileMode]::OpenOrCreate,
            [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None, 4096, [System.IO.FileOptions]::DeleteOnClose)
    }
    catch [System.IO.IOException] {
        throw "Another Submit-CertificateRequests run is using tracking file '$TrackingFile' (its lock file '$lockPath' is held). Wait for it to finish, or use a different -TrackingFile. ($($_.Exception.Message))"
    }
}

try {   # the lock is released in the finally at the end of the run, on every exit path

    if (-not (Test-Path $OutputFolder)) {
        if ($PSCmdlet.ShouldProcess($OutputFolder, 'Create output folder')) {
            New-Item -Path $OutputFolder -ItemType Directory -Force -Confirm:$false | Out-Null
            Write-BatchLog "Created output folder: $OutputFolder"
        }
    }

    # The delivery boundary: certificates may only be delivered beneath these folders. Each root
    # must be a real folder (not a pre-planted junction), and one that untrusted users could
    # delete or rename - and so swap for a junction during a delivery - is refused up front, the
    # same rule Assert-CertificateOutputPath applies per destination (fail closed; the operator
    # can accept the risk explicitly with -AllowUnprotectedOutputFolder). Only the create-subfolder
    # right on its own is reported for awareness rather than refused.
    $script:AllowUnprotectedOutput = [bool]$AllowUnprotectedOutputFolder
    $script:TrustedSids = Get-TrustedPrincipalSet -Extra $TrustedOutputPrincipal
    $retrieveRoots = @([System.IO.Path]::GetDirectoryName($TrackingFile), $OutputFolder)
    foreach ($root in ($retrieveRoots | Sort-Object -Unique)) {
        if (-not (Test-Path -LiteralPath $root -PathType Container)) { continue }
        # The same chain check every delivery repeats: reparse points, untrusted owners, swappable
        # folders and unreadable ACLs from the root up to the volume/share root, fail closed.
        Assert-ProtectedDirectoryChain -Name "Output location" -Directory $root -Path $root
        try {
            $creators = @(Get-UntrustedGrant -Path $root -Kind Create)   # @(): an empty result unrolls to $null otherwise
            if ($creators.Count) {
                Write-Verbose "Output location '$root' grants create-subfolder rights to $($creators -join ', ') (cannot set a reparse point on it; planted subfolders/junctions are refused and deliveries never overwrite)."
            }
        }
        catch { Write-Verbose "Create-rights check skipped for '$root': $_" }   # informational only; the Swap check above already failed closed
    }

    # CA connectivity test
    if (-not (Test-CAConnectivity -CAConfig $CAConfig)) {
        throw "Cannot reach CA. Aborting."
    }

    # Submit mode
    if ($Mode -in 'Submit', 'Both') {
        $requestFiles = @(Get-RequestFiles -Path $InputPath)

        if ($requestFiles.Count -eq 0) {
            Write-BatchLog "Nothing to submit. Skipping Submit phase." -Level Warning
        }
        else {
            $existingData = Import-TrackingData -Path $TrackingFile
            $tracking = [System.Collections.ArrayList]@($existingData)

            # A file counts as submitted when its row carries a RequestID - or records an issuance
            # without one (Issued / Undelivered with the RequestId line not parsed): resubmitting
            # any of these duplicates the request at the CA.
            $alreadySubmitted = @($tracking | Where-Object { $_.RequestID -or $_.Status -in 'Issued', 'Undelivered' } | Select-Object -ExpandProperty RequestFile)
            $runResults = [System.Collections.ArrayList]@()

            # Files sharing a base name (a.req + a.csr) would collide on a.cer;
            # those keep the full file name in their .cer name instead
            $duplicateBaseNames = @(
                $requestFiles |
                    Group-Object { [System.IO.Path]::GetFileNameWithoutExtension($_.Name) } |
                    Where-Object { $_.Count -gt 1 } |
                    Select-Object -ExpandProperty Name
            )

            Write-BatchLog "Found $($requestFiles.Count) request file(s) in $InputPath"

            foreach ($file in $requestFiles) {
                if ($file.Length -eq 0) {
                    Write-BatchLog "Skipping (empty file): $($file.Name)" -Level Warning
                    continue
                }

                if (-not $PSCmdlet.ShouldProcess($file.Name, "Submit certificate request to $CAConfig")) {
                    continue
                }

                if ($file.FullName -in $alreadySubmitted) {
                    $existing = @($tracking | Where-Object { $_.RequestFile -eq $file.FullName -and ($_.RequestID -or $_.Status -in 'Issued', 'Undelivered') }) |
                                Select-Object -Last 1
                    $prevId = $existing.RequestID
                    $prevStatus = $existing.Status

                    if ($Force) {
                        Write-BatchLog "Resubmitting (-Force): $($file.Name) [previous RequestID: $prevId, Status: $prevStatus]" -Level Warning
                    }
                    else {
                        $yes = New-Object System.Management.Automation.Host.ChoiceDescription '&Yes', 'Resubmit as a new certificate request'
                        $no  = New-Object System.Management.Automation.Host.ChoiceDescription '&No',  'Skip this file'
                        $choices = [System.Management.Automation.Host.ChoiceDescription[]]@($yes, $no)
                        $message = "File '$($file.Name)' was already submitted (RequestID: $prevId, Status: $prevStatus). Resubmit as a new request?"
                        try {
                            $decision = $Host.UI.PromptForChoice('Already submitted', $message, $choices, 1)
                        }
                        catch {
                            Write-BatchLog "Host cannot prompt; skipping already-submitted file: $($file.Name). Use -Force to resubmit." -Level Warning
                            continue
                        }
                        if ($decision -ne 0) {
                            Write-BatchLog "Skipping (already submitted): $($file.Name)"
                            continue
                        }
                        Write-BatchLog "Resubmitting on user confirmation: $($file.Name) [previous RequestID: $prevId, Status: $prevStatus]" -Level Warning
                    }
                }

                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
                $cerName = if ($baseName -in $duplicateBaseNames) { "$($file.Name).cer" } else { "$baseName.cer" }
                $cerPath = Join-Path $OutputFolder $cerName

                try {
                    $result = Submit-SingleRequest -RequestFile $file `
                        -CAConfig $CAConfig `
                        -CertificateTemplate $CertificateTemplate `
                        -CerPath $cerPath `
                        -KeepRspFile:$KeepRspFile `
                        -AllowedRoots @($OutputFolder)
                }
                catch {
                    Write-BatchLog "Error submitting $($file.Name): $_" -Level Error
                    $result = [PSCustomObject]@{
                        RequestFile    = $file.FullName
                        RequestID      = $null
                        SubmitTime     = (Get-Date -Format 'o')
                        Status         = 'Error'
                        OutputCertFile = ''
                        LastCheckTime  = (Get-Date -Format 'o')
                        ErrorMessage   = $_.ToString()
                    }
                }

                [void]$tracking.Add($result)
                [void]$runResults.Add($result)

                # Persist after every file so request IDs survive a mid-batch crash
                Export-TrackingData -Data @($tracking) -Path $TrackingFile
            }

            Write-Summary -RunData @($runResults) -AllData @($tracking) -RunLabel 'Submit'
        }
    }

    # Retrieve mode
    if ($Mode -in 'Retrieve', 'Both') {
        $tracking = @(Import-TrackingData -Path $TrackingFile)

        if ($tracking.Count -eq 0) {
            Write-BatchLog "No data in tracking file. Run Submit first." -Level Warning
        }
        else {
            # Retry anything with a RequestID that is not finally resolved - including
            # rows stuck in 'Unknown' or 'Error' that the CA may have issued since
            $unresolved = @($tracking | Where-Object { $_.RequestID -and $_.Status -notin 'Issued', 'Denied' })
            Write-BatchLog "Found $($unresolved.Count) unresolved request(s) with a RequestID"
            # Issued-but-undelivered rows WITHOUT a RequestID cannot be fetched again; they are
            # never resubmitted either, so they need the operator: the certificate itself is still
            # in its staging file (path in ErrorMessage) and the ID can be found in the CA database.
            foreach ($orphan in @($tracking | Where-Object { -not $_.RequestID -and $_.Status -eq 'Undelivered' })) {
                Write-BatchLog "Needs manual reconciliation: '$($orphan.RequestFile)' was issued but never delivered and has no RequestID. $($orphan.ErrorMessage)" -Level Warning
            }

            $runResults = [System.Collections.ArrayList]@()

            # The .cer path is recorded per row at submit time. An explicit -OutputFolder on a
            # Retrieve run overrides it (the file name is kept); the default is not applied here,
            # so a Retrieve from another working directory still lands where Submit put it.
            $redirectOutput = $PSBoundParameters.ContainsKey('OutputFolder')
            if ($redirectOutput) {
                Write-BatchLog "Retrieved certificates will be written to: $OutputFolder"
            }

            foreach ($record in $unresolved) {
                # Both fields below reach the certreq command line verbatim, and both come from the
                # CSV (editable, importable): a non-numeric RequestID or a path carrying a quote is
                # refused for this row rather than passed on. The row is left as it is and skipped.
                if ("$($record.RequestID)" -notmatch '^\d+$') {
                    Write-BatchLog "Skipping row for '$($record.RequestFile)': RequestID '$($record.RequestID)' is not numeric (edited tracking file?)." -Level Error
                    continue
                }

                if ($redirectOutput) {
                    $cerName = if ($record.OutputCertFile) {
                        [System.IO.Path]::GetFileName($record.OutputCertFile)
                    }
                    else {
                        [System.IO.Path]::GetFileNameWithoutExtension($record.RequestFile) + '.cer'
                    }
                    $record.OutputCertFile = Join-Path $OutputFolder $cerName
                }

                if (-not $record.OutputCertFile) {
                    Write-BatchLog "Skipping RequestID $($record.RequestID): the row has no OutputCertFile (pass -OutputFolder to redirect it)." -Level Error
                    continue
                }
                try {
                    Assert-CertificateOutputPath -Name "OutputCertFile of RequestID $($record.RequestID)" -Path $record.OutputCertFile -AllowedRoots $retrieveRoots
                }
                catch {
                    Write-BatchLog "Skipping RequestID $($record.RequestID): $_" -Level Error
                    continue
                }

                if ($PSCmdlet.ShouldProcess("RequestID $($record.RequestID)", "Retrieve certificate from $CAConfig")) {
                    try {
                        Get-IssuedCertificate -Record $record -CAConfig $CAConfig -KeepRspFile:$KeepRspFile -AllowedRoots $retrieveRoots
                    }
                    catch {
                        Write-BatchLog "Error retrieving RequestID $($record.RequestID): $_" -Level Error
                        $record.ErrorMessage = $_.ToString()
                        $record.LastCheckTime = (Get-Date -Format 'o')
                        $record.Status = 'Error'
                    }
                    [void]$runResults.Add($record)

                    # Persist after every record so status updates survive a mid-batch crash
                    Export-TrackingData -Data $tracking -Path $TrackingFile
                }
            }

            Write-Summary -RunData @($runResults) -AllData $tracking -RunLabel 'Retrieve'
        }
    }

    if ($script:SuppressLogFile) {
        Write-BatchLog "Done. (-WhatIf run: log file not written)"
    }
    else {
        Write-BatchLog "Done. Log file: $script:LogFile"
    }

}
finally {
    # Always release the tracking-file lock - a throw anywhere above would otherwise keep the
    # handle (and so the lock) open until the object is finalized, blocking the next run.
    if ($script:RunLock) { $script:RunLock.Dispose() }
}

#endregion
