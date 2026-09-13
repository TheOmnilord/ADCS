# Certificate substitution risks in the output folder

This note explains what an attacker gains by altering a delivered `.cer` file when the
output folder is not protected. It is the threat model behind the `-AllowUnprotectedOutputFolder`
CAUTION in `Submit-CertificateRequests.ps1`: "an untrusted user can redirect the privileged
write, or alter the certificate before or after delivery."

## The two tamper windows

The script delivers a certificate in three steps, and two of them leave a file an attacker can reach:

1. `certreq` writes the issued certificate into a **private staging file inside the output folder chain**, beside the eventual destination.
2. The script re-checks the destination, then performs a **no-overwrite rename** (`File.Move`) from the staging file to the final `.cer`.
3. A consumer (a person, a GPO step, another job) later reads that `.cer`.

"Delivery" is step 2. **Before** delivery is the life of the staging file in step 1; **after**
delivery is the life of the destination file between step 2 and the consumer's read in step 3.

The staging file is created inside the output folder, so it inherits that folder's file access
control entries. When an inheritable "files" entry grants an untrusted principal write, append,
write-attributes, delete or re-permission rights, the staging file and the delivered certificate
both carry that grant. That is the exact check `-AllowUnprotectedOutputFolder` downgrades from a
refusal to a warning, and it is what opens both windows.

## Case 1: identity use — substitution is a denial of service

A certificate binds an identity to the **public key from the original request**. Where the
certificate is later paired with a private key to prove possession (a server or client
authentication identity), the pairing is checked, and a wrong public key fails at that check:

- `certreq -accept foo.cer` on the machine that holds the key binds the certificate to the key **by matching the public key**. A substituted certificate with a different key gets no key handle, so the accept fails and nothing usable is installed.
- Importing a key-and-certificate pair into any serious store (CNG/KSP, the Windows certificate store, a Java keystore, OpenSSL) associates them by public-key match. A mismatch is rejected, or it leaves an orphan certificate with no key.

So substituting an unrelated certificate breaks the service. The requester's private key no
longer matches the delivered public key, the service does not load it, and the handshake or the
import fails. The value of the attack here is **integrity and availability**, not impersonation:

- The legitimate certificate never arrives intact, which causes an outage, a failed renewal that lets the old certificate expire, or a deployment with the wrong subject alternative name.
- The attacker can repeat this to deny a working certificate, or time it to a renewal window.

The attacker cannot impersonate the service this way, because that needs the private key for
whatever public key is in the planted file. Delivering the attacker's own certificate does not
give the requester the attacker's key either. In an identity flow the only way substitution
stays silent is a broken importer that binds a certificate to a key by filename or slot without
checking the key. That is a bug, not the normal path.

## Case 2: public-object use — substitution is a trust compromise

The substitution works **silently** where the certificate is never paired with a private key at
the consumer. The certificate is installed to trust, verify, encrypt to, or pin. There is no key
to mismatch, so nothing catches the swap.

1. **Trust anchors and CA certificates.** A `.cer` imported into Trusted Root, or worse into `NTAuthCertificates`. A substituted attacker CA certificate makes every client-authentication certificate that CA issues trusted, which is full forged-identity capability. Pure public-key use.

2. **Signature-verification certificates.** The verifier holds only the public certificate:
   - A **SAML identity-provider signing certificate** that a service provider loads to validate assertions. Swap it and the service provider accepts attacker-forged assertions, an authentication bypass.
   - **Code-signing or script-signing trust**, **JWT or OIDC signing keys published as certificates**, and **document-signing verification** have the same shape. Replace the "who we trust to sign" certificate, and attacker-signed artifacts verify.

3. **Encryption to a recipient.** The sender uses only the recipient's public certificate to encrypt: **S/MIME encryption, CMS or PKCS#7 enveloping**, or "encrypt this to the server's certificate." When the delivery folder is treated as the certificate to encrypt to for a host, an attacker substitutes a certificate whose private key they hold. Senders then encrypt to the attacker, who decrypts. This substitution needs no access to any legitimate private key, and it yields a confidentiality break rather than an outage.

4. **Certificate or public-key pinning, and expected-certificate allow-lists.** A client that stores the peer's expected certificate to validate future connections. Replace the pinned `.cer` and the client accepts more than it should, which enables a later machine-in-the-middle attack. Public-key use only.

5. **Publishing the certificate onward.** The `.cer` is pushed into Active Directory (`userCertificate`, AIA, CDP or NTAuth objects), federation metadata, or an allow-list sent to other parties. Every downstream consumer is a verifier or a truster, so the exposure fans out.

## The redirect sibling

"Redirect the privileged write" is the stronger risk that the same CAUTION names, and it does not
depend on the private-key question at all. An untrusted principal with delete or rename rights
can replace the folder with a junction while the script delivers, and one with only write-data or
write-attributes rights can turn an empty folder into a junction in place. Either change redirects
the privileged write. It can **capture the genuine certificate** as it is written, which reveals
what was issued to whom and when, or it can place the attacker's own file at the destination. That
is redirection rather than alteration, which is why the CAUTION lists the two outcomes separately.

## Summary

- **Identity use** (a private key is matched later): substitution breaks the certificate. The impact is a denial of service, and the failure is the payoff.
- **Public-object use** (trust, verify, encrypt to, pin, publish): no private key is matched at the consumer, so a valid but attacker-chosen certificate flows straight through and is used. This is where "alter the certificate before or after delivery" becomes a trust compromise.

The output folder can feed either kind of consumer, and the script cannot know which. So the
CAUTION states the capability, not one fixed outcome.
