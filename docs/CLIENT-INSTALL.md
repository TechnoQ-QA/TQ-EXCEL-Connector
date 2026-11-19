# Installing TQ Connector at a client

For the client's IT team (and Techno Q's consultant on site). TQ Connector's files are signed with **Techno Q's own
code-signing certificate**. Windows doesn't know that certificate until the client trusts it once. After that,
every TQ Connector file shows **Techno Q** as its verified publisher, and any changed or tampered file shows as
invalid.

## What Techno Q sends

| File | What it is |
|---|---|
| `TQConnectorSetup.exe` | The setup file, run once on each PC (for the signed-in Windows user, no administrator rights). |
| `TQConnectorSetup.exe.sha256.txt` | Its SHA-256 checksum. |
| `TechnoQ.cer` | The public part of Techno Q's code-signing certificate (no private key). |
| `TQConnector.settings.json`, `licence.json` | Optional: the NetSuite details and the licence. Placed next to the setup file, they're used without anyone typing them in. |

Pass the files on through the company's file share, Teams or email to IT, rather than a web download. Windows
SmartScreen only warns about files that came from the internet.

## 1. Check the file is genuine

```powershell
Get-FileHash .\TQConnectorSetup.exe -Algorithm SHA256
```

The hash must match the one in `TQConnectorSetup.exe.sha256.txt` and the one Techno Q gave you separately (for example in
the email or on the phone).

Check the certificate before trusting it. Its thumbprint must be
**F264951F57972A7C749F4163724D4B2B2F2DD4AB** (subject `CN=Techno Q Code Signing, O=Techno Q`, valid until 19 Nov 2031):

```powershell
certutil -dump .\TechnoQ.cer
```

The certificate can only sign programs. It isn't valid for websites, email or anything else, so trusting it
lets nothing but files signed by Techno Q through.

## 2. Trust Techno Q (once per company or per PC)

The certificate goes in two stores: **Trusted Root Certification Authorities**, so Windows accepts the chain, and
**Trusted Publishers**, so Windows treats Techno Q as a known publisher.

**Whole company, Group Policy:** Computer Configuration > Policies > Windows Settings > Security Settings > Public Key
Policies. Import `TechnoQ.cer` under *Trusted Root Certification Authorities* and again under *Trusted Publishers*.

**Whole company, Intune:** Devices > Configuration > Create > Templates > *Trusted certificate*, with `TechnoQ.cer` and
the computer certificate store *Root*. Then a second profile for the store *Trusted Publisher*, if your tenant offers
it; if not, use Group Policy or the command below.

**One PC, for all its users** (an administrator PowerShell):

```powershell
certutil -addstore Root .\TechnoQ.cer
```

```powershell
certutil -addstore TrustedPublisher .\TechnoQ.cer
```

**One PC, only for the signed-in user** (no administrator rights; Windows asks to confirm the first one):

```powershell
certutil -user -addstore Root .\TechnoQ.cer
```

```powershell
certutil -user -addstore TrustedPublisher .\TechnoQ.cer
```

Check: right-click `TQConnectorSetup.exe` > Properties > Digital Signatures > Details should say *This digital
signature is OK*. Or:

```powershell
Get-AuthenticodeSignature .\TQConnectorSetup.exe
```

`Status` should be **Valid**.

## 3. Antivirus

If the antivirus stops the setup file or `TQConnector.exe`, allow **Techno Q** as a trusted publisher by its
certificate, or allow the files by their SHA-256. In Kaspersky that's under Trusted applications / Exclusions, or a
policy in Kaspersky Security Center. The connector lives in `%LOCALAPPDATA%\TQ Connector`.

## 4. Install

Double-click `TQConnectorSetup.exe` as the person who'll use Excel, or install silently while signed in as that person:

```powershell
.\TQConnectorSetup.exe /VERYSILENT /SUPPRESSMSGBOXES /NORESTART
```

It installs for the signed-in Windows user, so deployment tools must run it **as the user, not as SYSTEM**. On the first
install, Windows asks the user to confirm trusting the connector's own certificate for https://localhost. That's a
separate, PC-specific certificate that Excel needs to talk to the connector. Then restart Excel: the TQ Connector tab
appears.

## Removing

Windows Settings > Apps > TQ Connector for NetSuite > Uninstall. To stop trusting Techno Q, delete the certificate with
thumbprint F264951F57972A7C749F4163724D4B2B2F2DD4AB from the two stores (certmgr.msc, or
`certutil -delstore Root F264951F57972A7C749F4163724D4B2B2F2DD4AB` and the same for `TrustedPublisher`).

---

# For Techno Q: the signing certificate

- It lives in the Windows certificate store of the PC that builds the setup (Current User > Personal), made on
  19 Nov 2026 with `New-SelfSignedCertificate -Type CodeSigningCert`. `installer\build.ps1` finds it by its subject
  and signs `TQConnector.exe`, the setup file and the uninstaller through `installer\sign.ps1`. Signatures are
  timestamped (DigiCert's free timestamp server), so files signed before the expiry date stay valid after it.
- **Back up the private key** to a USB stick kept somewhere safe: certmgr.msc > Personal > Certificates > Techno Q
  Code Signing > All Tasks > Export, *Yes, export the private key*, with a strong password of your own. Without the
  backup, a lost or reinstalled PC means a new certificate that every client has to trust again.
- **Keep the private key private.** Whoever has it can sign programs that client PCs trust. If it leaks, make a new
  certificate (same command, new thumbprint), rebuild, and ask clients to remove the old thumbprint and trust the new
  `TechnoQ.cer`.
- **Before 19 Nov 2031**, make a new certificate and send clients the new `TechnoQ.cer`.
- **False positives:** with each new version, if an antivirus flags it, submit the setup file as a false positive:
  Kaspersky at https://opentip.kaspersky.com, Microsoft at https://www.microsoft.com/en-us/wdsi/filesubmission.
- **A bought certificate later** (OV, or Microsoft Artifact Signing): set `TQ_SIGNTOOL` to its signing command
  (see build.ps1). Clients then no longer need step 2.
