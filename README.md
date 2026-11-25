# Office 365 Security and Admin Scripts

Practical scripts, tools, and simulation assets for Microsoft 365 administration, security operations, and troubleshooting.

## Start Here

- Project wiki: [Get Started](https://github.com/directorcia/Office365/wiki)

## Simulators

Use these interactive simulators to walk through common Microsoft 365 security scenarios.

| Simulator | What it helps with |
| --- | --- |
| [M365 Login](https://directorcia.github.io/Office365/m365-login-sim.html) | Visualizes Microsoft 365 sign-in behavior, authentication stages, and Conditional Access decision points. |
| [M365 Mail Security](https://directorcia.github.io/Office365/m365-mail-security-sim.html) | Demonstrates Exchange Online and Defender for Office 365 mail flow, filtering, and policy evaluation order. |
| [M365 License Coverage](https://directorcia.github.io/Office365/m365-license-coverage-sim.html) | Compares selected Microsoft 365 products with security and compliance requirements, highlighting illustrative gaps and potential overlaps. |
| [Image Model Comparison](https://directorcia.github.io/Office365/image-comparison.html) | Compares image-generation results from different models, including creation dates and image properties. |

## Games

- [Tigers in the Snow / Ardennes: Winter Offensive](https://directorcia.github.io/Office365/tigersinsnow/) — A browser-based retro wargame inspired by the Battle of the Bulge.

## Linux SharePoint certificate connection

`o365-connect-pnp-cert-linux.ps1` (version 1.1) creates a PnP certificate and connects to SharePoint Online on Linux.

- `-DeviceCodeTimeoutSeconds` (default 900) is how long provisioning waits for device-code sign-in. Polling retries transient network read timeouts until that window ends. The Graph PowerShell SDK's own device-code sign-in stops after 120 seconds, so the script acquires the token itself.
- `-SharePointPermission` accepts `Sites.FullControl.All` (default), `Sites.Read.All`, or `Sites.Selected`. Graph `Application.Read.All` is still granted.
- `-ProvisionEntraApp` can reuse a PFX already on disk (`-CertificateFilePath` or `-CertificatePath`, or `-CertificateThumbprint`) instead of `-GenerateLocalCertificate`.
- Saved profiles match tenant, tenant id, organization, default domain, and profile name without case sensitivity. A vanity domain or tenant GUID matches a profile stored as the tenant's `*.onmicrosoft.com` domain. When several profiles match, pass `-ProfileName`.

## Connect

- Twitter: [@directorcia](https://www.twitter.com/directorcia)
- Blog: [CIAOPS](https://blog.ciaops.com)
