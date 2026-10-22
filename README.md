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
| [Image Model Comparison](https://directorcia.github.io/Office365/image-comparison.html) | Compares image-generation results from different models, including creation dates and image properties. |

## SharePoint Online certificate connections on Linux

[o365-connect-spo-cert-linux.ps1](./o365-connect-spo-cert-linux.ps1) connects to
the SharePoint admin center using an encrypted PFX. Keep
[o365-connect-pnp-cert-linux.ps1](./o365-connect-pnp-cert-linux.ps1) beside it:
the SPO script reuses that template for certificate generation, application
provisioning, private file storage and connection.

Requires Linux and PowerShell 7.4.6+. Connections use PnP.PowerShell 3.0.0+,
not the Windows-only SharePoint Online Management Shell. Use PnP commands such
as `Get-PnPTenantSite`, not `Get-SPOSite`, after connecting in the same session.

```powershell
# One-time setup; prompts for a PFX password and administrator device-code sign-in.
./o365-connect-spo-cert-linux.ps1 -GenerateLocalCertificate -ProvisionEntraApp -Tenant contoso.onmicrosoft.com

# Subsequent connections; prompts for the encrypted PFX password.
./o365-connect-spo-cert-linux.ps1 -UseCertificateAuth -ProfileName O365-SPO-AppAuth

# Unattended connection with a password injected by a secret manager.
./o365-connect-spo-cert-linux.ps1 -UseCertificateAuth -ProfileName O365-SPO-AppAuth -Headless -CertificatePasswordEnvironmentVariable SPO_PFX_PASSWORD
```

Provisioning grants SharePoint `Sites.FullControl.All` and Microsoft Graph
`Application.Read.All`. Certificates and profiles use private XDG data/config
directories under `ciaops/spo`; passwords are never saved in profiles.
Headless connections require preinstalled modules and a supplied password.
Specify `-AdminUrl` for tenant GUIDs, custom domains or renamed SharePoint tenants.
Use `Get-Help ./o365-connect-spo-cert-linux.ps1 -Full` for all options, including
using an existing application/PFX without provisioning.

## Microsoft Teams certificate connections on Linux

[o365-connect-tms-cert-linux.ps1](./o365-connect-tms-cert-linux.ps1) generates
encrypted certificate files, optionally provisions an Entra application, and
connects to Teams using a loaded certificate object instead of the Windows
certificate store. Requires Linux, PowerShell 7.4+ and MicrosoftTeams 5.0.0+.
Missing modules are installed for the current user without sudo; certificate
generation alone needs no modules.

```powershell
# One-time setup; prompts for a PFX password and administrator device-code sign-in.
./o365-connect-tms-cert-linux.ps1 -GenerateLocalCertificate -ProvisionEntraApp -Tenant contoso.onmicrosoft.com

# Subsequent connections; prompts for the encrypted PFX password.
./o365-connect-tms-cert-linux.ps1 -UseCertificateAuth -ProfileName O365-TMS-AppAuth

# Unattended connection with a password injected by a secret manager.
./o365-connect-tms-cert-linux.ps1 -UseCertificateAuth -ProfileName O365-TMS-AppAuth -Headless -CertificatePasswordEnvironmentVariable TMS_PFX_PASSWORD

# Run Teams commands in the same PowerShell session.
Get-CsTenant
```

Provisioning requires an administrator able to consent to Graph application
permissions and assign directory roles (for example, a Privileged Role
Administrator or Global Administrator). It assigns **Teams Administrator** and
grants Graph `Organization.Read.All`, `User.Read.All`, `Group.ReadWrite.All`,
`AppCatalog.ReadWrite.All`, `TeamSettings.ReadWrite.All`, `Channel.Delete.All`,
`ChannelSettings.ReadWrite.All` and `ChannelMember.ReadWrite.All`. These follow
[Microsoft's Teams app-only requirements](https://learn.microsoft.com/en-us/microsoftteams/teams-powershell-application-authentication),
not the broader directory permissions in the Windows script. No Skype and Teams
Tenant Admin API permissions are added. Existing app permissions and certificates
are preserved.

Certificates and profiles use private XDG data/config directories under
`ciaops/tms` (directories 700, files 600); passwords are never saved in profiles.
Custom storage directories must be dedicated to this script because their
permissions are also restricted. Headless connections require preinstalled modules
and a supplied password. Use `-CertificatePassword` to pass a SecureString from a
vault instead of an environment variable when possible.

Use `-Tenant`, `-AppId` and `-CertificateFilePath` to connect with an existing app
and PFX without provisioning. `-Organization` is an alias for `-Tenant`. Windows
profiles containing only a thumbprint require an exported PFX. Connections stay
open in the current session; they end when that PowerShell process exits.
Not all Teams cmdlets support app-only authentication, and Administrative Unit
scenarios need additional permissions; see Microsoft's documentation above.
Use `Get-Help ./o365-connect-tms-cert-linux.ps1 -Full` for all options.

Offline Linux regression checks (no tenant or external modules required):
`pwsh -NoProfile -File ./tests/o365-connect-tms-cert-linux.Tests.ps1`.
The [test script](./tests/o365-connect-tms-cert-linux.Tests.ps1) exercises real
certificate generation and Unix permissions with mocked Microsoft service calls.

## Games

- [Tigers in the Snow / Ardennes: Winter Offensive](https://directorcia.github.io/Office365/tigersinsnow/) — A browser-based retro wargame inspired by the Battle of the Bulge.

## Connect

- Twitter: [@directorcia](https://www.twitter.com/directorcia)
- Blog: [CIAOPS](https://blog.ciaops.com)
