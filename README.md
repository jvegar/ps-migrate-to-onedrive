### Prerequisites

## Install PnP Module

```Powershell
Install-Module -Name PnP.PowerShell

```

## Create self-signed certificates

```Powershell
New-PnPAzureCertificate -OutPfx pnp.pfx -OutCert pnp.cer -CertificatePassword (ConvertTo-SecureString -String "pass@word1" -AsPlainText -Force)
```

```XML
<AppPermissionRequests AllowAppOnlyPolicy="true">
 <AppPermissionRequest Scope="http://sharepoint/content/sitecollection" Right="FullControl" />
</AppPermissionRequests>
```

### Executing script

```Shell
./main.ps1
```
