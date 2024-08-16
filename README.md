### Prerequisites

## Install required modules

```Powershell module
Install-Module -Name PnP.PowerShell
```

## Set config values

1. Rename `config.psd1.sample` to `config.psd1`.
2. Update the required config values.

## Configure app permissions

Use the above XML file to set app permissions level to site.

```XML
<AppPermissionRequests AllowAppOnlyPolicy="true">
 <AppPermissionRequest Scope="http://sharepoint/content/sitecollection/web" Right="FullControl" />
</AppPermissionRequests>
```

### Executing script

```Shell
./main.ps1
```
