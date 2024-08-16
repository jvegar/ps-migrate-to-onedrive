### Prerequisites

## Install PnP Module

```Powershell
Install-Module -Name PnP.PowerShell

```

## Configure app permission to site level

```XML
<AppPermissionRequests AllowAppOnlyPolicy="true">
 <AppPermissionRequest Scope="http://sharepoint/content/sitecollection/web" Right="FullControl" />
</AppPermissionRequests>
```

### Executing script

```Shell
./main.ps1
```
