# Sharepoint Scripts

## PnP Module

App registration is now mandatory for PnP module authentication to Office 365.

### 1. Install PowerShell 7.x
PowerShell 7.x is required for the PnP module. Confirm you are running PowerShell 7.x by running ```$host``` and confirm the version is 7.x.x.
![image](https://github.com/user-attachments/assets/b58672d0-d3fd-4bfd-9be1-25efac138779)

If the version is 5.x.x [install the PowerShell 7.x MSI package](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.4#installing-the-msi-package).

### 2. Update/install the PnP module
To install the module fresh:
```
Install-Module -Name PnP.PowerShell
```

To date an existing installation to the latest:
```
Update-Module -Name PnP.PowerShell
```

### 3. Create the app registration

This will prompt for admin consent to grant permissions to the app registration. A global administrator is required to run this command or approve the consent request.

https://pnp.github.io/powershell/articles/registerapplication.html#automatically-create-an-app-registration-for-interactive-login

```
Register-PnPEntraIDAppForInteractiveLogin -Tenant [yourtenant].onmicrosoft.com -Interactive -ApplicationName "PnP Module" 
```

### 4. Connect using Interactive Authentication and specify the ClientId

Now when connecting to a site, specify ```-ClientId``` with the application ID from the app registration.

https://pnp.github.io/powershell/articles/authentication.html#interactive-authentication

```
Connect-PnPOnline [yourtenant].sharepoint.com -Interactive -ClientId <client id of your Entra ID Application Registration>
```
