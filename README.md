# Sharepoint Migration Scripts

## PnP Module

App registration is now mandatory for PnP module authentication to Office 365.

1. Create the app registration with a Global Admin user

https://pnp.github.io/powershell/articles/registerapplication.html#automatically-create-an-app-registration-for-interactive-login

```
Register-PnPEntraIDAppForInteractiveLogin -Tenant [yourtenant].onmicrosoft.com -Interactive -ApplicationName "PnP Module" 
```

2. Connect using Interactive Authentication and specify the ClientId

https://pnp.github.io/powershell/articles/authentication.html#interactive-authentication

```
Connect-PnPOnline [yourtenant].sharepoint.com -Interactive -ClientId <client id of your Entra ID Application Registration>
```
