# Set-ExchangeAppMailSend.ps1

Configures Exchange Online RBAC for Applications to scope `Mail.Send` permissions for an Entra ID app registration to a specific mailbox. Supports both setup and teardown.

## Background

This script came out of a problem encountered while building an automated guest user lifecycle notification workflow in Azure Automation. The workflow runs on a hybrid worker and uses an app registration with certificate-based authentication to make Microsoft Graph API calls. Part of the workflow needed to send notification emails from a shared mailbox.

The natural assumption was that granting `Mail.Send` as an application permission on the app registration in Entra ID and obtaining admin consent would be sufficient. It wasn't.

Even with `Mail.Send` consented, every attempt to send via `POST /users/{mailbox}/sendMail` returned `403 Forbidden`. The reason is that Exchange Online operates its own authorization layer independently of Entra ID. Granting `Mail.Send` in Entra tells Entra the app is allowed to send mail, but Exchange still needs to be told separately which mailboxes the app is permitted to send from. Without that, Exchange blocks the request regardless of what Entra says.

The first attempted fix was `New-ApplicationAccessPolicy`, which is the legacy mechanism for scoping app permissions to specific mailboxes. That failed because the cmdlet requires a mail-enabled security group as its scope -- shared mailboxes are not accepted as a policy scope directly.

Further research revealed that Application Access Policies are being replaced by RBAC for Applications, which is Microsoft's current recommended approach. RBAC for Applications solves the problem cleanly -- it allows you to create a management scope targeting a specific mailbox and assign the `Application Mail.Send` role to the app's service principal within that scope. The app can then send from that mailbox and only that mailbox, with no mail-enabled security group required.

An added benefit is that this approach enforces least privilege at the Exchange layer. Even though `Mail.Send` as an Entra application permission technically allows sending as any mailbox in the tenant, the RBAC assignment restricts the app to only the scoped mailbox. If the app's credentials were ever compromised, the blast radius is limited to that one mailbox.

## Prerequisites

- Exchange Online PowerShell (`ExchangeOnlineManagement` module)
- Organization Management role in Exchange Online (required to create service principals, management scopes, and role assignments)
- The app registration must already exist in Entra ID
- The target mailbox must already exist in Exchange Online

## Parameters

| Parameter | Required | Description |
|---|---|---|
| `-Mode` | Yes | `Setup` or `Teardown` |
| `-AppId` | Yes | The app registration's Application (client) ID from Entra ID |
| `-ObjectId` | Yes | The Enterprise Application Object ID from Entra ID |
| `-DisplayName` | Yes | Display name for the Exchange service principal |
| `-MailboxAddress` | Yes | Primary SMTP address of the mailbox to scope sending to |
| `-ScopeName` | Yes | A unique name for the management scope |
| `-RemoveServicePrincipal` | No | Switch used in Teardown mode to also remove the Exchange service principal. Only use when fully decommissioning the app from Exchange. |

## Usage

### Setup
```powershell
.\Set-ExchangeAppMailSend.ps1 `
    -Mode Setup `
    -AppId "your-app-id" `
    -ObjectId "your-object-id" `
    -DisplayName "your-app-display-name" `
    -MailboxAddress "your-mailbox@yourdomain.com" `
    -ScopeName "Your Scope Name"
```

### Teardown - keep service principal
```powershell
.\Set-ExchangeAppMailSend.ps1 `
    -Mode Teardown `
    -AppId "your-app-id" `
    -ObjectId "your-object-id" `
    -DisplayName "your-app-display-name" `
    -MailboxAddress "your-mailbox@yourdomain.com" `
    -ScopeName "Your Scope Name"
```

### Teardown - fully decommission
```powershell
.\Set-ExchangeAppMailSend.ps1 `
    -Mode Teardown `
    -AppId "your-app-id" `
    -ObjectId "your-object-id" `
    -DisplayName "your-app-display-name" `
    -MailboxAddress "your-mailbox@yourdomain.com" `
    -ScopeName "Your Scope Name" `
    -RemoveServicePrincipal
```

## Notes

- The script checks whether the service principal already exists in Exchange before attempting to create it, making it safe to re-run for adding additional mailbox scopes to an existing app registration.
- In Teardown mode, `-RemoveServicePrincipal` will not proceed if other role assignments still exist on the service principal. It will list the remaining assignments and exit cleanly.
- `Mail.Send` granted in Entra ID is not required when using RBAC for Applications. Exchange Online controls the permission entirely via the management role assignment. If `Mail.Send` exists on the app registration in Entra, it can be removed once RBAC for Applications is configured.

## References

- [RBAC for Applications in Exchange Online - Microsoft Learn](https://learn.microsoft.com/en-us/exchange/permissions-exo/application-rbac)
- [New-ApplicationAccessPolicy - Microsoft Learn](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/new-applicationaccesspolicy?view=exchange-ps)
