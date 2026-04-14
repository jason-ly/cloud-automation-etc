# cloud-automation-etc

A collection of scripts, tools, and projects from my work and learning in cloud security and automation. Content spans Microsoft 365, Azure, AWS, Terraform, Docker, and whatever else I'm currently building or breaking.

## Background
I'm a junior cybersecurity engineer working across identity governance, Microsoft 365 security, and cloud automation. This repo is where I document things I've built, learned, or figured out the hard way.

## What's here

### sanitized-work-scripts
Real scripts from production work, sanitized for public sharing. These solve actual problems I'd encountered in enterprise M365 and Azure environments.

| Script | Description |
|---|---|
| `Set-ExchangeAppMailSend.ps1` | Configures Exchange Online RBAC for Applications to scope `Mail.Send` permissions for an app registration to a specific mailbox. Supports setup and teardown modes. |

## Prerequisites
Scripts in `sanitized-work-scripts` generally require:
- Exchange Online PowerShell (`ExchangeOnlineManagement` module)
- Microsoft Graph PowerShell SDK (`Microsoft.Graph` module)
- Appropriate Entra ID and Exchange Online permissions

Each script's parameters are documented inline.
