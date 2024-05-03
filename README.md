# Prio365 Setup Repository

This repository contains PowerShell scripts for setting up Prio365 in the Microsoft Tenant environment and managing Azure Runbooks for Prio365 tasks.

## Setup by PowerShell

The PowerShell scripts provided in this repository automate the setup process for Prio365 in the Microsoft Tenant environment. 
The scripts perform tasks such as creating resource groups, automation accounts, and configuring permissions for managing Exchange Online mailboxes.

### Usage

The Script will create a ressource Group called "prio365" and create a automation-account in west-europe. If you wanna change this. Change the PowerShell Code.

1. Ensure you have the necessary permissions and prerequisites installed, such as Azure PowerShell modules.
2. Run the PowerShell scripts in the following order:
   - `Setup-Prio365.ps1`: Sets up the Prio365 environment in the Microsoft Tenant.

### Azure Runbook Code for Power Shell Scripts
   - `AddUserToMailbox.ps1`: Adds a user to a shared mailbox in Exchange Online.
   - `RemoveUserFromMailbox.ps1`: Removes a user from a shared mailbox in Exchange Online.
   - `CreateSharedMailbox.ps1`: Creates a new shared mailbox in Exchange Online.

## License

This repository is licensed under the [Prio365 Proprietary License v1.0](LICENSE.md). See the LICENSE.md file for details.
