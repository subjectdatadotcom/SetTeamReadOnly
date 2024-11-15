
# TeamsReadonly.ps1

## Overview
The **TeamsReadonly.ps1** script is designed to process Teams data from a CSV file, identify the associated SharePoint sites, and set their lock state to **ReadOnly**. The script also verifies the lock state to ensure the operation was successful.

---

## Prerequisites
1. **Administrative Permissions**:
   - Ensure you have administrative permissions for Microsoft Teams, Exchange Online, and SharePoint Online.

2. **Required Modules**:
   - [MicrosoftTeams](https://learn.microsoft.com/en-us/powershell/module/teams/?view=teams-ps)
   - [Microsoft.Online.SharePoint.PowerShell](https://learn.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online?view=sharepoint-ps)
   - [ExchangeOnlineManagement](https://learn.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell)

   The script automatically installs these modules if they are not already installed.

3. **CSV File**:
   - The input CSV file must include at least the following columns:
     - `GroupId`
     - `Source MailNickname`

4. **Authentication**:
   - Valid credentials for Microsoft Teams, Exchange Online, and SharePoint Online are required for authentication.

---

## How to Use

### Step 1: Prepare the CSV File
- Create a CSV file (e.g., `Teams.csv`) containing the required columns (`GroupId` and `Source MailNickname`).
- Place the file in the same directory as the script or specify its path using the `-CSVFilePath` parameter.

### Step 2: Run the Script
You can execute the script in one of the following ways:

#### Default CSV Path
```powershell
.\TeamsReadonly.ps1
```
This assumes the CSV file (`Teams.csv`) is in the same directory as the script.

#### Custom CSV Path
```powershell
.\TeamsReadonly.ps1 -CSVFilePath "C:\Path\To\YourFile.csv"
```

### Step 3: Monitor Output
- The script outputs the status of each operation, including:
  - Teams being processed
  - SharePoint sites being updated
  - Errors encountered (if any)

---

## Features
- **Module Management**:
  Automatically installs and imports required PowerShell modules.

- **Connection Handling**:
  Connects to:
  - Microsoft Teams
  - Exchange Online
  - SharePoint Online Admin

- **Read-Only Lock State**:
  Sets the SharePoint sites associated with Teams to `ReadOnly`.

- **Error Handling**:
  Captures and logs exceptions during the process.

---

## Parameters
| Parameter     | Description                                                                           | Required | Default Value        |
|---------------|---------------------------------------------------------------------------------------|----------|----------------------|
| `CSVFilePath` | Path to the CSV file containing Teams data (`GroupId`, `Source MailNickname` columns).| No       | `Teams.csv` in the script's directory |

---

## Example Scenarios
1. **Default Execution**:
   Run the script with the default `Teams.csv` file:
   ```powershell
   .\TeamsReadonly.ps1
   ```

2. **Custom File Path**:
   Run the script with a specific CSV file:
   ```powershell
   .\TeamsReadonly.ps1 -CSVFilePath "C:\Data\TeamsList.csv"
   ```

---

## Notes
- Ensure you have proper permissions for the modules and services used in this script.
- The `Set-SPOSite` command will lock the SharePoint sites, making them **ReadOnly**. Double-check the site URLs before running the script.
- Always test with a small subset of Teams before processing large batches.

---

## Author
**SubjectData**

For questions or issues, feel free to contact the author.

---
