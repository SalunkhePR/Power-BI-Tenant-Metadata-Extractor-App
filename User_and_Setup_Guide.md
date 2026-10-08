# Power BI Metadata Extractor

## User & Setup Guide

---

### 1. Prerequisites & System Requirements

Before using the application, please make sure the following are in place:

| Requirement                 | Details                                                                                                      |
| --------------------------- | ------------------------------------------------------------------------------------------------------------ |
| **Operating System**  | Windows 10 or Windows 11                                                                                     |
| **PowerShell**        | Built into Windows (no separate install needed)                                                              |
| **Microsoft Account** | A Power BI**Administrator** account                                                                    |
| **Tenant Setting**    | “Enhance admin API responses with detailed metadata” must be**Enabled** in the Power BI Admin portal |

> **Note:** The application is delivered as a single `.exe` file. You do not need to install Python.

---

### 2. One-Time Setup Commands

The first time you use the application on a computer, you need to install one official Microsoft component.

**Open PowerShell** and run these two commands one after the other:

```powershell
Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser -Force
```

```powershell
Import-Module MicrosoftPowerBIMgmt
```

These commands install and activate the official Microsoft PowerShell library needed for the app to securely authenticate, connect, and perform automated actions in your Power BI workspace.

- If PowerShell asks about an untrusted repository, type `Y` and press Enter.
- You only need to do this **once** per computer.

---

### 3. How to Execute

Follow these simple steps every time you want to run an extraction:

1. Double-click **PowerBI_Metadata_Extractor.exe**
2. Confirm or change the **Output Folder** (default is your Desktop or Documents folder)
3. Optionally keep the checkbox **“Open Excel file when finished”** selected
4. Click **Start Extraction**
5. When the Microsoft sign-in window appears, log in with your **Power BI Administrator** account
6. Wait for the progress bar to complete
7. When finished, a confirmation message will show the location of the generated files

The Activity Log at the bottom of the window shows live progress.

---

### 4. Output File Locations

After a successful run, the application creates two files in the folder you selected:

| File                                                    | Description                                                                                                                                      |
| ------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------ |
| **PowerBI_Complete_Architecture_Map_Output.xlsx** | Main report – formatted Excel workbook with 13 sheets (workspaces, reports, datasets, data sources, measures, M code, relationships, RLS, etc.) |
| **PowerBI_Raw_Metadata.json**                     | Technical backup of the raw metadata returned by Power BI                                                                                        |

**Default location** (if you do not change the folder):

- `Desktop\PowerBI_Metadata\`or
- `Documents\PowerBI_Metadata\`

You can choose any other folder using the **Browse…** button before starting the extraction.

---

### 5. Troubleshooting & Support

| Issue                                           | What to do                                                                                                                                                             |
| ----------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Prerequisites panel shows the module is missing | Open PowerShell and run the two setup commands from Section 2, then restart the application                                                                            |
| Microsoft sign-in window does not appear        | Close the app, run`Import-Module MicrosoftPowerBIMgmt` in PowerShell, then try again                                                                                 |
| “Access denied” or empty results              | Confirm you are signed in with a**Power BI Administrator** account and that the tenant setting “Enhance admin API responses with detailed metadata” is Enabled |
| Extraction is slow                              | Large tenants with many workspaces take longer. This is normal – please wait for the process to finish                                                                |
| Excel file cannot be opened / is locked         | Close any previous copy of the Excel file and run the extraction again                                                                                                 |
| Application does not start                      | Make sure you are using the latest`.exe` and that Windows has not blocked it (Right-click → Properties → Unblock, if available)                                    |

---

**Thank you for using Power BI Metadata Extractor.**
