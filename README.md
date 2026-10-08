# Power BI Complete Metadata Extractor — v5.2

Extract a full architecture map of your Power BI / Microsoft Fabric tenant into a professionally formatted Excel workbook + raw JSON.

**What you get (13 sheets):**

| Sheet | Content |
|-------|---------|
| Workspaces | ID, name, type, state |
| Reports | Name, dataset link, creator, endorsement, dates |
| Dashboards + Dashboard Tiles | All tiles with pinned report links |
| Datasets | Storage mode, configured-by, schema status |
| Data Sources & Lineage | Source type + connection string (parsed from M code) |
| Tables & Columns | Including calculated columns & DAX formulas |
| Measures (DAX) | Full measure expressions |
| **M Expressions** | **Full Advanced Editor code** for every table + parameters |
| RLS Roles | Roles, table filters, members |
| Model Relationships | From DAX `INFO.VIEW.RELATIONSHIPS()` + Scanner fallback |
| PQ Implicit Joins | `Table.NestedJoin` / `Table.Join` found inside M |
| Artifact Lineage | Dependencies between artifacts |

---

## Prerequisites

### 1. Software (on the machine that will run the tool)

| Requirement | Details |
|-------------|---------|
| **Windows 10 / 11** | Required (uses PowerShell for interactive Microsoft login) |
| **Python 3.9+** | Recommended 3.10 or 3.11 |
| **Python packages** | See installation below |
| **PowerShell module** | `MicrosoftPowerBIMgmt` |

### 2. Power BI / Fabric Permissions & Tenant Settings (Critical)

The account you sign in with **must be a Power BI Administrator** (or Microsoft 365 Global Admin / Fabric Admin).

Go to **Power BI Service → ⚙ Settings → Admin portal → Tenant settings** and ensure the following are **Enabled**:

| Tenant Setting | Why needed |
|----------------|------------|
| **Enhance admin API responses with detailed metadata** | Required for schema, M expressions, lineage |
| **Allow Admin APIs** / Service principals can call Admin APIs | Core scanner endpoint |
| Developer settings → Allow users to create and use apps | Future-proofing |

Without the first setting the Admin Scanner returns almost empty results.

### 3. Installation

Open **PowerShell** or **Command Prompt** and run:

```powershell
# 1. Install Python packages
pip install requests pandas openpyxl customtkinter

# 2. Install Power BI PowerShell module
Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser -Force
```

---

## How to Run

### Option A — Desktop GUI (Recommended)

```powershell
cd path\to\PowerBI_Metadata_Extractor_App
python app.py
```

1. The app window opens.
2. Choose (or keep) the **Output Folder**.
3. Optionally tick **Open Excel when finished**.
4. Click **Start Extraction**.
5. A Microsoft login popup appears → sign in with a **Power BI Admin** account.
6. Watch the progress bar and live log.
7. When finished the Excel + JSON files are written to the chosen folder.

### Option B — Command Line

```powershell
python core_extractor.py

# Or with custom folder
python core_extractor.py -o "C:\MyReports\PBI_Metadata" --open
```

| Argument | Description |
|----------|-------------|
| `-o` / `--output` | Output folder (default = Desktop or Documents `\PowerBI_Metadata`) |
| `--open` | Open the Excel file when done (Windows only) |

---

## Output Files

By default the tool creates:

```
%USERPROFILE%\Desktop\PowerBI_Metadata\
    ├── PowerBI_Complete_Architecture_Map_Output.xlsx
    └── PowerBI_Raw_Metadata.json
```

(If Desktop is not writable it falls back to Documents, then to the current directory.)

You can change the folder in the GUI or with the `-o` flag.

---

## What the Tool Does Under the Hood

1. **Authentication** – Interactive Microsoft login via `MicrosoftPowerBIMgmt` module → obtains access token.
2. **Phase 1 – Admin Scanner API**  
   - Lists all workspaces (excludes Personal groups).  
   - Scans in batches of 100.  
   - Requests full lineage + datasource details + dataset schema + expressions.
3. **Deep parsing**  
   - Extracts every M expression (Advanced Editor code).  
   - Resolves Power Query parameters.  
   - Detects 15+ connector types (SQL, Snowflake, BigQuery, Databricks, Fabric, SharePoint, etc.).  
   - Finds implicit joins (`Table.NestedJoin` / `Table.Join`).  
   - Tries live DAX `INFO.VIEW.RELATIONSHIPS()` for accurate model relationships.
4. **Phase 2** – Supplemental standard API call for any missing datasources.
5. **Export** – Styled Excel (frozen headers, auto-width, wrap on long columns) + raw JSON.

---

## Common Errors & Fixes

| Error message / symptom | Cause | Solution |
|-------------------------|-------|----------|
| `[Auth] Login cancelled or failed` | User closed popup or module missing | Re-run `Install-Module MicrosoftPowerBIMgmt …` and try again |
| HTTP 401 / 403 on admin endpoints | Missing Admin role or tenant setting off | Grant Power BI Admin role + enable “Enhance admin API responses…” |
| Empty / almost empty Excel | Tenant setting disabled or no access | Same as above |
| Very slow | Large tenant (hundreds of workspaces) | Normal – batches of 100; just wait |
| `No module named 'customtkinter'` | GUI dependency missing | `pip install customtkinter` |
| Path / permission error | Folder not writable | Choose another output folder |

---

## Project Structure

```
PowerBI_Metadata_Extractor_App/
├── app.py                 ← Desktop GUI (CustomTkinter)
├── core_extractor.py      ← Core engine (also usable as CLI)
└── README.md              ← This file
```

---

## Tips

- Run the tool during off-peak hours on very large tenants.
- Keep the Excel file closed while the tool is writing (otherwise you get a permission error).
- The **M Expressions** sheet is extremely useful for documentation, migration, or impact analysis.
- Personal workspaces (`My Workspace`) are intentionally excluded by the Admin API filter.

---

## License / Disclaimer

This is an internal / community utility.  
Use at your own risk. Always test on a non-production tenant first.  
Requires appropriate administrative permissions in your Microsoft 365 / Power BI environment.
