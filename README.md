# Email to Spreadsheet Framework

A modular framework that transforms emails into structured spreadsheet data. Supports **offline** (manual wizard), **online** (Azure Function webhook), and **no-Graph** (local file) modes.

## Features

- ✅ **OFFLINE MODE**: Interactive wizard with Graph detection and local file fallback
- ✅ **ONLINE MODE**: Azure Function HTTP trigger for one-click profile execution
- ✅ **NO-GRAPH MODE**: Process local `.eml`, `.msg`, or CSV files without Microsoft Graph
- ✅ **Profile-based**: Reusable JSON configurations for different workflows
- ✅ **Keyword/Regex Rules**: Automatically map content to spreadsheet columns
- ✅ **Multi-output**: Excel (with template support), CSV
- ✅ **OneDrive Integration**: Auto-upload results (when Graph available)

---

## Quick Start

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Run Offline Wizard

```bash
python run_wizard.py
```

The wizard will:
- Detect if you have Microsoft Graph access
- Offer local file fallback if Graph is unavailable
- Guide you through profile creation or execution

### 3. (Optional) Test Graph Permissions

```bash
python permissions_diagnostic.py
```

This script checks which Graph permissions you have and provides remediation steps.

---

## Usage Modes

### Mode 1: Offline with Microsoft Graph

**Requirements**:
- Azure AD app registration with delegated permissions
- User account with mailbox access

**Permissions needed**:
- `Mail.Read` (read emails)
- `Files.ReadWrite` (upload to OneDrive)

**How to run**:
```bash
python run_wizard.py
```

Select "Microsoft Graph" as input source. The wizard will:
1. Authenticate via device code flow
2. Fetch emails from your Outlook folder
3. Apply keyword rules
4. Generate Excel/CSV
5. Upload to OneDrive or save locally

---

### Mode 2: Offline without Graph (Local Files)

**Requirements**: None (no Azure/Graph needed)

**How to run**:
1. Export emails to `.eml` files (from Outlook, Gmail, etc.)
2. Place them in `./input_emails/`
3. Run wizard:
   ```bash
   python run_wizard.py
   ```
4. Select "Local .eml files" as input source

**Supported formats**:
- `.eml` (standard email format)
- `.csv` (with columns: subject, from, to, date, body)

---

### Mode 3: Online (Azure Function)

**Requirements**:
- Azure Function App
- Azure AD app registration with **application permissions** (client credentials)

**Permissions needed**:
- `Mail.Read` (application permission)
- `Files.ReadWrite` (application permission)

**Deployment**:
1. Deploy the `function_app/` directory to Azure
2. Set environment variables in Azure Function App settings:
   - `TENANT_ID`
   - `CLIENT_ID`
   - `CLIENT_SECRET`
   - `MAILBOX_USER` (target mailbox)
3. Create profiles in `profiles/` directory and deploy them

**Usage**:
```bash
# Run a profile via HTTP POST
curl -X POST "https://YOUR_FUNCTION_APP.azurewebsites.net/api/run?profile=example_graph"
```

---

## Configuration

### Profile Schema

Profiles are JSON files in `profiles/` that define:

```json
{
  "name": "my_profile",
  "description": "Process billing emails",
  "input_source": "graph | local_eml | local_csv",
  "email_selection": {
    "folder_name": "Inbox",
    "newest_n": 25
  },
  "schema": {
    "columns": [
      {"name": "Subject", "type": "text"},
      {"name": "From", "type": "text"},
      {"name": "Billing", "type": "text"}
    ]
  },
  "rules": [
    {
      "column": "Billing",
      "keywords": ["invoice", "billing", "payment"],
      "value": "Yes",
      "unmatched_value": ""
    }
  ],
  "output": {
    "format": "excel",
    "destination": "onedrive",
    "onedrive_path": "/EmailReports",
    "filename_template": "report_{timestamp}.xlsx"
  }
}
```

### Example Profiles

- **`profiles/example_graph.json`**: Uses Microsoft Graph to fetch emails
- **`profiles/example_local.json`**: Uses local .eml files (no Graph)

---

## Setup: Microsoft Graph Permissions

### For Local/Offline Use (Delegated Permissions)

1. Go to **Azure Portal** → **App Registrations** → Your App
2. Navigate to **API permissions**
3. Add these **Delegated** permissions:
   - `Microsoft Graph` → `Mail.Read`
   - `Microsoft Graph` → `Files.ReadWrite`
   - `Microsoft Graph` → `User.Read`
4. If your org requires it, click **Grant admin consent**

### For Azure Function/Online Use (Application Permissions)

1. Go to **Azure Portal** → **App Registrations** → Your App
2. Navigate to **API permissions**
3. Add these **Application** permissions:
   - `Microsoft Graph` → `Mail.Read`
   - `Microsoft Graph` → `Files.ReadWrite`
4. **Grant admin consent** (required for application permissions)
5. Create a **client secret** under **Certificates & secrets**

### Environment Variables (Production)

Set these in your Azure Function App or local `.env`:

```bash
TENANT_ID=your-tenant-id
CLIENT_ID=your-client-id
CLIENT_SECRET=your-client-secret
MAILBOX_USER=user@company.com
TARGET_FOLDER_ID=folder-id-if-needed
ONEDRIVE_PATH=/EmailReports
```

**IMPORTANT**: Never commit secrets to GitHub. Use environment variables in production.

---

## Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                      EXECUTION MODES                         │
├─────────────────────────────────────────────────────────────┤
│  OFFLINE (run_wizard.py)  │  ONLINE (Azure Function)        │
│  - Interactive wizard     │  - HTTP trigger /api/run        │
│  - Graph or local files   │  - Client credentials auth      │
└─────────────────────────────────────────────────────────────┘
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                    CORE ENGINE (core/)                       │
│  - ExecutionEngine: Orchestrates job execution              │
│  - ProfileLoader: Loads/validates JSON profiles             │
└─────────────────────────────────────────────────────────────┘
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                    ADAPTERS (adapters/)                      │
│  Input:  GraphEmailAdapter | LocalEmailAdapter             │
│  Output: ExcelWriter | CSVWriter | OneDriveAdapter          │
└─────────────────────────────────────────────────────────────┘
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                      JOBS (jobs/)                            │
│  - email_to_table: Extract emails → tabular data            │
│  - excel_to_biready: Transform data for BI                  │
│  - append_to_master: Merge into master dataset              │
└─────────────────────────────────────────────────────────────┘
```

---

## Troubleshooting

### "Graph authentication failed"

**Solution**:
1. Run `python permissions_diagnostic.py` to check permissions
2. Ensure you have the required Graph permissions in Azure AD
3. If using client credentials, ensure `CLIENT_SECRET` is correct

### "No emails found"

**Solutions**:
- Check folder name spelling (case-sensitive in some cases)
- Verify folder exists in your mailbox
- Ensure you have emails in that folder
- Try using search query instead

### "Profile validation errors"

**Solutions**:
- Check JSON syntax (use a validator)
- Ensure `input_source` is one of: `graph`, `local_eml`, `local_csv`
- Verify all required fields are present

### Local .eml files not loading

**Solutions**:
- Ensure files are valid `.eml` format
- Check directory path is correct
- Verify files are readable (permissions)

---

## File Structure

```
.
├── adapters/               # Data source/output adapters
│   ├── graph_email.py      # Microsoft Graph email fetcher
│   ├── local_email.py      # Local .eml/.csv loader (no Graph)
│   ├── excel_writer.py     # Excel output
│   ├── csv_writer.py       # CSV output
│   └── onedrive_storage.py # OneDrive upload/download
├── core/                   # Core engine
│   ├── engine.py           # Main execution engine
│   └── profile_loader.py   # Profile JSON loader
├── jobs/                   # Job modules
│   ├── email_to_table.py   # Email → table transformation
│   ├── excel_to_biready.py # BI transformations
│   └── append_to_master.py # Append to master dataset
├── profiles/               # Profile configurations
│   ├── example_graph.json  # Example with Graph
│   └── example_local.json  # Example without Graph
├── function_app/           # Azure Function (online mode)
│   ├── Notifications/      # Webhook for Graph subscriptions
│   ├── RunProfile/         # HTTP trigger to run profiles
│   └── shared/             # Shared modules
├── config/                 # Configuration files
│   ├── app_settings.example.json  # Template (safe to commit)
│   └── keyword_map.json    # Legacy keyword config
├── run_wizard.py           # Offline wizard entry point
├── permissions_diagnostic.py  # Test Graph permissions
├── main.py                 # Legacy local script
└── README.md               # This file
```

---

## Migration from Old Version

If you were using `main.py` with `config.json`:

1. **Create a profile**:
   ```bash
   python run_wizard.py
   ```
   Choose option 2 (Create new profile) and answer prompts.

2. **Or manually convert** `config.json` to a profile:
   - `keywords` → `rules` with keyword matching
   - `outlook_folder` → `email_selection.folder_name`
   - `onedrive_folder` → `output.onedrive_path`

3. **Run the new wizard** instead of `main.py`

---

## Contributing

1. Add new adapters in `adapters/`
2. Add new jobs in `jobs/`
3. Create example profiles in `profiles/`
4. Update this README

---

## License

MIT License - See LICENSE file for details

---

## Support

**Common Issues**:
- Run `python permissions_diagnostic.py` for permission issues
- Check Azure AD app registration for correct permissions
- Verify environment variables are set correctly
- Use local file mode if Graph is unavailable

**Need Help?**
- Check the troubleshooting section above
- Review example profiles in `profiles/`
- Test with the diagnostic script
