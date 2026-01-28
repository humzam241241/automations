# 📧 Email Automation Pro v2.0

Transform emails into Excel spreadsheets and interactive BI dashboards automatically - with or without Microsoft Graph access!

## ✨ What's New in v2.0

### 🎯 Smart Input/Output Defaults
- **Email input (.eml, Graph)** → Automatically creates Excel spreadsheet
- **Excel/CSV input** → Automatically creates BI dashboard
- **Pipeline mode** → Email → Excel → BI Dashboard (all-in-one!)

### 📊 Interactive BI Dashboards
- Beautiful HTML dashboards with Chart.js
- Auto-generates charts based on your data:
  - Pie charts for categories
  - Bar charts for distributions
  - Line charts for trends
  - Timeline charts for dates
- Opens automatically in your browser
- No Excel required to view!

### 🔍 Smart Keyword Matching
Search for **"date"** and find:
- The word "date" anywhere
- 2024-01-15, January 15, 2024
- All date formats automatically!

Search for **"amount"** and find:
- The word "amount"
- $100, €50, £25
- 1,234.56
- All currency and number formats!

### 💻 Modern Web-Style UI
- Apple-inspired clean design
- Gradient backgrounds
- Card-based layout
- Real-time activity log
- Profile management
- Delete profiles with one click

### 📁 Multiple Input Sources
1. **Microsoft Graph** - Direct from Outlook (requires permissions)
2. **Local .eml files** - Export emails from any client
3. **Excel files (.xlsx)** - Process existing spreadsheets
4. **CSV files** - Standard CSV format

### 🎨 Auto-Column Detection
- Upload Excel/CSV file
- Columns detected automatically
- No manual typing needed!

## 🚀 Quick Start

### Option 1: Use the GUI (Recommended)

```bash
python run_gui_v2.py
```

Or just double-click **`run_gui.bat`**

### Option 2: Use the Wizard (CLI)

```bash
python run_wizard.py
```

## 📋 Installation

1. **Install Python 3.8+**
2. **Install dependencies:**

```bash
pip install -r requirements.txt
```

3. **Configure Graph API (optional):**

Copy `config/app_settings.example.json` to `config/app_settings.json` and fill in your Azure app details:

```json
{
  "client_id": "your-app-client-id",
  "tenant_id": "your-tenant-id",
  "client_secret": "your-secret-if-using-azure-function"
}
```

**Don't have Graph access?** No problem! Use local .eml files or Excel/CSV files instead.

## 🎯 Usage Examples

### Example 1: Process Emails → Create Excel

**Input:** 25 emails from Outlook folder
**Output:** Excel spreadsheet with columns for Subject, From, Date, Keywords

1. Open GUI
2. Click "New Profile"
3. Select **"Microsoft Graph"** as input
4. Enter columns: `Subject,From,Date,Priority`
5. Choose **"Excel File"** as output
6. Run profile

Result: `output/emails_20240115.xlsx`

### Example 2: Excel File → BI Dashboard

**Input:** Existing Excel file with email data
**Output:** Interactive HTML dashboard with charts

1. Open GUI
2. Click "New Profile"
3. Select **"Excel File (.xlsx)"** as input
4. Browse to your file
5. Check **"Auto-detect columns"**
6. Output automatically sets to **"BI Dashboard"**
7. Run profile

Result: Dashboard opens in browser automatically! 🎉

### Example 3: End-to-End Pipeline

**Input:** Emails from Inbox
**Output:** Excel + BI Dashboard

1. Open GUI
2. Click "New Profile"
3. Select **"Microsoft Graph"** as input
4. Enter columns or load from template
5. Choose **"Excel + BI Dashboard"** as output
6. Run profile

Result:
- `output/emails_20240115.xlsx` (Excel file)
- `output/dashboards/dashboard_20240115.html` (Interactive dashboard)

## 🔐 Microsoft Graph Permissions

### Required Permissions

| Permission | Type | Purpose |
|------------|------|---------|
| `User.Read` | Delegated | Identity check |
| `Mail.Read` | Delegated/Application | Read emails |
| `Files.ReadWrite` | Delegated/Application | Upload to OneDrive |

### Check Your Permissions

```bash
python permissions_diagnostic.py
```

This will test:
- ✅ Can access /me
- ✅ Can read mail
- ✅ Can access OneDrive

And tell you exactly which permissions to request from IT!

## 📁 Working Without Graph (Local Mode)

### Export Emails from Outlook

1. **Select emails** in Outlook
2. **Drag and drop** to a folder
3. **Files saved as .eml**
4. Point the tool to that folder!

```python
{
  "name": "Process Local Emails",
  "input_source": "local_eml",
  "email_selection": {
    "directory": "./input_emails",
    "pattern": "*.eml"
  },
  ...
}
```

### Export from Gmail

1. Open email
2. Click **⋮** (More)
3. Choose **Download message**
4. Saves as .eml file!

## 🎨 Profile Configuration

Profiles are JSON files stored in `profiles/` directory.

### Minimal Profile Example

```json
{
  "name": "Simple Email Export",
  "input_source": "graph",
  "email_selection": {
    "folder_name": "Inbox",
    "newest_n": 25
  },
  "schema": {
    "columns": [
      {"name": "Subject", "type": "text"},
      {"name": "From", "type": "text"},
      {"name": "Date", "type": "date"}
    ]
  },
  "output": {
    "format": "excel",
    "destination": "local",
    "local_path": "./output"
  }
}
```

### Advanced Profile with Rules

```json
{
  "name": "Priority Email Tracking",
  "input_source": "graph",
  "email_selection": {
    "folder_name": "Inbox",
    "search_query": "hasAttachments:true"
  },
  "schema": {
    "columns": [
      {"name": "Subject", "type": "text"},
      {"name": "From", "type": "text"},
      {"name": "Date", "type": "date"},
      {"name": "Priority", "type": "text"},
      {"name": "Has_Invoice", "type": "text"}
    ]
  },
  "rules": [
    {
      "column": "Priority",
      "match_type": "keyword",
      "keywords": ["urgent", "important", "asap"],
      "search_in": ["subject", "body"],
      "value_if_matched": "High",
      "priority": 10
    },
    {
      "column": "Has_Invoice",
      "match_type": "datatype",
      "datatype": "number",
      "search_in": ["attachments_text"],
      "value_if_matched": "Yes"
    }
  ],
  "output": {
    "format": "excel",
    "destination": "both",
    "local_path": "./output",
    "also_export_bi": true
  }
}
```

## 🔧 Advanced Features

### Pipeline Chaining

Process data through multiple steps:

```json
{
  "pipeline": [
    "email_to_table",
    "excel_to_biready",
    "append_to_master"
  ]
}
```

### Master Dataset

Append new data to existing dataset with deduplication:

```json
{
  "master_dataset": {
    "path": "./master/all_emails.xlsx",
    "deduplicate_column": "Subject"
  }
}
```

### BI Transformations

Clean and transform data:

```json
{
  "bi_transformations": [
    {"type": "trim_whitespace"},
    {"type": "remove_empty_rows"},
    {"type": "cast_dates", "columns": ["Date"]}
  ]
}
```

### Smart Keyword Rules

#### Exact Word Match (Default)

```json
{
  "match_type": "keyword",
  "keywords": ["invoice"],
  "word_boundary": true
}
```

Finds: "invoice", "Invoice", "INVOICE"
Skips: "invoices", "invoice123"

#### Substring Match

```json
{
  "match_type": "keyword",
  "keywords": ["invoice"],
  "substring": true
}
```

Finds: "invoice", "invoices", "invoice123", "pro-invoice"

#### Datatype Match

```json
{
  "match_type": "datatype",
  "datatype": "date"
}
```

Automatically finds:
- 2024-01-15
- Jan 15, 2024
- 15/01/2024
- All date formats!

Available datatypes:
- `date` - Any date format
- `number` - Numbers and currencies
- `email` - Email addresses
- `url` - Web URLs
- `phone` - Phone numbers

#### Search Scopes

```json
{
  "search_in": ["subject", "body"]
}
```

Options:
- `subject` - Email subject only
- `body` - Email body only
- `from` - Sender address
- `to` - Recipient addresses
- `attachments_text` - Text extracted from attachments
- `attachments_names` - Attachment filenames
- `all` - Search everywhere (default)

## 📊 BI Dashboard Features

Auto-generated dashboards include:

### 📈 Charts
- **Doughnut charts** for categorical data (Priority, Status, etc.)
- **Bar charts** for top senders/recipients
- **Line charts** for trends over time
- **Timeline charts** for date-based data

### 📋 Data Table
- First 100 rows displayed
- Sortable columns
- Badge highlighting for Yes/No values
- Responsive design

### 📱 Responsive
- Works on desktop, tablet, and mobile
- Modern web design
- Gradient backgrounds
- Card-based layout

## 🏗️ Architecture

```
EMAILtoEXCELLprogram/
├── core/
│   ├── engine.py           # Main execution engine
│   └── profile_loader.py   # Profile management
├── adapters/
│   ├── graph_email.py      # Microsoft Graph
│   ├── local_email.py      # .eml file parsing
│   ├── excel_csv_email.py  # Excel/CSV input
│   ├── excel_writer.py     # Excel output
│   ├── csv_writer.py       # CSV output
│   └── onedrive_storage.py # OneDrive upload
├── jobs/
│   ├── email_to_table.py   # Rule engine
│   ├── smart_keyword_matcher.py  # Advanced matching
│   ├── excel_to_biready.py # BI transformations
│   └── append_to_master.py # Dataset merging
├── profiles/               # Profile configurations
├── output/                 # Output files
│   └── dashboards/         # HTML dashboards
└── run_gui_v2.py           # Modern GUI
```

## 🤝 Contributing

This is a Sanofi internal project. For questions or improvements, contact the automation team.

## 📝 License

Internal use only - Sanofi Confidential

## 🆘 Troubleshooting

### "Microsoft Graph permission denied"

**Solution:** Use local .eml files or CSV/Excel input instead:
1. Export emails from Outlook (drag to folder)
2. Use "Local .eml Files" as input source
3. No permissions needed!

### "Font error in GUI"

**Fixed in v2.0!** Auto-detects your OS and uses appropriate font:
- Windows: Segoe UI
- macOS: SF Pro
- Linux: Ubuntu

### "Can't see my columns in Excel input"

**Solution:** Check "Auto-detect columns" when selecting the file!

### "BI Dashboard shows no charts"

**Reason:** Data might not have categorical or numeric columns suitable for charts.

**Solution:** Dashboard still shows the data table - perfect for reviewing data!

### "Profile won't save"

**Check:**
1. Profile name is filled
2. At least one column is specified (or auto-detect is enabled)
3. Input source is selected

## 🎓 Training Videos

_(Coming soon: Internal Sanofi training portal)_

## 📞 Support

- **Internal Wiki:** [Link to Sanofi automation docs]
- **Teams Channel:** #email-automation
- **Email:** automation-team@sanofi.com

---

Made with ❤️ by the Sanofi Automation Team
