# 📁 Email Automation Pro - Project Structure

## Overview

Email Automation Pro v2.0 - Transform emails into Excel spreadsheets and interactive BI dashboards with smart defaults and modern UI.

## 🗂️ Directory Structure

```
EMAILtoEXCELLprogram/
│
├── 📱 GUI & Entry Points
│   ├── run_gui.py              # Modern web-style GUI (main interface)
│   ├── run_gui.bat             # Windows launcher for GUI
│   ├── run_wizard.py           # CLI wizard (alternative interface)
│   └── run_app.bat             # Choose GUI or CLI
│
├── 🧠 Core Engine
│   └── core/
│       ├── engine.py           # Main execution engine with pipeline support
│       └── profile_loader.py   # Load/save/validate profiles
│
├── 🔌 Adapters (I/O)
│   └── adapters/
│       ├── graph_email.py      # Microsoft Graph (Outlook) integration
│       ├── local_email.py      # Local .eml file parsing
│       ├── excel_csv_email.py  # Excel/CSV input adapter
│       ├── excel_writer.py     # Excel output writer
│       ├── csv_writer.py       # CSV output writer
│       └── onedrive_storage.py # OneDrive upload/download
│
├── ⚙️ Jobs (Business Logic)
│   └── jobs/
│       ├── email_to_table.py   # Rule engine & smart matching
│       ├── smart_keyword_matcher.py  # Datatype detection
│       ├── excel_to_biready.py # BI transformations
│       └── append_to_master.py # Dataset merging
│
├── 📊 BI Export
│   └── bi_dashboard_export.py  # Interactive HTML dashboard generator
│
├── 📝 Configuration
│   ├── profiles/               # Profile JSON configurations
│   │   ├── example_graph.json
│   │   ├── example_local.json
│   │   └── example_advanced.json
│   └── config/
│       ├── app_settings.example.json  # Template (commit to git)
│       └── app_settings.json          # User secrets (in .gitignore)
│
├── 🔧 Utilities
│   └── permissions_diagnostic.py  # Test Graph permissions
│
├── ☁️ Azure Functions (Optional)
│   └── function_app/
│       ├── RunProfile/         # HTTP trigger to run profiles
│       ├── Notifications/      # Webhook for real-time processing
│       └── shared/             # Shared modules
│
├── 📖 Documentation
│   ├── README.md               # Main user guide (start here!)
│   ├── QUICK_START.md          # Get started in 5 minutes
│   ├── QUICK_REFERENCE.md      # Quick reference card
│   ├── GUI_GUIDE.md            # GUI-specific guide
│   ├── WORKFLOW_GUIDE.md       # Step-by-step workflows
│   ├── SMART_DEFAULTS.md       # How auto-output works
│   ├── WHATS_NEW_V2.md         # Version 2.0 features
│   ├── COMPLETE_FEATURE_LIST.md # All features documented
│   └── IMPLEMENTATION_SUMMARY_V2.md # Technical details
│
└── 🔒 Git Configuration
    ├── .gitignore              # Excludes secrets & user data
    └── requirements.txt        # Python dependencies
```

## 📦 Core Components

### Execution Flow

```
User Input (GUI/CLI)
    ↓
Profile Selection/Creation
    ↓
Core Engine (engine.py)
    ↓
┌─────────────────────────────────┐
│ 1. Load Input (adapters)        │
│    • Graph, .eml, Excel, CSV    │
└─────────────────────────────────┘
    ↓
┌─────────────────────────────────┐
│ 2. Execute Pipeline (jobs)      │
│    • email_to_table             │
│    • excel_to_biready           │
│    • append_to_master           │
└─────────────────────────────────┘
    ↓
┌─────────────────────────────────┐
│ 3. Write Output (adapters)      │
│    • Excel file                 │
│    • CSV file                   │
│    • OneDrive                   │
└─────────────────────────────────┘
    ↓
┌─────────────────────────────────┐
│ 4. BI Export (if enabled)       │
│    • Generate HTML dashboard    │
│    • Open in browser            │
└─────────────────────────────────┘
```

## 🎯 Key Files by Use Case

### "I want to process emails"
1. Launch: `run_gui.py` or `run_gui.bat`
2. Create profile using GUI
3. Select Microsoft Graph or local .eml input
4. Run profile

**Key files:**
- `run_gui.py` - GUI interface
- `adapters/graph_email.py` - Outlook integration
- `adapters/local_email.py` - .eml parsing
- `jobs/email_to_table.py` - Rule engine

### "I want to analyze existing data"
1. Launch: `run_gui.py`
2. Select Excel/CSV input
3. Auto-detect columns
4. BI Dashboard auto-selected!

**Key files:**
- `run_gui.py` - GUI interface
- `adapters/excel_csv_email.py` - Excel/CSV reading
- `bi_dashboard_export.py` - Dashboard generation

### "I want to customize rules"
1. Edit profile JSON in `profiles/` folder
2. Add rules with smart matching
3. Use datatypes for auto-detection

**Key files:**
- `profiles/*.json` - Your configurations
- `jobs/email_to_table.py` - Rule engine
- `jobs/smart_keyword_matcher.py` - Datatype detection

### "I want to deploy to Azure"
1. Configure `function_app/`
2. Deploy to Azure Functions
3. Use HTTP trigger or webhooks

**Key files:**
- `function_app/RunProfile/` - HTTP trigger
- `function_app/Notifications/` - Webhook
- `function_app/shared/` - Shared code

## 🔐 Security & Secrets

### Files in .gitignore (NOT committed)
- `config/app_settings.json` - Your Graph credentials
- `token_cache.bin` - Auth tokens
- `processed_emails.json` - Runtime state
- `output/` - Generated files
- `input_emails/` - User emails
- `*.xlsx`, `*.csv` - Data files

### Files in Git (Safe to commit)
- `config/app_settings.example.json` - Template only
- All source code
- All documentation
- Example profiles (no real data)

## 📊 Dependencies

### Core Requirements (`requirements.txt`)
```
msal>=1.20.0              # Microsoft authentication
requests>=2.28.0          # HTTP requests
openpyxl>=3.0.10          # Excel operations
python-dateutil>=2.8.2    # Date parsing
```

### Optional Dependencies
```
azure-functions           # For Azure deployment
```

## 🚀 Getting Started

### Quick Setup
```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Configure Graph API (optional)
copy config\app_settings.example.json config\app_settings.json
# Edit with your Azure app details

# 3. Launch GUI
python run_gui.py
```

### First Profile
1. Click "New Profile"
2. Select input source
3. Add columns (or auto-detect)
4. Output auto-selected based on input!
5. Run profile

## 📝 Profile Structure

Profiles are JSON files in `profiles/` directory:

```json
{
  "name": "My Profile",
  "input_source": "graph|local_eml|excel_file|csv_file",
  "email_selection": { ... },
  "schema": {
    "columns": [
      {"name": "Subject", "type": "text"},
      {"name": "From", "type": "text"}
    ]
  },
  "rules": [ ... ],
  "output": {
    "format": "excel|csv",
    "destination": "local|onedrive|bi_dashboard"
  }
}
```

## 🎨 UI Components

### Modern GUI (`run_gui.py`)
- **GradientFrame** - Canvas-based gradient backgrounds
- **ModernCard** - Card layout with shadows
- **ModernButton** - Hover effects and animations
- **Status Indicators** - Real-time connection status
- **Activity Log** - Color-coded execution log

### Themes
- Windows: Segoe UI
- macOS: SF Pro
- Linux: Ubuntu

## 🧪 Testing

### Manual Testing
```bash
# Run GUI
python run_gui.py

# Run CLI wizard
python run_wizard.py

# Test permissions
python permissions_diagnostic.py
```

### Unit Tests (Future)
```bash
# Coming soon
pytest tests/
```

## 📖 Documentation Guide

### For New Users
1. **Start here:** README.md
2. **Quick setup:** QUICK_START.md
3. **Step-by-step:** WORKFLOW_GUIDE.md
4. **Quick tips:** QUICK_REFERENCE.md

### For Developers
1. **Architecture:** IMPLEMENTATION_SUMMARY_V2.md
2. **Features:** COMPLETE_FEATURE_LIST.md
3. **Smart defaults:** SMART_DEFAULTS.md
4. **Project structure:** This file!

### For IT/Admin
1. **Permissions:** permissions_diagnostic.py output
2. **Azure setup:** function_app/README.md
3. **Security:** .gitignore review
4. **Requirements:** requirements.txt

## 🔄 Version History

### v2.0 (Current)
- ✅ Smart input/output defaults
- ✅ Modern web-style UI
- ✅ Interactive BI dashboards
- ✅ Excel/CSV input support
- ✅ Smart keyword matching
- ✅ Datatype detection
- ✅ Profile management (delete, etc.)

### v1.0
- Basic email → Excel
- Graph integration
- Rule engine
- OneDrive upload

## 🤝 Contributing

This is an internal Sanofi project. For improvements:
1. Create feature branch
2. Test thoroughly
3. Update documentation
4. Submit for review

## 📞 Support

- **Teams:** #email-automation
- **Email:** automation-team@sanofi.com
- **Wiki:** [Internal Sanofi docs]

## 📄 License

Internal use only - Sanofi Confidential

---

**Made with ❤️ by the Sanofi Automation Team**
**Version 2.0 - January 2024**
