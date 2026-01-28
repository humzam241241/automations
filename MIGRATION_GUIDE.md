# Migration Guide: Old → New Architecture

## What Changed?

The codebase has been refactored from a monolithic script into a **modular framework** with:
- **Profile-based configuration** (JSON profiles instead of hardcoded config)
- **Multiple execution modes** (offline wizard, Azure Function, local files)
- **No-Graph fallback** (process local .eml files without Graph permissions)
- **Cleaner separation** of concerns (adapters, jobs, core engine)

---

## Quick Migration

### If you were using `main.py`:

**Old way**:
```bash
python main.py
```

**New way**:
```bash
python run_wizard.py
```

The wizard will:
1. Detect your Graph access
2. Guide you through profile creation
3. Execute the profile

---

## Converting `config.json` to Profile

### Old `config.json`:
```json
{
  "keywords": ["time", "date", "location"],
  "outlook_folder": "auto",
  "onedrive_folder": "/EmailReports",
  "template_path": "",
  "check_interval": 3600
}
```

### New Profile (`profiles/my_profile.json`):
```json
{
  "name": "my_profile",
  "description": "Migrated from config.json",
  "input_source": "graph",
  "email_selection": {
    "folder_name": "auto",
    "newest_n": 25
  },
  "schema": {
    "columns": [
      {"name": "Subject", "type": "text"},
      {"name": "From", "type": "text"},
      {"name": "Date", "type": "datetime"},
      {"name": "Time", "type": "text"},
      {"name": "Location", "type": "text"}
    ]
  },
  "rules": [
    {
      "column": "Time",
      "keywords": ["time"],
      "value": "Yes"
    },
    {
      "column": "Location",
      "keywords": ["location"],
      "value": "Yes"
    }
  ],
  "output": {
    "format": "excel",
    "destination": "onedrive",
    "onedrive_path": "/EmailReports",
    "filename_template": "email_report_{timestamp}.xlsx"
  }
}
```

**To create this automatically**:
```bash
python run_wizard.py
# Choose option 2: Create new profile
# Follow the prompts
```

---

## Converting `config/keyword_map.json` to Profile

### Old `config/keyword_map.json`:
```json
{
  "columns": [
    {
      "header": "Billing",
      "keywords": ["invoice", "billing"],
      "value": "Yes",
      "unmatched_value": ""
    }
  ],
  "attachment_column": "Attachments",
  "notes_column": "Notes"
}
```

### New Profile:
```json
{
  "name": "billing_profile",
  "input_source": "graph",
  "schema": {
    "columns": [
      {"name": "Subject", "type": "text"},
      {"name": "From", "type": "text"},
      {"name": "Billing", "type": "text"},
      {"name": "Attachments", "type": "text"},
      {"name": "Notes", "type": "text"}
    ]
  },
  "rules": [
    {
      "column": "Billing",
      "keywords": ["invoice", "billing"],
      "value": "Yes",
      "unmatched_value": ""
    }
  ],
  "output": {
    "format": "excel",
    "destination": "onedrive",
    "onedrive_path": "/EmailReports"
  }
}
```

**Note**: Attachments and Notes are now part of the schema columns.

---

## Azure Function Migration

### Old Function:
- `function_app/Notifications/__init__.py` (webhook-driven)

### New Functions:
- `function_app/Notifications/__init__.py` (unchanged - webhook)
- `function_app/RunProfile/__init__.py` (NEW - HTTP trigger for profiles)

### How to Use New Function:

**Deploy profiles** to Azure Function:
1. Upload `profiles/*.json` to your Function App
2. Set environment variables (see README)
3. Call the new endpoint:

```bash
POST https://YOUR_APP.azurewebsites.net/api/run?profile=my_profile
```

---

## Key Benefits of New Architecture

| Feature | Old | New |
|---------|-----|-----|
| **Configuration** | Hardcoded in files | Reusable JSON profiles |
| **Offline Mode** | Device code auth only | Wizard with Graph detection |
| **No Graph Access** | ❌ Not supported | ✅ Use local .eml files |
| **Multiple Workflows** | Single config | Multiple profiles |
| **Online Execution** | Webhook only | Webhook + HTTP trigger |
| **Testing** | Manual | `permissions_diagnostic.py` |

---

## File Mapping

| Old File | New File(s) | Notes |
|----------|-------------|-------|
| `main.py` | `run_wizard.py` | Interactive wizard |
| `config.json` | `profiles/*.json` | Profile-based config |
| `email_processor.py` | `adapters/graph_email.py`, `jobs/email_to_table.py` | Split by concern |
| `excel_generator.py` | `adapters/excel_writer.py` | Reusable adapter |
| `onedrive_uploader.py` | `adapters/onedrive_storage.py` | Reusable adapter |
| `auth.py` | `run_wizard.py` (integrated) | Used internally |

---

## Common Migration Issues

### "My old scripts stopped working"
- Use `run_wizard.py` instead of `main.py`
- Old files are still there, but the new framework is recommended

### "I don't have Graph permissions"
- Use **no-Graph mode** with local .eml files
- Export emails from Outlook to .eml format
- Place in `input_emails/` directory

### "My Azure Function isn't working"
- Make sure you deployed the new `RunProfile/` function
- Update environment variables to use the new names
- Test with `permissions_diagnostic.py`

---

## Backward Compatibility

The old files (`main.py`, `config.json`) are **still present** and will work, but:
- Not actively maintained
- Missing new features (no-Graph mode, profiles, etc.)
- **Recommended**: Migrate to the new framework

---

## Next Steps

1. **Test locally**:
   ```bash
   python run_wizard.py
   ```

2. **Create your first profile**:
   - Use the wizard or copy `profiles/example_graph.json`

3. **Test Graph permissions**:
   ```bash
   python permissions_diagnostic.py
   ```

4. **Deploy to Azure** (optional):
   - Update environment variables
   - Deploy `function_app/`
   - Upload profiles

---

## Need Help?

- See [README.md](README.md) for full documentation
- Check `profiles/example_*.json` for examples
- Run `permissions_diagnostic.py` for permission issues
