# Quick Reference - New Features

## 🚀 Run the Framework

### Option 1: Using run_app.bat (Recommended)
```cmd
run_app.bat
```
- Checks dependencies
- Checks Graph configuration  
- Launches interactive wizard

### Option 2: Direct Python
```cmd
python run_wizard.py
```

---

## 📋 Rule Syntax Reference

### Basic Rule
```json
{
  "column": "Billing",
  "keywords": ["invoice", "payment"],
  "value": "Yes",
  "unmatched_value": ""
}
```

### Word-Boundary Matching (Default)
```json
{
  "column": "VPN",
  "keywords": ["vpn"],
  "word_boundary": true,
  "value": "Found"
}
```
- ✓ Matches: "Need VPN access"
- ✗ Skips: "VPNACCESS123"

### Substring Matching
```json
{
  "column": "Code",
  "keywords": ["ABC"],
  "word_boundary": false,
  "value": "Found"
}
```
- ✓ Matches: "ABC123" (substring)

### Target Specific Fields
```json
{
  "column": "Urgent",
  "keywords": ["urgent", "asap"],
  "search_in": ["subject"],
  "value": "Yes"
}
```

**Valid search_in values:**
- `["all"]` - subject + body + attachments + from + to (default)
- `["subject"]` - subject line only
- `["body"]` - email body only
- `["attachments_text"]` - attachment content
- `["attachments_names"]` - attachment filenames
- `["from"]` - sender address
- `["to"]` - recipient addresses

### Priority System
```json
[
  {
    "column": "Category",
    "keywords": ["urgent invoice"],
    "priority": 100,
    "value": "Urgent Billing"
  },
  {
    "column": "Category",
    "keywords": ["invoice"],
    "priority": 10,
    "value": "Billing"
  },
  {
    "column": "Category",
    "keywords": ["email"],
    "priority": 1,
    "value": "General"
  }
]
```
- Higher priority wins
- First match at same priority wins
- First match locks the column (no overwrites)

### Regex Matching
```json
{
  "column": "Invoice_Number",
  "regex": "INV-\\d{5,}",
  "value": "FOUND",
  "search_in": ["body"]
}
```

---

## 📊 Profile Schema

### Email Selection (Graph)

**Folder Mode:**
```json
{
  "input_source": "graph",
  "email_selection": {
    "folder_name": "Inbox",
    "newest_n": 25
  }
}
```

**Search Query Mode:**
```json
{
  "input_source": "graph",
  "email_selection": {
    "search_query": "from:billing@company.com AND subject:invoice",
    "after_date": "2026-01-01",
    "newest_n": 50
  }
}
```

### Email Selection (Local)

**.eml files:**
```json
{
  "input_source": "local_eml",
  "email_selection": {
    "directory": "./input_emails",
    "pattern": "*.eml"
  }
}
```

**CSV file:**
```json
{
  "input_source": "local_csv",
  "email_selection": {
    "csv_path": "./emails.csv"
  }
}
```

---

## 🔧 Pipeline Examples

### Simple (Default)
```json
{
  "pipeline": ["email_to_table"]
}
```

### With Master Dataset
```json
{
  "pipeline": ["email_to_table", "append_to_master"],
  "master_dataset": {
    "path": "./output/master.xlsx",
    "deduplicate_column": "Subject"
  }
}
```

### Full BI Pipeline
```json
{
  "pipeline": [
    "email_to_table",
    "excel_to_biready",
    "append_to_master"
  ],
  "bi_transformations": [
    {
      "type": "clean",
      "column": "Amount",
      "rules": {"trim": true}
    },
    {
      "type": "filter",
      "column": "From",
      "condition": "not_empty"
    }
  ]
}
```

---

## 🎯 Environment Variables

### For Graph API Access

**Windows (cmd):**
```cmd
set CLIENT_ID=your-client-id-here
set AUTHORITY=https://login.microsoftonline.com/your-tenant-id
```

**Windows (PowerShell):**
```powershell
$env:CLIENT_ID="your-client-id-here"
$env:AUTHORITY="https://login.microsoftonline.com/your-tenant-id"
```

**Unix/Mac:**
```bash
export CLIENT_ID="your-client-id-here"
export AUTHORITY="https://login.microsoftonline.com/your-tenant-id"
```

### Without Graph (Local Mode)
- Don't set CLIENT_ID/AUTHORITY
- Wizard will auto-detect and switch to local-only mode

---

## 🔍 Explain Mode

Run with explain mode to see detailed match information:

```cmd
python run_wizard.py
```

Select a profile and it will show:

```json
{
  "explanations": [
    {
      "Billing": {
        "matched": true,
        "match_type": "keyword_word",
        "matched_term": "invoice",
        "search_in": ["subject", "body"],
        "snippet": "...Please review the invoice for..."
      }
    }
  ]
}
```

---

## 📁 Schema from Template

In the wizard:
```
Define schema:
  A. Type column headers manually
  B. Load headers from template .xlsx
Choice [A/B]: B
Template .xlsx path: ./templates/my_template.xlsx
```

The wizard reads the first row of the first sheet and uses those as column headers.

---

## 🚨 Troubleshooting

### "Graph mode is disabled"
**Cause**: CLIENT_ID or AUTHORITY not set  
**Fix**: Set environment variables OR use local mode

### "Mail access: DENIED"
**Cause**: Missing Mail.Read permission  
**Fix**: Contact IT admin to grant permission

### "OneDrive access: DENIED"
**Cause**: Missing Files.ReadWrite permission  
**Fix**: Use local output OR contact IT admin

### "Profile requires Mail.Read permission"
**Cause**: Profile needs Graph but permission missing  
**Fix**: 
1. Create new profile with local input, OR
2. Contact IT admin to grant permission

### Token cache location changed
**Issue**: Need to re-authenticate after upgrade  
**Cause**: Cache moved to user directory  
**Fix**: Just re-run wizard, it will prompt for auth

---

## 💡 Tips & Best Practices

### 1. Use Word Boundaries by Default
```json
{
  "word_boundary": true
}
```
More accurate than substring matching.

### 2. Set Priorities for Important Rules
```json
{
  "column": "Urgent",
  "keywords": ["urgent"],
  "priority": 100
}
```

### 3. Target Specific Fields
```json
{
  "search_in": ["subject"]
}
```
Faster and more precise than searching everything.

### 4. Use Master Datasets for Deduplication
```json
{
  "pipeline": ["email_to_table", "append_to_master"],
  "master_dataset": {
    "path": "./output/master.xlsx",
    "deduplicate_column": "Subject"
  }
}
```

### 5. Test Locally First
Start with local .eml files to test your rules before using Graph.

---

## 📚 Example Workflows

### Workflow 1: Invoice Tracking
```json
{
  "name": "invoice_tracker",
  "input_source": "graph",
  "email_selection": {
    "search_query": "from:billing@company.com AND subject:invoice"
  },
  "schema": {
    "columns": [
      {"name": "Subject", "type": "text"},
      {"name": "From", "type": "text"},
      {"name": "Invoice_Number", "type": "text"},
      {"name": "Amount", "type": "text"}
    ]
  },
  "rules": [
    {
      "column": "Invoice_Number",
      "regex": "INV-\\d{5,}",
      "search_in": ["subject", "body"]
    },
    {
      "column": "Amount",
      "regex": "\\$[0-9,]+\\.\\d{2}",
      "search_in": ["body"]
    }
  ],
  "pipeline": ["email_to_table", "append_to_master"],
  "master_dataset": {
    "path": "./output/invoices.xlsx",
    "deduplicate_column": "Invoice_Number"
  },
  "output": {
    "format": "excel",
    "destination": "local",
    "local_path": "./output"
  }
}
```

### Workflow 2: Department Categorization
```json
{
  "name": "dept_categorizer",
  "input_source": "graph",
  "email_selection": {
    "folder_name": "Inbox",
    "newest_n": 100
  },
  "schema": {
    "columns": [
      {"name": "Subject", "type": "text"},
      {"name": "Category", "type": "text"}
    ]
  },
  "rules": [
    {
      "column": "Category",
      "keywords": ["invoice", "billing"],
      "value": "Finance",
      "priority": 10,
      "search_in": ["subject"]
    },
    {
      "column": "Category",
      "keywords": ["laptop", "password"],
      "value": "IT",
      "priority": 10,
      "search_in": ["subject"]
    },
    {
      "column": "Category",
      "keywords": ["policy", "benefits"],
      "value": "HR",
      "priority": 10,
      "search_in": ["subject"]
    }
  ],
  "output": {
    "format": "excel",
    "destination": "onedrive",
    "onedrive_path": "/Reports"
  }
}
```

---

## 🎓 Learning Path

1. **Start with local mode**
   - Place .eml files in `input_emails/`
   - Run `run_app.bat`
   - Create simple profile with basic keywords

2. **Add word boundaries**
   - Set `word_boundary: true`
   - Test precision improvement

3. **Target specific fields**
   - Use `search_in: ["subject"]`
   - See how it filters better

4. **Add priorities**
   - Create rules with different priorities
   - See first-match-wins behavior

5. **Try Graph mode**
   - Set CLIENT_ID and AUTHORITY
   - Authenticate and fetch from Outlook

6. **Use master datasets**
   - Add `append_to_master` to pipeline
   - Deduplicate by a column

7. **Build complex workflows**
   - Combine search queries + regex + priorities
   - Use full pipelines

---

## 📞 Support

- Check `CHANGES_SUMMARY.md` for detailed documentation
- See `profiles/example_*.json` for working examples
- Run `python permissions_diagnostic.py` for Graph permission issues
