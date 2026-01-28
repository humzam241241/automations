# 🚀 Quick Reference Card - Email Automation Pro v2.0

## ⚡ 30-Second Start

```bash
python run_gui_v2.py
```

## 🎯 Smart Defaults (Auto-Magic!)

| Input | Auto Output | Override? |
|-------|-------------|-----------|
| 📧 Emails | Excel | Yes |
| 📊 Excel | BI Dashboard | Yes |
| 📄 CSV | BI Dashboard | Yes |

## 🔄 Common Workflows

### 1. Daily Email → Excel
```
Graph → Inbox → Auto: Excel ✅
```

### 2. Analyze Data → Dashboard
```
Excel file → Auto: BI Dashboard ✅
```

### 3. Full Pipeline
```
Emails → Excel + BI Dashboard ✅
```

## 🎨 GUI Shortcuts

| Action | How |
|--------|-----|
| New profile | Click "New" |
| Run profile | Select + Click "Run" |
| Delete profile | Select + Click "Delete" |
| Refresh list | Click "Refresh" |
| Get help | Click "?" |

## 📋 Profile Creation (4 Steps)

1. **Name** → Enter profile name
2. **Input** → Select source
3. **Columns** → Type or auto-detect
4. **Output** → Auto-selected! ✅

## 🔍 Smart Matching

### Find Dates
```json
{"match_type": "datatype", "datatype": "date"}
```
Finds: 2024-01-15, Jan 15, 2024, etc.

### Find Numbers
```json
{"match_type": "datatype", "datatype": "number"}
```
Finds: $100, €50, 1,234.56, etc.

### Find Exact Words
```json
{"match_type": "keyword", "keywords": ["invoice"]}
```
Finds: invoice, Invoice, INVOICE

## 📊 Input Sources

| Source | When to Use |
|--------|-------------|
| 📧 Graph | Have permissions |
| 📁 .eml | No permissions |
| 📊 Excel | Existing data |
| 📄 CSV | Standard format |

## 🎯 Output Options

| Output | Best For |
|--------|----------|
| Excel | Archiving, editing |
| BI Dashboard | Analysis, sharing |
| Both | Complete workflow |

## 🚨 Status Indicators

- 🟢 **Green** = Available
- 🔴 **Red** = Unavailable

Check before running!

## 📁 File Locations

```
profiles/           → Your profiles
output/             → Excel files
output/dashboards/  → BI dashboards
```

## 🔧 Quick Fixes

### "No emails found"
→ Check folder name or use .eml files

### "Permission denied"
→ Use local .eml or Excel input

### "No charts in dashboard"
→ Data table still works!

### "Can't save profile"
→ Check name and columns filled

## 💡 Pro Tips

1. **Auto-detect columns** for Excel/CSV
2. **Use datatypes** for smart matching
3. **Check status** before running
4. **Save profiles** for reuse
5. **Share dashboards** as HTML

## 🎓 Learning Path

**Day 1:** Try Excel → BI Dashboard
**Day 2:** Create email profile
**Day 3:** Add smart rules
**Day 4:** Use pipeline mode

## 📞 Help

- **Docs:** README_v2.md
- **Workflows:** WORKFLOW_GUIDE.md
- **Features:** COMPLETE_FEATURE_LIST.md
- **Support:** automation-team@sanofi.com

## ⌨️ Keyboard Tips

- **Enter** = Confirm
- **Esc** = Cancel
- **Tab** = Next field
- **F1** = Help (coming soon)

## 🎯 Common Profiles

### Daily Reports
```
Input: Graph (Inbox)
Columns: Subject, From, Date
Output: Excel
```

### Invoice Tracking
```
Input: Graph (Invoices folder)
Rules: Find amounts (datatype: number)
Output: Excel + BI
```

### Data Analysis
```
Input: Excel file
Auto-detect: ✅
Output: BI Dashboard
```

## 📊 Dashboard Features

- 📈 Auto charts
- 📋 Data table
- 🎨 Modern design
- 📱 Responsive
- 🔗 Shareable

## 🔄 Update Workflow

```bash
git pull origin main
pip install -r requirements.txt
python run_gui_v2.py
```

## 🎉 Quick Wins

**Fastest result:**
1. Open GUI
2. Select Excel file
3. Auto-detect
4. Run
5. Dashboard opens!

**Time: 20 seconds** ⚡

---

## 📋 Cheat Sheet

### Profile JSON Template
```json
{
  "name": "My Profile",
  "input_source": "graph",
  "email_selection": {
    "folder_name": "Inbox",
    "newest_n": 25
  },
  "schema": {
    "columns": [
      {"name": "Subject", "type": "text"},
      {"name": "From", "type": "text"}
    ]
  },
  "rules": [],
  "output": {
    "format": "excel",
    "destination": "local",
    "local_path": "./output"
  }
}
```

### Rule Template
```json
{
  "column": "Priority",
  "match_type": "keyword",
  "keywords": ["urgent"],
  "search_in": ["subject", "body"],
  "value_if_matched": "High",
  "priority": 100
}
```

### Datatype Rule
```json
{
  "column": "Date_Found",
  "match_type": "datatype",
  "datatype": "date",
  "value_if_matched": "Yes"
}
```

---

## 🎯 Remember

✅ **Smart defaults** = Less work
✅ **Auto-detect** = No typing
✅ **Datatypes** = Smart matching
✅ **Pipeline** = Both outputs
✅ **Status dots** = Check first

---

**Keep this card handy!** 📌

*Email Automation Pro v2.0*
*Made with ❤️ by Sanofi Automation Team*
