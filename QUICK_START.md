# ⚡ Quick Start Guide

## 🎯 Just Want to Get Started?

### For Windows Users (Easiest)

1. **Double-click** `run_gui.bat`
2. The GUI will launch automatically
3. Click **➕ Create New** to make your first profile
4. Fill out the 5 tabs
5. Click **💾 Save & Close**
6. Select your profile and click **▶ Run**

**That's it!** 🎉

---

## 🖥️ All Platforms

### Option 1: GUI (Recommended)
```bash
python run_gui.py
```

### Option 2: Command-line Wizard
```bash
python run_wizard.py
```

### Option 3: Batch Menu (Windows)
```cmd
run_app.bat
```
Choose option 1 for GUI or option 2 for command-line.

---

## 🚦 First Time Setup

### If You Have Graph Access
1. Set environment variables:
   ```cmd
   set CLIENT_ID=your-azure-app-client-id
   set AUTHORITY=https://login.microsoftonline.com/your-tenant-id
   ```
2. Launch GUI → It will detect and connect to Graph
3. Create profiles with Outlook/OneDrive

### If You DON'T Have Graph Access
1. Don't set any environment variables
2. Launch GUI → It will work in local mode
3. Use `.eml` files or CSV as input
4. Save output to local files

**You can use this framework with or without Microsoft Graph!**

---

## 📝 Your First Profile

### Example: Extract Billing Emails

**Via GUI:**
1. Launch GUI
2. Click **➕ Create New**
3. Fill in:
   - **Name**: `billing_emails`
   - **Input**: Microsoft Graph → Folder: `Inbox`
   - **Schema**: `Subject,From,Date,Amount`
   - **Rules**: (use the pre-filled example)
   - **Output**: Local file → `./output`
4. Save and run!

**Result**: Excel file with billing emails in `./output/`

---

## 🎓 What's Next?

### Learn the Basics
- Read [GUI_GUIDE.md](GUI_GUIDE.md) for detailed GUI walkthrough
- Read [QUICK_REFERENCE.md](QUICK_REFERENCE.md) for rule syntax
- Check [profiles/](profiles/) folder for example profiles

### Try Advanced Features
- Word-boundary keyword matching
- Target specific fields with `search_in`
- Priority system for rules
- Pipeline with master datasets
- Load schema from template .xlsx

### Get Help
- **Activity Log** in GUI shows what's happening
- **CHANGES_SUMMARY.md** - Technical documentation
- **README.md** - Comprehensive guide

---

## 🐛 Troubleshooting

### "GUI won't launch"
```bash
# Install dependencies
python -m pip install -r requirements.txt

# Try command-line instead
python run_wizard.py
```

### "Graph authentication failed"
Either:
- **Set** `CLIENT_ID` and `AUTHORITY` environment variables, OR
- **Use local mode** (works without Graph)

### "Rules not matching"
- Check the Activity Log for errors
- Try command-line wizard with explain mode: `python run_wizard.py`
- Read [QUICK_REFERENCE.md](QUICK_REFERENCE.md) for rule syntax

---

## 🎯 Common Use Cases

### Use Case 1: Track Invoices
- **Input**: Outlook folder with invoices
- **Rules**: Extract invoice numbers with regex
- **Output**: Excel with invoice tracking

### Use Case 2: Categorize Emails
- **Input**: Inbox
- **Rules**: Keywords to categorize (IT, HR, Finance)
- **Output**: Excel with categories

### Use Case 3: Process Exported Emails
- **Input**: Local .eml files
- **Rules**: Any keyword/regex rules
- **Output**: Local Excel (no Graph needed!)

---

## 🚀 Pro Tips

1. **Start local, go online later**
   - Test with .eml files first
   - Once rules work, switch to Graph

2. **Use templates**
   - Create template .xlsx with your columns
   - Load headers automatically

3. **Check the log**
   - GUI Activity Log shows everything
   - Watch it while running to debug

4. **Test incrementally**
   - Start with simple keywords
   - Add complexity gradually

5. **Share profiles**
   - Profiles are JSON files
   - Easy to share with team

---

## 📚 Documentation Map

| Document | Purpose |
|----------|---------|
| **QUICK_START.md** (this file) | Get started in 5 minutes |
| **GUI_GUIDE.md** | Detailed GUI walkthrough |
| **QUICK_REFERENCE.md** | Rule syntax reference |
| **README.md** | Complete documentation |
| **CHANGES_SUMMARY.md** | Technical implementation details |
| **MIGRATION_GUIDE.md** | Upgrade from old version |

---

## ⏱️ Time to First Result

- **GUI + Local files**: ~3 minutes
- **GUI + Graph**: ~5 minutes (auth time)
- **Command-line**: ~5 minutes

---

Ready to automate your email processing? Launch the GUI and get started! 🚀

```bash
python run_gui.py
```
