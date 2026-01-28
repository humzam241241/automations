# GUI User Guide

## 🚀 Getting Started

### Launch the GUI

**Windows:**
1. Double-click `run_app.bat`
2. Select option 1 (GUI Interface)

**Or directly:**
```bash
python run_gui.py
```

---

## 📊 Main Interface

### Status Panel (Left Side)
Shows the current state of Microsoft Graph connectivity:

- **✅ Graph Connected** - Full access to Outlook and OneDrive
- **⚠ Graph Unavailable** - Local mode only (use .eml files)
- **📬 Mail: ✅ OK** - Can fetch emails from Outlook
- **📬 Mail: ❌ Denied** - Missing Mail.Read permission
- **💾 OneDrive: ✅ OK** - Can upload to OneDrive
- **💾 OneDrive: ❌ Denied** - Missing Files.ReadWrite permission

### Profile List
- Shows all available profiles
- Click to select a profile
- Green **▶ Run Profile** button executes selected profile
- Blue **➕ Create New** button opens profile creation wizard
- Grey **🔄 Refresh** button reloads profile list

### Activity Log (Right Side)
Real-time log of all actions:
- Blue: Information messages
- Green: Success messages
- Orange: Warnings
- Red: Errors

---

## ✏️ Creating a Profile

Click **➕ Create New** to open the profile creation wizard with 5 tabs:

### Tab 1: Basic Info
- **Profile Name**: Unique identifier (required)
- **Description**: Optional description of what this profile does

### Tab 2: Input Source
Choose where emails come from:

**📧 Microsoft Graph** (requires permissions)
- Folder name (e.g., Inbox, Sent Items)
- Newest N messages (default: 25)

**📁 Local .eml Files** (no Graph needed)
- Directory path (default: `./input_emails`)
- Browse button to select folder

**📊 CSV File** (no Graph needed)
- CSV file path
- Browse button to select file

### Tab 3: Schema
Define the columns for your spreadsheet:

**✍ Manual Entry:**
- Enter column names separated by commas
- Example: `Subject,From,Date,Amount,Status`

**📋 Load from Template:**
- Click "Browse..." to select an .xlsx template
- Columns are read from the first row
- Automatically populates the manual entry field

### Tab 4: Rules
Define keyword/regex matching rules (optional):

**JSON format:**
```json
[
  {
    "column": "Billing",
    "keywords": ["invoice", "billing"],
    "value": "Yes",
    "word_boundary": true,
    "search_in": ["subject", "body"],
    "priority": 10
  }
]
```

**Pre-filled with example** - edit or replace as needed.

### Tab 5: Output
Choose output format and destination:

**Format:**
- Excel (.xlsx) - Recommended for most use cases
- CSV (.csv) - For compatibility with other tools

**Destination:**
- **OneDrive** (requires Files.ReadWrite permission)
  - Specify OneDrive path (e.g., `/EmailReports`)
- **Local file** (no permissions needed)
  - Specify local directory (e.g., `./output`)
  - Browse button to select folder

---

## ▶️ Running a Profile

1. **Select a profile** from the list (single click)
2. Click **▶ Run Profile** button
3. Watch the Activity Log for progress
4. Success dialog shows results:
   - Number of emails processed
   - Output file location

### Permission Checks
The GUI automatically checks if you have the required permissions:
- Profile needs mail but permission denied → Shows error with instructions
- Profile needs OneDrive but permission denied → Shows error with instructions

---

## 🎨 Interface Colors

### Status Indicators
- **Green (✅)**: Working correctly
- **Red (❌)**: Permission denied or error
- **Orange (⚠)**: Warning or unavailable

### Buttons
- **Green**: Primary action (Run, Save)
- **Blue**: Secondary action (Create New)
- **Grey**: Utility action (Refresh)
- **Red**: Destructive action (Cancel)

### Log Messages
- **White**: Normal information
- **Green**: Success
- **Orange**: Warning
- **Red**: Error

---

## 🔍 Common Scenarios

### Scenario 1: I have full Graph permissions
1. Launch GUI → Status shows all green
2. Create profile with Graph input + OneDrive output
3. Run profile → Emails fetched and uploaded

### Scenario 2: I don't have Graph permissions
1. Launch GUI → Status shows orange warning
2. Create profile with local .eml input + local output
3. Place .eml files in `./input_emails/`
4. Run profile → Files processed locally

### Scenario 3: I have mail but not OneDrive
1. Launch GUI → Mail: ✅ OK, OneDrive: ❌ Denied
2. Create profile with Graph input + **local output**
3. Run profile → Emails fetched, saved locally

### Scenario 4: I want to test rules without Graph
1. Export some test emails as .eml files
2. Create profile with local .eml input
3. Define your rules in Tab 4
4. Run profile and check output
5. Refine rules and re-run

---

## ⌨️ Keyboard Shortcuts

- **Double-click profile** → Selects profile
- **Enter on selected profile** → (not implemented, use button)
- **Escape** → (not implemented)

---

## 🛠️ Troubleshooting

### "GUI failed to launch"
**Possible causes:**
- Python not installed
- tkinter not available (rare on Windows, install `python-tk` on Linux)

**Solution:**
- Use command-line wizard instead (option 2 in `run_app.bat`)
- Or: `python run_wizard.py`

### "Authentication failed"
**Cause:** Missing or incorrect CLIENT_ID/AUTHORITY environment variables

**Solution:**
```cmd
set CLIENT_ID=your-client-id
set AUTHORITY=https://login.microsoftonline.com/your-tenant-id
```
Then restart the GUI.

### "Profile requires Mail.Read permission"
**Cause:** Your Azure AD app doesn't have Mail.Read permission

**Solution:**
1. Contact IT admin to grant permission, OR
2. Create a profile with local .eml input instead

### "Invalid JSON in Rules"
**Cause:** Syntax error in rules JSON

**Solution:**
- Check commas, brackets, quotes
- Use a JSON validator (jsonlint.com)
- Start with the pre-filled example and modify

### Rules not matching
**Cause:** Various (word boundaries, search_in, priority)

**Solution:**
1. Check `word_boundary` setting
2. Verify `search_in` includes the right fields
3. Test with command-line wizard's explain mode for detailed output

---

## 📚 Tips & Tricks

### Tip 1: Start with Examples
The GUI pre-fills examples in the Rules tab. Modify them rather than starting from scratch.

### Tip 2: Test Locally First
Before using Graph, test your profile with local .eml files to verify rules work correctly.

### Tip 3: Use Templates for Consistency
If multiple people create profiles, share a template .xlsx so everyone uses the same columns.

### Tip 4: Check the Activity Log
The Activity Log shows exactly what's happening. Watch it while running profiles to debug issues.

### Tip 5: Profile Names
Use descriptive names like `invoice_tracker` or `dept_categorizer` instead of `profile1`.

### Tip 6: Refresh After Manual Edits
If you manually edit a profile JSON file, click **🔄 Refresh** to reload the list.

---

## 🎓 Workflow Example

### Creating an Invoice Tracking Profile

1. **Launch GUI**
   ```bash
   python run_gui.py
   ```

2. **Click ➕ Create New**

3. **Tab 1: Basic Info**
   - Name: `invoice_tracker`
   - Description: `Tracks invoices from billing@company.com`

4. **Tab 2: Input Source**
   - Select: 📧 Microsoft Graph
   - Folder: `Inbox`
   - Newest: `50`

5. **Tab 3: Schema**
   - Manual entry: `Subject,From,Date,Invoice_Number,Amount,Status`

6. **Tab 4: Rules**
   ```json
   [
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
   ]
   ```

7. **Tab 5: Output**
   - Format: Excel
   - Destination: Local file
   - Directory: `./output`

8. **Click 💾 Save & Close**

9. **Select** `invoice_tracker` from list

10. **Click ▶ Run Profile**

11. **Check Activity Log** for results

---

## 🆚 GUI vs Command-line

| Feature | GUI | Command-line |
|---------|-----|--------------|
| **Ease of Use** | ⭐⭐⭐⭐⭐ | ⭐⭐⭐ |
| **Visual Feedback** | ✅ | ❌ |
| **Status Display** | ✅ Real-time | ❌ Text only |
| **Profile Creation** | ✅ Tab-based | ⭐ Step-by-step |
| **Rules Editing** | ⭐ JSON editor | ⭐ Prompted entry |
| **Explain Mode** | ❌ Not shown | ✅ Detailed output |
| **Scriptable** | ❌ | ✅ |
| **Remote SSH** | ❌ Requires X11 | ✅ |

**Use GUI when:**
- You prefer visual interfaces
- You're creating/managing multiple profiles
- You want real-time status updates
- You're on Windows/Mac with desktop access

**Use command-line when:**
- You're connecting via SSH
- You're automating with scripts
- You need detailed explain mode output
- You prefer terminal interfaces

---

## 🔄 Updates & Maintenance

### Refreshing Profiles
Click **🔄 Refresh** after:
- Creating a profile via command-line
- Manually editing profile JSON files
- Copying profiles from another machine

### Updating the Application
1. Pull latest code from GitHub
2. Restart the GUI
3. Profiles are preserved (they're just JSON files)

---

## 💡 Advanced Features

### Running Profiles Programmatically
While the GUI can't be scripted, you can use the command-line for automation:

```bash
python -c "from core.engine import ExecutionEngine; from core.profile_loader import ProfileLoader; engine = ExecutionEngine(access_token='YOUR_TOKEN'); profile = ProfileLoader().load_profile('invoice_tracker'); result = engine.run_profile(profile); print(result)"
```

Or create a simple script:
```python
from core.engine import ExecutionEngine
from core.profile_loader import ProfileLoader

loader = ProfileLoader()
profile = loader.load_profile("invoice_tracker")

engine = ExecutionEngine(access_token=None)  # Or provide token
result = engine.run_profile(profile)

print(f"Processed {result['emails_processed']} emails")
print(f"Output: {result['output']['path']}")
```

---

## 📞 Support

If you encounter issues:
1. Check the **Activity Log** for error messages
2. Refer to this guide for common scenarios
3. Check `CHANGES_SUMMARY.md` for technical details
4. Use command-line wizard with explain mode for debugging rules

---

## 🎉 Quick Reference

| Action | Steps |
|--------|-------|
| **Create Profile** | Click ➕ Create New → Fill 5 tabs → Save |
| **Run Profile** | Select profile → Click ▶ Run |
| **Check Status** | Look at left panel (green = OK) |
| **View Results** | Watch Activity Log (right panel) |
| **Reload Profiles** | Click 🔄 Refresh |
| **Switch to CLI** | Close GUI, run `python run_wizard.py` |

---

Enjoy the GUI! 🚀
