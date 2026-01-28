# 🚀 Unified Launcher Guide

## ✨ Single Entry Point

We've unified everything into **one launcher**: `start.bat`

### How to Start

**Just double-click:** `start.bat`

Or from command line:
```bash
start.bat
```

---

## 📋 Launch Menu

When you run `start.bat`, you'll see:

```
============================================================

         EMAIL AUTOMATION PRO - UNIFIED LAUNCHER

============================================================

 Choose how you want to run the application:

 [1] Web Application (Recommended)
     - Modern browser interface
     - Runs on localhost:3000
     - Multi-user support
     - File browser with visual navigation

 [2] Desktop GUI
     - Native window interface
     - Traditional desktop app

 [3] CLI Wizard
     - Command-line interface
     - Step-by-step prompts

 [4] Exit

============================================================

Enter your choice (1-4):
```

---

## 🎯 Choose Your Mode

### Option 1: Web Application (Recommended) 🌐

**Best for:**
- Modern web experience
- Multi-user environments
- Visual file browsing
- Remote access (localhost)

**Features:**
- Runs in your browser
- Full file system browser
- Multi-user login
- Real-time activity log
- Modern UI with animations

**Press:** `1`

Then your browser opens to: http://localhost:3000

---

### Option 2: Desktop GUI 🖥️

**Best for:**
- Traditional desktop app users
- Single-user scenarios
- Quick local tasks

**Features:**
- Native window interface
- Tkinter-based
- Fast and lightweight

**Press:** `2`

Then the desktop window opens

---

### Option 3: CLI Wizard 💻

**Best for:**
- Command-line lovers
- Scripting and automation
- Step-by-step guidance

**Features:**
- Terminal-based interface
- Interactive prompts
- Profile creation wizard

**Press:** `3`

Then follow the prompts

---

### Option 4: Exit 🚪

**Press:** `4`

Exits the launcher

---

## 🗂️ Full File System Browser

### What Changed?

**Before:** Limited to one directory
**Now:** Browse your ENTIRE computer!

### Features

#### 1. **See All Your Drives**
When you open the file browser, you'll see:
```
💾 C:\ Drive
💾 D:\ Drive
💾 E:\ Drive
```

#### 2. **Navigate Anywhere**
- Click folders to open them
- Click "⬆️ Parent Folder" to go up
- No restrictions!

#### 3. **Smart Starting Location**
Opens at:
1. Your Desktop (most common)
2. OneDrive Desktop (if available)
3. Home folder (fallback)

#### 4. **Visual File Icons**
- 💾 Drives
- 📁 Folders
- 📊 Excel files
- 📄 CSV files
- 📧 Email files (.eml)
- ✉️ Outlook files (.msg)

---

## 📂 How to Use File Browser

### For Files (Excel, CSV)

1. Click **"📁 Browse"** button
2. Navigate to any folder
3. Click folders to open them
4. Click a file to select it
5. File path auto-fills!

**Example:**
```
File Path: [_______________________] [📁 Browse]
                                          ↑ Click

Browser Opens:
┌─────────────────────────────────────────────┐
│ 📁 Browse Files                          × │
├─────────────────────────────────────────────┤
│ Current Location:                           │
│ C:\Users\You\Desktop\Reports               │
├─────────────────────────────────────────────┤
│ ⬆️ Parent Folder                           │
│ 📁 2024_Reports              Jan 28 10:30  │
│ 📁 Archives                  Jan 27 15:20  │
│ 📊 sales_data.xlsx  2.5 MB   Jan 26 09:15  │ ← Click this!
│ 📄 emails.csv       156 KB   Jan 25 14:45  │
└─────────────────────────────────────────────┘

Result:
File Path: [C:\Users\You\Desktop\Reports\sales_data.xlsx]
```

---

### For Folders (.eml directories)

1. Click **"📁 Browse"** button
2. Navigate to desired folder
3. **Option A:** Click **"✓ Select This Folder"** button at top
4. **Option B:** Double-click the folder name
5. Folder path auto-fills!

**Example:**
```
Directory Path: [_______________________] [📁 Browse]
                                              ↑ Click

Browser Opens:
┌─────────────────────────────────────────────┐
│ 📁 Browse Files                          × │
├─────────────────────────────────────────────┤
│ Current Location:                           │
│ C:\Users\You\Documents                     │
├─────────────────────────────────────────────┤
│ ✓ Select This Folder                       │ ← Click to select current
├─────────────────────────────────────────────┤
│ ⬆️ Parent Folder                           │
│ 📁 EmailExports              Jan 28 10:30  │ ← Double-click to select
│ 📁 Projects                  Jan 27 15:20  │
└─────────────────────────────────────────────┘

Result:
Directory Path: [C:\Users\You\Documents\EmailExports]
```

---

## 🎯 Quick Examples

### Example 1: Start Web App

1. Double-click `start.bat`
2. Press `1`
3. Browser opens to localhost:3000
4. Login and use!

### Example 2: Browse for Excel File

1. In web app, create new profile
2. Select "Excel File" input
3. Click "📁 Browse"
4. Navigate: Desktop → Reports → 2024
5. Click `sales_Q1.xlsx`
6. Done! Path auto-filled

### Example 3: Browse for Email Folder

1. Create new profile
2. Select "Local .eml Files"
3. Click "📁 Browse"
4. Navigate to your email export folder
5. Click "✓ Select This Folder" at top
6. Done! Folder path auto-filled

---

## 🆚 Old vs New

### Old Way (3 separate batch files)
```
run_app.bat       → Choose in terminal
run_gui.bat       → Desktop GUI
start_web.bat     → Web app
```

### New Way (1 unified launcher)
```
start.bat         → Choose with menu
                    ↓
    [1] Web App
    [2] Desktop GUI
    [3] CLI Wizard
```

---

## 💡 Pro Tips

### Tip 1: Bookmark the Menu Choice
If you always use web app:
- Create a shortcut to `start.bat`
- Right-click → Properties
- Target: `start.bat 1` (auto-selects option 1)

### Tip 2: Pin to Taskbar
- Right-click `start.bat`
- Send to → Desktop (create shortcut)
- Drag shortcut to taskbar
- Quick access!

### Tip 3: File Browser Shortcuts
- **Single-click** folder = Navigate into it
- **Double-click** folder = Select it (when browsing for directory)
- **Click "⬆️ Parent"** = Go up one level
- **Click drive** = Jump to C:\, D:\, etc.

### Tip 4: Recent Files
The file browser remembers where you were!
- If you typed a path before, it starts there
- Makes repeat selections super fast

---

## 🔧 Technical Details

### Unified Launcher
- **File:** `start.bat`
- **Replaces:** run_app.bat, run_gui.bat, start_web.bat
- **Size:** ~80 lines
- **Color:** Cyan text on black background

### File Browser API
- **Endpoint:** POST `/api/browse-directory`
- **Access:** Full file system (no restrictions)
- **Security:** Permission error handling
- **Features:** Drive enumeration, folder navigation

### Deleted Files
- ❌ run_app.bat (replaced)
- ❌ run_gui.bat (replaced)
- ❌ start_web.bat (replaced)

---

## ❓ FAQ

**Q: Can I still use the old batch files?**
A: They're deleted, but `start.bat` does everything they did!

**Q: Can I browse the entire C:\ drive?**
A: Yes! You have full access to your entire computer.

**Q: What if I get "Permission denied"?**
A: Some system folders are protected. Use "Go to Home Folder" button to reset.

**Q: Can I type the path manually?**
A: Yes! The browse button is optional. Type if you prefer!

**Q: Does it work on Mac/Linux?**
A: The Python scripts work everywhere. The .bat file is Windows-only.
  For Mac/Linux: `python app.py` or `python run_gui.py` or `python run_wizard.py`

**Q: Can I create a desktop shortcut?**
A: Yes! Right-click `start.bat` → Send to → Desktop

---

## 🎉 Summary

**Unified Launcher:**
- ✅ One file to rule them all: `start.bat`
- ✅ Choose your mode with a menu
- ✅ Clean and simple
- ✅ No more confusion about which file to run

**File Browser:**
- ✅ Browse entire file system
- ✅ All drives accessible
- ✅ Visual navigation
- ✅ Smart file filtering
- ✅ Folder selection support
- ✅ No more typing long paths!

---

**Ready to Go!**

```bash
# Just run:
start.bat

# Then choose your mode!
```

---

**Made with ❤️ by the Sanofi Automation Team**
**Unified Launcher - January 2026**
