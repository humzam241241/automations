# 🌐 Email Automation Pro - Web Application Guide

## 🚀 Quick Start

### Launch the Web App

**Option 1: Using the batch file (Windows)**
```bash
start_web.bat
```

**Option 2: Using Python directly**
```bash
python app.py
```

**Option 3: Manual**
```bash
cd "path\to\EMAILtoEXCELLprogram"
python app.py
```

The app will start on **http://localhost:3000** and your browser will open automatically!

---

## ✨ New Web Features

### 🔐 User Management
- **Login/Create User**: Enter username to login or create new account
- **Load Existing User**: Click "Load User" to see all users and select one
- **User Sessions**: Stay logged in during your session
- **Multi-User Support**: Each user can have their own profiles

### 🎨 Modern Web Interface
- **True Web Application**: Runs in your browser (Chrome, Edge, Firefox)
- **Real-time Updates**: Activity log updates in real-time
- **Responsive Design**: Works on desktop, tablet, and mobile
- **Modern UI**: Card-based design with gradients and animations

### 📊 Dashboard Features
- **Connection Status**: Live status indicators for Graph, Mail, OneDrive
- **Profile List**: See all your profiles in the sidebar
- **Quick Actions**: Run, delete profiles with one click
- **Activity Log**: Color-coded logs (green=success, red=error, yellow=warning)

---

## 🎯 How to Use

### 1. First Time Setup

1. **Start the app**: Run `start_web.bat` or `python app.py`
2. **Browser opens**: Navigate to http://localhost:3000
3. **Enter username**: Type your name and click "Login"
4. **You're in!**: The dashboard loads

### 2. Create Your First Profile

1. Click the **"+"** button in the Profiles panel
2. Fill in the form:
   - **Profile Name**: e.g., "Daily Reports"
   - **Input Source**: Select email, .eml, Excel, or CSV
   - **Columns**: Type columns OR check "Auto-detect"
   - **Output**: Smart default selected automatically!
3. Click **"Create Profile"**
4. Done! Profile appears in the list

### 3. Run a Profile

1. **Select profile** from the list (turns purple)
2. View details in the Dashboard panel
3. Click **"▶ Run Profile"**
4. Watch the Activity Log for progress
5. Success! Output files created

### 4. Smart Defaults in Action

**Email Input:**
- Select "Microsoft Graph" or "Local .eml Files"
- Output **automatically** set to "Excel"
- Perfect for processing emails!

**Excel/CSV Input:**
- Select "Excel File" or "CSV File"
- Output **automatically** set to "BI Dashboard"
- Perfect for analysis!

**Want Both?**
- Select "Excel + BI Dashboard"
- Get Excel file AND interactive dashboard!

---

## 🔑 User Management

### Login Flow

```
1. Open http://localhost:3000
2. See login screen
3. Enter username
   └─ New user? → Account created automatically
   └─ Existing user? → Logged in immediately
4. Dashboard loads
```

### Load Existing User

```
1. Click "Load User" button
2. See list of all users with last login time
3. Click any user to login as them
4. Access their profiles and settings
```

### User Sessions

- Stay logged in while browser is open
- Logout anytime with "Logout" button
- Each user has separate profile space
- Safe: No passwords needed (localhost only)

---

## 📊 Dashboard Sections

### Status Bar (Top)
- **Microsoft Graph**: Green = connected, Red = offline
- **Mail Access**: Shows mail permission status
- **OneDrive**: Shows drive permission status

### Profiles Panel (Left)
- List of all your profiles
- Click to select and view details
- Purple highlight = active profile
- "+" button creates new profile

### Dashboard Panel (Right)
- **Profile Details**: See input, columns, output
- **Action Buttons**: Run or delete profile
- **Activity Log**: Real-time execution feedback

---

## 🎨 Modern UI Features

### Animations
- Smooth page transitions
- Button hover effects
- Card lift on hover
- Fade-in modals

### Colors
- **Purple gradient**: Primary branding
- **Green**: Success messages
- **Red**: Errors
- **Yellow**: Warnings
- **Gray**: Neutral info

### Responsive
- Desktop: 2-column layout
- Tablet: Single column
- Mobile: Stacked layout
- Always readable!

---

## 🔧 Advanced Features

### Auto-Detect Columns

When creating a profile with Excel/CSV input:

1. Select "Excel File" or "CSV File"
2. Enter file path
3. Check "Auto-detect columns from file"
4. Click "Detect Columns" button
5. Columns populated automatically!

### Profile Management

**View Profile:**
- Click profile in list
- See all details in dashboard
- Input source, columns, output shown

**Run Profile:**
- Select profile
- Click "▶ Run Profile"
- Watch activity log
- Files created automatically

**Delete Profile:**
- Select profile
- Click "🗑 Delete"
- Confirm deletion
- Profile removed

### Multiple Users

Each user gets:
- Separate session
- Own activity log
- Independent profiles
- Personal settings

---

## 🌐 Technical Details

### Technology Stack
- **Backend**: Flask (Python)
- **Frontend**: HTML5, CSS3, JavaScript
- **Port**: 3000 (localhost)
- **Session**: Flask sessions (secure cookies)

### Files Structure

```
EMAILtoEXCELLprogram/
├── app.py                  # Flask application
├── templates/
│   └── index.html         # Web interface
├── static/
│   ├── css/
│   │   └── style.css     # Modern styles
│   └── js/
│       └── app.js         # Frontend logic
├── users.json             # User database
└── start_web.bat          # Easy launcher
```

### API Endpoints

```
POST   /api/login              - User login
POST   /api/logout             - User logout
GET    /api/current-user       - Get current user
GET    /api/users              - List all users
GET    /api/check-graph        - Check Graph status
GET    /api/profiles           - List profiles
POST   /api/profiles           - Create profile
GET    /api/profiles/<name>    - Get profile details
DELETE /api/profiles/<name>    - Delete profile
POST   /api/profiles/<name>/run - Run profile
POST   /api/detect-columns     - Auto-detect columns
```

---

## 🆚 Web App vs Desktop GUI

| Feature | Web App | Desktop GUI |
|---------|---------|-------------|
| **Interface** | Browser | tkinter window |
| **Access** | localhost:3000 | Desktop app |
| **Users** | Multi-user | Single user |
| **Look** | Modern web design | Native OS |
| **Platform** | Any browser | Windows/Mac/Linux |
| **Sessions** | Cookie-based | None |
| **Updates** | Refresh page | Restart app |

---

## 🔐 Security Notes

### Localhost Only

The web app runs on **localhost (127.0.0.1)** which means:
- ✅ Only accessible from your computer
- ✅ Not exposed to internet
- ✅ Safe for internal use
- ❌ Cannot access from other devices

### No Password

For simplicity:
- No passwords required
- Username only
- Localhost = trusted environment
- Production: Add auth if exposing to network

### Data Storage

- **User data**: `users.json` (local file)
- **Profiles**: `profiles/` directory
- **Sessions**: Secure Flask cookies
- **Output**: `output/` directory

---

## 🐛 Troubleshooting

### "Cannot connect to localhost:3000"

**Solution:**
1. Check if app is running
2. Look for "Running on http://localhost:3000" message
3. Open browser manually to http://localhost:3000
4. Check firewall settings

### "Port 3000 already in use"

**Solution:**
1. Close other apps using port 3000
2. Or change port in `app.py`:
   ```python
   app.run(host='0.0.0.0', port=3001, debug=True)
   ```

### "Profile won't create"

**Check:**
- Profile name filled
- Columns specified (or auto-detect checked)
- Input source selected
- Valid file paths (for Excel/CSV)

### "Graph not connected"

**This is OK!** You can still:
- Use local .eml files
- Use Excel/CSV input
- Create profiles
- Run non-Graph profiles

---

## 📱 Mobile Access

### Desktop Browser
Best experience! Full features, fast, responsive.

### Tablet Browser
Works great! Touch-friendly, responsive layout.

### Mobile Phone
Basic support. Some features cramped on small screens.

**Recommended:** Use on desktop or tablet for best experience.

---

## 🚀 Quick Tips

### Tip 1: Keyboard Shortcuts
- **Enter**: Login from login screen
- **Ctrl+C**: Stop server (in terminal)
- **F5**: Refresh page

### Tip 2: Multiple Tabs
- Open multiple browser tabs
- Each tab = separate session
- Run different profiles simultaneously

### Tip 3: Bookmarks
- Bookmark http://localhost:3000
- Quick access anytime
- Add to browser favorites

### Tip 4: Dark Theme
- Activity log = dark theme
- Easy on eyes
- Monospace font for logs

---

## 🔄 Updates & Maintenance

### Restart Server

To apply code changes:
1. Press **Ctrl+C** in terminal
2. Run `python app.py` again
3. Refresh browser

### Clear Sessions

To logout all users:
1. Stop server
2. Delete cookie in browser
3. Restart server

### Backup Users

Copy `users.json` to safe location to backup user database.

---

## 📞 Support

### Common Questions

**Q: Can I access from another computer?**
A: No, localhost only. To allow network access, change `host='0.0.0.0'` and configure firewall.

**Q: Do I need to be online?**
A: No! Works completely offline (except Microsoft Graph features).

**Q: Can multiple people use it?**
A: Yes! Each person creates their own username.

**Q: Will it remember my profiles?**
A: Yes! Profiles saved in `profiles/` directory.

---

## 🎉 Enjoy Your Web App!

**Key Features:**
- ✅ True web application
- ✅ Runs on localhost:3000
- ✅ Multi-user support
- ✅ Modern design
- ✅ Real-time updates
- ✅ Smart defaults
- ✅ Activity logging

**Start using:**
```bash
start_web.bat
```

**Or directly:**
```bash
python app.py
```

**Then open:** http://localhost:3000

---

**Made with ❤️ by the Sanofi Automation Team**

**Web App Version - January 2026**
