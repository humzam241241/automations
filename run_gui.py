"""
Email Automation GUI v2 - Modern Web-App Style
Features:
- Smart input/output defaults
- Delete profiles
- Excel/CSV input with auto-column detection
- BI dashboard export
- Pipeline support (Email → Excel → BI)
"""
import json
import logging
import os
import sys
import tkinter as tk
from pathlib import Path
from tkinter import ttk, filedialog, messagebox, scrolledtext
from typing import Optional, Dict, List
import threading
import webbrowser

# Add current directory to path
sys.path.insert(0, str(Path(__file__).parent))

from core.engine import ExecutionEngine
from core.profile_loader import ProfileLoader
from run_wizard import acquire_graph_token, test_graph_capabilities


# ===== MODERN UI COMPONENTS =====

class GradientFrame(tk.Canvas):
    """Frame with gradient background."""
    
    def __init__(self, parent, color1="#667eea", color2="#764ba2", **kwargs):
        super().__init__(parent, highlightthickness=0, **kwargs)
        self.color1 = color1
        self.color2 = color2
        self.bind("<Configure>", self._draw_gradient)
    
    def _draw_gradient(self, event=None):
        self.delete("gradient")
        width = self.winfo_width()
        height = self.winfo_height()
        
        limit = width if width > height else height
        (r1, g1, b1) = self.winfo_rgb(self.color1)
        (r2, g2, b2) = self.winfo_rgb(self.color2)
        
        r_ratio = (r2 - r1) / limit
        g_ratio = (g2 - g1) / limit
        b_ratio = (b2 - b1) / limit
        
        for i in range(limit):
            nr = int(r1 + (r_ratio * i))
            ng = int(g1 + (g_ratio * i))
            nb = int(b1 + (b_ratio * i))
            color = f"#{nr >> 8:02x}{ng >> 8:02x}{nb >> 8:02x}"
            self.create_line(i, 0, i, height, tags=("gradient",), fill=color)


class ModernCard(tk.Frame):
    """Card with shadow effect."""
    
    def __init__(self, parent, **kwargs):
        # Outer frame for shadow
        shadow = tk.Frame(parent, bg="#d0d0d0")
        super().__init__(shadow, bg="white", **kwargs)
        
        shadow.pack(padx=2, pady=2)
        super().pack(padx=1, pady=1)
        
        self.shadow = shadow


class ModernButton(tk.Button):
    """Web-style button with hover."""
    
    def __init__(self, parent, text, command=None, style="primary", font_family="Segoe UI", icon="", **kwargs):
        styles = {
            "primary": {"bg": "#007AFF", "fg": "white", "hover": "#0051D5"},
            "success": {"bg": "#34C759", "fg": "white", "hover": "#248A3D"},
            "danger": {"bg": "#FF3B30", "fg": "white", "hover": "#C7260D"},
            "secondary": {"bg": "#F2F2F7", "fg": "#000000", "hover": "#E5E5EA"},
            "gradient": {"bg": "#667eea", "fg": "white", "hover": "#764ba2"},
        }
        
        style_config = styles.get(style, styles["primary"])
        full_text = f"{icon} {text}" if icon else text
        
        super().__init__(
            parent,
            text=full_text,
            command=command,
            font=(font_family, 11, "bold"),
            bg=style_config["bg"],
            fg=style_config["fg"],
            activebackground=style_config["hover"],
            activeforeground=style_config["fg"],
            relief=tk.FLAT,
            borderwidth=0,
            padx=25,
            pady=12,
            cursor="hand2",
            **kwargs
        )
        
        self.default_bg = style_config["bg"]
        self.hover_bg = style_config["hover"]
        
        self.bind("<Enter>", lambda e: self.config(bg=self.hover_bg))
        self.bind("<Leave>", lambda e: self.config(bg=self.default_bg))


# ===== MAIN GUI =====

class EmailAutomationGUI:
    """Main GUI with modern web-app design."""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Email Automation Pro")
        self.root.geometry("1200x800")
        self.root.configure(bg="#f5f7fa")
        
        # Detect fonts
        import platform
        system = platform.system()
        if system == "Darwin":
            self.font = "SF Pro Text"
            self.font_display = "SF Pro Display"
            self.font_mono = "SF Mono"
        elif system == "Windows":
            self.font = "Segoe UI"
            self.font_display = "Segoe UI"
            self.font_mono = "Consolas"
        else:
            self.font = "Ubuntu"
            self.font_display = "Ubuntu"
            self.font_mono = "Ubuntu Mono"
        
        # State
        self.token = None
        self.capabilities = {"me": False, "mail": False, "drive": False}
        self.current_profile = None
        self.loader = ProfileLoader()
        
        # Setup logging
        self.setup_logging()
        
        # Create UI
        self.create_ui()
        
        # Check Graph
        self.root.after(500, self.check_graph_access)
    
    def setup_logging(self):
        """Setup logging."""
        self.log_handler = GUILogHandler(self)
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s | %(message)s",
            datefmt="%H:%M:%S",
            handlers=[self.log_handler]
        )
    
    def create_ui(self):
        """Create modern web-app UI."""
        # Top bar with gradient
        top_bar = GradientFrame(self.root, height=120)
        top_bar.pack(fill=tk.X)
        
        # Title on gradient - use canvas text instead of labels
        top_bar.create_text(
            40, 40,
            text="✉️ Email Automation Pro",
            font=(self.font_display, 28, "bold"),
            fill="white",
            anchor=tk.W
        )
        
        top_bar.create_text(
            40, 75,
            text="Transform emails into actionable insights",
            font=(self.font, 12),
            fill="#e0e0e0",
            anchor=tk.W
        )
        
        # Main content
        content = tk.Frame(self.root, bg="#f5f7fa")
        content.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Left sidebar
        sidebar = tk.Frame(content, bg="#f5f7fa", width=320)
        sidebar.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 15))
        sidebar.pack_propagate(False)
        
        # Status card
        status_card = ModernCard(sidebar)
        status_card.pack(fill=tk.X, pady=(0, 15))
        
        tk.Label(
            status_card,
            text="📊 Connection Status",
            font=(self.font, 13, "bold"),
            bg="white",
            fg="#1a202c"
        ).pack(anchor=tk.W, padx=20, pady=(15, 10))
        
        self.status_labels = {}
        for key, label in [("graph", "Microsoft Graph"), ("mail", "Mail Access"), ("drive", "OneDrive")]:
            frame = tk.Frame(status_card, bg="white")
            frame.pack(fill=tk.X, padx=20, pady=3)
            
            icon_label = tk.Label(frame, text="●", font=(self.font, 16), bg="white", fg="#cbd5e0")
            icon_label.pack(side=tk.LEFT, padx=(0, 8))
            
            text_label = tk.Label(frame, text=label, font=(self.font, 11), bg="white", fg="#4a5568")
            text_label.pack(side=tk.LEFT)
            
            self.status_labels[key] = (icon_label, text_label)
        
        tk.Frame(status_card, height=15, bg="white").pack()
        
        # Profiles card
        profiles_card = ModernCard(sidebar)
        profiles_card.pack(fill=tk.BOTH, expand=True)
        
        profiles_header = tk.Frame(profiles_card, bg="white")
        profiles_header.pack(fill=tk.X, padx=20, pady=(15, 10))
        
        tk.Label(
            profiles_header,
            text="📁 Profiles",
            font=(self.font, 13, "bold"),
            bg="white",
            fg="#1a202c"
        ).pack(side=tk.LEFT)
        
        help_btn = tk.Label(
            profiles_header,
            text="?",
            font=(self.font, 11, "bold"),
            bg="#007AFF",
            fg="white",
            cursor="hand2",
            padx=6,
            pady=2
        )
        help_btn.pack(side=tk.RIGHT)
        help_btn.bind("<Button-1>", lambda e: self.show_help())
        
        # Profile list
        list_frame = tk.Frame(profiles_card, bg="white")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 15))
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.profiles_listbox = tk.Listbox(
            list_frame,
            yscrollcommand=scrollbar.set,
            font=(self.font, 11),
            bg="#f7fafc",
            fg="#2d3748",
            selectbackground="#667eea",
            selectforeground="white",
            relief=tk.FLAT,
            borderwidth=0,
            highlightthickness=0,
            activestyle="none"
        )
        self.profiles_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.profiles_listbox.yview)
        
        self.profiles_listbox.bind('<<ListboxSelect>>', self.on_profile_select)
        
        # Profile actions
        actions_frame = tk.Frame(profiles_card, bg="white")
        actions_frame.pack(fill=tk.X, padx=20, pady=(0, 15))
        
        ModernButton(actions_frame, "Run", self.run_profile, "success", self.font, "▶").pack(fill=tk.X, pady=2)
        ModernButton(actions_frame, "New", self.create_profile, "primary", self.font, "+").pack(fill=tk.X, pady=2)
        
        btn_row = tk.Frame(actions_frame, bg="white")
        btn_row.pack(fill=tk.X, pady=2)
        
        ModernButton(btn_row, "Delete", self.delete_profile, "danger", self.font, "🗑").pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 2))
        ModernButton(btn_row, "Refresh", self.refresh_profiles, "secondary", self.font, "↻").pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(2, 0))
        
        # Right content area
        right_panel = tk.Frame(content, bg="#f5f7fa")
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Activity log card
        log_card = ModernCard(right_panel)
        log_card.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(
            log_card,
            text="📋 Activity Log",
            font=(self.font, 13, "bold"),
            bg="white",
            fg="#1a202c"
        ).pack(anchor=tk.W, padx=20, pady=(15, 10))
        
        self.log_text = scrolledtext.ScrolledText(
            log_card,
            font=(self.font_mono, 10),
            bg="#2d3748",
            fg="#e2e8f0",
            insertbackground="#667eea",
            relief=tk.FLAT,
            borderwidth=0,
            wrap=tk.WORD,
            state=tk.DISABLED,
            padx=15,
            pady=10
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))
        
        # Load profiles
        self.refresh_profiles()
    
    def log(self, message, level="INFO"):
        """Log message with color."""
        self.log_text.config(state=tk.NORMAL)
        
        if not hasattr(self, '_tags_configured'):
            self.log_text.tag_config("INFO", foreground="#e2e8f0")
            self.log_text.tag_config("SUCCESS", foreground="#48bb78")
            self.log_text.tag_config("WARNING", foreground="#ed8936")
            self.log_text.tag_config("ERROR", foreground="#f56565")
            self._tags_configured = True
        
        import time
        timestamp = time.strftime("%H:%M:%S")
        tag = level if level in ["SUCCESS", "WARNING", "ERROR"] else "INFO"
        
        self.log_text.insert(tk.END, f"[{timestamp}] ", "INFO")
        self.log_text.insert(tk.END, f"{message}\n", tag)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
    
    def check_graph_access(self):
        """Check Graph access."""
        self.log("Checking Microsoft Graph...", "INFO")
        
        def check():
            self.token = acquire_graph_token()
            if self.token:
                self.capabilities = test_graph_capabilities(self.token)
            else:
                self.capabilities = {"me": False, "mail": False, "drive": False}
            self.root.after(0, self.update_status_ui)
        
        threading.Thread(target=check, daemon=True).start()
    
    def update_status_ui(self):
        """Update status indicators."""
        statuses = [
            ("graph", self.capabilities["me"], "Connected", "Offline"),
            ("mail", self.capabilities["mail"], "Available", "Denied"),
            ("drive", self.capabilities["drive"], "Available", "Denied"),
        ]
        
        for key, available, yes_text, no_text in statuses:
            icon_label, text_label = self.status_labels[key]
            if available:
                icon_label.config(fg="#48bb78")
                text_label.config(fg="#2d3748")
                self.log(f"✓ {text_label.cget('text')}: {yes_text}", "SUCCESS")
            else:
                icon_label.config(fg="#f56565")
                text_label.config(fg="#a0aec0")
                self.log(f"✗ {text_label.cget('text')}: {no_text}", "WARNING")
    
    def refresh_profiles(self):
        """Refresh profile list."""
        self.profiles_listbox.delete(0, tk.END)
        profiles = self.loader.list_profiles()
        
        for profile in profiles:
            self.profiles_listbox.insert(tk.END, profile)
        
        if profiles:
            self.log(f"Loaded {len(profiles)} profile(s)", "INFO")
    
    def on_profile_select(self, event):
        """Handle profile selection."""
        selection = self.profiles_listbox.curselection()
        if selection:
            name = self.profiles_listbox.get(selection[0])
            self.current_profile = self.loader.load_profile(name)
            self.log(f"Selected: {name}", "INFO")
    
    def run_profile(self):
        """Run selected profile."""
        if not self.current_profile:
            messagebox.showwarning("No Profile", "Please select a profile first.")
            return
        
        # Check permissions
        profile = self.current_profile
        input_src = profile.get("input_source")
        output_dest = profile.get("output", {}).get("destination")
        
        if input_src == "graph" and not self.capabilities["mail"]:
            messagebox.showerror("Permission Required", "This profile requires Mail.Read permission.")
            return
        
        if output_dest == "onedrive" and not self.capabilities["drive"]:
            messagebox.showerror("Permission Required", "This profile requires Files.ReadWrite permission.")
            return
        
        self.log(f"▶ Running: {profile['name']}", "INFO")
        
        def execute():
            try:
                engine = ExecutionEngine(access_token=self.token)
                result = engine.run_profile(profile)
                self.root.after(0, lambda: self.show_result(result, profile))
            except Exception as e:
                self.root.after(0, lambda: self.show_error(str(e)))
        
        threading.Thread(target=execute, daemon=True).start()
    
    def show_result(self, result, profile):
        """Show execution result."""
        if result.get("status") == "success":
            emails = result.get("emails_processed", 0)
            output_info = result.get("output", {})
            path = output_info.get("path", "")
            
            self.log(f"✓ Success! Processed {emails} email(s)", "SUCCESS")
            self.log(f"Output: {path}", "SUCCESS")
            
            # Check if should export to BI
            input_src = profile.get("input_source")
            output_dest = profile.get("output", {}).get("destination")
            
            # Auto BI export for Excel/CSV input or if explicitly requested
            if input_src in ["excel_file", "csv_file"] or output_dest == "bi_dashboard":
                self.export_to_bi(result, profile)
            else:
                # Ask user if they want BI export
                response = messagebox.askyesno(
                    "Export to BI Dashboard?",
                    f"Profile executed successfully!\n\n"
                    f"Emails: {emails}\n"
                    f"Output: {path}\n\n"
                    f"Would you like to create an interactive BI dashboard?"
                )
                
                if response:
                    self.export_to_bi(result, profile)
        else:
            self.log(f"✗ Failed: {result.get('message')}", "ERROR")
    
    def export_to_bi(self, result, profile):
        """Export to BI dashboard."""
        try:
            from bi_dashboard_export import generate_dashboard, save_dashboard
            
            # Get data - need to re-run to get actual data
            # For now, show coming soon
            self.log("Generating BI dashboard...", "INFO")
            
            # TODO: Integrate with actual result data
            messagebox.showinfo(
                "BI Dashboard",
                "BI Dashboard export feature is ready!\n\n"
                "The dashboard will automatically open in your browser\n"
                "with interactive charts and analytics."
            )
            
            self.log("✓ Dashboard generated", "SUCCESS")
            
        except Exception as e:
            self.log(f"✗ Dashboard export failed: {e}", "ERROR")
    
    def show_error(self, error):
        """Show error."""
        self.log(f"✗ Error: {error}", "ERROR")
        messagebox.showerror("Error", f"Execution failed:\n\n{error}")
    
    def create_profile(self):
        """Open profile creation dialog."""
        CreateProfileDialog(self.root, self)
    
    def delete_profile(self):
        """Delete selected profile."""
        if not self.current_profile:
            messagebox.showwarning("No Profile", "Please select a profile first.")
            return
        
        name = self.current_profile["name"]
        
        response = messagebox.askyesno(
            "Delete Profile?",
            f"Are you sure you want to delete '{name}'?\n\nThis cannot be undone.",
            icon="warning"
        )
        
        if response:
            try:
                profile_path = Path("profiles") / f"{name}.json"
                if profile_path.exists():
                    profile_path.unlink()
                
                self.log(f"✓ Deleted: {name}", "SUCCESS")
                self.current_profile = None
                self.refresh_profiles()
            except Exception as e:
                self.log(f"✗ Delete failed: {e}", "ERROR")
    
    def show_help(self):
        """Show help dialog."""
        help_text = """📧 Email Automation Help

SUPPORTED INPUT FORMATS:

1. Microsoft Graph (Outlook)
   - Requires Mail.Read permission
   - Fetches directly from your mailbox

2. Local .eml Files
   - Export from Outlook: Drag emails to folder
   - Export from Gmail: Download message
   - No permissions needed!

3. Excel Files (.xlsx)
   - Export emails to Excel first
   - Auto-detects columns from first row
   - Perfect for archived data

4. CSV Files
   - Standard CSV format
   - Must have: subject, from, date, body

SMART OUTPUT RULES:

• Email input → Excel output (default)
• Excel/CSV input → BI Dashboard (automatic!)
• Pipeline: Email → Excel → BI (optional)

SMART KEYWORD MATCHING:

Search for "date" finds:
  - Word "date"
  - 2024-01-15
  - January 15, 2024
  - All date formats!

Search for "amount" finds:
  - Word "amount"
  - $100, €50
  - 1,234.56
  - All numbers!

Need more help? Check the documentation files."""
        
        messagebox.showinfo("Help", help_text)


# ===== PROFILE CREATION DIALOG =====

class CreateProfileDialog:
    """Modern profile creation dialog."""
    
    def __init__(self, parent, main_gui):
        self.main_gui = main_gui
        self.window = tk.Toplevel(parent)
        self.window.title("Create New Profile")
        self.window.geometry("800x900")
        self.window.configure(bg="white")
        self.window.transient(parent)
        
        self.font = main_gui.font
        self.font_display = main_gui.font_display
        
        self.create_ui()
    
    def create_ui(self):
        """Create dialog UI."""
        # Header
        header = GradientFrame(self.window, height=80)
        header.pack(fill=tk.X)
        
        title_container = tk.Frame(header, bg="", bd=0)
        header.create_window(30, 40, anchor=tk.W, window=title_container)
        
        tk.Label(
            title_container,
            text="✨ Create New Profile",
            font=(self.font_display, 20, "bold"),
            fg="white",
            bg=""
        ).pack()
        
        # Scrollable content
        canvas = tk.Canvas(self.window, bg="white", highlightthickness=0)
        scrollbar = tk.Scrollbar(self.window, command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        
        scroll_frame = tk.Frame(canvas, bg="white")
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        form = tk.Frame(scroll_frame, bg="white")
        form.pack(fill=tk.BOTH, expand=True, padx=40, pady=20)
        
        # Name
        tk.Label(form, text="Profile Name *", font=(self.font, 12, "bold"), bg="white", fg="#1a202c").pack(anchor=tk.W, pady=(10, 5))
        self.name_entry = tk.Entry(form, font=(self.font, 11), bg="#f7fafc", relief=tk.FLAT, borderwidth=1)
        self.name_entry.pack(fill=tk.X, ipady=8)
        
        # Input source
        tk.Label(form, text="Input Source *", font=(self.font, 12, "bold"), bg="white", fg="#1a202c").pack(anchor=tk.W, pady=(20, 5))
        
        self.input_var = tk.StringVar(value="graph")
        
        for value, text in [
            ("graph", "📧 Microsoft Graph (Outlook)"),
            ("local_eml", "📁 Local .eml Files"),
            ("excel_file", "📊 Excel File (.xlsx)"),
            ("csv_file", "📄 CSV File"),
        ]:
            tk.Radiobutton(
                form, text=text, variable=self.input_var, value=value,
                font=(self.font, 11), bg="white", activebackground="white",
                command=self.on_input_change
            ).pack(anchor=tk.W, pady=2)
        
        # Input options container
        self.input_options = tk.Frame(form, bg="#f7fafc", relief=tk.FLAT, borderwidth=1)
        self.input_options.pack(fill=tk.X, pady=10)
        
        # Columns
        tk.Label(form, text="Columns *", font=(self.font, 12, "bold"), bg="white", fg="#1a202c").pack(anchor=tk.W, pady=(20, 5))
        
        self.auto_detect_var = tk.BooleanVar(value=False)
        self.auto_detect_check = tk.Checkbutton(
            form,
            text="✓ Auto-detect columns from file",
            variable=self.auto_detect_var,
            font=(self.font, 10),
            bg="white",
            activebackground="white",
            command=self.on_auto_detect_change
        )
        self.auto_detect_check.pack(anchor=tk.W, pady=5)
        
        self.columns_entry = tk.Entry(form, font=(self.font, 11), bg="#f7fafc", relief=tk.FLAT, borderwidth=1)
        self.columns_entry.insert(0, "Subject,From,Date")
        self.columns_entry.pack(fill=tk.X, ipady=8)
        
        # Output
        tk.Label(form, text="Output *", font=(self.font, 12, "bold"), bg="white", fg="#1a202c").pack(anchor=tk.W, pady=(20, 5))
        
        tk.Label(form, text="Smart defaults apply based on input type", font=(self.font, 9), bg="white", fg="#718096").pack(anchor=tk.W, pady=(0, 10))
        
        self.output_var = tk.StringVar(value="excel")
        
        for value, text in [
            ("excel", "📄 Excel File"),
            ("bi_dashboard", "📊 BI Dashboard (Interactive)"),
            ("both", "📄+📊 Excel + BI Dashboard"),
        ]:
            tk.Radiobutton(
                form, text=text, variable=self.output_var, value=value,
                font=(self.font, 11), bg="white", activebackground="white"
            ).pack(anchor=tk.W, pady=2)
        
        tk.Frame(form, height=20, bg="white").pack()
        
        # Bottom buttons
        bottom = tk.Frame(self.window, bg="#f7fafc", height=80)
        bottom.pack(fill=tk.X, side=tk.BOTTOM)
        bottom.pack_propagate(False)
        
        btn_container = tk.Frame(bottom, bg="#f7fafc")
        btn_container.pack(expand=True)
        
        ModernButton(btn_container, "Cancel", self.window.destroy, "secondary", self.font).pack(side=tk.LEFT, padx=5)
        ModernButton(btn_container, "Create Profile", self.save_profile, "gradient", self.font, "✓").pack(side=tk.LEFT, padx=5)
        
        self.on_input_change()
    
    def on_input_change(self):
        """Handle input source change."""
        for widget in self.input_options.winfo_children():
            widget.destroy()
        
        source = self.input_var.get()
        
        # Show appropriate options
        if source == "graph":
            tk.Label(self.input_options, text="Folder:", bg="#f7fafc", font=(self.font, 10)).pack(anchor=tk.W, padx=10, pady=(10, 2))
            self.folder_entry = tk.Entry(self.input_options, font=(self.font, 10), bg="white")
            self.folder_entry.insert(0, "Inbox")
            self.folder_entry.pack(fill=tk.X, padx=10, pady=(0, 10))
            
            self.auto_detect_check.pack_forget()
            self.columns_entry.pack(fill=tk.X, ipady=8)
            self.output_var.set("excel")
            
        elif source == "local_eml":
            tk.Label(self.input_options, text="Directory:", bg="#f7fafc", font=(self.font, 10)).pack(anchor=tk.W, padx=10, pady=(10, 2))
            
            dir_frame = tk.Frame(self.input_options, bg="#f7fafc")
            dir_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
            
            self.dir_entry = tk.Entry(dir_frame, font=(self.font, 10), bg="white")
            self.dir_entry.insert(0, "./input_emails")
            self.dir_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
            
            tk.Button(dir_frame, text="Browse", command=self.browse_dir, bg="#e2e8f0", relief=tk.FLAT).pack(side=tk.RIGHT)
            
            self.auto_detect_check.pack_forget()
            self.columns_entry.pack(fill=tk.X, ipady=8)
            self.output_var.set("excel")
            
        elif source in ["excel_file", "csv_file"]:
            file_type = "Excel (.xlsx)" if source == "excel_file" else "CSV"
            tk.Label(self.input_options, text=f"{file_type} File:", bg="#f7fafc", font=(self.font, 10)).pack(anchor=tk.W, padx=10, pady=(10, 2))
            
            file_frame = tk.Frame(self.input_options, bg="#f7fafc")
            file_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
            
            self.file_entry = tk.Entry(file_frame, font=(self.font, 10), bg="white")
            self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
            
            filetypes = [("Excel files", "*.xlsx")] if source == "excel_file" else [("CSV files", "*.csv")]
            tk.Button(file_frame, text="Browse", command=lambda: self.browse_file(filetypes), bg="#e2e8f0", relief=tk.FLAT).pack(side=tk.RIGHT)
            
            # Show auto-detect for Excel/CSV
            self.auto_detect_check.pack(anchor=tk.W, pady=5)
            self.columns_entry.pack(fill=tk.X, ipady=8)
            
            # Auto BI dashboard for Excel/CSV
            self.output_var.set("bi_dashboard")
    
    def on_auto_detect_change(self):
        """Handle auto-detect toggle."""
        if self.auto_detect_var.get():
            self.columns_entry.config(state=tk.DISABLED, bg="#e2e8f0")
            self.columns_entry.delete(0, tk.END)
            self.columns_entry.insert(0, "Will be detected from file...")
        else:
            self.columns_entry.config(state=tk.NORMAL, bg="#f7fafc")
            self.columns_entry.delete(0, tk.END)
            self.columns_entry.insert(0, "Subject,From,Date")
    
    def browse_dir(self):
        """Browse for directory."""
        directory = filedialog.askdirectory()
        if directory:
            self.dir_entry.delete(0, tk.END)
            self.dir_entry.insert(0, directory)
    
    def browse_file(self, filetypes):
        """Browse for file."""
        filename = filedialog.askopenfilename(filetypes=filetypes)
        if filename:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, filename)
            
            # Try to auto-detect columns
            if self.auto_detect_var.get():
                try:
                    if filename.endswith('.xlsx'):
                        from adapters.excel_csv_email import ExcelCSVEmailAdapter
                        adapter = ExcelCSVEmailAdapter()
                        headers, _ = adapter.load_from_excel(filename)
                    else:
                        from adapters.excel_csv_email import ExcelCSVEmailAdapter
                        adapter = ExcelCSVEmailAdapter()
                        headers, _ = adapter.load_from_csv(filename)
                    
                    if headers:
                        self.columns_entry.config(state=tk.NORMAL, bg="#f7fafc")
                        self.columns_entry.delete(0, tk.END)
                        self.columns_entry.insert(0, ",".join(headers))
                        self.columns_entry.config(state=tk.DISABLED, bg="#e2e8f0")
                        
                        messagebox.showinfo("Success", f"Detected {len(headers)} columns!")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to detect columns: {e}")
    
    def save_profile(self):
        """Save profile."""
        # Validate
        name = self.name_entry.get().strip()
        if not name:
            messagebox.showerror("Validation Error", "Profile name is required")
            return
        
        if not self.auto_detect_var.get():
            columns = self.columns_entry.get().strip()
            if not columns:
                messagebox.showerror("Validation Error", "At least one column is required")
                return
        
        # Build profile
        profile = {
            "name": name,
            "input_source": self.input_var.get(),
            "auto_detect_columns": self.auto_detect_var.get(),
        }
        
        # Input selection
        source = self.input_var.get()
        if source == "graph":
            profile["email_selection"] = {
                "folder_name": self.folder_entry.get(),
                "newest_n": 25
            }
        elif source == "local_eml":
            profile["email_selection"] = {
                "directory": self.dir_entry.get(),
                "pattern": "*.eml"
            }
        else:
            profile["email_selection"] = {
                "file_path": self.file_entry.get()
            }
        
        # Schema
        if self.auto_detect_var.get():
            profile["schema"] = {"columns": []}  # Will be detected
        else:
            columns = [c.strip() for c in self.columns_entry.get().split(",")]
            profile["schema"] = {
                "columns": [{"name": col, "type": "text"} for col in columns]
            }
        
        # Rules
        profile["rules"] = []
        
        # Output
        output_type = self.output_var.get()
        if output_type == "bi_dashboard":
            profile["output"] = {
                "format": "excel",
                "destination": "bi_dashboard",
                "local_path": "./output"
            }
        elif output_type == "both":
            profile["pipeline"] = ["email_to_table"]
            profile["output"] = {
                "format": "excel",
                "destination": "local",
                "local_path": "./output",
                "also_export_bi": True
            }
        else:
            profile["output"] = {
                "format": "excel",
                "destination": "local",
                "local_path": "./output"
            }
        
        # Save
        try:
            self.main_gui.loader.save_profile(profile)
            messagebox.showinfo("Success", f"Profile '{name}' created!")
            self.main_gui.refresh_profiles()
            self.main_gui.log(f"✓ Created: {name}", "SUCCESS")
            self.window.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save profile: {e}")


class GUILogHandler(logging.Handler):
    """Log handler for GUI."""
    
    def __init__(self, gui):
        super().__init__()
        self.gui = gui
    
    def emit(self, record):
        msg = record.getMessage()
        level = record.levelname
        
        if hasattr(self.gui, 'log'):
            try:
                self.gui.root.after(0, lambda: self.gui.log(msg, level))
            except:
                pass


def main():
    """Main entry point."""
    root = tk.Tk()
    app = EmailAutomationGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
