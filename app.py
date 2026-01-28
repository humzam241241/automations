"""
Email Automation Pro - Web Application
Flask-based web interface running on localhost:3000
"""
import json
import logging
import os
from pathlib import Path
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file, session
from flask_cors import CORS
import threading
import secrets

from core.engine import ExecutionEngine
from core.profile_loader import ProfileLoader
from jobs.column_suggester import suggest_columns
from run_wizard import acquire_graph_token, test_graph_capabilities
from jobs.ai_helper import AIHelper
from jobs.column_suggester import ColumnSuggester

# Initialize Flask app
app = Flask(__name__)
app.secret_key = secrets.token_hex(32)
CORS(app)

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s | %(message)s')

# Global state
profile_loader = ProfileLoader()
current_user = None
users_file = Path("users.json")

# ===== USER MANAGEMENT =====

def load_users():
    """Load users from file."""
    if users_file.exists():
        with users_file.open('r') as f:
            return json.load(f)
    return {}

def save_users(users):
    """Save users to file."""
    with users_file.open('w') as f:
        json.dump(users, f, indent=2)

def get_current_user():
    """Get current user from session."""
    return session.get('username')

# ===== ROUTES =====

@app.route('/')
def index():
    """Main page."""
    return render_template('index.html')

@app.route('/api/login', methods=['POST'])
def login():
    """User login."""
    data = request.json
    username = data.get('username', '').strip()
    
    if not username:
        return jsonify({'success': False, 'message': 'Username required'})
    
    users = load_users()
    
    # Create user if doesn't exist
    if username not in users:
        users[username] = {
            'created_at': datetime.now().isoformat(),
            'last_login': datetime.now().isoformat()
        }
        save_users(users)
    else:
        users[username]['last_login'] = datetime.now().isoformat()
        save_users(users)
    
    session['username'] = username
    return jsonify({'success': True, 'username': username})

@app.route('/api/logout', methods=['POST'])
def logout():
    """User logout."""
    session.pop('username', None)
    return jsonify({'success': True})

@app.route('/api/current-user')
def current_user_info():
    """Get current user info."""
    username = get_current_user()
    if username:
        users = load_users()
        return jsonify({
            'logged_in': True,
            'username': username,
            'info': users.get(username, {})
        })
    return jsonify({'logged_in': False})

@app.route('/api/users')
def list_users():
    """List all users."""
    users = load_users()
    return jsonify({
        'users': [
            {
                'username': username,
                'created_at': info.get('created_at'),
                'last_login': info.get('last_login')
            }
            for username, info in users.items()
        ]
    })

@app.route('/api/check-graph')
def check_graph():
    """Check Microsoft Graph connection."""
    token = acquire_graph_token()
    if token:
        capabilities = test_graph_capabilities(token)
        return jsonify({
            'available': True,
            'capabilities': capabilities
        })
    return jsonify({'available': False})

@app.route('/api/profiles')
def list_profiles():
    """List all profiles."""
    profiles = profile_loader.list_profiles()
    return jsonify({'profiles': profiles})

@app.route('/api/profiles/<name>')
def get_profile(name):
    """Get profile details."""
    profile = profile_loader.load_profile(name)
    if profile:
        return jsonify({'success': True, 'profile': profile})
    return jsonify({'success': False, 'message': 'Profile not found'}), 404

@app.route('/api/profiles', methods=['POST'])
def create_profile():
    """Create new profile."""
    username = get_current_user()
    if not username:
        return jsonify({'success': False, 'message': 'Not logged in'}), 401
    
    profile = request.json
    profile['created_by'] = username
    profile['created_at'] = datetime.now().isoformat()
    
    try:
        profile_loader.save_profile(profile)
        return jsonify({'success': True, 'message': 'Profile created'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 400

@app.route('/api/profiles/<name>', methods=['DELETE'])
def delete_profile(name):
    """Delete profile."""
    username = get_current_user()
    if not username:
        return jsonify({'success': False, 'message': 'Not logged in'}), 401
    
    try:
        profile_path = Path("profiles") / f"{name}.json"
        if profile_path.exists():
            profile_path.unlink()
            return jsonify({'success': True, 'message': 'Profile deleted'})
        return jsonify({'success': False, 'message': 'Profile not found'}), 404
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 400

@app.route('/api/profiles/<name>/run', methods=['POST'])
def run_profile(name):
    """Run profile."""
    username = get_current_user()
    if not username:
        return jsonify({'success': False, 'message': 'Not logged in'}), 401
    
    profile = profile_loader.load_profile(name)
    if not profile:
        return jsonify({'success': False, 'message': 'Profile not found'}), 404
    
    def execute():
        try:
            token = acquire_graph_token()
            engine = ExecutionEngine(access_token=token)
            result = engine.run_profile(profile)
            return result
        except Exception as e:
            logging.error(f"Profile execution failed: {e}")
            return {'status': 'error', 'message': str(e)}
    
    # Execute in background
    result = execute()
    return jsonify(result)

@app.route('/api/synonyms')
def get_synonyms():
    """Get all synonyms."""
    try:
        config_path = Path("config/synonyms.json")
        if config_path.exists():
            with open(config_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return jsonify({
                    'success': True,
                    'synonyms': data.get('synonyms', {})
                })
        return jsonify({'success': True, 'synonyms': {}})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 400

@app.route('/api/synonyms', methods=['POST'])
def save_synonyms():
    """Save synonyms."""
    try:
        data = request.json
        synonyms = data.get('synonyms', {})
        
        config_path = Path("config/synonyms.json")
        
        # Load existing config to preserve other fields
        existing = {}
        if config_path.exists():
            with open(config_path, 'r', encoding='utf-8') as f:
                existing = json.load(f)
        
        # Update synonyms
        existing['synonyms'] = synonyms
        existing['_description'] = "Column name synonyms - add your own mappings here!"
        
        # Save
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(existing, f, indent=2)
        
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 400

@app.route('/api/synonyms/reset', methods=['POST'])
def reset_synonyms():
    """Reset synonyms to defaults."""
    try:
        default_synonyms = {
            "qc": ["quality criticality", "quality crit", "criticality", "qc level"],
            "quality criticality": ["qc", "quality crit", "criticality"],
            "wo": ["work order", "work order number", "workorder", "wo #", "wo number"],
            "work order": ["wo", "wo #", "work order number", "workorder"],
            "order": ["order number", "order #", "order no", "order id"],
            "order type": ["order type", "type", "wo type"],
            "equipment": ["equipment id", "equip", "eq", "equipment number", "equip #"],
            "building": ["bldg", "bldg #", "building number", "building #", "bld"],
            "date": ["due date", "finish date", "completion date", "date completed"],
            "due date": ["initial due date", "target date", "deadline", "due"],
            "status": ["completion status", "state", "current status"],
            "completion status": ["status", "complete status", "completed"],
            "description": ["desc", "details", "description of technical object"],
            "main work center": ["work center", "mwc", "center", "main center"],
            "comments": ["notes", "remarks", "reconciliation comments"],
            "mi": ["mi #", "mi number", "mi no"],
            "mm": ["mm #", "mm number", "mm no"],
            "capa": ["capa #", "capa number", "capa no"],
            "from": ["sender", "sent by", "from address"],
            "to": ["recipient", "sent to", "to address"],
            "subject": ["title", "topic", "re", "subj"]
        }
        
        config_path = Path("config/synonyms.json")
        config_data = {
            "_description": "Column name synonyms - add your own mappings here!",
            "synonyms": default_synonyms,
            "abbreviations": {
                "qc": "Quality Criticality",
                "wo": "Work Order",
                "eq": "Equipment",
                "bldg": "Building",
                "mwc": "Main Work Center"
            }
        }
        
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=2)
        
        return jsonify({'success': True, 'synonyms': default_synonyms})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 400

@app.route('/api/ai-suggest', methods=['POST'])
def ai_suggest():
    """Return an AI suggestion for a low-confidence column."""
    helper = AIHelper()
    if not helper.is_configured():
        return jsonify({'success': False, 'message': 'AI helper not configured'}), 503

    data = request.json or {}
    column = data.get('column')
    email_data = data.get('email_data', {})

    if not column:
        return jsonify({'success': False, 'message': 'Column name is required'}), 400

    suggestion = helper.suggest(column, email_data)
    if not suggestion:
        return jsonify({'success': False, 'message': 'No suggestion available'})

    return jsonify({'success': True, 'suggestion': suggestion})

@app.route('/api/suggest-columns', methods=['POST'])
def suggest_columns():
    """Suggest column names based on sample emails."""
    data = request.json or {}
    profile = data.get('profile')
    if not profile:
        return jsonify({'success': False, 'message': 'Profile data required'}), 400

    try:
        token = None
        if profile.get('input_source') == 'graph':
            token = acquire_graph_token()
        engine = ExecutionEngine(access_token=token)
        emails = engine._load_emails(profile)
        if not emails:
            return jsonify({'success': False, 'message': 'No emails were loaded for suggestions.'})

        suggester = ColumnSuggester(emails, max_samples=25)
        suggestions = suggester.suggest()
        return jsonify({'success': True, 'suggestions': suggestions})
    except Exception as e:
        logging.error(f"Column suggestion failed: {e}")
        return jsonify({'success': False, 'message': str(e)}), 400

@app.route('/api/detect-columns', methods=['POST'])
def detect_columns():
    """Auto-detect columns from file."""
    data = request.json
    file_path = data.get('file_path')
    input_type = data.get('input_type')
    
    if not file_path or not os.path.exists(file_path):
        return jsonify({'success': False, 'message': 'File not found'}), 400
    
    try:
        if input_type == 'excel_file':
            from adapters.excel_csv_email import ExcelCSVEmailAdapter
            adapter = ExcelCSVEmailAdapter()
            headers, _ = adapter.load_from_excel(file_path)
        elif input_type == 'csv_file':
            from adapters.excel_csv_email import ExcelCSVEmailAdapter
            adapter = ExcelCSVEmailAdapter()
            headers, _ = adapter.load_from_csv(file_path)
        else:
            return jsonify({'success': False, 'message': 'Unsupported file type'}), 400
        
        return jsonify({'success': True, 'columns': headers})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 400

@app.route('/api/suggest-columns', methods=['POST'])
def suggest_columns_api():
    """
    Suggest column names from a sample of emails based on the profile payload.
    Expects a profile-like JSON with input_source and email_selection.
    """
    try:
        payload = request.json or {}
        profile = payload.get("profile", payload)  # allow both top-level and nested

        # Create a lightweight engine to load sample emails
        engine = ExecutionEngine()
        emails = engine._load_emails(profile)

        # Limit sample size
        emails = emails[:50]

        suggestions = suggest_columns(emails, top_n=15)
        return jsonify({'success': True, 'suggestions': suggestions})
    except Exception as e:
        logging.error(f"Suggest columns failed: {e}")
        return jsonify({'success': False, 'message': str(e)}), 400

@app.route('/api/open-file-dialog', methods=['POST'])
def open_file_dialog():
    """Open native file dialog using tkinter."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        
        data = request.json
        dialog_type = data.get('type', 'file')
        file_types = data.get('file_types', [])
        initial_dir = data.get('initial_dir', '')
        
        # Create hidden root window
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        root.focus_force()
        
        # Find the REAL Downloads folder (check OneDrive first)
        if not initial_dir:
            possible_downloads = [
                Path.home() / 'OneDrive - Sanofi' / 'Downloads',
                Path.home() / 'OneDrive' / 'Downloads',
                Path.home() / 'Downloads',
                Path(os.environ.get('USERPROFILE', '')) / 'Downloads',
            ]
            
            for dl in possible_downloads:
                if dl.exists() and any(dl.iterdir()):
                    initial_dir = str(dl)
                    break
            
            if not initial_dir:
                initial_dir = str(Path.home() / 'Downloads')
        
        if dialog_type == 'directory':
            # FOR FOLDER SELECTION: Let user pick ANY file, then get its folder
            # This way they can SEE their files while browsing!
            result = filedialog.askopenfilename(
                title='Select ANY file inside the folder you want (we will use that folder)',
                initialdir=initial_dir,
                filetypes=[
                    ('All Files', '*.*'),
                    ('Email Files', '*.eml *.msg'),
                    ('Excel Files', '*.xlsx *.xls'),
                    ('CSV Files', '*.csv'),
                ]
            )
            
            if result:
                # Extract the folder from the selected file
                path = str(Path(result).parent)
            else:
                path = None
        else:
            # File picker with all important types visible
            multiple = data.get('multiple', False)
            
            if file_types:
                filetypes = [
                    (ft['name'], ft['pattern']) for ft in file_types
                ]
            else:
                # Default: show all important file types
                filetypes = [
                    ('All Files', '*.*'),  # Put "All Files" FIRST so user sees everything
                    ('Email Files', '*.eml *.msg'),
                    ('Excel Files', '*.xlsx *.xls'),
                    ('CSV Files', '*.csv'),
                    ('PDF Files', '*.pdf'),
                    ('Word Documents', '*.docx *.doc'),
                    ('Text Files', '*.txt'),
                ]
            
            if multiple:
                # MULTIPLE FILE SELECTION
                paths = filedialog.askopenfilenames(
                    title='Select Files (Ctrl+Click for multiple)',
                    initialdir=initial_dir,
                    filetypes=filetypes
                )
                root.destroy()
                
                if paths:
                    return jsonify({
                        'success': True,
                        'paths': list(paths),
                        'path': paths[0] if len(paths) == 1 else '; '.join(paths)
                    })
                else:
                    return jsonify({
                        'success': False,
                        'message': 'No files selected'
                    })
            else:
                # Single file selection
                path = filedialog.askopenfilename(
                    title='Select File',
                    initialdir=initial_dir,
                    filetypes=filetypes
                )
                root.destroy()
                
                if path:
                    return jsonify({
                        'success': True,
                        'path': path,
                        'paths': [path]
                    })
                else:
                    return jsonify({
                        'success': False,
                        'message': 'No file selected'
                    })
        
        root.destroy()
        return jsonify({'success': False, 'message': 'No selection made'})
    
    except Exception as e:
        logging.error(f"File dialog failed: {e}")
        return jsonify({
            'success': False,
            'message': str(e)
        }), 400

@app.route('/api/browse-directory', methods=['POST'])
def browse_directory():
    """Browse file system - FULL ACCESS."""
    data = request.json
    current_path = data.get('path', '')
    
    try:
        # If no path provided, start at user's home or desktop
        if not current_path:
            # Try Desktop first (most common place for files)
            desktop = Path.home() / 'Desktop'
            if desktop.exists():
                path = desktop
            else:
                # Try OneDrive Desktop
                onedrive_desktop = Path.home() / 'OneDrive - Sanofi' / 'Desktop'
                if onedrive_desktop.exists():
                    path = onedrive_desktop
                else:
                    path = Path.home()
        else:
            path = Path(current_path)
        
        # Validate path exists
        if not path.exists():
            # Try parent directory
            path = path.parent if path.parent.exists() else Path.home()
        
        # If path is a file, go to its directory
        if path.is_file():
            path = path.parent
        
        items = []
        
        # Add drives on Windows (C:\, D:\, etc.)
        if os.name == 'nt':
            # Check if we're at the root level
            if len(path.parts) <= 1:
                # Show available drives
                import string
                for letter in string.ascii_uppercase:
                    drive = Path(f'{letter}:\\')
                    if drive.exists():
                        try:
                            items.append({
                                'name': f'{letter}:\\ Drive',
                                'path': str(drive),
                                'type': 'drive',
                                'size': '',
                                'modified': ''
                            })
                        except:
                            pass
        
        # Add parent directory (if not at root)
        if path.parent != path:
            items.append({
                'name': '⬆️ Parent Folder',
                'path': str(path.parent),
                'type': 'directory',
                'size': '',
                'modified': ''
            })
        
        # List directories first, then files
        try:
            all_items = list(path.iterdir())
        except PermissionError:
            return jsonify({
                'success': False, 
                'message': 'Permission denied. Try another folder.',
                'current_path': str(path)
            }), 403
        except Exception as e:
            return jsonify({
                'success': False,
                'message': f'Cannot access folder: {str(e)}',
                'current_path': str(path)
            }), 400
        
        # Separate and sort
        directories = []
        files = []
        
        for item in all_items:
            try:
                if item.is_dir():
                    directories.append(item)
                elif item.is_file():
                    files.append(item)
            except:
                pass  # Skip items we can't access
        
        directories.sort(key=lambda x: x.name.lower())
        files.sort(key=lambda x: x.name.lower())
        
        # Add directories
        for directory in directories:
            try:
                items.append({
                    'name': directory.name,
                    'path': str(directory),
                    'type': 'directory',
                    'size': '',
                    'modified': datetime.fromtimestamp(directory.stat().st_mtime).strftime('%Y-%m-%d %H:%M')
                })
            except:
                pass  # Skip directories we can't stat
        
        # Add files
        for file in files:
            try:
                stat = file.stat()
                size_bytes = stat.st_size
                
                # Format size
                if size_bytes >= 1024 * 1024:
                    size_str = f'{size_bytes / (1024 * 1024):.2f} MB'
                elif size_bytes >= 1024:
                    size_str = f'{size_bytes / 1024:.1f} KB'
                else:
                    size_str = f'{size_bytes} B'
                
                items.append({
                    'name': file.name,
                    'path': str(file),
                    'type': 'file',
                    'extension': file.suffix.lower(),
                    'size': size_str,
                    'modified': datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M')
                })
            except:
                pass  # Skip files we can't stat
        
        return jsonify({
            'success': True,
            'current_path': str(path),
            'items': items,
            'is_windows': os.name == 'nt'
        })
    
    except Exception as e:
        logging.error(f"Browse directory failed: {e}")
        return jsonify({
            'success': False, 
            'message': str(e),
            'current_path': str(Path.home())
        }), 400

# ===== MAIN =====

if __name__ == '__main__':
    print("="*60)
    print()
    print("   Email Automation Pro - Web Application")
    print()
    print("   Server running on: http://localhost:3000")
    print()
    print("   Press Ctrl+C to stop the server")
    print()
    print("="*60)
    print()
    
    # Run without the Flask reloader so the window stays open on exit/failure.
    app.run(host='0.0.0.0', port=3000, debug=False, use_reloader=False)
