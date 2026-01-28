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
from run_wizard import acquire_graph_token, test_graph_capabilities

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
    
    app.run(host='0.0.0.0', port=3000, debug=True)
