// Email Automation Pro - Web App JavaScript

let currentProfile = null;
let currentUser = null;

// ===== INITIALIZATION =====

document.addEventListener('DOMContentLoaded', () => {
    checkLogin();
});

// ===== USER MANAGEMENT =====

async function checkLogin() {
    try {
        const response = await fetch('/api/current-user');
        const data = await response.json();
        
        if (data.logged_in) {
            currentUser = data.username;
            showApp();
            initialize();
        } else {
            showLogin();
        }
    } catch (error) {
        console.error('Login check failed:', error);
        showLogin();
    }
}

function showLogin() {
    document.getElementById('loginModal').style.display = 'flex';
    document.getElementById('app').style.display = 'none';
}

function showApp() {
    document.getElementById('loginModal').style.display = 'none';
    document.getElementById('app').style.display = 'block';
    document.getElementById('currentUser').textContent = currentUser;
}

async function login() {
    const username = document.getElementById('usernameInput').value.trim();
    
    if (!username) {
        alert('Please enter a username');
        return;
    }
    
    try {
        const response = await fetch('/api/login', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({username})
        });
        
        const data = await response.json();
        
        if (data.success) {
            currentUser = data.username;
            showApp();
            initialize();
        } else {
            alert(data.message);
        }
    } catch (error) {
        console.error('Login failed:', error);
        alert('Login failed. Please try again.');
    }
}

async function showUserList() {
    try {
        const response = await fetch('/api/users');
        const data = await response.json();
        
        const userListDiv = document.getElementById('userList');
        userListDiv.innerHTML = '';
        
        if (data.users.length === 0) {
            userListDiv.innerHTML = '<p style="color: #999;">No existing users</p>';
        } else {
            data.users.forEach(user => {
                const userItem = document.createElement('div');
                userItem.className = 'user-item';
                userItem.innerHTML = `
                    <div class="user-item-name">${user.username}</div>
                    <div class="user-item-meta">Last login: ${formatDate(user.last_login)}</div>
                `;
                userItem.onclick = () => loadUser(user.username);
                userListDiv.appendChild(userItem);
            });
        }
        
        document.getElementById('userListContainer').style.display = 'block';
    } catch (error) {
        console.error('Failed to load users:', error);
    }
}

async function loadUser(username) {
    document.getElementById('usernameInput').value = username;
    await login();
}

async function logout() {
    try {
        await fetch('/api/logout', {method: 'POST'});
        currentUser = null;
        currentProfile = null;
        showLogin();
        document.getElementById('usernameInput').value = '';
        document.getElementById('userListContainer').style.display = 'none';
    } catch (error) {
        console.error('Logout failed:', error);
    }
}

// ===== INITIALIZATION =====

async function initialize() {
    log('Initializing...', 'info');
    await checkGraphStatus();
    await loadProfiles();
    log('Ready!', 'success');
}

async function checkGraphStatus() {
    try {
        const response = await fetch('/api/check-graph');
        const data = await response.json();
        
        const graphIcon = document.getElementById('graphStatus');
        const mailIcon = document.getElementById('mailStatus');
        const driveIcon = document.getElementById('driveStatus');
        
        if (data.available) {
            graphIcon.className = 'status-icon connected';
            log('✓ Microsoft Graph connected', 'success');
            
            if (data.capabilities.mail) {
                mailIcon.className = 'status-icon connected';
                log('✓ Mail access OK', 'success');
            } else {
                mailIcon.className = 'status-icon disconnected';
                log('✗ Mail access denied', 'warning');
            }
            
            if (data.capabilities.drive) {
                driveIcon.className = 'status-icon connected';
                log('✓ OneDrive access OK', 'success');
            } else {
                driveIcon.className = 'status-icon disconnected';
                log('✗ OneDrive access denied', 'warning');
            }
        } else {
            graphIcon.className = 'status-icon disconnected';
            mailIcon.className = 'status-icon disconnected';
            driveIcon.className = 'status-icon disconnected';
            log('✗ Microsoft Graph not available', 'warning');
            log('ℹ You can still use local files and Excel/CSV input', 'info');
        }
    } catch (error) {
        console.error('Graph check failed:', error);
        log('✗ Failed to check Graph status', 'error');
    }
}

// ===== PROFILES =====

async function loadProfiles() {
    try {
        const response = await fetch('/api/profiles');
        const data = await response.json();
        
        const profileList = document.getElementById('profileList');
        profileList.innerHTML = '';
        
        if (data.profiles.length === 0) {
            profileList.innerHTML = '<div style="padding: 20px; text-align: center; color: #999;">No profiles yet. Click + to create one!</div>';
        } else {
            data.profiles.forEach(profileName => {
                const item = document.createElement('div');
                item.className = 'profile-item';
                item.innerHTML = `
                    <div class="profile-item-name">${profileName}</div>
                    <div class="profile-item-meta">Click to view details</div>
                `;
                item.onclick = () => selectProfile(profileName);
                profileList.appendChild(item);
            });
        }
        
        log(`Loaded ${data.profiles.length} profile(s)`, 'info');
    } catch (error) {
        console.error('Failed to load profiles:', error);
        log('✗ Failed to load profiles', 'error');
    }
}

async function selectProfile(name) {
    try {
        const response = await fetch(`/api/profiles/${name}`);
        const data = await response.json();
        
        if (data.success) {
            currentProfile = data.profile;
            
            // Update UI
            document.querySelectorAll('.profile-item').forEach(item => {
                item.classList.remove('active');
                if (item.querySelector('.profile-item-name').textContent === name) {
                    item.classList.add('active');
                }
            });
            
            // Show details
            const detailsDiv = document.getElementById('profileDetails');
            detailsDiv.innerHTML = `
                <h3>${currentProfile.name}</h3>
                <div class="profile-detail-item">
                    <div class="profile-detail-label">Input Source:</div>
                    <div class="profile-detail-value">${formatInputSource(currentProfile.input_source)}</div>
                </div>
                <div class="profile-detail-item">
                    <div class="profile-detail-label">Columns:</div>
                    <div class="profile-detail-value">${formatColumns(currentProfile.schema)}</div>
                </div>
                <div class="profile-detail-item">
                    <div class="profile-detail-label">Output:</div>
                    <div class="profile-detail-value">${formatOutput(currentProfile.output)}</div>
                </div>
            `;
            
            document.getElementById('profileActions').style.display = 'block';
            log(`Selected: ${name}`, 'info');
        }
    } catch (error) {
        console.error('Failed to load profile:', error);
        log('✗ Failed to load profile details', 'error');
    }
}

async function deleteProfile() {
    if (!currentProfile) return;
    
    if (!confirm(`Are you sure you want to delete "${currentProfile.name}"?`)) {
        return;
    }
    
    try {
        const response = await fetch(`/api/profiles/${currentProfile.name}`, {
            method: 'DELETE'
        });
        
        const data = await response.json();
        
        if (data.success) {
            log(`✓ Deleted: ${currentProfile.name}`, 'success');
            currentProfile = null;
            document.getElementById('profileActions').style.display = 'none';
            await loadProfiles();
        } else {
            log(`✗ Delete failed: ${data.message}`, 'error');
        }
    } catch (error) {
        console.error('Delete failed:', error);
        log('✗ Failed to delete profile', 'error');
    }
}

async function runProfile() {
    if (!currentProfile) return;
    
    log(`▶ Running: ${currentProfile.name}`, 'info');
    
    try {
        const response = await fetch(`/api/profiles/${currentProfile.name}/run`, {
            method: 'POST'
        });
        
        const result = await response.json();
        
        if (result.status === 'success') {
            log(`✓ Success! Processed ${result.emails_processed} emails`, 'success');
            log(`Output: ${result.output.path}`, 'success');
            
            // Check if BI dashboard
            if (result.output.bi_dashboard && result.output.bi_dashboard.enabled) {
                log(`📊 BI Dashboard: ${result.output.bi_dashboard.path}`, 'success');
                log('Dashboard opened in browser!', 'info');
            }
        } else {
            log(`✗ Failed: ${result.message}`, 'error');
        }
    } catch (error) {
        console.error('Run failed:', error);
        log('✗ Profile execution failed', 'error');
    }
}

// ===== CREATE PROFILE =====

function showCreateProfile() {
    document.getElementById('createProfileModal').style.display = 'flex';
    updateInputOptions();
}

function closeCreateProfile() {
    document.getElementById('createProfileModal').style.display = 'none';
    // Reset form
    document.getElementById('profileName').value = '';
    document.getElementById('columns').value = '';
    document.getElementById('autoDetect').checked = false;
}

function updateInputOptions() {
    const inputSource = document.getElementById('inputSource').value;
    const optionsDiv = document.getElementById('inputOptions');
    const outputType = document.getElementById('outputType');
    const outputHint = document.getElementById('outputHint');
    
    // Smart defaults
    if (inputSource === 'graph' || inputSource === 'local_eml') {
        outputType.value = 'excel';
        outputHint.textContent = '✅ Auto-selected: Excel (email input)';
    } else if (inputSource === 'excel_file' || inputSource === 'csv_file') {
        outputType.value = 'bi_dashboard';
        outputHint.textContent = '✅ Auto-selected: BI Dashboard (data input)';
    }
    
    // Show appropriate input options
    if (inputSource === 'graph') {
        optionsDiv.innerHTML = `
            <label>Folder Name</label>
            <input type="text" id="folderName" value="Inbox">
        `;
    } else if (inputSource === 'local_eml') {
        optionsDiv.innerHTML = `
            <label>Directory Path</label>
            <div class="file-input-group">
                <input type="text" id="emlDirectory" value="./input_emails">
                <button onclick="browseDirectory('emlDirectory')" class="btn-browse">📁 Browse</button>
            </div>
        `;
    } else if (inputSource === 'excel_file' || inputSource === 'csv_file') {
        const fileType = inputSource === 'excel_file' ? '.xlsx' : '.csv';
        optionsDiv.innerHTML = `
            <label>File Path</label>
            <div class="file-input-group">
                <input type="text" id="filePath" placeholder="C:\\path\\to\\file${fileType}">
                <button onclick="browseFile('filePath', '${inputSource}')" class="btn-browse">📁 Browse</button>
            </div>
            <button onclick="detectColumns()" class="btn-secondary" style="margin-top: 10px; width: 100%;">🔍 Detect Columns</button>
        `;
    }
}

function toggleAutoDetect() {
    const autoDetect = document.getElementById('autoDetect').checked;
    const columnsSection = document.getElementById('columnsSection');
    const columnsInput = document.getElementById('columns');
    
    if (autoDetect) {
        columnsInput.disabled = true;
        columnsInput.placeholder = 'Will be detected from file...';
        columnsInput.style.background = '#f7fafc';
    } else {
        columnsInput.disabled = false;
        columnsInput.placeholder = 'Subject,From,Date,Priority';
        columnsInput.style.background = 'white';
    }
}

async function detectColumns() {
    const filePath = document.getElementById('filePath').value.trim();
    const inputSource = document.getElementById('inputSource').value;
    
    if (!filePath) {
        alert('Please enter a file path');
        return;
    }
    
    try {
        const response = await fetch('/api/detect-columns', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({
                file_path: filePath,
                input_type: inputSource
            })
        });
        
        const data = await response.json();
        
        if (data.success) {
            document.getElementById('columns').value = data.columns.join(',');
            document.getElementById('autoDetect').checked = true;
            toggleAutoDetect();
            alert(`Detected ${data.columns.length} columns!`);
        } else {
            alert(`Failed to detect columns: ${data.message}`);
        }
    } catch (error) {
        console.error('Column detection failed:', error);
        alert('Failed to detect columns');
    }
}

async function saveProfile() {
    const name = document.getElementById('profileName').value.trim();
    const inputSource = document.getElementById('inputSource').value;
    const columns = document.getElementById('columns').value.trim();
    const outputType = document.getElementById('outputType').value;
    const autoDetect = document.getElementById('autoDetect').checked;
    
    if (!name) {
        alert('Profile name is required');
        return;
    }
    
    if (!autoDetect && !columns) {
        alert('Please specify columns or enable auto-detect');
        return;
    }
    
    // Build profile
    const profile = {
        name: name,
        input_source: inputSource,
        auto_detect_columns: autoDetect,
        email_selection: {},
        schema: {
            columns: autoDetect ? [] : columns.split(',').map(c => ({name: c.trim(), type: 'text'}))
        },
        rules: [],
        output: {
            format: outputType === 'bi_dashboard' ? 'excel' : 'excel',
            destination: outputType === 'bi_dashboard' ? 'bi_dashboard' : 'local',
            local_path: './output',
            also_export_bi: outputType === 'both'
        }
    };
    
    // Add input-specific options
    if (inputSource === 'graph') {
        profile.email_selection = {
            folder_name: document.getElementById('folderName').value,
            newest_n: 25
        };
    } else if (inputSource === 'local_eml') {
        profile.email_selection = {
            directory: document.getElementById('emlDirectory').value,
            pattern: '*.eml'
        };
    } else if (inputSource === 'excel_file' || inputSource === 'csv_file') {
        profile.email_selection = {
            file_path: document.getElementById('filePath').value
        };
    }
    
    try {
        const response = await fetch('/api/profiles', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(profile)
        });
        
        const data = await response.json();
        
        if (data.success) {
            log(`✓ Created: ${name}`, 'success');
            closeCreateProfile();
            await loadProfiles();
        } else {
            alert(`Failed to create profile: ${data.message}`);
        }
    } catch (error) {
        console.error('Save failed:', error);
        alert('Failed to create profile');
    }
}

// ===== UTILITIES =====

function log(message, type = 'info') {
    const logDiv = document.getElementById('activityLog');
    const timestamp = new Date().toLocaleTimeString();
    const entry = document.createElement('div');
    entry.className = `log-entry ${type}`;
    entry.textContent = `[${timestamp}] ${message}`;
    logDiv.appendChild(entry);
    logDiv.scrollTop = logDiv.scrollHeight;
}

function formatInputSource(source) {
    const map = {
        'graph': '📧 Microsoft Graph',
        'local_eml': '📁 Local .eml Files',
        'excel_file': '📊 Excel File',
        'csv_file': '📄 CSV File'
    };
    return map[source] || source;
}

function formatColumns(schema) {
    if (!schema || !schema.columns) return 'None';
    return schema.columns.map(c => c.name).join(', ');
}

function formatOutput(output) {
    if (!output) return 'Excel';
    if (output.destination === 'bi_dashboard') return '📊 BI Dashboard';
    if (output.also_export_bi) return '📄+📊 Excel + BI Dashboard';
    return '📄 Excel File';
}

function formatDate(dateStr) {
    if (!dateStr) return 'Never';
    const date = new Date(dateStr);
    return date.toLocaleDateString() + ' ' + date.toLocaleTimeString();
}

// ===== FILE BROWSER =====

let currentBrowserPath = '';
let browserTargetInput = '';
let browserFileType = '';

async function browseDirectory(targetInputId) {
    browserTargetInput = targetInputId;
    browserFileType = 'directory';
    
    // Get initial path from input or use home directory
    const input = document.getElementById(targetInputId);
    const initialPath = input.value.trim() || '';
    
    await loadDirectory(initialPath || '');
    document.getElementById('fileBrowserModal').style.display = 'flex';
}

async function browseFile(targetInputId, fileType) {
    browserTargetInput = targetInputId;
    browserFileType = fileType;
    
    // Get initial path from input or use home directory
    const input = document.getElementById(targetInputId);
    let initialPath = input.value.trim() || '';
    
    // If path is a file, get its directory
    if (initialPath && !initialPath.endsWith('\\') && !initialPath.endsWith('/')) {
        initialPath = initialPath.substring(0, initialPath.lastIndexOf('\\') || initialPath.lastIndexOf('/'));
    }
    
    await loadDirectory(initialPath || '');
    document.getElementById('fileBrowserModal').style.display = 'flex';
}

async function loadDirectory(path) {
    try {
        const response = await fetch('/api/browse-directory', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({path: path})
        });
        
        const data = await response.json();
        
        if (data.success) {
            currentBrowserPath = data.current_path;
            document.getElementById('currentPath').textContent = data.current_path;
            
            const contentDiv = document.getElementById('browserContent');
            contentDiv.innerHTML = '';
            
            // If browsing for directory, add "Select This Folder" button at top
            if (browserFileType === 'directory') {
                const selectBtn = document.createElement('div');
                selectBtn.style.cssText = 'padding: 16px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; text-align: center; font-weight: 600; cursor: pointer; border-bottom: 2px solid #e2e8f0;';
                selectBtn.innerHTML = '✓ Select This Folder';
                selectBtn.onclick = () => selectDirectory(data.current_path);
                contentDiv.appendChild(selectBtn);
            }
            
            if (data.items.length === 0) {
                const emptyDiv = document.createElement('div');
                emptyDiv.style.cssText = 'padding: 20px; text-align: center; color: #999;';
                emptyDiv.textContent = 'Empty directory';
                contentDiv.appendChild(emptyDiv);
                return;
            }
            
            data.items.forEach(item => {
                const itemDiv = document.createElement('div');
                itemDiv.className = `browser-item ${item.type}`;
                
                // Set icons
                let icon = '📁';
                if (item.type === 'drive') {
                    icon = '💾';
                } else if (item.type === 'directory') {
                    icon = '📁';
                } else if (item.type === 'file') {
                    if (item.extension === '.xlsx') icon = '📊';
                    else if (item.extension === '.csv') icon = '📄';
                    else if (item.extension === '.eml') icon = '📧';
                    else if (item.extension === '.msg') icon = '✉️';
                    else if (item.extension === '.txt') icon = '📝';
                    else if (item.extension === '.pdf') icon = '📕';
                    else icon = '📄';
                }
                
                itemDiv.innerHTML = `
                    <div class="browser-icon">${icon}</div>
                    <div class="browser-info">
                        <div class="browser-name">${escapeHtml(item.name)}</div>
                        <div class="browser-meta">
                            <span>${item.size}</span>
                            <span>${item.modified}</span>
                        </div>
                    </div>
                `;
                
                // Handle clicks
                if (item.type === 'directory' || item.type === 'drive') {
                    // Folders are always clickable for navigation
                    itemDiv.style.cursor = 'pointer';
                    
                    // If browsing for directory, allow selecting folders
                    if (browserFileType === 'directory') {
                        // Double-click to select folder
                        itemDiv.ondblclick = () => selectDirectory(item.path);
                        // Single click to navigate into it
                        itemDiv.onclick = () => loadDirectory(item.path);
                        // Visual hint for selectable folders
                        itemDiv.style.backgroundColor = '#f0f9ff';
                        itemDiv.title = 'Double-click to select, single-click to open';
                    } else {
                        // Just navigate
                        itemDiv.onclick = () => loadDirectory(item.path);
                    }
                } else if (item.type === 'file') {
                    // Check if file type matches what we're looking for
                    let allowSelection = false;
                    
                    if (browserFileType === 'excel_file' && item.extension === '.xlsx') {
                        allowSelection = true;
                    } else if (browserFileType === 'csv_file' && item.extension === '.csv') {
                        allowSelection = true;
                    } else if (browserFileType === 'directory') {
                        // When browsing for directory, don't allow file selection
                        allowSelection = false;
                    }
                    
                    if (allowSelection) {
                        itemDiv.onclick = () => selectFile(item.path);
                        itemDiv.style.cursor = 'pointer';
                        // Highlight selectable files
                        itemDiv.style.backgroundColor = '#f0f9ff';
                    } else {
                        itemDiv.style.opacity = '0.5';
                        itemDiv.style.cursor = 'not-allowed';
                    }
                }
                
                contentDiv.appendChild(itemDiv);
            });
        } else {
            // Show error but don't close the browser
            log(`⚠ ${data.message}`, 'warning');
            const contentDiv = document.getElementById('browserContent');
            contentDiv.innerHTML = `
                <div style="padding: 20px; text-align: center; color: #f56565;">
                    <div style="font-size: 48px; margin-bottom: 10px;">⚠️</div>
                    <div style="font-weight: bold; margin-bottom: 5px;">Cannot Access Folder</div>
                    <div style="font-size: 14px;">${escapeHtml(data.message)}</div>
                    <button onclick="loadDirectory('')" style="margin-top: 15px;" class="btn-secondary">
                        Go to Home Folder
                    </button>
                </div>
            `;
        }
    } catch (error) {
        console.error('Failed to load directory:', error);
        log('✗ Failed to load directory', 'error');
    }
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function selectFile(path) {
    const input = document.getElementById(browserTargetInput);
    input.value = path;
    closeFileBrowser();
    log(`✓ Selected: ${path}`, 'info');
}

function selectDirectory(path) {
    const input = document.getElementById(browserTargetInput);
    input.value = path;
    closeFileBrowser();
    log(`✓ Selected folder: ${path}`, 'info');
}

function closeFileBrowser() {
    document.getElementById('fileBrowserModal').style.display = 'none';
}

// Allow Enter key to login
document.addEventListener('keypress', (e) => {
    if (e.key === 'Enter' && document.getElementById('loginModal').style.display === 'flex') {
        login();
    }
});
