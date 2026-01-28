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

function editProfile() {
    if (!currentProfile) return;
    
    // Set modal to edit mode
    document.getElementById('profileModalTitle').textContent = '✏️ Edit Profile';
    document.getElementById('saveProfileBtn').textContent = 'Save Changes';
    document.getElementById('editingProfileName').value = currentProfile.name;
    
    // Populate form with current profile data
    document.getElementById('profileName').value = currentProfile.name;
    document.getElementById('inputSource').value = currentProfile.input_source || 'graph';
    
    // Trigger input options update
    updateInputOptions();
    
    // Fill in input-specific fields after a short delay
    setTimeout(() => {
        if (currentProfile.email_selection) {
            if (currentProfile.input_source === 'graph') {
                const folderInput = document.getElementById('folderName');
                if (folderInput) folderInput.value = currentProfile.email_selection.folder_name || 'Inbox';
            } else if (currentProfile.input_source === 'local_eml') {
                const dirInput = document.getElementById('emlDirectory');
                if (dirInput) dirInput.value = currentProfile.email_selection.directory || '';
            } else if (currentProfile.input_source === 'excel_file' || currentProfile.input_source === 'csv_file') {
                const fileInput = document.getElementById('filePath');
                if (fileInput) fileInput.value = currentProfile.email_selection.file_path || '';
            }
        }
    }, 100);
    
    // Load columns with types
    profileColumns = [];
    if (currentProfile.schema && currentProfile.schema.columns) {
        currentProfile.schema.columns.forEach(col => {
            profileColumns.push({
                name: col.name || col.header || '',
                type: col.extract_type || col.type || 'auto'
            });
        });
    }
    renderColumns();
    updateHiddenInputs();
    
    // Set auto-detect
    document.getElementById('autoDetect').checked = currentProfile.auto_detect_columns || false;
    toggleAutoDetect();
    
    // Set output type
    if (currentProfile.output) {
        if (currentProfile.output.also_export_bi) {
            document.getElementById('outputType').value = 'both';
        } else if (currentProfile.output.destination === 'bi_dashboard') {
            document.getElementById('outputType').value = 'bi_dashboard';
        } else {
            document.getElementById('outputType').value = 'excel';
        }
    }
    
    // Show modal
    document.getElementById('createProfileModal').style.display = 'flex';
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
    // Reset to create mode
    document.getElementById('profileModalTitle').textContent = '✨ Create New Profile';
    document.getElementById('saveProfileBtn').textContent = 'Create Profile';
    document.getElementById('editingProfileName').value = '';
    
    // Reset form
    document.getElementById('profileName').value = '';
    document.getElementById('inputSource').value = 'graph';
    document.getElementById('autoDetect').checked = false;
    document.getElementById('outputType').value = 'excel';
    
    // Clear and reset the column builder
    clearColumns();
    updateInputOptions();
    
    document.getElementById('createProfileModal').style.display = 'flex';
}

function closeCreateProfile() {
    document.getElementById('createProfileModal').style.display = 'none';
    // Reset form
    document.getElementById('profileName').value = '';
    document.getElementById('autoDetect').checked = false;
    document.getElementById('editingProfileName').value = '';
    clearColumns();
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
            <label>Email Files <small style="color:#666;">(Ctrl+Click for multiple)</small></label>
            <div class="file-input-group" style="margin-bottom: 8px;">
                <input type="text" id="emlFiles" placeholder="Select .eml files..." readonly style="cursor: pointer;" onclick="browseFile('emlFiles', 'local_eml')">
                <button onclick="browseFile('emlFiles', 'local_eml')" class="btn-browse">📁 Select Files</button>
            </div>
            <div style="text-align: center; color: #999; margin: 5px 0;">— OR —</div>
            <div class="file-input-group">
                <input type="text" id="emlDirectory" value="./input_emails" placeholder="Enter folder path">
                <button onclick="browseDirectory('emlDirectory')" class="btn-browse">📂 Select Folder</button>
            </div>
            <small class="file-hint">Supports: .eml, PDF attachments, Word docs inside emails</small>
        `;
    } else if (inputSource === 'excel_file' || inputSource === 'csv_file') {
        const fileType = inputSource === 'excel_file' ? '.xlsx, .csv, .pdf' : '.csv, .xlsx, .pdf';
        optionsDiv.innerHTML = `
            <label>Input Files <small style="color:#666;">(Ctrl+Click to select multiple)</small></label>
            <div class="file-input-group">
                <input type="text" id="filePath" placeholder="Select files (${fileType})" readonly style="cursor: pointer;" onclick="browseFile('filePath', '${inputSource}')">
                <button onclick="browseFile('filePath', '${inputSource}')" class="btn-browse">📁 Browse Files</button>
            </div>
            <small class="file-hint">Supports: Excel, CSV, PDF, Word documents</small>
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
        columnsInput.placeholder = 'Subject,From,Date,MI #,CAPA #,Priority';
        columnsInput.style.background = 'white';
    }
}

function toggleColumnHelp() {
    const helpBox = document.getElementById('columnHelp');
    if (helpBox.style.display === 'none') {
        helpBox.style.display = 'block';
    } else {
        helpBox.style.display = 'none';
    }
}

// ===== COLUMN BUILDER =====

let profileColumns = [];  // Array of {name: string, type: string}

function addColumn() {
    const nameInput = document.getElementById('newColumnName');
    const typeSelect = document.getElementById('newColumnType');
    
    const name = nameInput.value.trim();
    const type = typeSelect.value;
    
    if (!name) {
        nameInput.focus();
        return;
    }
    
    // Check for duplicates
    if (profileColumns.some(c => c.name.toLowerCase() === name.toLowerCase())) {
        alert(`Column "${name}" already exists`);
        return;
    }
    
    // Add to array
    profileColumns.push({ name, type });
    
    // Update UI
    renderColumns();
    updateHiddenInputs();
    
    // Clear input
    nameInput.value = '';
    nameInput.focus();
}

function removeColumn(index) {
    profileColumns.splice(index, 1);
    renderColumns();
    updateHiddenInputs();
}

function renderColumns() {
    const list = document.getElementById('columnList');
    
    if (profileColumns.length === 0) {
        list.innerHTML = '<span style="color: #999; font-size: 13px;">No columns added yet. Add columns below.</span>';
        return;
    }
    
    const typeLabels = {
        'auto': '🤖 Auto',
        'number': '🔢 Number',
        'text': '📝 Text',
        'date': '📅 Date',
        'time': '🕐 Time',
        'amount': '💰 Amount',
        'yesno': '✅ Yes/No',
        'email_field': '📧 Email'
    };
    
    list.innerHTML = profileColumns.map((col, index) => `
        <div class="column-tag" data-type="${col.type}">
            <span class="column-tag-name">${escapeHtml(col.name)}</span>
            <span class="column-tag-type">${typeLabels[col.type] || col.type}</span>
            <button class="column-tag-remove" onclick="removeColumn(${index})" title="Remove">&times;</button>
        </div>
    `).join('');
}

function updateHiddenInputs() {
    // Update hidden inputs for form submission
    const columnsInput = document.getElementById('columns');
    const typesInput = document.getElementById('columnTypes');
    
    columnsInput.value = profileColumns.map(c => c.name).join(',');
    typesInput.value = JSON.stringify(profileColumns);
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function clearColumns() {
    profileColumns = [];
    renderColumns();
    updateHiddenInputs();
}

function loadColumnsFromString(columnsStr, typesStr) {
    // Parse columns and types
    profileColumns = [];
    
    if (typesStr) {
        try {
            profileColumns = JSON.parse(typesStr);
        } catch (e) {
            // Fallback: parse comma-separated columns as auto type
            if (columnsStr) {
                profileColumns = columnsStr.split(',').map(name => ({
                    name: name.trim(),
                    type: 'auto'
                })).filter(c => c.name);
            }
        }
    } else if (columnsStr) {
        // Legacy: comma-separated columns
        profileColumns = columnsStr.split(',').map(name => ({
            name: name.trim(),
            type: 'auto'
        })).filter(c => c.name);
    }
    
    renderColumns();
    updateHiddenInputs();
}

// Handle Enter key in column name input
document.addEventListener('DOMContentLoaded', () => {
    setTimeout(() => {
        const input = document.getElementById('newColumnName');
        if (input) {
            input.addEventListener('keypress', (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    addColumn();
                }
            });
        }
        // Initialize empty column list
        renderColumns();
    }, 100);
});

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
    const outputType = document.getElementById('outputType').value;
    const autoDetect = document.getElementById('autoDetect').checked;
    const editingName = document.getElementById('editingProfileName').value;
    const isEditing = editingName !== '';
    
    if (!name) {
        alert('Profile name is required');
        return;
    }
    
    if (!autoDetect && profileColumns.length === 0) {
        alert('Please add at least one column or enable auto-detect');
        return;
    }
    
    // Build schema columns with types
    const schemaColumns = profileColumns.map(col => ({
        name: col.name,
        type: col.type || 'auto',
        extract_type: col.type || 'auto'  // Used by extraction engine
    }));
    
    // Build profile
    const profile = {
        name: name,
        input_source: inputSource,
        auto_detect_columns: autoDetect,
        email_selection: {},
        schema: {
            columns: autoDetect ? [] : schemaColumns
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
        const folderInput = document.getElementById('folderName');
        profile.email_selection = {
            folder_name: folderInput ? folderInput.value : 'Inbox',
            newest_n: 25
        };
    } else if (inputSource === 'local_eml') {
        const dirInput = document.getElementById('emlDirectory');
        const filesInput = document.getElementById('emlFiles');
        
        // Check if specific files were selected
        let filePaths = [];
        if (filesInput && filesInput.dataset.paths) {
            try {
                filePaths = JSON.parse(filesInput.dataset.paths);
            } catch (e) {
                if (filesInput.value) {
                    filePaths = filesInput.value.split(';').map(p => p.trim()).filter(p => p);
                }
            }
        }
        
        profile.email_selection = {
            directory: dirInput ? dirInput.value : '',
            pattern: '*.eml',
            file_paths: filePaths  // NEW: specific files selected
        };
    } else if (inputSource === 'excel_file' || inputSource === 'csv_file') {
        const fileInput = document.getElementById('filePath');
        
        // Get multiple file paths
        let filePaths = [];
        if (fileInput && fileInput.dataset.paths) {
            try {
                filePaths = JSON.parse(fileInput.dataset.paths);
            } catch (e) {
                if (fileInput.value) {
                    filePaths = fileInput.value.split(';').map(p => p.trim()).filter(p => p);
                }
            }
        } else if (fileInput && fileInput.value) {
            filePaths = fileInput.value.split(';').map(p => p.trim()).filter(p => p);
        }
        
        profile.email_selection = {
            file_path: filePaths[0] || '',  // Backward compat
            file_paths: filePaths  // NEW: multiple files
        };
    }
    
    try {
        // If editing and name changed, delete old profile first
        if (isEditing && editingName !== name) {
            await fetch(`/api/profiles/${editingName}`, { method: 'DELETE' });
        }
        
        const response = await fetch('/api/profiles', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(profile)
        });
        
        const data = await response.json();
        
        if (data.success) {
            log(`✓ ${isEditing ? 'Updated' : 'Created'}: ${name}`, 'success');
            closeCreateProfile();
            await loadProfiles();
            
            // If we were editing, update current profile
            if (isEditing) {
                currentProfile = profile;
            }
        } else {
            alert(`Failed to ${isEditing ? 'update' : 'create'} profile: ${data.message}`);
        }
    } catch (error) {
        console.error('Save failed:', error);
        alert(`Failed to ${isEditing ? 'update' : 'create'} profile`);
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
    
    // Get current path from input (if any)
    const input = document.getElementById(targetInputId);
    const currentPath = input.value.trim();
    
    // Open native folder picker dialog - ACTUAL Windows Explorer!
    try {
        const response = await fetch('/api/open-file-dialog', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({
                type: 'directory',
                initial_dir: currentPath || ''
            })
        });
        
        const data = await response.json();
        
        if (data.success && data.path) {
            input.value = data.path;
            log(`✓ Selected folder: ${data.path}`, 'success');
        } else if (data.message !== 'No file selected') {
            log(`⚠ ${data.message}`, 'warning');
        }
    } catch (error) {
        console.error('Failed to open folder dialog:', error);
        log('✗ Failed to open folder picker', 'error');
    }
}

async function browseFile(targetInputId, fileType, allowMultiple = true) {
    browserTargetInput = targetInputId;
    
    // Get current path from input (if any) to start in the right folder
    const input = document.getElementById(targetInputId);
    const currentPath = input.value.trim();
    let initialDir = '';
    
    if (currentPath) {
        // Extract directory from first file path
        const firstPath = currentPath.split(';')[0].trim();
        const lastSlash = Math.max(firstPath.lastIndexOf('\\'), firstPath.lastIndexOf('/'));
        if (lastSlash > 0) {
            initialDir = firstPath.substring(0, lastSlash);
        }
    }
    
    // Define file types for the dialog
    let fileTypes = [];
    
    if (fileType === 'excel_file') {
        fileTypes = [
            {name: 'Excel & CSV Files', pattern: '*.xlsx *.xls *.csv'},
            {name: 'PDF Files', pattern: '*.pdf'},
            {name: 'All Files', pattern: '*.*'}
        ];
    } else if (fileType === 'csv_file') {
        fileTypes = [
            {name: 'CSV & Excel Files', pattern: '*.csv *.xlsx *.xls'},
            {name: 'Text Files', pattern: '*.txt'},
            {name: 'PDF Files', pattern: '*.pdf'},
            {name: 'All Files', pattern: '*.*'}
        ];
    } else if (fileType === 'local_eml') {
        fileTypes = [
            {name: 'Email Files', pattern: '*.eml *.msg'},
            {name: 'PDF Files', pattern: '*.pdf'},
            {name: 'Word Documents', pattern: '*.docx *.doc'},
            {name: 'All Files', pattern: '*.*'}
        ];
    } else {
        // Show all important file types
        fileTypes = [
            {name: 'All Supported', pattern: '*.eml *.msg *.xlsx *.xls *.csv *.pdf *.docx'},
            {name: 'Email Files', pattern: '*.eml *.msg'},
            {name: 'Excel Files', pattern: '*.xlsx *.xls'},
            {name: 'CSV Files', pattern: '*.csv'},
            {name: 'PDF Files', pattern: '*.pdf'},
            {name: 'All Files', pattern: '*.*'}
        ];
    }
    
    // Open native file picker dialog - supports MULTIPLE selection
    try {
        const response = await fetch('/api/open-file-dialog', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({
                type: 'file',
                file_types: fileTypes,
                initial_dir: initialDir,
                multiple: allowMultiple  // Enable multi-select
            })
        });
        
        const data = await response.json();
        
        if (data.success && (data.path || data.paths)) {
            // Store paths - show summary in input
            if (data.paths && data.paths.length > 1) {
                input.value = data.paths.join('; ');
                input.dataset.paths = JSON.stringify(data.paths);
                log(`Selected ${data.paths.length} files`, 'info');
            } else {
                input.value = data.path;
                input.dataset.paths = JSON.stringify([data.path]);
            }
            log(`✓ Selected: ${data.path}`, 'success');
            
            // Auto-detect columns if checkbox is checked
            const autoDetect = document.getElementById('autoDetect');
            if (autoDetect && autoDetect.checked) {
                await detectColumns();
            }
        } else if (data.message !== 'No file selected') {
            log(`⚠ ${data.message}`, 'warning');
        }
    } catch (error) {
        console.error('Failed to open file dialog:', error);
        log('✗ Failed to open file picker', 'error');
    }
}

// Fallback: Custom browser (for advanced users who want to see the file tree)
async function showCustomBrowser(targetInputId, fileType) {
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
            
            // Add info banner
            const infoBanner = document.createElement('div');
            infoBanner.style.cssText = 'padding: 12px; background: #fef3c7; border-bottom: 2px solid #f59e0b; color: #92400e; font-size: 13px;';
            
            if (browserFileType === 'directory') {
                infoBanner.innerHTML = '📁 <strong>Folder Selection Mode:</strong> Click folders to navigate, then select folder below';
                contentDiv.appendChild(infoBanner);
                
                // Add "Select This Folder" button
                const selectBtn = document.createElement('div');
                selectBtn.style.cssText = 'padding: 16px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; text-align: center; font-weight: 600; cursor: pointer; border-bottom: 2px solid #e2e8f0;';
                selectBtn.innerHTML = '✓ Select This Folder';
                selectBtn.onclick = () => selectDirectory(data.current_path);
                contentDiv.appendChild(selectBtn);
            } else {
                infoBanner.innerHTML = '📄 <strong>Selectable:</strong> .eml, .msg, .xlsx, .xls, .csv, .pdf, .txt, .json, .xml | <strong>Grayed out:</strong> Other file types';
                contentDiv.appendChild(infoBanner);
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
                    // Check if file type is selectable
                    const selectableExtensions = [
                        '.eml', '.msg',       // Email files
                        '.xlsx', '.xls',      // Excel files
                        '.csv',               // CSV files
                        '.pdf',               // PDF files
                        '.txt',               // Text files
                        '.json',              // JSON files
                        '.xml'                // XML files
                    ];
                    
                    const isSelectable = selectableExtensions.includes(item.extension);
                    
                    if (isSelectable) {
                        // File is selectable
                        itemDiv.onclick = () => selectFile(item.path);
                        itemDiv.style.cursor = 'pointer';
                        
                        // Highlight files that match expected type
                        let isPreferredType = false;
                        if (browserFileType === 'excel_file' && (item.extension === '.xlsx' || item.extension === '.xls')) {
                            isPreferredType = true;
                        } else if (browserFileType === 'csv_file' && item.extension === '.csv') {
                            isPreferredType = true;
                        } else if (item.extension === '.eml' || item.extension === '.msg') {
                            isPreferredType = true;
                        } else if (item.extension === '.pdf') {
                            isPreferredType = true;
                        }
                        
                        if (isPreferredType) {
                            // Strongly highlight preferred files
                            itemDiv.style.backgroundColor = '#e6f3ff';
                            itemDiv.style.border = '2px solid #667eea';
                            itemDiv.style.fontWeight = '600';
                        } else {
                            // Lightly highlight other selectable files
                            itemDiv.style.backgroundColor = '#f9fafb';
                        }
                    } else {
                        // File is NOT selectable - show it but grayed out
                        itemDiv.style.opacity = '0.4';
                        itemDiv.style.cursor = 'not-allowed';
                        itemDiv.title = 'This file type cannot be selected';
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
