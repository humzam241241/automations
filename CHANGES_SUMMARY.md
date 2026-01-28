# Changes Summary - Modular Framework Upgrade

## ✅ All Requirements Implemented

### A) Correctness & Robustness Fixes

#### 1. Fixed `exp` → `explanation` bug
- **File**: `jobs/email_to_table.py`
- **Change**: Fixed return statement in `email_to_row()` to return `explanation` instead of undefined `exp`
- **Impact**: Explain mode now works correctly

#### 2. Headers from profile schema (not first email)
- **File**: `core/engine.py`
- **Change**: New method `_records_to_table()` extracts headers from `profile["schema"]["columns"]` to preserve order
- **Impact**: All emails now have consistent column order matching schema

#### 3. First-match-wins with priority
- **File**: `jobs/email_to_table.py`
- **Change**: 
  - Rules are sorted by `priority` (higher wins), then by definition order
  - Once a column has a value, subsequent rules skip it
  - Add optional `priority` field to rules (default: 0)
- **Impact**: No more rule overwrites; predictable rule application

---

### B) Rule Engine Upgrades

#### 4. Word-boundary matching
- **File**: `jobs/email_to_table.py`
- **Change**: 
  - Default: `word_boundary: true` uses regex `\b` word boundaries
  - Optional: `word_boundary: false` for substring matching
- **Example**:
  ```json
  {
    "column": "IT",
    "keywords": ["vpn"],
    "word_boundary": true
  }
  ```
  - Matches: "Need VPN access"
  - Skips: "VPNACCESS123" (not a word boundary)

#### 5. `search_in` field
- **File**: `jobs/email_to_table.py`
- **Change**: Rules can specify where to search:
  - `["all"]` (default): subject + body + attachments + from + to
  - `["subject"]`: subject only
  - `["body"]`: body only
  - `["attachments_text"]`: attachment content
  - `["attachments_names"]`: attachment filenames
  - `["from"]`: sender email
  - `["to"]`: recipient emails
- **Example**:
  ```json
  {
    "column": "Urgent",
    "keywords": ["urgent"],
    "search_in": ["subject"],
    "priority": 100
  }
  ```

#### 6. Enhanced explain mode
- **File**: `jobs/email_to_table.py`
- **Change**: Explanation dict now includes:
  - `matched`: bool
  - `match_type`: "keyword_word" | "keyword_substring" | "regex"
  - `matched_term`: The actual keyword/regex that matched
  - `search_in`: Where it searched
  - `snippet`: 60-char context snippet around match
- **Example output**:
  ```json
  {
    "Billing": {
      "matched": true,
      "match_type": "keyword_word",
      "matched_term": "invoice",
      "search_in": ["subject", "body"],
      "snippet": "...Please review the invoice for January..."
    }
  }
  ```

---

### C) Pipeline Support

#### 7. `profile["pipeline"]` field
- **File**: `core/engine.py`
- **Change**: 
  - Profiles can define `"pipeline": ["email_to_table", "append_to_master", "excel_to_biready"]`
  - Jobs execute in sequence
  - Default: `["email_to_table"]` for backward compatibility
- **Supported jobs**:
  - `email_to_table`: Extract emails → records
  - `append_to_master`: Merge with master dataset
  - `excel_to_biready`: BI transformations

#### 8. Dict-based record representation
- **File**: `jobs/email_to_table.py`, `core/engine.py`
- **Change**:
  - New function: `email_to_record()` returns `(record_dict, explanation)`
  - Records are `{column_name: value}` dicts throughout pipeline
  - Conversion to headers+rows only at output time
  - Old `email_to_row()` kept for backward compatibility
- **Impact**: No misalignment between headers and values

---

### D) Wizard Enhancements

#### 9. Template header loading + search query builder
- **File**: `run_wizard.py`
- **Changes**:
  
  **A) Schema from template .xlsx:**
  ```
  Define schema:
    A. Type column headers manually
    B. Load headers from template .xlsx
  Choice [A/B]: B
  Template .xlsx path: ./templates/invoice_template.xlsx
  Loaded 8 columns: Subject, From, Date, Amount, ...
  ```
  
  **B) Search query mode:**
  ```
  Email selection mode:
    A. Folder + newest N messages
    B. Search query (from/subject/date/attachments)
  Choice [A/B]: B
  
  Build search query:
    From contains: billing@company.com
    Subject contains: invoice
    After date (YYYY-MM-DD): 2026-01-01
    Has attachments? (y/n): y
  ```
  
  Generates:
  ```json
  {
    "search_query": "from:billing@company.com AND subject:invoice AND hasAttachments:true",
    "after_date": "2026-01-01"
  }
  ```

#### 10. Removed hardcoded credentials
- **File**: `run_wizard.py`
- **Changes**:
  - Removed hardcoded `client_id` and `authority`
  - If `CLIENT_ID` or `AUTHORITY` env vars missing:
    - Skip Graph auth
    - Print clear message: "Graph mode is disabled"
    - Force local-only inputs/outputs
  - Users **must** set environment variables to use Graph

#### 11. Token cache in user directory
- **File**: `run_wizard.py`
- **Change**: 
  - Windows: `%LOCALAPPDATA%\EmailAutomations\token_cache.bin`
  - Unix: `~/.local/share/EmailAutomations/token_cache.bin`
  - No longer pollutes repo root
- **Function**: `get_token_cache_path()`

---

### E) Capability Checks

#### 12. Separate mail/drive capability tests
- **File**: `run_wizard.py`
- **Function**: `test_graph_capabilities(token)`
- **Tests**:
  - `/me` - Basic Graph access
  - `/me/mailFolders` - Mail.Read permission
  - `/me/drive/root` - Files.ReadWrite permission
- **Output**:
  ```
  ✓ Microsoft Graph connected
    ✓ Mail access: OK
    ✗ OneDrive access: DENIED (cannot upload to OneDrive)
  ```
- **Profile validation**:
  - If profile needs mail but `capabilities["mail"]` is False → block execution
  - If profile needs drive but `capabilities["drive"]` is False → block execution
  - Clear error messages with remediation steps

---

## Updated Files

### Modified Files (3 core + 2 supporting):
1. ✅ **`jobs/email_to_table.py`** - Enhanced rule engine, word boundaries, search_in, explain mode
2. ✅ **`core/engine.py`** - Pipeline support, dict-based records, master dataset
3. ✅ **`run_wizard.py`** - Template headers, search queries, capability checks, token cache
4. ✅ **`run_app.bat`** - Updated to launch new wizard with env var checks
5. ✅ **`profiles/example_graph.json`** - Updated with new rule features

### New Files:
6. ✅ **`profiles/example_advanced.json`** - Shows pipeline, search query, master dataset

---

## New Rule Features (Examples)

### Word-Boundary Matching
```json
{
  "column": "VPN_Issue",
  "keywords": ["vpn"],
  "word_boundary": true,
  "value": "Yes"
}
```
- ✓ Matches: "Need VPN help"
- ✗ Skips: "VPNACCESS"

### Search-In Targeting
```json
{
  "column": "Urgent_Subject",
  "keywords": ["urgent", "asap"],
  "search_in": ["subject"],
  "priority": 100
}
```
- Only searches in subject line
- High priority ensures it wins over other rules

### Priority System
```json
[
  {
    "column": "Category",
    "keywords": ["billing"],
    "priority": 10
  },
  {
    "column": "Category",
    "keywords": ["invoice"],
    "priority": 5
  }
]
```
- "billing" rule wins if both match (higher priority)
- First match at same priority wins

---

## Pipeline Examples

### Simple Pipeline (default)
```json
{
  "pipeline": ["email_to_table"]
}
```

### Master Dataset Pipeline
```json
{
  "pipeline": ["email_to_table", "append_to_master"],
  "master_dataset": {
    "path": "./output/master.xlsx",
    "deduplicate_column": "Subject"
  }
}
```

### Full BI Pipeline
```json
{
  "pipeline": ["email_to_table", "excel_to_biready", "append_to_master"],
  "bi_transformations": [
    {"type": "clean", "column": "Amount", "rules": {"trim": true}},
    {"type": "filter", "column": "From", "condition": "not_empty"}
  ]
}
```

---

## How to Use with run_app.bat

### 1. Set Environment Variables (if using Graph)
```cmd
set CLIENT_ID=your-client-id
set AUTHORITY=https://login.microsoftonline.com/your-tenant-id
```

### 2. Run the batch file
```cmd
run_app.bat
```

### 3. Wizard Flow
1. Checks dependencies
2. Checks Graph configuration
3. Launches wizard
4. Detects Graph capabilities
5. Guides profile creation/execution

### 4. Without Graph (Local Mode)
- Don't set CLIENT_ID/AUTHORITY
- Wizard detects and switches to local-only mode
- Use .eml files or CSV input
- Local file output only

---

## Backward Compatibility

### ✅ Old code still works:
- `email_to_row()` function maintained
- Profiles without `pipeline` default to `["email_to_table"]`
- Profiles without `word_boundary` default to `true`
- Profiles without `search_in` default to `["all"]`
- Old profiles run unchanged

### ⚠ Breaking changes:
- **Token cache location moved** (from `./token_cache.bin` to user directory)
  - Old cache will not be used
  - User must re-authenticate
- **Hardcoded credentials removed**
  - Must set `CLIENT_ID` and `AUTHORITY` env vars
  - Or use local-only mode

---

## Testing Checklist

### ✅ Basic Functionality
- [x] Word-boundary matching works
- [x] Substring fallback works (`word_boundary: false`)
- [x] `search_in` filters correctly
- [x] Priority system works (first-match-wins)
- [x] Explain mode returns detailed info with snippets

### ✅ Wizard
- [x] Detects missing CLIENT_ID/AUTHORITY
- [x] Tests mail/drive capabilities separately
- [x] Loads headers from template .xlsx
- [x] Builds search queries
- [x] Blocks execution when permissions missing

### ✅ Pipeline
- [x] email_to_table job works
- [x] append_to_master job integrates
- [x] Dict-based records prevent misalignment
- [x] Headers preserve schema order

### ✅ Backward Compatibility
- [x] Old profiles still execute
- [x] `email_to_row()` still works
- [x] Default behaviors preserved

---

## Next Steps

1. **Test with real data**:
   ```cmd
   run_app.bat
   ```

2. **Try the advanced profile**:
   ```cmd
   python run_wizard.py
   # Choose option 1: Run existing profile
   # Enter: example_advanced
   ```

3. **Create your own profiles** with new features:
   - Use `search_in` to target specific fields
   - Set priorities for important rules
   - Try word-boundary vs substring matching
   - Build pipelines for complex workflows

4. **Push to GitHub**:
   ```cmd
   git add .
   git commit -m "feat: Modular framework with enhanced rules and pipeline support"
   git push origin main
   ```

---

## Summary

**All 12 requirements implemented successfully!**

- ✅ Bug fixes (exp→explanation, headers from schema, first-match-wins)
- ✅ Enhanced rule engine (word boundaries, search_in, explain mode)
- ✅ Pipeline support (job chaining, dict records)
- ✅ Wizard improvements (template headers, search queries, no hardcoded creds)
- ✅ Capability checks (separate mail/drive tests, clear errors)
- ✅ Works with run_app.bat

The framework is now truly modular, department-reusable, and works offline/online with no-Graph fallback.
