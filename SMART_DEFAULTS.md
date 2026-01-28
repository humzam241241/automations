# 🎯 Smart Input/Output Defaults

## Overview

Email Automation Pro v2.0 introduces intelligent defaults that automatically choose the best output format based on your input type.

## Default Behavior

### 📧 Email Inputs → Excel Output

**When you select:**
- Microsoft Graph (Outlook)
- Local .eml files

**You automatically get:**
- Excel spreadsheet (`.xlsx`)
- Structured table with your columns
- Ready to open in Excel

**Why?** Emails need to be structured into tables first - Excel is the perfect format for this!

### 📊 Excel/CSV Inputs → BI Dashboard

**When you select:**
- Excel file (`.xlsx`)
- CSV file (`.csv`)

**You automatically get:**
- Interactive HTML dashboard
- Auto-generated charts
- Opens in your browser

**Why?** Your data is already structured! Skip Excel and go straight to beautiful visualizations.

## Manual Overrides

### Want Excel + BI Dashboard?

Choose **"Excel + BI Dashboard"** in the output options:

```json
{
  "output": {
    "format": "excel",
    "destination": "local",
    "also_export_bi": true
  }
}
```

### Want to Process Emails → Excel → BI?

Use the pipeline option:

1. Select Email input
2. Choose "Excel + BI Dashboard" output
3. Run profile

**You get:**
1. `output/emails_20240115.xlsx` (for archiving)
2. `output/dashboards/dashboard_20240115.html` (for analysis)

## Examples

### Example 1: Daily Email Report
```
Input: 50 new emails from "Daily Reports" folder
Auto Output: Excel spreadsheet
Time Saved: No need to choose format!
```

### Example 2: Analyze Exported Data
```
Input: email_archive.xlsx (2000 rows)
Auto Output: BI Dashboard with charts
Time Saved: Skip manual chart creation!
```

### Example 3: Complete Workflow
```
Input: Inbox emails
Manual Selection: "Excel + BI Dashboard"
Output: Both files
Use Case: Archive in Excel, analyze in browser
```

## Smart Matching Integration

When you search for keywords, the system automatically finds:

### Dates
Search: `date`

Finds:
- The word "date"
- `2024-01-15`
- `January 15, 2024`
- `15/01/2024`
- `Jan 15, 2024`
- All date formats!

### Numbers/Amounts
Search: `amount`, `price`, `cost`

Finds:
- The exact words
- `$100.00`
- `€50,00`
- `£25.99`
- `1,234.56`
- All currency formats!

### Email Addresses
Search: `email`, `contact`

Finds:
- The exact words
- `user@domain.com`
- Any email pattern

### URLs
Search: `link`, `website`

Finds:
- The exact words
- `https://example.com`
- `www.example.com`
- All URL patterns

## Configuration

### Profile Auto-Detection

The system detects input type from your profile:

```json
{
  "input_source": "graph"
  // Auto sets output.format = "excel"
}
```

```json
{
  "input_source": "excel_file"
  // Auto sets output.destination = "bi_dashboard"
}
```

### Override in GUI

1. Create new profile
2. Select input source
3. **Output is pre-selected automatically!**
4. Want different? Just change it!

### Override in JSON

```json
{
  "input_source": "excel_file",
  "output": {
    "format": "excel",
    "destination": "local"
    // Overrides the auto BI dashboard
  }
}
```

## Benefits

✅ **Faster workflow** - No decisions needed for common cases
✅ **Smart defaults** - Based on data science best practices
✅ **Still flexible** - Override anytime
✅ **Beginner friendly** - New users get good results immediately
✅ **Power user approved** - Advanced options still available

## Technical Details

### Engine Logic (core/engine.py)

```python
# Auto BI export logic
output_config = profile.get("output", {})
should_export_bi = (
    output_config.get("destination") == "bi_dashboard" or
    output_config.get("also_export_bi", False) or
    profile.get("input_source") in ["excel_file", "csv_file"]
)
```

### GUI Logic (run_gui_v2.py)

```python
def on_input_change(self):
    source = self.input_var.get()
    
    if source in ["excel_file", "csv_file"]:
        # Auto select BI dashboard
        self.output_var.set("bi_dashboard")
    else:
        # Auto select Excel
        self.output_var.set("excel")
```

## FAQ

**Q: Can I still save Excel when using Excel input?**
A: Yes! Choose "Excel + BI Dashboard" output option.

**Q: What if I only want Excel for Excel input?**
A: Just change the output dropdown to "Excel File".

**Q: Does this change existing profiles?**
A: No! Existing profiles work exactly as before.

**Q: Can I disable auto-defaults?**
A: Yes, just manually select your preferred output each time.

**Q: What about OneDrive uploads?**
A: Works with all output types! Just choose OneDrive as destination.

## Best Practices

### For Daily Monitoring
```
Email Input → Excel → Save to archive folder
Check Excel later or use BI export on demand
```

### For Analysis
```
CSV/Excel Input → BI Dashboard → Share HTML link
No Excel installation needed to view!
```

### For End-of-Month Reports
```
Email Input → Excel + BI Dashboard → Both formats
Excel for IT/finance records
Dashboard for presentations
```

## Future Enhancements

🔮 **Coming soon:**
- PDF output for reports
- Power BI integration
- Slack/Teams notifications with dashboard links
- Scheduled auto-export

---

**Happy Automating!** 🚀
