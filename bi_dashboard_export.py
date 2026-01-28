"""
BI Dashboard Export - Creates interactive HTML dashboards from email data.
"""
import json
from pathlib import Path
from typing import List, Dict
from datetime import datetime
import base64
from io import BytesIO


def generate_dashboard(headers: List[str], rows: List[List[str]], profile_name: str) -> str:
    """
    Generate an interactive HTML dashboard.
    
    Args:
        headers: Column headers
        rows: Data rows
        profile_name: Name of the profile
    
    Returns:
        HTML string for dashboard
    """
    # Analyze data
    stats = analyze_data(headers, rows)
    
    # Generate visualizations
    charts = generate_charts(headers, rows, stats)
    
    # Build HTML
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{profile_name} - Email Analytics Dashboard</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }}
        
        .dashboard {{
            max-width: 1400px;
            margin: 0 auto;
        }}
        
        .header {{
            background: white;
            border-radius: 16px;
            padding: 30px;
            margin-bottom: 20px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.1);
        }}
        
        .header h1 {{
            font-size: 32px;
            color: #1a202c;
            margin-bottom: 10px;
        }}
        
        .header .subtitle {{
            color: #718096;
            font-size: 14px;
        }}
        
        .stats-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 20px;
        }}
        
        .stat-card {{
            background: white;
            border-radius: 16px;
            padding: 25px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.1);
            transition: transform 0.3s ease, box-shadow 0.3s ease;
        }}
        
        .stat-card:hover {{
            transform: translateY(-5px);
            box-shadow: 0 15px 40px rgba(0,0,0,0.15);
        }}
        
        .stat-card .label {{
            font-size: 12px;
            color: #718096;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 10px;
        }}
        
        .stat-card .value {{
            font-size: 36px;
            font-weight: bold;
            color: #1a202c;
        }}
        
        .stat-card .change {{
            font-size: 14px;
            color: #48bb78;
            margin-top: 5px;
        }}
        
        .charts-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(500px, 1fr));
            gap: 20px;
            margin-bottom: 20px;
        }}
        
        .chart-card {{
            background: white;
            border-radius: 16px;
            padding: 25px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.1);
        }}
        
        .chart-card h3 {{
            font-size: 18px;
            color: #1a202c;
            margin-bottom: 20px;
        }}
        
        .table-card {{
            background: white;
            border-radius: 16px;
            padding: 25px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.1);
            overflow-x: auto;
        }}
        
        table {{
            width: 100%;
            border-collapse: collapse;
        }}
        
        th {{
            background: #f7fafc;
            padding: 12px;
            text-align: left;
            font-weight: 600;
            color: #2d3748;
            border-bottom: 2px solid #e2e8f0;
        }}
        
        td {{
            padding: 12px;
            border-bottom: 1px solid #e2e8f0;
            color: #4a5568;
        }}
        
        tr:hover {{
            background: #f7fafc;
        }}
        
        .badge {{
            display: inline-block;
            padding: 4px 12px;
            border-radius: 12px;
            font-size: 12px;
            font-weight: 600;
        }}
        
        .badge-success {{
            background: #c6f6d5;
            color: #22543d;
        }}
        
        .badge-warning {{
            background: #fed7d7;
            color: #742a2a;
        }}
        
        @media (max-width: 768px) {{
            .charts-grid {{
                grid-template-columns: 1fr;
            }}
        }}
    </style>
</head>
<body>
    <div class="dashboard">
        <div class="header">
            <h1>📊 {profile_name}</h1>
            <div class="subtitle">Email Analytics Dashboard • Generated {datetime.now().strftime('%B %d, %Y at %I:%M %p')}</div>
        </div>
        
        <div class="stats-grid">
            <div class="stat-card">
                <div class="label">Total Emails</div>
                <div class="value">{stats['total_emails']}</div>
            </div>
            <div class="stat-card">
                <div class="label">Unique Senders</div>
                <div class="value">{stats['unique_senders']}</div>
            </div>
            <div class="stat-card">
                <div class="label">Date Range</div>
                <div class="value" style="font-size: 16px;">{stats['date_range']}</div>
            </div>
            <div class="stat-card">
                <div class="label">Columns</div>
                <div class="value">{len(headers)}</div>
            </div>
        </div>
        
        <div class="charts-grid">
            {charts}
        </div>
        
        <div class="table-card">
            <h3>📋 Data Table (First 100 rows)</h3>
            <table>
                <thead>
                    <tr>
                        {''.join(f'<th>{h}</th>' for h in headers)}
                    </tr>
                </thead>
                <tbody>
                    {''.join(generate_table_row(row, headers) for row in rows[:100])}
                </tbody>
            </table>
        </div>
    </div>
</body>
</html>"""
    
    return html


def analyze_data(headers: List[str], rows: List[List[str]]) -> Dict:
    """Analyze data and generate statistics."""
    stats = {
        'total_emails': len(rows),
        'unique_senders': 0,
        'date_range': 'N/A',
    }
    
    # Find sender column
    sender_col_idx = None
    for i, header in enumerate(headers):
        if header.lower() in ['from', 'sender', 'from address']:
            sender_col_idx = i
            break
    
    if sender_col_idx is not None:
        senders = set()
        for row in rows:
            if sender_col_idx < len(row) and row[sender_col_idx]:
                senders.add(row[sender_col_idx])
        stats['unique_senders'] = len(senders)
    
    # Find date range
    date_col_idx = None
    for i, header in enumerate(headers):
        if header.lower() in ['date', 'datetime', 'timestamp', 'received']:
            date_col_idx = i
            break
    
    if date_col_idx is not None:
        dates = []
        for row in rows:
            if date_col_idx < len(row) and row[date_col_idx]:
                dates.append(row[date_col_idx])
        
        if dates:
            dates_sorted = sorted(dates)
            stats['date_range'] = f"{dates_sorted[0][:10]} to {dates_sorted[-1][:10]}"
    
    return stats


def generate_charts(headers: List[str], rows: List[List[str]], stats: Dict) -> str:
    """Generate chart HTML."""
    charts_html = []
    
    # Category distribution chart (for columns with "Yes/No" or similar values)
    for col_idx, header in enumerate(headers):
        if header.lower() not in ['subject', 'body', 'from', 'to', 'date']:
            value_counts = {}
            for row in rows:
                if col_idx < len(row):
                    val = row[col_idx] if row[col_idx] else 'Empty'
                    value_counts[val] = value_counts.get(val, 0) + 1
            
            if len(value_counts) <= 10 and len(value_counts) > 0:  # Only show if manageable number of categories
                chart_id = f"chart_{col_idx}"
                labels = list(value_counts.keys())
                values = list(value_counts.values())
                
                chart_html = f"""
                <div class="chart-card">
                    <h3>{header} Distribution</h3>
                    <canvas id="{chart_id}"></canvas>
                    <script>
                        new Chart(document.getElementById('{chart_id}'), {{
                            type: 'doughnut',
                            data: {{
                                labels: {json.dumps(labels)},
                                datasets: [{{
                                    data: {json.dumps(values)},
                                    backgroundColor: [
                                        '#667eea', '#764ba2', '#f093fb', '#4facfe',
                                        '#43e97b', '#fa709a', '#fee140', '#30cfd0'
                                    ]
                                }}]
                            }},
                            options: {{
                                responsive: true,
                                plugins: {{
                                    legend: {{ position: 'bottom' }}
                                }}
                            }}
                        }});
                    </script>
                </div>
                """
                charts_html.append(chart_html)
    
    # Sender frequency chart
    sender_col_idx = None
    for i, header in enumerate(headers):
        if header.lower() in ['from', 'sender']:
            sender_col_idx = i
            break
    
    if sender_col_idx is not None:
        sender_counts = {}
        for row in rows:
            if sender_col_idx < len(row) and row[sender_col_idx]:
                sender = row[sender_col_idx]
                sender_counts[sender] = sender_counts.get(sender, 0) + 1
        
        # Top 10 senders
        top_senders = sorted(sender_counts.items(), key=lambda x: x[1], reverse=True)[:10]
        if top_senders:
            labels = [s[0][:30] for s in top_senders]
            values = [s[1] for s in top_senders]
            
            chart_html = f"""
            <div class="chart-card">
                <h3>📬 Top Senders</h3>
                <canvas id="sender_chart"></canvas>
                <script>
                    new Chart(document.getElementById('sender_chart'), {{
                        type: 'bar',
                        data: {{
                            labels: {json.dumps(labels)},
                            datasets: [{{
                                label: 'Emails Sent',
                                data: {json.dumps(values)},
                                backgroundColor: '#667eea'
                            }}]
                        }},
                        options: {{
                            responsive: true,
                            indexAxis: 'y',
                            plugins: {{
                                legend: {{ display: false }}
                            }}
                        }}
                    }});
                </script>
            </div>
            """
            charts_html.append(chart_html)
    
    return '\n'.join(charts_html) if charts_html else '<div class="chart-card"><p>No visualizations available for this dataset.</p></div>'


def generate_table_row(row: List[str], headers: List[str]) -> str:
    """Generate HTML for table row."""
    cells = []
    for i, cell in enumerate(row):
        # Add badges for Yes/No values
        if str(cell).lower() in ['yes', 'true', '1']:
            cells.append(f'<td><span class="badge badge-success">Yes</span></td>')
        elif str(cell).lower() in ['no', 'false', '0']:
            cells.append(f'<td><span class="badge badge-warning">No</span></td>')
        else:
            # Truncate long text
            display_text = str(cell)[:100] + '...' if len(str(cell)) > 100 else str(cell)
            cells.append(f'<td>{display_text}</td>')
    
    return f"<tr>{''.join(cells)}</tr>"


def save_dashboard(html: str, output_path: str) -> str:
    """
    Save dashboard to file.
    
    Args:
        html: HTML content
        output_path: Output directory or file path
    
    Returns:
        Path to saved file
    """
    path = Path(output_path)
    
    if path.suffix == '.html':
        file_path = path
    else:
        # It's a directory
        path.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_path = path / f"dashboard_{timestamp}.html"
    
    file_path.write_text(html, encoding='utf-8')
    return str(file_path.absolute())
