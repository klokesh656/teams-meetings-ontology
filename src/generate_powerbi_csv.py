#!/usr/bin/env python3
"""
Generate CSV files for Power Automate and Power BI from batch analysis results.

Outputs:
1. va_risk_summary.csv - Summary view for Power BI dashboards
2. pending_suggestions.csv - All suggestions pending review (for Power Automate emails)
3. critical_alerts.csv - Critical/High risk cases for immediate stakeholder action
4. va_signals_detail.csv - Detailed signal data for deep analysis
5. stakeholder_review.csv - Formatted for email notifications via Power Automate
"""

import json
import csv
import os
from datetime import datetime
from pathlib import Path

# Paths
OUTPUT_DIR = Path(__file__).parent.parent / "output"
BATCH_RESULTS = OUTPUT_DIR / "batch_analysis_results_20260107_011511.json"
ANALYSIS_HISTORY = OUTPUT_DIR / "checkin_analysis_history.json"
PENDING_SUGGESTIONS = OUTPUT_DIR / "pending_suggestions.json"


def load_json(filepath):
    """Load JSON file safely."""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"Warning: Could not load {filepath}: {e}")
        return None


def generate_va_risk_summary():
    """Generate VA risk summary for Power BI dashboard."""
    print("\n📊 Generating VA Risk Summary...")
    
    data = load_json(BATCH_RESULTS)
    if not data:
        return
    
    # Aggregate by VA
    va_stats = {}
    for record in data.get('processed', []):
        va = record['va_name']
        if va not in va_stats:
            va_stats[va] = {
                'va_name': va,
                'total_checkins': 0,
                'critical_count': 0,
                'high_count': 0,
                'medium_count': 0,
                'low_count': 0,
                'total_signals': 0,
                'total_suggestions': 0,
                'latest_date': record['date'],
                'latest_risk': record['risk_level'],
                'clients': set()
            }
        
        stats = va_stats[va]
        stats['total_checkins'] += 1
        stats['total_signals'] += record['signals_count']
        stats['total_suggestions'] += record['suggestions_count']
        
        if record['client_name'] and record['client_name'] != 'Unknown':
            stats['clients'].add(record['client_name'][:50])  # Truncate long names
        
        # Count by risk level
        risk = record['risk_level'].lower()
        if risk == 'critical':
            stats['critical_count'] += 1
        elif risk == 'high':
            stats['high_count'] += 1
        elif risk == 'medium':
            stats['medium_count'] += 1
        else:
            stats['low_count'] += 1
        
        # Track latest
        if record['date'] >= stats['latest_date']:
            stats['latest_date'] = record['date']
            stats['latest_risk'] = record['risk_level']
    
    # Calculate risk scores and write CSV
    output_file = OUTPUT_DIR / "va_risk_summary.csv"
    with open(output_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow([
            'VA Name', 'Total Check-ins', 'Critical Count', 'High Count', 
            'Medium Count', 'Low Count', 'Total Signals', 'Total Suggestions',
            'Risk Score', 'Current Risk Level', 'Latest Check-in Date', 
            'Known Clients', 'Attention Required'
        ])
        
        for va, stats in sorted(va_stats.items(), key=lambda x: (
            x[1]['critical_count'] * 100 + x[1]['high_count'] * 10
        ), reverse=True):
            # Calculate risk score (weighted)
            risk_score = (
                stats['critical_count'] * 100 +
                stats['high_count'] * 25 +
                stats['medium_count'] * 5 +
                stats['low_count'] * 1
            )
            
            attention = 'CRITICAL' if stats['critical_count'] > 0 else \
                       'HIGH' if stats['high_count'] > 0 else \
                       'MONITOR' if stats['medium_count'] > 2 else 'OK'
            
            writer.writerow([
                stats['va_name'],
                stats['total_checkins'],
                stats['critical_count'],
                stats['high_count'],
                stats['medium_count'],
                stats['low_count'],
                stats['total_signals'],
                stats['total_suggestions'],
                risk_score,
                stats['latest_risk'].upper(),
                stats['latest_date'],
                '; '.join(list(stats['clients'])[:3]) if stats['clients'] else 'Unknown',
                attention
            ])
    
    print(f"   ✅ Saved: {output_file}")
    return len(va_stats)


def generate_pending_suggestions_csv():
    """Generate pending suggestions for Power Automate workflow."""
    print("\n📋 Generating Pending Suggestions CSV...")
    
    data = load_json(PENDING_SUGGESTIONS)
    if not data:
        # Try to extract from analysis history
        data = load_json(ANALYSIS_HISTORY)
        if not data:
            print("   ⚠️ No pending suggestions found")
            return 0
    
    output_file = OUTPUT_DIR / "pending_suggestions_review.csv"
    suggestion_count = 0
    
    with open(output_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow([
            'Suggestion ID', 'VA Name', 'Client', 'Meeting Date', 'Risk Level',
            'Issue', 'Category', 'Urgency', 'Suggestion', 'Rationale',
            'Status', 'Review Link'
        ])
        
        if isinstance(data, dict) and 'suggestions' in data:
            suggestions = data['suggestions']
        elif isinstance(data, list):
            suggestions = data
        else:
            suggestions = []
        
        for sugg in suggestions:
            if isinstance(sugg, dict):
                writer.writerow([
                    sugg.get('id', f'SUGG-{suggestion_count+1:04d}'),
                    sugg.get('va_name', 'Unknown'),
                    sugg.get('client_name', 'Unknown'),
                    sugg.get('meeting_date', ''),
                    sugg.get('risk_level', 'medium').upper(),
                    sugg.get('issue', '')[:200],
                    sugg.get('category', ''),
                    sugg.get('urgency', 'monitor'),
                    sugg.get('suggestion', '')[:500],
                    sugg.get('rationale', '')[:300],
                    sugg.get('status', 'pending'),
                    f"Review in Outlier Insights Engine"
                ])
                suggestion_count += 1
    
    print(f"   ✅ Saved: {output_file} ({suggestion_count} suggestions)")
    return suggestion_count


def generate_critical_alerts():
    """Generate critical alerts for immediate stakeholder attention."""
    print("\n🚨 Generating Critical Alerts CSV...")
    
    data = load_json(BATCH_RESULTS)
    if not data:
        return 0
    
    # Also load detailed analysis
    history = load_json(ANALYSIS_HISTORY) or {'analyses': []}
    
    # Create lookup for detailed info
    detail_lookup = {}
    for analysis in history.get('analyses', []):
        key = f"{analysis.get('va_name', '')}_{analysis.get('meeting_date', '')}"
        detail_lookup[key] = analysis
    
    output_file = OUTPUT_DIR / "critical_alerts.csv"
    alert_count = 0
    
    with open(output_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow([
            'Priority', 'VA Name', 'Client', 'Meeting Date', 'Risk Level',
            'Signals Count', 'Executive Summary', 'Key Findings',
            'Escalation Reason', 'Immediate Actions Required',
            'Stakeholder', 'Review Status'
        ])
        
        for record in data.get('processed', []):
            risk = record['risk_level'].lower()
            if risk in ['critical', 'high']:
                key = f"{record['va_name']}_{record['date']}"
                detail = detail_lookup.get(key, {})
                
                # Determine priority
                priority = 'P1 - CRITICAL' if risk == 'critical' else 'P2 - HIGH'
                
                # Get summary and findings
                summary = detail.get('executive_summary', 'Analysis pending review')[:500]
                findings = '; '.join(detail.get('key_findings', [])[:3])[:400]
                escalation = detail.get('escalation_reason', '')[:300]
                
                # Get immediate actions from suggestions
                actions = []
                for sugg in detail.get('ai_suggestions', [])[:2]:
                    if sugg.get('urgency') in ['immediate', 'within_48h']:
                        actions.append(sugg.get('suggestion', '')[:150])
                
                writer.writerow([
                    priority,
                    record['va_name'],
                    record['client_name'][:50] if record['client_name'] else 'Unknown',
                    record['date'],
                    risk.upper(),
                    record['signals_count'],
                    summary,
                    findings,
                    escalation,
                    ' | '.join(actions) if actions else 'Review detailed report',
                    'Pending Assignment',
                    'PENDING REVIEW'
                ])
                alert_count += 1
    
    print(f"   ✅ Saved: {output_file} ({alert_count} alerts)")
    return alert_count


def generate_signals_detail():
    """Generate detailed signals data for Power BI analysis."""
    print("\n📈 Generating Signals Detail CSV...")
    
    history = load_json(ANALYSIS_HISTORY)
    if not history:
        return 0
    
    output_file = OUTPUT_DIR / "va_signals_detail.csv"
    signal_count = 0
    
    with open(output_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow([
            'Analysis ID', 'VA Name', 'Client', 'Meeting Date', 'Overall Risk',
            'Signal ID', 'Signal Category', 'Evidence', 'Confidence',
            'VA Status', 'Client Health'
        ])
        
        for analysis in history.get('analyses', []):
            for signal in analysis.get('detected_signals', []):
                # Determine signal category from ID
                signal_id = signal.get('signal_id', '')
                if signal_id.startswith('VA'):
                    category = 'VA Issue'
                elif signal_id.startswith('CL'):
                    category = 'Client Issue'
                elif signal_id.startswith('RH'):
                    category = 'Relationship/HR Issue'
                else:
                    category = 'Other'
                
                writer.writerow([
                    analysis.get('analysis_id', ''),
                    analysis.get('va_name', 'Unknown'),
                    analysis.get('client_name', 'Unknown')[:50],
                    analysis.get('meeting_date', ''),
                    analysis.get('overall_risk_level', 'unknown').upper(),
                    signal_id,
                    category,
                    signal.get('evidence', '')[:300],
                    signal.get('confidence', 'medium'),
                    analysis.get('va_status', ''),
                    analysis.get('client_health', '')
                ])
                signal_count += 1
    
    print(f"   ✅ Saved: {output_file} ({signal_count} signals)")
    return signal_count


def generate_stakeholder_review():
    """Generate stakeholder review file for Power Automate email workflow."""
    print("\n📧 Generating Stakeholder Review CSV...")
    
    data = load_json(BATCH_RESULTS)
    history = load_json(ANALYSIS_HISTORY) or {'analyses': []}
    
    if not data:
        return 0
    
    # Create lookup
    detail_lookup = {}
    for analysis in history.get('analyses', []):
        key = f"{analysis.get('va_name', '')}_{analysis.get('meeting_date', '')}"
        detail_lookup[key] = analysis
    
    output_file = OUTPUT_DIR / "stakeholder_review.csv"
    
    with open(output_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow([
            'Review ID', 'Date Generated', 'VA Name', 'Client', 'Check-in Date',
            'Risk Level', 'Risk Score', 'Signals Detected', 'Suggestions Count',
            'Summary', 'Top Recommendation', 'Stakeholder Email', 'Review URL',
            'Approval Status', 'Approved By', 'Approval Date', 'Notes'
        ])
        
        review_count = 0
        for record in data.get('processed', []):
            key = f"{record['va_name']}_{record['date']}"
            detail = detail_lookup.get(key, {})
            
            # Calculate risk score
            risk_weights = {'critical': 100, 'high': 75, 'medium': 25, 'low': 5}
            risk_score = risk_weights.get(record['risk_level'].lower(), 10)
            
            # Get top recommendation
            top_rec = ''
            for sugg in detail.get('ai_suggestions', []):
                if sugg.get('urgency') == 'immediate':
                    top_rec = sugg.get('suggestion', '')[:200]
                    break
            if not top_rec and detail.get('ai_suggestions'):
                top_rec = detail.get('ai_suggestions', [{}])[0].get('suggestion', '')[:200]
            
            review_id = f"REV-{datetime.now().strftime('%Y%m%d')}-{review_count+1:04d}"
            
            writer.writerow([
                review_id,
                datetime.now().strftime('%Y-%m-%d %H:%M'),
                record['va_name'],
                record['client_name'][:50] if record['client_name'] else 'Unknown',
                record['date'],
                record['risk_level'].upper(),
                risk_score,
                record['signals_count'],
                record['suggestions_count'],
                detail.get('executive_summary', '')[:300],
                top_rec,
                '',  # Stakeholder email - to be filled in Power Automate
                '',  # Review URL - to be filled with SharePoint link
                'PENDING',
                '',  # Approved by
                '',  # Approval date
                ''   # Notes
            ])
            review_count += 1
    
    print(f"   ✅ Saved: {output_file} ({review_count} reviews)")
    return review_count


def generate_kpi_summary():
    """Generate KPI summary for Power BI executive dashboard."""
    print("\n📊 Generating KPI Summary CSV...")
    
    data = load_json(BATCH_RESULTS)
    history = load_json(ANALYSIS_HISTORY) or {'analyses': []}
    
    if not data:
        return
    
    # Calculate KPIs
    processed = data.get('processed', [])
    total_analyses = len(processed)
    
    risk_counts = {'critical': 0, 'high': 0, 'medium': 0, 'low': 0}
    total_signals = 0
    total_suggestions = 0
    unique_vas = set()
    unique_clients = set()
    escalations_needed = 0
    
    for record in processed:
        risk_counts[record['risk_level'].lower()] += 1
        total_signals += record['signals_count']
        total_suggestions += record['suggestions_count']
        unique_vas.add(record['va_name'])
        if record['client_name'] and record['client_name'] != 'Unknown':
            unique_clients.add(record['client_name'][:30])
    
    # Count escalations from history
    for analysis in history.get('analyses', []):
        if analysis.get('escalation_needed'):
            escalations_needed += 1
    
    output_file = OUTPUT_DIR / "kpi_dashboard_summary.csv"
    
    with open(output_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['KPI', 'Value', 'Category', 'Trend', 'Target', 'Status'])
        
        kpis = [
            ('Total Check-ins Analyzed', total_analyses, 'Volume', '↑', '-', 'INFO'),
            ('Unique VAs Monitored', len(unique_vas), 'Coverage', '→', '-', 'INFO'),
            ('Unique Clients', len(unique_clients), 'Coverage', '→', '-', 'INFO'),
            ('Critical Risk Cases', risk_counts['critical'], 'Risk', '↓', '0', 
             'CRITICAL' if risk_counts['critical'] > 0 else 'OK'),
            ('High Risk Cases', risk_counts['high'], 'Risk', '↓', '< 5', 
             'WARNING' if risk_counts['high'] > 5 else 'OK'),
            ('Medium Risk Cases', risk_counts['medium'], 'Risk', '→', '-', 'MONITOR'),
            ('Low Risk Cases', risk_counts['low'], 'Risk', '↑', '> 50%', 'INFO'),
            ('Total Signals Detected', total_signals, 'Analysis', '→', '-', 'INFO'),
            ('Total Suggestions Generated', total_suggestions, 'Action', '→', '-', 'INFO'),
            ('Escalations Required', escalations_needed, 'Urgency', '↓', '0', 
             'CRITICAL' if escalations_needed > 10 else 'WARNING' if escalations_needed > 5 else 'OK'),
            ('Critical Risk Rate (%)', round(risk_counts['critical']/total_analyses*100, 1) if total_analyses else 0, 
             'Risk', '↓', '< 5%', 'CRITICAL' if risk_counts['critical']/total_analyses*100 > 5 else 'OK'),
            ('High Risk Rate (%)', round(risk_counts['high']/total_analyses*100, 1) if total_analyses else 0, 
             'Risk', '↓', '< 15%', 'WARNING' if risk_counts['high']/total_analyses*100 > 15 else 'OK'),
            ('Avg Signals per Check-in', round(total_signals/total_analyses, 1) if total_analyses else 0, 
             'Quality', '→', '-', 'INFO'),
            ('Report Generated', datetime.now().strftime('%Y-%m-%d %H:%M'), 'Meta', '-', '-', 'INFO'),
        ]
        
        for kpi in kpis:
            writer.writerow(kpi)
    
    print(f"   ✅ Saved: {output_file}")


def generate_all_in_one():
    """Generate a single comprehensive CSV with all check-in data."""
    print("\n📦 Generating All-in-One Dataset CSV...")
    
    data = load_json(BATCH_RESULTS)
    history = load_json(ANALYSIS_HISTORY) or {'analyses': []}
    
    if not data:
        return
    
    # Create detail lookup
    detail_lookup = {}
    for analysis in history.get('analyses', []):
        key = f"{analysis.get('va_name', '')}_{analysis.get('meeting_date', '')}"
        detail_lookup[key] = analysis
    
    output_file = OUTPUT_DIR / "checkin_analysis_complete.csv"
    
    with open(output_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow([
            'VA Name', 'Client', 'Meeting Date', 'Risk Level', 'Risk Score',
            'VA Status', 'Client Health', 'Signals Count', 'Suggestions Count',
            'Signal IDs', 'Key Finding 1', 'Key Finding 2', 'Key Finding 3',
            'Top Suggestion', 'Escalation Needed', 'Escalation Reason',
            'Executive Summary', 'Positive Indicators', 'Analysis ID', 'Analyzed At'
        ])
        
        for record in data.get('processed', []):
            key = f"{record['va_name']}_{record['date']}"
            detail = detail_lookup.get(key, {})
            
            # Risk score
            risk_weights = {'critical': 100, 'high': 75, 'medium': 25, 'low': 5}
            risk_score = risk_weights.get(record['risk_level'].lower(), 10)
            
            # Signal IDs
            signal_ids = [s.get('signal_id', '') for s in detail.get('detected_signals', [])]
            
            # Key findings
            findings = detail.get('key_findings', ['', '', ''])
            while len(findings) < 3:
                findings.append('')
            
            # Top suggestion
            top_sugg = detail.get('ai_suggestions', [{}])[0].get('suggestion', '')[:250] if detail.get('ai_suggestions') else ''
            
            # Positive indicators
            positives = '; '.join(detail.get('positive_indicators', [])[:2])[:200]
            
            writer.writerow([
                record['va_name'],
                record['client_name'][:50] if record['client_name'] else 'Unknown',
                record['date'],
                record['risk_level'].upper(),
                risk_score,
                detail.get('va_status', ''),
                detail.get('client_health', ''),
                record['signals_count'],
                record['suggestions_count'],
                ', '.join(signal_ids),
                findings[0][:150] if findings[0] else '',
                findings[1][:150] if len(findings) > 1 and findings[1] else '',
                findings[2][:150] if len(findings) > 2 and findings[2] else '',
                top_sugg,
                'Yes' if detail.get('escalation_needed') else 'No',
                detail.get('escalation_reason', '')[:200],
                detail.get('executive_summary', '')[:300],
                positives,
                detail.get('analysis_id', ''),
                detail.get('analyzed_at', '')
            ])
    
    print(f"   ✅ Saved: {output_file}")


def main():
    """Generate all CSV files."""
    print("=" * 70)
    print("🚀 GENERATING CSV FILES FOR POWER AUTOMATE & POWER BI")
    print("=" * 70)
    print(f"   Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Generate all CSVs
    va_count = generate_va_risk_summary()
    generate_pending_suggestions_csv()
    alert_count = generate_critical_alerts()
    signal_count = generate_signals_detail()
    review_count = generate_stakeholder_review()
    generate_kpi_summary()
    generate_all_in_one()
    
    print("\n" + "=" * 70)
    print("✅ CSV GENERATION COMPLETE")
    print("=" * 70)
    print(f"""
📁 Output Files Created in {OUTPUT_DIR}:

   FOR POWER BI DASHBOARDS:
   ├── va_risk_summary.csv          - VA risk scores & trends
   ├── va_signals_detail.csv        - Signal-level analysis
   ├── kpi_dashboard_summary.csv    - Executive KPIs
   └── checkin_analysis_complete.csv - Full dataset

   FOR POWER AUTOMATE EMAIL WORKFLOW:
   ├── critical_alerts.csv          - P1/P2 alerts for immediate action
   ├── stakeholder_review.csv       - Review items with approval tracking
   └── pending_suggestions_review.csv - All suggestions pending review

📧 POWER AUTOMATE SETUP:
   1. Upload CSV files to SharePoint/OneDrive
   2. Create Flow triggered on file update
   3. Parse CSV and send email with:
      - Link to SharePoint spreadsheet
      - Summary of critical alerts
      - Review deadline

📊 POWER BI SETUP:
   1. Import CSV files as data sources
   2. Create relationships on VA Name + Date
   3. Build dashboards:
      - Risk Distribution pie chart
      - VA Risk Trend over time
      - Signal Category breakdown
      - KPI scorecards
    """)


if __name__ == "__main__":
    main()
