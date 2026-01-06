"""
Auto Outlier Analyzer
=====================
Automatically processes new check-in meeting transcripts for outlier-level insights.

This script:
1. Monitors for new check-in transcripts
2. Parses VA and client information from filenames/content
3. Runs the Outlier Insights Engine on each
4. Generates reports and suggestions for stakeholder review
5. Creates a daily summary report

Run daily after transcript sync to replace weekly Outlier meetings.
"""

import os
import re
import json
from datetime import datetime, timedelta
from pathlib import Path
from dotenv import load_dotenv

# Import the insights engine
from outlier_insights_engine import (
    analyze_checkin_for_outliers,
    save_analysis_result,
    save_pending_suggestions,
    generate_outlier_report,
    list_pending_suggestions,
    load_knowledge_base
)

load_dotenv()

TRANSCRIPTS_DIR = Path('transcripts')
OUTPUT_DIR = Path('output')
PROCESSED_LOG_FILE = OUTPUT_DIR / 'processed_checkins.json'


def load_processed_log():
    """Load log of already processed transcripts"""
    if PROCESSED_LOG_FILE.exists():
        with open(PROCESSED_LOG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {"processed": [], "last_run": None}


def save_processed_log(log):
    """Save processed log"""
    OUTPUT_DIR.mkdir(exist_ok=True)
    with open(PROCESSED_LOG_FILE, 'w', encoding='utf-8') as f:
        json.dump(log, f, indent=2)


def parse_checkin_filename(filename):
    """
    Parse VA and client information from check-in transcript filename.
    Expected patterns:
    - "VA Name - Client Name - Check-in - Date.txt"
    - "Check-in Meeting - VA Name - Client.txt"
    - "2025-01-07 - VA Name - Client - Check-in.txt"
    """
    # Remove extension
    name = Path(filename).stem
    
    # Try to extract date
    date_match = re.search(r'(\d{4}-\d{2}-\d{2})', name)
    date = date_match.group(1) if date_match else None
    
    # Remove date from name for parsing
    if date:
        name = name.replace(date, '').strip(' -_')
    
    # Try to identify VA and client from common patterns
    # Pattern 1: "VA Name - Client Name - Check-in"
    parts = re.split(r'\s*[-_]\s*', name)
    parts = [p.strip() for p in parts if p.strip() and p.lower() not in ['check', 'in', 'checkin', 'meeting', 'transcript']]
    
    va_name = None
    client_name = None
    
    if len(parts) >= 2:
        va_name = parts[0]
        client_name = parts[1]
    elif len(parts) == 1:
        va_name = parts[0]
        client_name = "Unknown"
    
    return {
        "va_name": va_name,
        "client_name": client_name,
        "date": date or datetime.now().strftime('%Y-%m-%d'),
        "filename": filename
    }


def extract_info_from_transcript(transcript_text, filename):
    """
    Extract VA and client information from transcript content.
    Falls back to filename parsing if content parsing fails.
    """
    info = parse_checkin_filename(filename)
    
    # Try to find VA name from common patterns in transcript
    va_patterns = [
        r"(?:VA|Virtual Assistant)[:\s]+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)",
        r"Check-?in with[:\s]+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)",
        r"Meeting with[:\s]+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)",
    ]
    
    for pattern in va_patterns:
        match = re.search(pattern, transcript_text[:1000])
        if match:
            info["va_name"] = match.group(1)
            break
    
    # Try to find client name
    client_patterns = [
        r"(?:Client|Company)[:\s]+([A-Z][A-Za-z\s&]+)",
        r"working (?:for|with|at)[:\s]+([A-Z][A-Za-z\s&]+)",
    ]
    
    for pattern in client_patterns:
        match = re.search(pattern, transcript_text[:1000])
        if match:
            info["client_name"] = match.group(1).strip()
            break
    
    return info


def find_new_checkin_transcripts(since_date=None):
    """
    Find check-in transcripts that haven't been processed yet.
    """
    processed_log = load_processed_log()
    processed_files = set(processed_log.get("processed", []))
    
    # Find all transcript files
    new_transcripts = []
    
    for transcript_file in TRANSCRIPTS_DIR.glob('**/*.txt'):
        filename = transcript_file.name.lower()
        
        # Check if it's a check-in meeting (not an outlier discussion)
        is_checkin = any(term in filename for term in ['check-in', 'checkin', 'check_in'])
        is_outlier = 'outlier' in filename
        
        if is_checkin and not is_outlier:
            # Check if already processed
            if str(transcript_file) not in processed_files:
                # Check file modification date if since_date provided
                if since_date:
                    file_mtime = datetime.fromtimestamp(transcript_file.stat().st_mtime)
                    if file_mtime < since_date:
                        continue
                
                new_transcripts.append(transcript_file)
    
    return new_transcripts


def process_checkin_transcript(transcript_path):
    """
    Process a single check-in transcript for outlier insights.
    """
    print(f"\n{'='*70}")
    print(f"📄 Processing: {transcript_path.name}")
    print(f"{'='*70}")
    
    # Read transcript
    with open(transcript_path, 'r', encoding='utf-8') as f:
        transcript_text = f.read()
    
    # Extract VA and client info
    info = extract_info_from_transcript(transcript_text, transcript_path.name)
    
    if not info["va_name"] or info["va_name"] == "Unknown":
        print(f"   ⚠️ Could not identify VA name from transcript")
        # Prompt for manual entry or skip
        info["va_name"] = input("   Enter VA name (or press Enter to skip): ").strip()
        if not info["va_name"]:
            print("   ⏭️ Skipping this transcript")
            return None
    
    if not info["client_name"] or info["client_name"] == "Unknown":
        print(f"   ⚠️ Could not identify client name from transcript")
        info["client_name"] = input("   Enter client name (or press Enter to skip): ").strip()
        if not info["client_name"]:
            print("   ⏭️ Skipping this transcript")
            return None
    
    print(f"   👤 VA: {info['va_name']}")
    print(f"   🏢 Client: {info['client_name']}")
    print(f"   📅 Date: {info['date']}")
    
    # Run analysis
    print(f"\n   🤖 Running Outlier Insights Analysis...")
    analysis = analyze_checkin_for_outliers(
        transcript_text,
        info["va_name"],
        info["client_name"],
        info["date"]
    )
    
    if analysis:
        # Save results
        save_analysis_result(analysis)
        if analysis.get("ai_suggestions"):
            save_pending_suggestions(analysis)
        
        # Generate report
        report = generate_outlier_report(analysis)
        
        # Save individual report
        report_filename = f"outlier_report_{info['va_name'].replace(' ', '_')}_{info['date']}.txt"
        report_path = OUTPUT_DIR / report_filename
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(report)
        
        print(f"\n   ✅ Analysis complete")
        print(f"   📁 Report saved: {report_path}")
        
        # Print summary
        print(f"\n   📊 QUICK SUMMARY:")
        print(f"      Status: {analysis.get('va_status', 'unknown').upper()}")
        print(f"      Risk Level: {analysis.get('overall_risk_level', 'unknown').upper()}")
        print(f"      Signals Detected: {len(analysis.get('detected_signals', []))}")
        print(f"      Suggestions Generated: {len(analysis.get('ai_suggestions', []))}")
        
        if analysis.get('escalation_needed'):
            print(f"\n   ⚠️ ESCALATION NEEDED: {analysis.get('escalation_reason', 'See report')}")
        
        return analysis
    else:
        print(f"   ❌ Analysis failed")
        return None


def generate_daily_summary(analyses):
    """
    Generate a daily summary report of all analyzed check-ins.
    This replaces the weekly Outlier meeting with daily automated insights.
    """
    if not analyses:
        return "No analyses to summarize."
    
    today = datetime.now().strftime('%Y-%m-%d')
    
    # Categorize by risk level
    critical = [a for a in analyses if a.get('overall_risk_level') == 'critical']
    high = [a for a in analyses if a.get('overall_risk_level') == 'high']
    medium = [a for a in analyses if a.get('overall_risk_level') == 'medium']
    low = [a for a in analyses if a.get('overall_risk_level') == 'low']
    
    # Count escalations
    escalations = [a for a in analyses if a.get('escalation_needed')]
    
    # Count signals
    all_signals = []
    for a in analyses:
        all_signals.extend([s['signal_id'] for s in a.get('detected_signals', [])])
    
    signal_counts = {}
    for s in all_signals:
        signal_counts[s] = signal_counts.get(s, 0) + 1
    
    summary = f"""
╔══════════════════════════════════════════════════════════════════════════════╗
║              DAILY OUTLIER INSIGHTS SUMMARY - {today}                     ║
║                 Replacing Weekly Outlier Discussion Meeting                  ║
╚══════════════════════════════════════════════════════════════════════════════╝

📊 EXECUTIVE DASHBOARD
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   Check-ins Analyzed: {len(analyses)}
   
   Risk Distribution:
   🔴 Critical: {len(critical)}
   🟠 High:     {len(high)}
   🟡 Medium:   {len(medium)}
   🟢 Low:      {len(low)}
   
   Escalations Required: {len(escalations)}
   Total AI Suggestions: {sum(len(a.get('ai_suggestions', [])) for a in analyses)}

"""
    
    # Critical issues first
    if critical or escalations:
        summary += """
⚠️ REQUIRES IMMEDIATE ATTENTION
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""
        for a in critical + [e for e in escalations if e not in critical]:
            summary += f"""
   🔴 {a['va_name']} at {a['client_name']}
      {a.get('executive_summary', 'See full report')}
      Escalation: {a.get('escalation_reason', 'Critical risk level')}
"""
    
    # High risk
    if high:
        summary += """
🟠 HIGH RISK - Action Within 48h
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""
        for a in high:
            if a not in critical:
                summary += f"""
   {a['va_name']} at {a['client_name']}
      {a.get('executive_summary', 'See full report')}
"""
    
    # Top signals detected
    if signal_counts:
        summary += """
📈 TOP RISK SIGNALS DETECTED TODAY
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""
        from outlier_insights_engine import CHURN_RISK_SIGNALS
        for signal_id, count in sorted(signal_counts.items(), key=lambda x: -x[1])[:5]:
            # Get signal description
            for cat in CHURN_RISK_SIGNALS.values():
                if signal_id in cat:
                    desc = cat[signal_id]['signal']
                    break
            else:
                desc = signal_id
            summary += f"   {signal_id}: {desc} ({count}x)\n"
    
    # Pending approvals
    pending = list_pending_suggestions()
    if pending:
        summary += f"""
📝 PENDING STAKEHOLDER APPROVALS: {len(pending)}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   Run: python src/outlier_insights_engine.py pending
   To review and approve/reject AI suggestions
"""
    
    summary += f"""
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}
This report replaces the weekly Outlier Discussion meeting.
Review individual reports in output/ folder for detailed suggestions.
══════════════════════════════════════════════════════════════════════════════════
"""
    
    return summary


def run_auto_analysis(since_days=1, interactive=False):
    """
    Main function: Run automated analysis on new check-in transcripts.
    
    Args:
        since_days: Only process transcripts from the last N days
        interactive: If True, prompt for missing VA/client info
    """
    print("\n" + "="*70)
    print("🚀 AUTO OUTLIER ANALYZER")
    print("   Automated Check-in Meeting Analysis for Outlier-Level Insights")
    print("="*70)
    
    since_date = datetime.now() - timedelta(days=since_days)
    print(f"\n📅 Looking for transcripts since: {since_date.strftime('%Y-%m-%d')}")
    
    # Find new transcripts
    new_transcripts = find_new_checkin_transcripts(since_date)
    
    if not new_transcripts:
        print("\n✅ No new check-in transcripts to process.")
        print("   All transcripts are already analyzed or no check-ins found.")
        return []
    
    print(f"\n📁 Found {len(new_transcripts)} new check-in transcript(s) to analyze:")
    for t in new_transcripts:
        print(f"   • {t.name}")
    
    # Process each transcript
    analyses = []
    processed_log = load_processed_log()
    
    for transcript_path in new_transcripts:
        try:
            if interactive:
                analysis = process_checkin_transcript(transcript_path)
            else:
                # Auto mode - extract info and analyze
                with open(transcript_path, 'r', encoding='utf-8') as f:
                    transcript_text = f.read()
                
                info = extract_info_from_transcript(transcript_text, transcript_path.name)
                
                if info["va_name"] and info["client_name"]:
                    print(f"\n{'='*70}")
                    print(f"📄 Auto-processing: {transcript_path.name}")
                    print(f"   👤 VA: {info['va_name']} | 🏢 Client: {info['client_name']}")
                    
                    analysis = analyze_checkin_for_outliers(
                        transcript_text,
                        info["va_name"],
                        info["client_name"],
                        info["date"]
                    )
                    
                    if analysis:
                        save_analysis_result(analysis)
                        if analysis.get("ai_suggestions"):
                            save_pending_suggestions(analysis)
                        analyses.append(analysis)
                        print(f"   ✅ Done - Risk: {analysis.get('overall_risk_level', 'unknown').upper()}")
                else:
                    print(f"\n⚠️ Skipping {transcript_path.name} - Could not extract VA/client info")
                    continue
            
            if analysis:
                analyses.append(analysis)
            
            # Mark as processed
            processed_log["processed"].append(str(transcript_path))
            
        except Exception as e:
            print(f"\n❌ Error processing {transcript_path.name}: {e}")
    
    # Update processed log
    processed_log["last_run"] = datetime.now().isoformat()
    save_processed_log(processed_log)
    
    # Generate daily summary
    if analyses:
        print("\n" + "="*70)
        print("📊 GENERATING DAILY SUMMARY")
        print("="*70)
        
        summary = generate_daily_summary(analyses)
        print(summary)
        
        # Save summary
        summary_path = OUTPUT_DIR / f"daily_outlier_summary_{datetime.now().strftime('%Y%m%d')}.txt"
        with open(summary_path, 'w', encoding='utf-8') as f:
            f.write(summary)
        print(f"\n📁 Summary saved: {summary_path}")
    
    return analyses


def main():
    """Main entry point"""
    import sys
    
    if len(sys.argv) > 1:
        if sys.argv[1] == '--interactive':
            run_auto_analysis(since_days=7, interactive=True)
        elif sys.argv[1] == '--days':
            days = int(sys.argv[2]) if len(sys.argv) > 2 else 1
            run_auto_analysis(since_days=days)
        elif sys.argv[1] == '--help':
            print("""
Auto Outlier Analyzer - Automated Check-in Analysis

Usage:
  python src/auto_outlier_analyzer.py                   # Process last 24h of transcripts
  python src/auto_outlier_analyzer.py --days 7          # Process last 7 days
  python src/auto_outlier_analyzer.py --interactive     # Interactive mode (prompt for info)
  python src/auto_outlier_analyzer.py --help            # Show this help

This script:
1. Finds new check-in meeting transcripts
2. Extracts VA and client information
3. Runs AI analysis for churn risk detection
4. Generates suggestions based on past Outlier meeting patterns
5. Creates a daily summary report

Run this after syncing transcripts to get automated Outlier-level insights.
""")
        else:
            print(f"Unknown argument: {sys.argv[1]}")
    else:
        run_auto_analysis(since_days=1)


if __name__ == '__main__':
    main()
