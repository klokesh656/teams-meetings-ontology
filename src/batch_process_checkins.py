"""
Batch Process Existing Check-in Meetings
=========================================
Processes all existing check-in meeting transcripts through the Outlier Insights Engine.
Extracts VA and client names from filenames and runs AI analysis on each.
"""

import os
import re
import json
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv

from outlier_insights_engine import (
    analyze_checkin_for_outliers,
    save_analysis_result,
    save_pending_suggestions,
    generate_outlier_report,
    load_knowledge_base
)

load_dotenv()

TRANSCRIPTS_DIR = Path('transcripts')
OUTPUT_DIR = Path('output')


def parse_checkin_filename(filename):
    """
    Parse VA name, client name, and date from check-in transcript filename.
    
    Patterns:
    - 20251212_010807_Catch Up with Jep.vtt
    - 20251106_Integration_Team_Check-in_Louise_x_Crystal.vtt
    - 20251104_144106_Integration Team Check-in  Shey x Catherine.vtt
    - unknown_Integration Team Check-in Shey xCarla.vtt
    """
    name = Path(filename).stem
    
    # Extract date from filename
    date_match = re.match(r'^(\d{8})', name)
    if date_match:
        date_str = date_match.group(1)
        date = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:8]}"
    else:
        date = None
    
    # Extract VA name from various patterns
    va_name = None
    integration_person = None  # Shey or Louise
    
    # Pattern 1: "Shey x <VA>" or "Louise x <VA>" (with or without spaces)
    match = re.search(r'(Shey|Louise)[_ ]*x[_ ]*([A-Za-z]+)', name, re.IGNORECASE)
    if match:
        integration_person = match.group(1)
        va_name = match.group(2)
    
    # Pattern 2: "Catch Up with <VA>" or "Catch up with <VA>"
    if not va_name:
        match = re.search(r'Catch[_ ]+[Uu]p[_ ]+with[_ ]+([A-Za-z]+)', name)
        if match:
            va_name = match.group(1)
    
    # Pattern 3: "Check-in  Shey x <VA>" with double space
    if not va_name:
        match = re.search(r'Check-?in\s+(Shey|Louise)\s*x\s*([A-Za-z]+)', name, re.IGNORECASE)
        if match:
            integration_person = match.group(1)
            va_name = match.group(2)
    
    # Pattern 4: "Daily EOD Check in with <VA>"
    if not va_name:
        match = re.search(r'(?:Daily|EOD|SOD)[_ ]+(?:EOD|SOD)?[_ ]*Check[_ ]+in[_ ]+with[_ ]+([A-Za-z]+)', name, re.IGNORECASE)
        if match:
            va_name = match.group(1)
    
    # Pattern 5: "Quick Catch Up - <VA>"
    if not va_name:
        match = re.search(r'Quick[_ ]+Catch[_ ]+Up[_ ]*-[_ ]*([A-Za-z]+)', name, re.IGNORECASE)
        if match:
            va_name = match.group(1)
    
    # Pattern 6: Last word before date suffix (e.g., "Shey x Mary-20251203")
    if not va_name:
        match = re.search(r'x\s*([A-Za-z]+)(?:-\d|\.)', name, re.IGNORECASE)
        if match:
            va_name = match.group(1)
    
    # Try to extract client from GTM/Onboarding patterns
    client_name = None
    
    # Pattern: "GTM Orientation <VA> x <Client>" or "Onboarding x <VA> (<Client>)"
    client_match = re.search(r'(?:GTM|Orientation|Onboarding|Readiness)[_ ]+(?:Check)?[_ ]*(?:x)?[_ ]*[A-Za-z]+[_ ]+(?:x|&)?[_ ]*([A-Za-z]+(?:[_ ]+[A-Za-z]+)?)', name, re.IGNORECASE)
    if client_match:
        potential_client = client_match.group(1)
        # Filter out common non-client words
        if potential_client.lower() not in ['check', 'in', 'team', 'integration', 'meeting']:
            client_name = potential_client
    
    # Default client to "Unknown" - will need to extract from transcript content
    if not client_name:
        client_name = "Unknown"
    
    return {
        "va_name": va_name,
        "client_name": client_name,
        "integration_person": integration_person,
        "date": date,
        "filename": filename
    }


def extract_client_from_transcript(transcript_text, va_name):
    """
    Try to extract client name from transcript content.
    """
    # Common patterns in transcripts
    patterns = [
        rf"{va_name}.*(?:works? (?:for|with|at)|assigned to|client is)[:\s]+([A-Z][A-Za-z\s&]+)",
        r"(?:client|company|working (?:for|with|at))[:\s]+([A-Z][A-Za-z\s&]+)",
        r"(?:assigned to|supporting)[:\s]+([A-Z][A-Za-z\s&]+)",
    ]
    
    for pattern in patterns:
        match = re.search(pattern, transcript_text[:3000], re.IGNORECASE)
        if match:
            client = match.group(1).strip()
            # Filter out common false positives
            if client.lower() not in ['the', 'a', 'an', 'our', 'your', 'their']:
                return client[:30]  # Limit length
    
    return None


def find_all_checkin_transcripts():
    """
    Find all check-in related transcripts in the transcripts folder.
    """
    checkin_patterns = [
        '*Check-in*',
        '*Check_in*',
        '*Checkin*',
        '*Catch*up*',
        '*Catch*Up*',
    ]
    
    transcripts = set()
    for pattern in checkin_patterns:
        transcripts.update(TRANSCRIPTS_DIR.glob(pattern))
    
    # Filter out Outlier discussions and interviews
    filtered = []
    for t in transcripts:
        name_lower = t.name.lower()
        if 'outlier' not in name_lower and 'interview' not in name_lower:
            filtered.append(t)
    
    return sorted(filtered, key=lambda x: x.name)


def batch_analyze_existing_checkins(limit=None, skip_already_processed=True):
    """
    Process all existing check-in transcripts.
    
    Args:
        limit: Maximum number of transcripts to process (None = all)
        skip_already_processed: Skip files already in analysis history
    """
    print("\n" + "="*70)
    print("🚀 BATCH PROCESSING EXISTING CHECK-IN MEETINGS")
    print("="*70)
    print(f"   Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Find all check-in transcripts
    transcripts = find_all_checkin_transcripts()
    print(f"\n📁 Found {len(transcripts)} check-in transcripts")
    
    if limit:
        transcripts = transcripts[:limit]
        print(f"   Processing first {limit} transcripts")
    
    # Load existing analysis history to skip duplicates
    processed_vas = set()
    if skip_already_processed:
        history_file = OUTPUT_DIR / 'checkin_analysis_history.json'
        if history_file.exists():
            with open(history_file, 'r', encoding='utf-8') as f:
                history = json.load(f)
            for a in history.get('analyses', []):
                key = f"{a.get('va_name', '')}_{a.get('meeting_date', '')}"
                processed_vas.add(key)
            print(f"   Skipping {len(processed_vas)} already processed")
    
    # Process each transcript
    results = {
        "processed": [],
        "skipped": [],
        "failed": [],
        "critical": [],
        "high_risk": []
    }
    
    for idx, transcript_path in enumerate(transcripts, 1):
        print(f"\n[{idx}/{len(transcripts)}] {transcript_path.name[:60]}...")
        
        # Parse filename
        info = parse_checkin_filename(transcript_path.name)
        
        if not info["va_name"]:
            print(f"   ⚠️ Could not extract VA name - skipping")
            results["skipped"].append({"file": transcript_path.name, "reason": "No VA name"})
            continue
        
        # Check if already processed
        key = f"{info['va_name']}_{info['date']}"
        if key in processed_vas:
            print(f"   ⏭️ Already processed - skipping")
            results["skipped"].append({"file": transcript_path.name, "reason": "Already processed"})
            continue
        
        # Read transcript
        try:
            with open(transcript_path, 'r', encoding='utf-8') as f:
                transcript_text = f.read()
        except Exception as e:
            print(f"   ❌ Error reading file: {e}")
            results["failed"].append({"file": transcript_path.name, "error": str(e)})
            continue
        
        # Try to extract client from transcript if unknown
        if info["client_name"] == "Unknown":
            extracted_client = extract_client_from_transcript(transcript_text, info["va_name"])
            if extracted_client:
                info["client_name"] = extracted_client
        
        date = info["date"] or datetime.now().strftime('%Y-%m-%d')
        
        print(f"   👤 VA: {info['va_name']}")
        print(f"   🏢 Client: {info['client_name']}")
        print(f"   📅 Date: {date}")
        
        # Run analysis
        try:
            analysis = analyze_checkin_for_outliers(
                transcript_text,
                info["va_name"],
                info["client_name"],
                date
            )
            
            if analysis:
                # Save results
                save_analysis_result(analysis)
                if analysis.get("ai_suggestions"):
                    save_pending_suggestions(analysis)
                
                # Track results
                results["processed"].append({
                    "va_name": info["va_name"],
                    "client_name": info["client_name"],
                    "date": date,
                    "risk_level": analysis.get("overall_risk_level", "unknown"),
                    "signals_count": len(analysis.get("detected_signals", [])),
                    "suggestions_count": len(analysis.get("ai_suggestions", []))
                })
                
                risk_level = analysis.get("overall_risk_level", "low")
                if risk_level == "critical":
                    results["critical"].append(analysis)
                    print(f"   🔴 CRITICAL RISK - {len(analysis.get('detected_signals', []))} signals")
                elif risk_level == "high":
                    results["high_risk"].append(analysis)
                    print(f"   🟠 HIGH RISK - {len(analysis.get('detected_signals', []))} signals")
                else:
                    print(f"   ✅ {risk_level.upper()} - {len(analysis.get('detected_signals', []))} signals")
            else:
                results["failed"].append({"file": transcript_path.name, "error": "Analysis returned None"})
                print(f"   ❌ Analysis failed")
                
        except Exception as e:
            results["failed"].append({"file": transcript_path.name, "error": str(e)})
            print(f"   ❌ Error: {e}")
    
    # Generate summary
    print("\n" + "="*70)
    print("📊 BATCH PROCESSING SUMMARY")
    print("="*70)
    print(f"\n   ✅ Processed: {len(results['processed'])}")
    print(f"   ⏭️ Skipped: {len(results['skipped'])}")
    print(f"   ❌ Failed: {len(results['failed'])}")
    print(f"\n   🔴 Critical Risk: {len(results['critical'])}")
    print(f"   🟠 High Risk: {len(results['high_risk'])}")
    
    # Show critical cases
    if results["critical"]:
        print("\n" + "-"*70)
        print("🚨 CRITICAL RISK CASES REQUIRING IMMEDIATE ATTENTION:")
        print("-"*70)
        for a in results["critical"]:
            print(f"\n   {a['va_name']} at {a['client_name']} ({a['meeting_date']})")
            print(f"   {a.get('executive_summary', 'See report')[:100]}...")
    
    # Save batch results
    results_file = OUTPUT_DIR / f"batch_analysis_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    with open(results_file, 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=2, default=str)
    print(f"\n   📁 Results saved: {results_file}")
    
    # Count total pending suggestions
    pending_file = OUTPUT_DIR / 'pending_suggestions.json'
    if pending_file.exists():
        with open(pending_file, 'r', encoding='utf-8') as f:
            pending = json.load(f)
        pending_count = len([s for s in pending.get('suggestions', []) if s['status'] == 'pending'])
        print(f"\n   📋 Total pending suggestions for review: {pending_count}")
        print(f"   Run: python src/outlier_insights_engine.py pending")
    
    return results


def main():
    """Main entry point"""
    import sys
    
    limit = None
    if len(sys.argv) > 1:
        if sys.argv[1] == '--limit':
            limit = int(sys.argv[2]) if len(sys.argv) > 2 else 10
        elif sys.argv[1] == '--all':
            limit = None
        elif sys.argv[1] == '--help':
            print("""
Batch Process Existing Check-in Meetings

Usage:
  python src/batch_process_checkins.py              # Process all (skip already done)
  python src/batch_process_checkins.py --limit 10   # Process first 10 only
  python src/batch_process_checkins.py --all        # Process all transcripts
  python src/batch_process_checkins.py --help       # Show this help
""")
            return
    
    batch_analyze_existing_checkins(limit=limit)


if __name__ == '__main__':
    main()
