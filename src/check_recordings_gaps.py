"""
Check recordings folder for check-in meetings that need transcription.
Compares recordings with existing transcripts to find gaps.
"""

import os
import re
from datetime import datetime
from pathlib import Path

RECORDINGS_DIR = Path('recordings')
TRANSCRIPTS_DIR = Path('transcripts')

def get_date_from_name(name):
    """Extract date from filename (YYYYMMDD format)"""
    match = re.match(r'(\d{8})', name)
    if match:
        return match.group(1)
    return None

def normalize_subject(name):
    """Normalize subject for comparison"""
    # Remove date prefix, extension, and common patterns
    name = re.sub(r'^\d{8}_\d{6}_', '', name)
    name = re.sub(r'\.(vtt|mp4)$', '', name)
    name = re.sub(r'-\d{8}_\d+$', '', name)  # Remove trailing date patterns
    name = re.sub(r'[^a-zA-Z0-9\s]', '', name)
    return name.lower().strip()[:40]

def main():
    print("=" * 70)
    print("RECORDINGS vs TRANSCRIPTS - GAP ANALYSIS")
    print("=" * 70)
    print(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Get all transcripts
    transcripts = {}
    for f in TRANSCRIPTS_DIR.glob('*.vtt'):
        date = get_date_from_name(f.name)
        subject = normalize_subject(f.name)
        key = f"{date}_{subject[:20]}" if date else subject[:30]
        transcripts[key] = f.name
    
    print(f"\n📄 Total transcripts: {len(transcripts)}")
    
    # Get all recordings (mp4 files and directories)
    recordings = []
    
    # Check for mp4 files directly in recordings folder
    for f in RECORDINGS_DIR.glob('*.mp4'):
        recordings.append({
            'name': f.name,
            'path': str(f),
            'size': f.stat().st_size,
            'date': get_date_from_name(f.name)
        })
    
    # Check for directories (some recordings are in folders)
    for d in RECORDINGS_DIR.iterdir():
        if d.is_dir():
            # Check for mp4 files inside
            mp4_files = list(d.glob('*.mp4'))
            if mp4_files:
                for mp4 in mp4_files:
                    recordings.append({
                        'name': d.name,
                        'path': str(mp4),
                        'size': mp4.stat().st_size,
                        'date': get_date_from_name(d.name)
                    })
            else:
                # Folder might contain the recording
                recordings.append({
                    'name': d.name,
                    'path': str(d),
                    'size': 0,
                    'date': get_date_from_name(d.name)
                })
    
    print(f"🎥 Total recordings: {len(recordings)}")
    
    # Filter for check-in/integration meetings
    checkin_keywords = ['check-in', 'check in', 'checkin', 'integration', 'shey', 'louise']
    
    checkin_recordings = []
    for r in recordings:
        name_lower = r['name'].lower()
        if any(kw in name_lower for kw in checkin_keywords):
            checkin_recordings.append(r)
    
    print(f"📋 Check-in/Integration recordings: {len(checkin_recordings)}")
    
    # Find recordings without transcripts
    missing_transcripts = []
    has_transcript = []
    
    for rec in checkin_recordings:
        date = rec['date']
        subject = normalize_subject(rec['name'])
        
        # Check if we have a matching transcript
        found = False
        for t_key, t_name in transcripts.items():
            t_date = get_date_from_name(t_name)
            t_subject = normalize_subject(t_name)
            
            # Match by date and partial subject
            if date and t_date and date == t_date:
                if subject[:15] in t_subject or t_subject[:15] in subject:
                    found = True
                    break
        
        if found:
            has_transcript.append(rec)
        else:
            missing_transcripts.append(rec)
    
    # Sort by date (most recent first)
    missing_transcripts.sort(key=lambda x: x['date'] or '0', reverse=True)
    
    print(f"\n✅ With transcripts: {len(has_transcript)}")
    print(f"❌ Missing transcripts: {len(missing_transcripts)}")
    
    # Display missing by month
    print("\n" + "=" * 70)
    print("RECORDINGS MISSING TRANSCRIPTS (Check-in/Integration)")
    print("=" * 70)
    
    by_month = {}
    for r in missing_transcripts:
        date = r['date']
        if date:
            month = date[:6]  # YYYYMM
            month_display = f"{date[:4]}-{date[4:6]}"
        else:
            month = 'unknown'
            month_display = 'Unknown'
        
        if month_display not in by_month:
            by_month[month_display] = []
        by_month[month_display].append(r)
    
    for month in sorted(by_month.keys(), reverse=True):
        recs = by_month[month]
        print(f"\n📅 {month}: {len(recs)} recordings")
        
        # Group by person
        by_person = {}
        for r in recs:
            name = r['name']
            if ' x ' in name.lower():
                parts = name.lower().split(' x ')
                person = parts[-1].split('-')[0].split('(')[0].strip()[:20].title()
            else:
                person = 'Other'
            
            if person not in by_person:
                by_person[person] = []
            by_person[person].append(r)
        
        for person in sorted(by_person.keys()):
            person_recs = by_person[person]
            print(f"   👤 {person}: {len(person_recs)}")
            for r in person_recs[:3]:
                size_mb = r['size'] / (1024*1024) if r['size'] else 0
                print(f"      - {r['name'][:55]}... ({size_mb:.1f} MB)")
            if len(person_recs) > 3:
                print(f"      ... and {len(person_recs) - 3} more")
    
    # Summary stats
    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    print(f"\n📊 Check-in/Integration Recordings Analysis:")
    print(f"   Total recordings: {len(checkin_recordings)}")
    print(f"   ✅ With transcripts: {len(has_transcript)} ({100*len(has_transcript)//max(len(checkin_recordings),1)}%)")
    print(f"   ❌ Need transcription: {len(missing_transcripts)} ({100*len(missing_transcripts)//max(len(checkin_recordings),1)}%)")
    
    # Estimate transcription time/cost
    total_size = sum(r['size'] for r in missing_transcripts if r['size'])
    total_gb = total_size / (1024*1024*1024)
    # Rough estimate: 1 hour of audio = ~100MB, Azure Speech = $1/hour
    est_hours = total_size / (100*1024*1024)
    est_cost = est_hours * 1.0
    
    print(f"\n💾 Total size of recordings needing transcription: {total_gb:.2f} GB")
    print(f"⏱️ Estimated audio hours: ~{est_hours:.0f} hours")
    print(f"💰 Estimated Azure Speech cost: ~${est_cost:.2f}")
    
    # List recent ones that should be prioritized
    print("\n" + "=" * 70)
    print("PRIORITY: Recent recordings (Dec 2025)")
    print("=" * 70)
    
    dec_recordings = [r for r in missing_transcripts if r['date'] and r['date'].startswith('202512')]
    print(f"\nDecember 2025 recordings without transcripts: {len(dec_recordings)}")
    
    for r in dec_recordings[:20]:
        size_mb = r['size'] / (1024*1024) if r['size'] else 0
        print(f"  📹 {r['name'][:60]}")
        print(f"     Size: {size_mb:.1f} MB | Date: {r['date']}")
    
    if len(dec_recordings) > 20:
        print(f"\n  ... and {len(dec_recordings) - 20} more")
    
    return missing_transcripts

if __name__ == "__main__":
    missing = main()
