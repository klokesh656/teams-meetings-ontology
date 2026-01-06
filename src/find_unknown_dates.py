"""
Find Dates for Unknown Meetings
===============================
Searches recordings folder and transcript content to find dates for meetings
that have 'unknown_' prefix in their transcript filenames.
"""

import re
import pandas as pd
from pathlib import Path
from datetime import datetime
from difflib import SequenceMatcher

OUTPUT_DIR = Path('output')
RECORDINGS_DIR = Path('recordings')
TRANSCRIPTS_DIR = Path('transcripts')

def similar(a, b):
    """Calculate similarity ratio between two strings"""
    return SequenceMatcher(None, a.lower(), b.lower()).ratio()

def extract_key_name(subject):
    """Extract the key identifier from a meeting subject"""
    # Remove common prefixes
    subject = re.sub(r'^(Integration Team Check-in|OurAssistants|Daily EOD Check in|Daily SOD Check in)\s*', '', subject, flags=re.IGNORECASE)
    # Remove 'Shey x', 'Louise x', etc.
    subject = re.sub(r'(Shey|Louise)\s*x?\s*', '', subject, flags=re.IGNORECASE)
    # Clean up
    subject = re.sub(r'\s+', ' ', subject).strip()
    return subject

def find_date_in_recordings(subject):
    """Search recordings folder for matching meeting"""
    key_name = extract_key_name(subject)
    
    best_match = None
    best_score = 0
    
    # Get all MP4 files and folders in recordings
    for item in RECORDINGS_DIR.iterdir():
        name = item.name
        
        # Extract date from name
        match = re.match(r'(\d{8})_(\d{6})_(.+)', name)
        if not match:
            continue
        
        date_str = match.group(1)
        recording_subject = match.group(3).replace('.mp4', '').replace('.wav', '')
        
        # Check similarity
        rec_key = extract_key_name(recording_subject)
        score = similar(key_name, rec_key)
        
        # Also check if subject words appear in recording name
        subject_words = set(subject.lower().split())
        name_words = set(name.lower().split())
        common_words = subject_words & name_words
        word_score = len(common_words) / max(len(subject_words), 1)
        
        combined_score = (score + word_score) / 2
        
        if combined_score > best_score and combined_score > 0.4:
            best_score = combined_score
            try:
                date = datetime.strptime(date_str, '%Y%m%d').strftime('%Y-%m-%d')
                best_match = (date, name, combined_score)
            except:
                pass
    
    return best_match

def find_date_in_transcript_content(vtt_path):
    """Search inside VTT file for date references"""
    try:
        with open(vtt_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Look for date patterns in the content
        # Pattern: December 5, 2025 or Dec 5, 2025 or 12/5/2025
        patterns = [
            r'(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})',
            r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2}),?\s+(\d{4})',
            r'(\d{1,2})/(\d{1,2})/(\d{4})',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, content)
            if match:
                return match.group(0)
        
        return None
    except:
        return None

def find_date_from_other_transcripts(subject, df):
    """Find if same VA has other meetings with known dates"""
    # Extract VA name
    key_name = extract_key_name(subject)
    
    # Look for similar subjects in known dates
    known_dates = df[df['Date'] != 'Unknown']
    
    for _, row in known_dates.iterrows():
        row_key = extract_key_name(row['Subject'])
        if similar(key_name, row_key) > 0.8:
            return row['Date'], row['Subject']
    
    return None, None

def main():
    print("="*70)
    print("FINDING DATES FOR UNKNOWN MEETINGS")
    print("="*70)
    
    # Load Excel
    df = pd.read_excel(OUTPUT_DIR / 'meeting_transcripts_fixed_20251231_012755.xlsx')
    unknown = df[df['Date'] == 'Unknown'].copy()
    print(f"\nMeetings with unknown dates: {len(unknown)}")
    
    found_dates = []
    
    for idx, row in unknown.iterrows():
        subject = row['Subject']
        transcript_file = row['Transcript File']
        
        print(f"\n[{len(found_dates)+1}/{len(unknown)}] {subject}")
        
        # Method 1: Search recordings folder
        recording_match = find_date_in_recordings(subject)
        if recording_match:
            date, matched_file, score = recording_match
            print(f"  ✓ Found in recordings: {date} (score: {score:.2f})")
            print(f"    Matched: {matched_file[:60]}...")
            found_dates.append({
                'idx': idx,
                'subject': subject,
                'date': date,
                'source': 'recordings',
                'matched_to': matched_file
            })
            continue
        
        # Method 2: Check transcript content for date references
        vtt_path = TRANSCRIPTS_DIR / transcript_file
        content_date = find_date_in_transcript_content(vtt_path)
        if content_date:
            print(f"  ✓ Found date reference in transcript: {content_date}")
            found_dates.append({
                'idx': idx,
                'subject': subject,
                'date': content_date,
                'source': 'transcript_content',
                'matched_to': None
            })
            continue
        
        # Method 3: Look for same VA in other meetings
        similar_date, similar_subject = find_date_from_other_transcripts(subject, df)
        if similar_date:
            print(f"  ~ Found similar meeting: {similar_subject[:50]}... ({similar_date})")
            # Don't auto-apply this - just note it
        
        print(f"  ✗ Could not find date")
    
    # Summary
    print("\n" + "="*70)
    print("SUMMARY")
    print("="*70)
    print(f"Total unknown: {len(unknown)}")
    print(f"Dates found: {len(found_dates)}")
    print(f"Still unknown: {len(unknown) - len(found_dates)}")
    
    if found_dates:
        print("\n" + "="*70)
        print("DATES FOUND")
        print("="*70)
        for item in found_dates:
            print(f"\n{item['subject']}")
            print(f"  Date: {item['date']}")
            print(f"  Source: {item['source']}")
            if item['matched_to']:
                print(f"  Matched to: {item['matched_to'][:60]}...")
        
        # Update Excel
        print("\n" + "="*70)
        print("UPDATING EXCEL")
        print("="*70)
        
        for item in found_dates:
            df.at[item['idx'], 'Date'] = item['date']
        
        # Save
        output_path = OUTPUT_DIR / f'meeting_transcripts_dates_fixed_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        df.to_excel(output_path, index=False)
        print(f"✅ Saved to: {output_path.name}")
        
        # Show remaining unknown
        still_unknown = df[df['Date'] == 'Unknown']['Subject'].tolist()
        if still_unknown:
            print(f"\nStill unknown ({len(still_unknown)}):")
            for s in still_unknown[:10]:
                print(f"  - {s}")
            if len(still_unknown) > 10:
                print(f"  ... and {len(still_unknown) - 10} more")

if __name__ == '__main__':
    main()
