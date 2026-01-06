"""
Fix Excel Meeting Data
======================
1. Extract date from transcript filename where missing
2. Add new columns:
   - Meeting Type: Check-in, Interview, Orientation, etc.
   - HR Person: Shey, Louise, or HR@our-assistants.com
   - VA Name: The virtual assistant being checked in with
"""

import re
import pandas as pd
from pathlib import Path
from datetime import datetime

# Paths
OUTPUT_DIR = Path('output')
EXCEL_PATH = OUTPUT_DIR / 'meeting_transcripts_latest_analyzed.xlsx'


def extract_date_from_filename(filename):
    """Extract date from transcript filename like 20251204_154721_..."""
    if not filename or pd.isna(filename):
        return None
    
    # Pattern: YYYYMMDD_HHMMSS_...
    match = re.match(r'(\d{8})_(\d{6})_', str(filename))
    if match:
        date_str = match.group(1)
        try:
            return datetime.strptime(date_str, '%Y%m%d').strftime('%Y-%m-%d')
        except:
            pass
    return None


def extract_time_from_filename(filename):
    """Extract time from transcript filename"""
    if not filename or pd.isna(filename):
        return None
    
    match = re.match(r'\d{8}_(\d{6})_', str(filename))
    if match:
        time_str = match.group(1)
        try:
            return f"{time_str[:2]}:{time_str[2:4]}:{time_str[4:6]}"
        except:
            pass
    return None


def determine_meeting_type(subject):
    """Determine meeting type from subject"""
    if not subject or pd.isna(subject):
        return "Unknown"
    
    subject_lower = str(subject).lower()
    
    if 'check-in' in subject_lower or 'checkin' in subject_lower:
        return "Check-in"
    elif 'interview' in subject_lower:
        return "Interview"
    elif 'orientation' in subject_lower:
        return "Orientation"
    elif 'onboarding' in subject_lower:
        return "Onboarding"
    elif 'readiness' in subject_lower:
        return "Readiness Check"
    elif 'gtm' in subject_lower:
        return "GTM"
    elif 'catch up' in subject_lower or 'catchup' in subject_lower:
        return "Catch-up"
    elif 'eod' in subject_lower:
        return "EOD Check-in"
    elif 'sod' in subject_lower:
        return "SOD Check-in"
    else:
        return "Other"


def extract_hr_person(subject):
    """Extract HR person from subject (Shey, Louise, or HR)"""
    if not subject or pd.isna(subject):
        return ""
    
    subject_lower = str(subject).lower()
    
    # Check for Shey
    if 'shey' in subject_lower:
        return "Shey"
    
    # Check for Louise
    if 'louise' in subject_lower:
        return "Louise"
    
    # Check for other HR indicators
    if 'hr' in subject_lower:
        return "HR"
    
    # For interviews, the HR person is typically the organizer
    # Check common patterns
    if 'integration team' in subject_lower:
        # Try to extract from pattern "Shey x Name" or "Louise x Name"
        match = re.search(r'(shey|louise)\s*x', subject_lower)
        if match:
            return match.group(1).title()
    
    return ""


def extract_va_name(subject):
    """Extract VA name from check-in meeting subject"""
    if not subject or pd.isna(subject):
        return ""
    
    subject_str = str(subject)
    
    # Pattern 1: "Check-in Shey x VAName" or "Check-in Louise x VAName"
    match = re.search(r'(?:shey|louise)\s*x\s*([A-Za-z]+(?:\s+[A-Za-z]+)?)', subject_str, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    
    # Pattern 2: "Check-in Shey xVAName" (no space after x)
    match = re.search(r'(?:shey|louise)\s*x([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)', subject_str)
    if match:
        return match.group(1).strip()
    
    # Pattern 3: For interviews, extract name after "Interview -"
    match = re.search(r'interview\s*[-:]\s*([A-Za-z]+(?:\s+[A-Za-z]+)?(?:\s+[A-Za-z]+)?)', subject_str, re.IGNORECASE)
    if match:
        name = match.group(1).strip()
        # Clean up common suffixes
        name = re.sub(r'\s*[-]\s*(VA|PM|MC|APM|PMA|HOA|Bookkeeper|Accountant).*', '', name, flags=re.IGNORECASE)
        return name
    
    # Pattern 4: "Catch up with Name"
    match = re.search(r'catch\s*up\s*with\s+([A-Za-z]+(?:\s+(?:and|&)\s+[A-Za-z]+)?)', subject_str, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    
    return ""


def fix_excel():
    """Main function to fix Excel data"""
    print("="*60)
    print("FIXING EXCEL MEETING DATA")
    print("="*60)
    
    # Load Excel
    df = pd.read_excel(EXCEL_PATH)
    print(f"Loaded {len(df)} rows from {EXCEL_PATH.name}")
    print(f"Columns: {df.columns.tolist()}")
    
    # Count issues before
    unknown_dates = (df['Date'] == 'Unknown').sum() if 'Date' in df.columns else 0
    missing_subjects = df['Subject'].isna().sum() if 'Subject' in df.columns else 0
    print(f"\nIssues found:")
    print(f"  - Rows with 'Unknown' date: {unknown_dates}")
    print(f"  - Rows with missing Subject: {missing_subjects}")
    
    # Fix missing Subject from transcript filename
    print("\n0. Fixing missing Subject from transcript filename...")
    fixed_subjects = 0
    for idx, row in df.iterrows():
        if pd.isna(row.get('Subject')) or row.get('Subject') == '':
            filename = row.get('Transcript File', '')
            if filename:
                # Extract subject from filename like "20251204_154721_Subject Name.vtt"
                match = re.match(r'\d{8}_\d{6}_(.+)\.vtt', str(filename))
                if match:
                    subject = match.group(1).strip()
                    df.at[idx, 'Subject'] = subject
                    fixed_subjects += 1
    print(f"   Fixed {fixed_subjects} subjects")
    
    # Fix dates from filename
    print("\n1. Fixing dates from transcript filenames...")
    fixed_dates = 0
    for idx, row in df.iterrows():
        if row.get('Date') == 'Unknown' or pd.isna(row.get('Date')):
            filename = row.get('Transcript File', '')
            new_date = extract_date_from_filename(filename)
            if new_date:
                df.at[idx, 'Date'] = new_date
                fixed_dates += 1
    print(f"   Fixed {fixed_dates} dates")
    
    # Fix times from filename
    print("\n2. Fixing times from transcript filenames...")
    fixed_times = 0
    for idx, row in df.iterrows():
        if pd.isna(row.get('Time')) or row.get('Time') == '':
            filename = row.get('Transcript File', '')
            new_time = extract_time_from_filename(filename)
            if new_time:
                df.at[idx, 'Time'] = new_time
                fixed_times += 1
    print(f"   Fixed {fixed_times} times")
    
    # Add Meeting Type column
    print("\n3. Adding Meeting Type column...")
    df['Meeting Type'] = df['Subject'].apply(determine_meeting_type)
    type_counts = df['Meeting Type'].value_counts()
    print(f"   Meeting types: {dict(type_counts)}")
    
    # Add HR Person column
    print("\n4. Adding HR Person column...")
    df['HR Person'] = df['Subject'].apply(extract_hr_person)
    hr_counts = df['HR Person'].value_counts()
    print(f"   HR persons: {dict(hr_counts)}")
    
    # Add VA Name column
    print("\n5. Adding VA Name column...")
    df['VA Name'] = df['Subject'].apply(extract_va_name)
    va_count = (df['VA Name'] != '').sum()
    print(f"   Extracted {va_count} VA names")
    
    # Reorder columns - put new columns near the front
    print("\n6. Reordering columns...")
    priority_cols = ['Subject', 'Date', 'Time', 'Meeting Type', 'HR Person', 'VA Name']
    other_cols = [c for c in df.columns if c not in priority_cols]
    new_order = priority_cols + other_cols
    # Only include columns that exist
    new_order = [c for c in new_order if c in df.columns]
    df = df[new_order]
    
    # Save updated Excel
    output_path = OUTPUT_DIR / f'meeting_transcripts_fixed_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    df.to_excel(output_path, index=False)
    print(f"\n✅ Saved updated Excel to: {output_path.name}")
    
    # Print summary
    print("\n" + "="*60)
    print("SUMMARY")
    print("="*60)
    print(f"Total rows: {len(df)}")
    print(f"Dates fixed: {fixed_dates}")
    print(f"Times fixed: {fixed_times}")
    print(f"\nMeeting Types:")
    for mtype, count in type_counts.items():
        print(f"  {mtype}: {count}")
    print(f"\nHR Persons:")
    for hr, count in hr_counts.items():
        if hr:
            print(f"  {hr}: {count}")
    
    # Show sample of check-in meetings
    checkins = df[df['Meeting Type'] == 'Check-in'][['Subject', 'Date', 'HR Person', 'VA Name']].head(15)
    print(f"\nSample Check-in Meetings:")
    print(checkins.to_string())
    
    return df


if __name__ == '__main__':
    fix_excel()
