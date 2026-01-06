"""
Fix Excel metadata for transcripts that are missing Date, Time, Meeting Type, HR Person, VA Name.
Parses the Transcript File name to extract these values.
"""
import os
import re
from pathlib import Path
from datetime import datetime
import pandas as pd

OUTPUT_DIR = Path('output')

def parse_transcript_filename(filename):
    """
    Parse transcript filename to extract metadata.
    
    Examples:
    - 20251106_220053_Integration Team Check-in  Louise x Irvy.vtt
    - 20251203_212035_Integration Team Check-in  Louise x Jeanvic-20251204.vtt
    - 20251217_Integration_Team_Check-in_Louise_x_Kim.vtt
    """
    result = {
        'date': None,
        'time': None,
        'meeting_type': None,
        'hr_person': None,
        'va_name': None,
        'subject': None
    }
    
    # Remove .vtt extension
    name = filename.replace('.vtt', '')
    
    # Try to extract date and time from start of filename
    # Pattern 1: 20251106_220053_... (date_time_rest)
    match1 = re.match(r'^(\d{8})_(\d{6})_(.+)$', name)
    # Pattern 2: 20251217_... (date_rest, no time)
    match2 = re.match(r'^(\d{8})_(.+)$', name)
    
    if match1:
        date_str = match1.group(1)
        time_str = match1.group(2)
        rest = match1.group(3)
        
        # Format date
        try:
            result['date'] = datetime.strptime(date_str, '%Y%m%d').strftime('%Y-%m-%d')
        except:
            result['date'] = date_str
            
        # Format time
        try:
            result['time'] = datetime.strptime(time_str, '%H%M%S').strftime('%H:%M:%S')
        except:
            result['time'] = time_str
            
    elif match2:
        date_str = match2.group(1)
        rest = match2.group(2)
        
        # Format date
        try:
            result['date'] = datetime.strptime(date_str, '%Y%m%d').strftime('%Y-%m-%d')
        except:
            result['date'] = date_str
    else:
        rest = name
    
    # Clean up the rest - remove trailing date patterns like -20251204 or -20251128_
    rest = re.sub(r'-\d{8}_?\d*$', '', rest)
    rest = re.sub(r'-\d{6}$', '', rest)
    
    # Replace underscores with spaces for pattern matching
    rest_spaced = rest.replace('_', ' ')
    
    # Set subject
    result['subject'] = rest_spaced.strip()
    
    # Try to extract meeting type, HR person, and VA name
    # Pattern: "Integration Team Check-in Louise x VA_Name"
    # Also handles: "Integration Team Check-in  Louise x VA_Name" (double space)
    
    checkin_patterns = [
        # "Integration Team Check-in  Louise x Joanne" (double space before name)
        r'^(Integration Team Check-in)\s+(\w+)\s+x\s+(.+)$',
        # "Check-in Louise x VA"
        r'^(Check-in)\s+(\w+)\s+x\s+(.+)$',
        # Any check-in pattern
        r'^(.+Check-in)\s+(\w+)\s+x\s+(.+)$',
        # "Integration Team Check-in Shey xAnn" (no space before x)
        r'^(Integration Team Check-in)\s+(\w+)\s*x\s*(\w+)$',
    ]
    
    for pattern in checkin_patterns:
        match = re.match(pattern, rest_spaced, re.IGNORECASE)
        if match:
            result['meeting_type'] = 'Integration Team Check-in'  # Standardize
            result['hr_person'] = match.group(2).strip()
            result['va_name'] = match.group(3).strip()
            break
    
    # Handle edge case: "Integration Team Check-in  Louise x Joanne-202512" 
    # where the trailing date wasn't fully removed
    if result['va_name']:
        result['va_name'] = re.sub(r'-\d+$', '', result['va_name']).strip()
    
    # If no check-in pattern matched, try interview pattern
    if not result['meeting_type']:
        interview_match = re.match(r'^(Interview)[\s\-:]+(.+)$', rest_spaced, re.IGNORECASE)
        if interview_match:
            result['meeting_type'] = 'Interview'
            result['va_name'] = interview_match.group(2).strip()
    
    # If still no meeting type, use the subject as meeting type
    if not result['meeting_type']:
        result['meeting_type'] = rest_spaced.strip()
    
    return result


def fix_excel_metadata():
    """Fix metadata in the Excel file"""
    # Find the latest Excel
    excel_path = OUTPUT_DIR / 'meeting_transcripts_latest_analyzed.xlsx'
    if not excel_path.exists():
        print(f"❌ Excel file not found: {excel_path}")
        return
    
    # Load Excel
    df = pd.read_excel(excel_path)
    print(f"✅ Loaded Excel: {excel_path.name}")
    print(f"   Total rows: {len(df)}")
    
    # Ensure columns exist
    required_cols = ['Meeting Type', 'HR Person', 'VA Name']
    for col in required_cols:
        if col not in df.columns:
            df[col] = None
    
    # Count rows that need fixing
    needs_fix = df[
        (df['Subject'].isna() | (df['Subject'] == '')) |
        (df['Time'].isna() | (df['Time'] == ''))
    ]
    print(f"📊 Rows needing metadata fix: {len(needs_fix)}")
    
    if len(needs_fix) == 0:
        print("✅ All rows have complete metadata!")
        return
    
    # Fix each row
    fixed_count = 0
    for idx, row in df.iterrows():
        transcript_file = row.get('Transcript File', '')
        if not transcript_file or pd.isna(transcript_file):
            continue
        
        # Parse filename
        metadata = parse_transcript_filename(transcript_file)
        
        # Update missing fields
        updated = False
        
        # Date
        if pd.isna(row.get('Date')) or row.get('Date') == '':
            if metadata['date']:
                df.at[idx, 'Date'] = metadata['date']
                updated = True
        
        # Time
        if pd.isna(row.get('Time')) or row.get('Time') == '':
            if metadata['time']:
                df.at[idx, 'Time'] = metadata['time']
                updated = True
        
        # Subject
        if pd.isna(row.get('Subject')) or row.get('Subject') == '':
            if metadata['subject']:
                df.at[idx, 'Subject'] = metadata['subject']
                updated = True
        
        # Meeting Type
        if pd.isna(row.get('Meeting Type')) or row.get('Meeting Type') == '':
            if metadata['meeting_type']:
                df.at[idx, 'Meeting Type'] = metadata['meeting_type']
                updated = True
        
        # HR Person (Organizer for check-ins)
        if pd.isna(row.get('HR Person')) or row.get('HR Person') == '':
            if metadata['hr_person']:
                df.at[idx, 'HR Person'] = metadata['hr_person']
                updated = True
        # Also update Organizer if HR Person is found
        if pd.isna(row.get('Organizer')) or row.get('Organizer') == '':
            if metadata['hr_person']:
                df.at[idx, 'Organizer'] = metadata['hr_person']
                updated = True
        
        # VA Name
        if pd.isna(row.get('VA Name')) or row.get('VA Name') == '':
            if metadata['va_name']:
                df.at[idx, 'VA Name'] = metadata['va_name']
                updated = True
        
        if updated:
            fixed_count += 1
    
    print(f"✅ Fixed {fixed_count} rows")
    
    # Reorder columns for better readability
    priority_cols = [
        'Date', 'Time', 'Meeting Type', 'HR Person', 'VA Name', 'Subject',
        'Transcript File', 'Sentiment Score', 'Churn Risk', 'Opportunity Score',
        'Execution Reliability', 'Operational Complexity', 'Events', 'Summary',
        'Key Concerns', 'Key Positives', 'Action Items'
    ]
    
    # Build final column order
    final_cols = []
    for col in priority_cols:
        if col in df.columns:
            final_cols.append(col)
    
    # Add remaining columns
    for col in df.columns:
        if col not in final_cols:
            final_cols.append(col)
    
    df = df[final_cols]
    
    # Save
    output_path = OUTPUT_DIR / f'meeting_transcripts_master_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    df.to_excel(output_path, index=False)
    print(f"💾 Saved to: {output_path}")
    
    # Also update latest
    df.to_excel(excel_path, index=False)
    print(f"💾 Updated: {excel_path.name}")
    
    # Show sample of fixed data
    print("\n📋 Sample of fixed Louise data:")
    louise = df[df['Transcript File'].str.contains('Louise', case=False, na=False)]
    if len(louise) > 0:
        sample = louise[['Date', 'Time', 'Meeting Type', 'HR Person', 'VA Name', 'Subject']].head(5)
        print(sample.to_string())


if __name__ == '__main__':
    fix_excel_metadata()
