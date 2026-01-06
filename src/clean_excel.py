"""
Clean up Excel - remove duplicates and fix remaining metadata issues.
"""
import os
import re
from pathlib import Path
from datetime import datetime
import pandas as pd

OUTPUT_DIR = Path('output')


def clean_va_name(va_name):
    """Clean VA name by removing trailing date patterns"""
    if pd.isna(va_name):
        return va_name
    # Remove patterns like -20251204, -2025120, -202512
    cleaned = re.sub(r'-\d{4,8}_?$', '', str(va_name))
    return cleaned.strip()


def clean_excel():
    """Clean up the Excel file"""
    excel_path = OUTPUT_DIR / 'meeting_transcripts_latest_analyzed.xlsx'
    if not excel_path.exists():
        print(f"❌ Excel file not found: {excel_path}")
        return
    
    # Load Excel
    df = pd.read_excel(excel_path)
    print(f"✅ Loaded Excel: {excel_path.name}")
    print(f"   Total rows before cleanup: {len(df)}")
    
    # 1. Clean VA Names
    if 'VA Name' in df.columns:
        df['VA Name'] = df['VA Name'].apply(clean_va_name)
        print("✅ Cleaned VA Names")
    
    # 2. Remove duplicate/incomplete transcripts
    # Files that end with just "Louise.vtt" without VA name are incomplete
    incomplete_mask = df['Transcript File'].str.match(
        r'.*Integration Team Check-in\s+Louise\.vtt$', 
        case=False, 
        na=False
    )
    incomplete_count = incomplete_mask.sum()
    if incomplete_count > 0:
        df = df[~incomplete_mask]
        print(f"✅ Removed {incomplete_count} incomplete transcript entries")
    
    # 3. Also look for other duplicate patterns
    # Some files have both underscore and space versions
    # Keep the more complete ones
    
    # Create a key for deduplication based on date + time + participants
    def make_dedup_key(row):
        date = str(row.get('Date', ''))
        time = str(row.get('Time', ''))[:8] if pd.notna(row.get('Time')) else ''
        hr = str(row.get('HR Person', ''))
        va = str(row.get('VA Name', ''))
        return f"{date}_{time}_{hr}_{va}".lower()
    
    df['_dedup_key'] = df.apply(make_dedup_key, axis=1)
    
    # Sort to prefer entries with more complete data (non-null VA Name)
    df = df.sort_values(
        by=['_dedup_key', 'VA Name', 'HR Person'], 
        ascending=[True, False, False],  # False puts non-null first
        na_position='last'
    )
    
    # Drop duplicates keeping first (most complete)
    before_dedup = len(df)
    df = df.drop_duplicates(subset=['_dedup_key'], keep='first')
    df = df.drop(columns=['_dedup_key'])
    
    dupes_removed = before_dedup - len(df)
    if dupes_removed > 0:
        print(f"✅ Removed {dupes_removed} duplicate entries")
    
    print(f"   Total rows after cleanup: {len(df)}")
    
    # Save
    output_path = OUTPUT_DIR / f'meeting_transcripts_master_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    df.to_excel(output_path, index=False)
    print(f"💾 Saved to: {output_path}")
    
    # Also update latest
    df.to_excel(excel_path, index=False)
    print(f"💾 Updated: {excel_path.name}")
    
    # Show Louise summary
    print("\n📋 Louise Check-ins Summary:")
    louise = df[df['Transcript File'].str.contains('Louise', case=False, na=False)]
    print(f"   Total Louise meetings: {len(louise)}")
    print(f"\n   Sample data:")
    print(louise[['Date', 'Time', 'Meeting Type', 'HR Person', 'VA Name', 'Sentiment Score']].head(10).to_string())


if __name__ == '__main__':
    clean_excel()
