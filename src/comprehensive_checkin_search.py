"""
Comprehensive Check-in Meeting Search
=====================================
Search for ALL check-in meetings from Shey, Louise, and HR@our-assistants.com
since November 1, 2025 from multiple sources:
1. Graph API - Calendar events
2. Graph API - Transcripts
3. Local recordings folder
4. Local transcripts folder
5. Current Excel file
"""

import os
import sys
import json
import requests
from datetime import datetime, timedelta
from pathlib import Path
from azure.identity import ClientSecretCredential
from dotenv import load_dotenv
import pandas as pd

load_dotenv()

# Configuration
TENANT_ID = os.getenv('AZURE_TENANT_ID')
CLIENT_ID = os.getenv('AZURE_CLIENT_ID')
CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET')

# HR User ID
HR_USER_ID = '81835016-79d5-4a15-91b1-c104e2cd9adb'

# Paths
OUTPUT_DIR = Path('output')
RECORDINGS_DIR = Path('recordings')
TRANSCRIPTS_DIR = Path('transcripts')

# Date range
START_DATE = '2025-11-01'
END_DATE = '2026-01-02'


class CheckinMeetingSearcher:
    def __init__(self):
        self.credential = ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
        self.token = None
        self.headers = None
        self.all_meetings = []
        
    def authenticate(self):
        """Get access token"""
        print("Authenticating with Microsoft Graph API...")
        self.token = self.credential.get_token('https://graph.microsoft.com/.default').token
        self.headers = {
            'Authorization': f'Bearer {self.token}',
            'Content-Type': 'application/json'
        }
        print("✅ Authenticated")
    
    def search_calendar_events(self):
        """Search HR calendar for check-in meetings"""
        print("\n" + "="*70)
        print("1. SEARCHING CALENDAR EVENTS")
        print("="*70)
        
        # Search for check-in meetings in calendar
        filter_query = f"start/dateTime ge '{START_DATE}T00:00:00Z' and start/dateTime le '{END_DATE}T00:00:00Z'"
        url = f"https://graph.microsoft.com/v1.0/users/{HR_USER_ID}/calendar/events"
        params = {
            '$filter': filter_query,
            '$select': 'id,subject,start,end,organizer,attendees,isOnlineMeeting,onlineMeetingUrl',
            '$top': 500,
            '$orderby': 'start/dateTime desc'
        }
        
        calendar_meetings = []
        
        while url:
            try:
                resp = requests.get(url, headers=self.headers, params=params if '?' not in url else None, timeout=30)
                if resp.status_code == 200:
                    data = resp.json()
                    events = data.get('value', [])
                    
                    for event in events:
                        subject = event.get('subject', '')
                        # Filter for check-in meetings with Shey or Louise
                        subject_lower = subject.lower()
                        if ('check-in' in subject_lower or 'checkin' in subject_lower) and \
                           ('shey' in subject_lower or 'louise' in subject_lower or 'integration' in subject_lower):
                            
                            start = event.get('start', {}).get('dateTime', '')[:10]
                            calendar_meetings.append({
                                'subject': subject,
                                'date': start,
                                'source': 'Calendar',
                                'hr_person': 'Shey' if 'shey' in subject_lower else ('Louise' if 'louise' in subject_lower else 'HR'),
                                'event_id': event.get('id', '')
                            })
                    
                    url = data.get('@odata.nextLink')
                    params = None
                else:
                    print(f"  Error: {resp.status_code}")
                    break
            except Exception as e:
                print(f"  Error: {e}")
                break
        
        print(f"  Found {len(calendar_meetings)} check-in meetings in calendar")
        return calendar_meetings
    
    def search_local_recordings(self):
        """Search local recordings folder"""
        print("\n" + "="*70)
        print("2. SEARCHING LOCAL RECORDINGS")
        print("="*70)
        
        recordings = []
        start_date = datetime.strptime(START_DATE, '%Y-%m-%d')
        
        for item in RECORDINGS_DIR.iterdir():
            name = item.name
            
            # Check if it's a check-in meeting
            name_lower = name.lower()
            if not ('check-in' in name_lower or 'checkin' in name_lower):
                continue
            if not ('shey' in name_lower or 'louise' in name_lower):
                continue
            
            # Extract date
            import re
            match = re.match(r'(\d{8})_(\d{6})_(.+)', name)
            if match:
                date_str = match.group(1)
                try:
                    date = datetime.strptime(date_str, '%Y%m%d')
                    if date >= start_date:
                        subject = match.group(3).replace('.mp4', '').replace('.wav', '')
                        
                        # Check for VTT
                        has_vtt = any(TRANSCRIPTS_DIR.glob(f"{date_str}_{match.group(2)}*.vtt"))
                        
                        recordings.append({
                            'subject': subject,
                            'date': date.strftime('%Y-%m-%d'),
                            'source': 'Recording',
                            'hr_person': 'Shey' if 'shey' in name_lower else 'Louise',
                            'has_transcript': has_vtt,
                            'filename': name
                        })
                except:
                    pass
        
        print(f"  Found {len(recordings)} check-in recordings since Nov 1")
        return recordings
    
    def search_local_transcripts(self):
        """Search local transcripts folder"""
        print("\n" + "="*70)
        print("3. SEARCHING LOCAL TRANSCRIPTS")
        print("="*70)
        
        transcripts = []
        start_date = datetime.strptime(START_DATE, '%Y-%m-%d')
        
        for vtt in TRANSCRIPTS_DIR.glob('*.vtt'):
            name = vtt.name
            name_lower = name.lower()
            
            # Check if it's a check-in meeting
            if not ('check-in' in name_lower or 'checkin' in name_lower):
                continue
            if not ('shey' in name_lower or 'louise' in name_lower):
                continue
            
            # Extract date
            import re
            match = re.match(r'(\d{8})_(\d{6})_(.+)\.vtt', name)
            if match:
                date_str = match.group(1)
                try:
                    date = datetime.strptime(date_str, '%Y%m%d')
                    if date >= start_date:
                        subject = match.group(3)
                        transcripts.append({
                            'subject': subject,
                            'date': date.strftime('%Y-%m-%d'),
                            'source': 'Transcript',
                            'hr_person': 'Shey' if 'shey' in name_lower else 'Louise',
                            'filename': name
                        })
                except:
                    pass
            elif name.startswith('unknown_'):
                # Unknown date transcripts
                subject = name.replace('unknown_', '').replace('.vtt', '')
                if ('check-in' in name_lower or 'checkin' in name_lower) and \
                   ('shey' in name_lower or 'louise' in name_lower):
                    transcripts.append({
                        'subject': subject,
                        'date': 'Unknown',
                        'source': 'Transcript',
                        'hr_person': 'Shey' if 'shey' in name_lower else 'Louise',
                        'filename': name
                    })
        
        print(f"  Found {len(transcripts)} check-in transcripts since Nov 1")
        return transcripts
    
    def search_excel(self):
        """Search current Excel file"""
        print("\n" + "="*70)
        print("4. SEARCHING EXCEL FILE")
        print("="*70)
        
        # Find latest Excel
        excel_files = list(OUTPUT_DIR.glob('*.xlsx'))
        if not excel_files:
            print("  No Excel files found")
            return []
        
        latest = max(excel_files, key=lambda x: x.stat().st_mtime)
        df = pd.read_excel(latest)
        print(f"  Using: {latest.name}")
        
        # Filter for check-ins
        excel_meetings = []
        
        for _, row in df.iterrows():
            meeting_type = str(row.get('Meeting Type', '')).lower()
            hr_person = row.get('HR Person', '')
            
            if 'check-in' in meeting_type and hr_person in ['Shey', 'Louise', 'HR']:
                date = row.get('Date', '')
                if date != 'Unknown':
                    try:
                        if isinstance(date, str):
                            date_obj = datetime.strptime(date, '%Y-%m-%d')
                        else:
                            date_obj = pd.to_datetime(date)
                        
                        if date_obj >= datetime.strptime(START_DATE, '%Y-%m-%d'):
                            excel_meetings.append({
                                'subject': row.get('Subject', ''),
                                'date': date_obj.strftime('%Y-%m-%d') if hasattr(date_obj, 'strftime') else str(date),
                                'source': 'Excel',
                                'hr_person': hr_person,
                                'va_name': row.get('VA Name', ''),
                                'has_analysis': pd.notna(row.get('Summary', ''))
                            })
                    except:
                        pass
        
        print(f"  Found {len(excel_meetings)} check-in meetings in Excel since Nov 1")
        return excel_meetings
    
    def consolidate_results(self, calendar, recordings, transcripts, excel):
        """Consolidate all results and find gaps"""
        print("\n" + "="*70)
        print("CONSOLIDATING RESULTS")
        print("="*70)
        
        # Create a master list keyed by (date, va_name)
        all_meetings = {}
        
        def extract_va(subject):
            """Extract VA name from subject"""
            import re
            subject_str = str(subject)
            match = re.search(r'(?:shey|louise)\s*x\s*([A-Za-z]+)', subject_str, re.IGNORECASE)
            if match:
                return match.group(1).strip()
            return subject_str[-20:] if len(subject_str) > 20 else subject_str
        
        # Add calendar events
        for m in calendar:
            va = extract_va(m['subject'])
            key = (m['date'], va.lower())
            if key not in all_meetings:
                all_meetings[key] = {
                    'date': m['date'],
                    'va_name': va,
                    'hr_person': m['hr_person'],
                    'subject': m['subject'],
                    'in_calendar': True,
                    'has_recording': False,
                    'has_transcript': False,
                    'in_excel': False,
                    'has_analysis': False
                }
            else:
                all_meetings[key]['in_calendar'] = True
        
        # Add recordings
        for m in recordings:
            va = extract_va(m['subject'])
            key = (m['date'], va.lower())
            if key not in all_meetings:
                all_meetings[key] = {
                    'date': m['date'],
                    'va_name': va,
                    'hr_person': m['hr_person'],
                    'subject': m['subject'],
                    'in_calendar': False,
                    'has_recording': True,
                    'has_transcript': m.get('has_transcript', False),
                    'in_excel': False,
                    'has_analysis': False
                }
            else:
                all_meetings[key]['has_recording'] = True
                all_meetings[key]['has_transcript'] = m.get('has_transcript', False)
        
        # Add transcripts
        for m in transcripts:
            va = extract_va(m['subject'])
            key = (m['date'], va.lower())
            if key not in all_meetings:
                all_meetings[key] = {
                    'date': m['date'],
                    'va_name': va,
                    'hr_person': m['hr_person'],
                    'subject': m['subject'],
                    'in_calendar': False,
                    'has_recording': False,
                    'has_transcript': True,
                    'in_excel': False,
                    'has_analysis': False
                }
            else:
                all_meetings[key]['has_transcript'] = True
        
        # Add Excel entries
        for m in excel:
            va = m.get('va_name', extract_va(m['subject']))
            key = (m['date'], str(va).lower())
            if key not in all_meetings:
                all_meetings[key] = {
                    'date': m['date'],
                    'va_name': va,
                    'hr_person': m['hr_person'],
                    'subject': m['subject'],
                    'in_calendar': False,
                    'has_recording': False,
                    'has_transcript': True,  # Must have transcript to be in Excel
                    'in_excel': True,
                    'has_analysis': m.get('has_analysis', False)
                }
            else:
                all_meetings[key]['in_excel'] = True
                all_meetings[key]['has_analysis'] = m.get('has_analysis', False)
        
        return list(all_meetings.values())
    
    def generate_report(self, all_meetings):
        """Generate comprehensive report"""
        print("\n" + "="*70)
        print("COMPREHENSIVE CHECK-IN MEETING REPORT")
        print(f"Period: {START_DATE} to {END_DATE}")
        print("="*70)
        
        # Sort by date
        all_meetings.sort(key=lambda x: (str(x['date']) if x['date'] != 'Unknown' else '9999-99-99', str(x['va_name'])))
        
        # Stats
        total = len(all_meetings)
        by_hr = {}
        with_transcript = 0
        with_analysis = 0
        in_excel = 0
        has_recording = 0
        
        for m in all_meetings:
            hr = m['hr_person']
            by_hr[hr] = by_hr.get(hr, 0) + 1
            if m['has_transcript']:
                with_transcript += 1
            if m['has_analysis']:
                with_analysis += 1
            if m['in_excel']:
                in_excel += 1
            if m['has_recording']:
                has_recording += 1
        
        print(f"\nTOTAL CHECK-IN MEETINGS: {total}")
        print(f"\nBy HR Person:")
        for hr, count in sorted(by_hr.items()):
            print(f"  {hr}: {count}")
        
        print(f"\nCoverage:")
        print(f"  Has Recording: {has_recording} ({100*has_recording/total:.0f}%)")
        print(f"  Has Transcript: {with_transcript} ({100*with_transcript/total:.0f}%)")
        print(f"  In Excel: {in_excel} ({100*in_excel/total:.0f}%)")
        print(f"  Has AI Analysis: {with_analysis} ({100*with_analysis/total:.0f}%)")
        
        # Gaps - meetings without transcripts
        gaps = [m for m in all_meetings if not m['has_transcript']]
        print(f"\n" + "="*70)
        print(f"GAPS - MEETINGS WITHOUT TRANSCRIPTS: {len(gaps)}")
        print("="*70)
        
        if gaps:
            print("\nDate         HR      VA Name              Subject")
            print("-"*70)
            for m in gaps[:30]:
                print(f"{m['date']:12} {m['hr_person']:7} {str(m['va_name'])[:20]:20} {m['subject'][:40]}")
            if len(gaps) > 30:
                print(f"... and {len(gaps)-30} more")
        
        # Missing from Excel
        not_in_excel = [m for m in all_meetings if m['has_transcript'] and not m['in_excel']]
        print(f"\n" + "="*70)
        print(f"HAS TRANSCRIPT BUT NOT IN EXCEL: {len(not_in_excel)}")
        print("="*70)
        
        if not_in_excel:
            for m in not_in_excel[:20]:
                print(f"  {m['date']} - {m['hr_person']} x {m['va_name']}")
        
        # Save to JSON
        output_path = OUTPUT_DIR / f'checkin_meetings_report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json'
        with open(output_path, 'w') as f:
            json.dump({
                'generated': datetime.now().isoformat(),
                'period': {'start': START_DATE, 'end': END_DATE},
                'summary': {
                    'total': total,
                    'by_hr': by_hr,
                    'has_recording': has_recording,
                    'has_transcript': with_transcript,
                    'in_excel': in_excel,
                    'has_analysis': with_analysis,
                    'gaps': len(gaps)
                },
                'meetings': all_meetings,
                'gaps': gaps
            }, f, indent=2, default=str)
        print(f"\n✅ Report saved to: {output_path.name}")
        
        return all_meetings


def main():
    searcher = CheckinMeetingSearcher()
    searcher.authenticate()
    
    # Search all sources
    calendar = searcher.search_calendar_events()
    recordings = searcher.search_local_recordings()
    transcripts = searcher.search_local_transcripts()
    excel = searcher.search_excel()
    
    # Consolidate
    all_meetings = searcher.consolidate_results(calendar, recordings, transcripts, excel)
    
    # Generate report
    searcher.generate_report(all_meetings)


if __name__ == '__main__':
    main()
