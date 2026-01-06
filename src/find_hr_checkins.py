"""
Find ALL HR Check-in Meetings
=============================
Search everywhere for hr@our-assistants.com check-in meetings since November 1, 2025:
1. Calendar events
2. OneDrive recordings (all users)
3. Local recordings folder
4. Local transcripts folder
5. Graph API transcripts
"""

import os
import sys
import json
import requests
import re
from datetime import datetime
from pathlib import Path
from azure.identity import ClientSecretCredential
from dotenv import load_dotenv

load_dotenv()

# Configuration
TENANT_ID = os.getenv('AZURE_TENANT_ID')
CLIENT_ID = os.getenv('AZURE_CLIENT_ID')
CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET')

HR_USER_ID = '81835016-79d5-4a15-91b1-c104e2cd9adb'  # hr@our-assistants.com

# Paths
OUTPUT_DIR = Path('output')
RECORDINGS_DIR = Path('recordings')
TRANSCRIPTS_DIR = Path('transcripts')

START_DATE = '2025-11-01'


class HRCheckInFinder:
    def __init__(self):
        self.credential = ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
        self.token = None
        self.headers = None
        self.checkin_meetings = []
        
    def authenticate(self):
        print("Authenticating with Microsoft Graph API...")
        self.token = self.credential.get_token('https://graph.microsoft.com/.default').token
        self.headers = {
            'Authorization': f'Bearer {self.token}',
            'Content-Type': 'application/json'
        }
        print("✅ Authenticated\n")
    
    def is_checkin_meeting(self, subject):
        """Check if meeting subject indicates a check-in"""
        subject_lower = subject.lower()
        return 'check-in' in subject_lower or 'checkin' in subject_lower
    
    def search_calendar(self):
        """Search HR calendar for ALL check-in meetings"""
        print("="*70)
        print("1. SEARCHING CALENDAR FOR ALL CHECK-IN MEETINGS")
        print("="*70)
        
        filter_query = f"start/dateTime ge '{START_DATE}T00:00:00Z'"
        url = f"https://graph.microsoft.com/v1.0/users/{HR_USER_ID}/calendar/events"
        params = {
            '$filter': filter_query,
            '$select': 'id,subject,start,end,organizer,isOnlineMeeting,onlineMeeting,attendees',
            '$top': 500,
            '$orderby': 'start/dateTime desc'
        }
        
        calendar_meetings = []
        page = 1
        
        while url:
            try:
                resp = requests.get(url, headers=self.headers, params=params if '?' not in url else None, timeout=30)
                if resp.status_code == 200:
                    data = resp.json()
                    events = data.get('value', [])
                    print(f"  Page {page}: {len(events)} events")
                    
                    for event in events:
                        subject = event.get('subject', '')
                        
                        # Check for check-in meetings
                        if self.is_checkin_meeting(subject):
                            start = event.get('start', {}).get('dateTime', '')[:10]
                            start_time = event.get('start', {}).get('dateTime', '')[11:19]
                            online_meeting = event.get('onlineMeeting', {})
                            join_url = online_meeting.get('joinUrl', '') if online_meeting else ''
                            
                            # Get attendees
                            attendees = event.get('attendees', [])
                            attendee_names = [a.get('emailAddress', {}).get('name', '') for a in attendees]
                            
                            calendar_meetings.append({
                                'subject': subject,
                                'date': start,
                                'time': start_time,
                                'event_id': event.get('id', ''),
                                'is_online': event.get('isOnlineMeeting', False),
                                'join_url': join_url,
                                'attendees': attendee_names,
                                'source': 'calendar'
                            })
                    
                    url = data.get('@odata.nextLink')
                    params = None
                    page += 1
                else:
                    print(f"  ❌ Error: {resp.status_code}")
                    break
            except Exception as e:
                print(f"  ❌ Exception: {e}")
                break
        
        print(f"\n  ✅ Found {len(calendar_meetings)} check-in meetings in calendar")
        return calendar_meetings
    
    def search_onedrive_recordings(self):
        """Search OneDrive for ALL check-in recordings"""
        print("\n" + "="*70)
        print("2. SEARCHING ONEDRIVE RECORDINGS")
        print("="*70)
        
        recordings = []
        
        # Search in HR user's OneDrive recordings folder
        url = f"https://graph.microsoft.com/v1.0/users/{HR_USER_ID}/drive/root:/Recordings:/children"
        params = {'$top': 500}
        
        print("  Searching HR's OneDrive Recordings folder...")
        page = 1
        
        while url:
            try:
                resp = requests.get(url, headers=self.headers, params=params if '?' not in url else None, timeout=30)
                if resp.status_code == 200:
                    data = resp.json()
                    items = data.get('value', [])
                    print(f"    Page {page}: {len(items)} items")
                    
                    for item in items:
                        name = item.get('name', '')
                        if self.is_checkin_meeting(name):
                            # Parse date from folder/file name
                            date_match = re.search(r'(\d{4})(\d{2})(\d{2})', name)
                            date = f"{date_match.group(1)}-{date_match.group(2)}-{date_match.group(3)}" if date_match else ''
                            
                            # Filter by date
                            if date >= START_DATE:
                                recordings.append({
                                    'name': name,
                                    'date': date,
                                    'item_id': item.get('id', ''),
                                    'web_url': item.get('webUrl', ''),
                                    'size': item.get('size', 0),
                                    'is_folder': 'folder' in item,
                                    'source': 'onedrive_hr'
                                })
                    
                    url = data.get('@odata.nextLink')
                    params = None
                    page += 1
                elif resp.status_code == 404:
                    print("    Recordings folder not found")
                    break
                else:
                    print(f"    Error: {resp.status_code}")
                    break
            except Exception as e:
                print(f"    Exception: {e}")
                break
        
        # Also search using general drive search
        print("\n  Searching all OneDrive with Graph search...")
        search_url = f"https://graph.microsoft.com/v1.0/users/{HR_USER_ID}/drive/root/search(q='check-in')"
        
        try:
            resp = requests.get(search_url, headers=self.headers, timeout=60)
            if resp.status_code == 200:
                data = resp.json()
                items = data.get('value', [])
                print(f"    Found {len(items)} items matching 'check-in'")
                
                for item in items:
                    name = item.get('name', '')
                    # Skip if not a video/audio or folder
                    if not any(ext in name.lower() for ext in ['.mp4', '.m4a', '.webm']) and 'folder' not in item:
                        continue
                    
                    date_match = re.search(r'(\d{4})(\d{2})(\d{2})', name)
                    date = f"{date_match.group(1)}-{date_match.group(2)}-{date_match.group(3)}" if date_match else ''
                    
                    if date >= START_DATE:
                        recordings.append({
                            'name': name,
                            'date': date,
                            'item_id': item.get('id', ''),
                            'web_url': item.get('webUrl', ''),
                            'size': item.get('size', 0),
                            'is_folder': 'folder' in item,
                            'source': 'onedrive_search'
                        })
        except Exception as e:
            print(f"    Exception: {e}")
        
        # Deduplicate by item_id
        seen_ids = set()
        unique_recordings = []
        for rec in recordings:
            if rec['item_id'] not in seen_ids:
                seen_ids.add(rec['item_id'])
                unique_recordings.append(rec)
        
        print(f"\n  ✅ Found {len(unique_recordings)} unique check-in recordings in OneDrive")
        return unique_recordings
    
    def search_local_recordings(self):
        """Search local recordings folder for check-in meetings"""
        print("\n" + "="*70)
        print("3. SEARCHING LOCAL RECORDINGS FOLDER")
        print("="*70)
        
        local_recordings = []
        
        if not RECORDINGS_DIR.exists():
            print("  Recordings directory not found")
            return local_recordings
        
        # Search all subdirectories
        for item in RECORDINGS_DIR.rglob('*'):
            if item.is_file() and item.suffix.lower() in ['.mp4', '.m4a', '.webm']:
                name = item.name
                if self.is_checkin_meeting(name) or self.is_checkin_meeting(item.parent.name):
                    date_match = re.search(r'(\d{4})(\d{2})(\d{2})', name)
                    if not date_match:
                        date_match = re.search(r'(\d{4})(\d{2})(\d{2})', item.parent.name)
                    date = f"{date_match.group(1)}-{date_match.group(2)}-{date_match.group(3)}" if date_match else ''
                    
                    if date >= START_DATE:
                        local_recordings.append({
                            'name': name,
                            'path': str(item),
                            'date': date,
                            'size_mb': round(item.stat().st_size / (1024*1024), 1),
                            'source': 'local'
                        })
        
        print(f"  ✅ Found {len(local_recordings)} check-in recordings locally")
        return local_recordings
    
    def search_local_transcripts(self):
        """Search local transcripts folder for check-in meetings"""
        print("\n" + "="*70)
        print("4. SEARCHING LOCAL TRANSCRIPTS FOLDER")
        print("="*70)
        
        local_transcripts = []
        
        if not TRANSCRIPTS_DIR.exists():
            print("  Transcripts directory not found")
            return local_transcripts
        
        for item in TRANSCRIPTS_DIR.glob('*.vtt'):
            name = item.name
            if self.is_checkin_meeting(name):
                date_match = re.search(r'(\d{4})(\d{2})(\d{2})', name)
                date = f"{date_match.group(1)}-{date_match.group(2)}-{date_match.group(3)}" if date_match else ''
                
                if date >= START_DATE:
                    local_transcripts.append({
                        'name': name,
                        'path': str(item),
                        'date': date,
                        'source': 'local_transcript'
                    })
        
        print(f"  ✅ Found {len(local_transcripts)} check-in transcripts locally")
        return local_transcripts
    
    def search_graph_transcripts(self):
        """Search for transcripts via Graph API online meetings"""
        print("\n" + "="*70)
        print("5. SEARCHING GRAPH API FOR TRANSCRIPTS")
        print("="*70)
        
        transcripts = []
        
        # List online meetings for HR user
        url = f"https://graph.microsoft.com/v1.0/users/{HR_USER_ID}/onlineMeetings"
        params = {'$top': 100}
        
        print("  Fetching online meetings...")
        page = 1
        meetings_checked = 0
        
        while url and meetings_checked < 500:
            try:
                resp = requests.get(url, headers=self.headers, params=params if '?' not in url else None, timeout=30)
                if resp.status_code == 200:
                    data = resp.json()
                    meetings = data.get('value', [])
                    print(f"    Page {page}: {len(meetings)} meetings")
                    
                    for meeting in meetings:
                        subject = meeting.get('subject', '')
                        if not subject:
                            continue
                            
                        meetings_checked += 1
                        
                        if self.is_checkin_meeting(subject):
                            meeting_id = meeting.get('id', '')
                            creation_time = meeting.get('creationDateTime', '')[:10]
                            
                            if creation_time >= START_DATE:
                                # Check for transcripts
                                trans_url = f"https://graph.microsoft.com/v1.0/users/{HR_USER_ID}/onlineMeetings/{meeting_id}/transcripts"
                                try:
                                    trans_resp = requests.get(trans_url, headers=self.headers, timeout=30)
                                    if trans_resp.status_code == 200:
                                        trans_data = trans_resp.json()
                                        trans_list = trans_data.get('value', [])
                                        has_transcript = len(trans_list) > 0
                                    else:
                                        has_transcript = False
                                except:
                                    has_transcript = False
                                
                                transcripts.append({
                                    'subject': subject,
                                    'meeting_id': meeting_id,
                                    'date': creation_time,
                                    'has_transcript': has_transcript,
                                    'source': 'graph_online_meeting'
                                })
                    
                    url = data.get('@odata.nextLink')
                    params = None
                    page += 1
                else:
                    print(f"    Error: {resp.status_code}")
                    break
            except Exception as e:
                print(f"    Exception: {e}")
                break
        
        print(f"\n  ✅ Found {len(transcripts)} check-in meetings via Graph API")
        return transcripts
    
    def extract_hr_person(self, subject):
        """Extract HR person name from subject like 'Integration Team Check-in Louise x Irvy'"""
        patterns = [
            r'Check-in\s+(\w+)\s+x\s+\w+',  # "Check-in Louise x Irvy"
            r'Check-in\s+(\w+)\s+x',         # "Check-in Shey xAnn"
        ]
        for pattern in patterns:
            match = re.search(pattern, subject, re.IGNORECASE)
            if match:
                return match.group(1)
        return None
    
    def extract_va_name(self, subject):
        """Extract VA name from subject"""
        patterns = [
            r'x\s+([A-Za-z]+(?:\s+[A-Za-z]+)?)',  # "x Irvy" or "x Jon Jevi"
        ]
        for pattern in patterns:
            match = re.search(pattern, subject, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        return None
    
    def run(self):
        """Main search process"""
        print("="*70)
        print("HR CHECK-IN MEETING FINDER")
        print(f"Searching for: hr@our-assistants.com check-in meetings")
        print(f"Since: {START_DATE}")
        print("="*70 + "\n")
        
        self.authenticate()
        
        results = {
            'search_date': datetime.now().isoformat(),
            'start_date': START_DATE,
            'calendar_meetings': [],
            'onedrive_recordings': [],
            'local_recordings': [],
            'local_transcripts': [],
            'graph_transcripts': [],
            'summary': {}
        }
        
        # Run all searches
        results['calendar_meetings'] = self.search_calendar()
        results['onedrive_recordings'] = self.search_onedrive_recordings()
        results['local_recordings'] = self.search_local_recordings()
        results['local_transcripts'] = self.search_local_transcripts()
        results['graph_transcripts'] = self.search_graph_transcripts()
        
        # Analyze by HR person
        print("\n" + "="*70)
        print("SUMMARY BY HR PERSON")
        print("="*70)
        
        hr_counts = {}
        all_meetings = (
            results['calendar_meetings'] + 
            results['onedrive_recordings'] + 
            results['local_recordings'] +
            results['local_transcripts'] +
            results['graph_transcripts']
        )
        
        for meeting in all_meetings:
            subject = meeting.get('subject', meeting.get('name', ''))
            hr_person = self.extract_hr_person(subject)
            if hr_person:
                hr_counts[hr_person] = hr_counts.get(hr_person, 0) + 1
        
        print("\n  HR Person          | Count")
        print("  " + "-"*30)
        for hr, count in sorted(hr_counts.items(), key=lambda x: -x[1]):
            print(f"  {hr:<18} | {count}")
        
        # Summary
        results['summary'] = {
            'total_calendar': len(results['calendar_meetings']),
            'total_onedrive': len(results['onedrive_recordings']),
            'total_local_recordings': len(results['local_recordings']),
            'total_local_transcripts': len(results['local_transcripts']),
            'total_graph': len(results['graph_transcripts']),
            'hr_breakdown': hr_counts
        }
        
        # Print overall summary
        print("\n" + "="*70)
        print("OVERALL SUMMARY")
        print("="*70)
        print(f"\n  Calendar meetings:     {results['summary']['total_calendar']}")
        print(f"  OneDrive recordings:   {results['summary']['total_onedrive']}")
        print(f"  Local recordings:      {results['summary']['total_local_recordings']}")
        print(f"  Local transcripts:     {results['summary']['total_local_transcripts']}")
        print(f"  Graph API meetings:    {results['summary']['total_graph']}")
        
        # Find meetings that need transcription
        print("\n" + "="*70)
        print("TRANSCRIPTION STATUS")
        print("="*70)
        
        # Get existing transcripts
        existing_transcripts = set()
        for t in results['local_transcripts']:
            existing_transcripts.add(t['name'].replace('.vtt', ''))
        
        needs_transcription = []
        for rec in results['local_recordings']:
            base_name = Path(rec['path']).stem
            if base_name not in existing_transcripts:
                needs_transcription.append(rec)
        
        print(f"\n  Recordings needing transcription: {len(needs_transcription)}")
        if needs_transcription:
            for rec in needs_transcription[:20]:
                print(f"    - {rec['name']} ({rec['size_mb']} MB)")
        
        # Save results
        output_file = OUTPUT_DIR / f'hr_checkin_meetings_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json'
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        print(f"\n💾 Results saved to: {output_file}")
        
        return results


if __name__ == '__main__':
    finder = HRCheckInFinder()
    finder.run()
