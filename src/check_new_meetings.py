"""
Check for New Meetings with Auto-Recording and Transcription
=============================================================
Searches for new online meetings since the last check (January 2, 2026).
Focuses on meetings with auto-recording and transcription enabled.
"""

import os
import sys
import json
import requests
from datetime import datetime
from pathlib import Path
from azure.identity import ClientSecretCredential
from dotenv import load_dotenv

load_dotenv()

# Configuration
TENANT_ID = os.getenv('AZURE_TENANT_ID')
CLIENT_ID = os.getenv('AZURE_CLIENT_ID')
CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET')

# User IDs
HR_USER_ID = '81835016-79d5-4a15-91b1-c104e2cd9adb'  # hr@our-assistants.com

# Paths
OUTPUT_DIR = Path('output')
RECORDINGS_DIR = Path('recordings')
TRANSCRIPTS_DIR = Path('transcripts')

# Last check date
LAST_CHECK_DATE = '2026-01-02'


class NewMeetingChecker:
    def __init__(self):
        self.credential = ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
        self.token = None
        self.headers = None
        self.results = {
            'check_date': datetime.now().isoformat(),
            'last_check': LAST_CHECK_DATE,
            'new_calendar_meetings': [],
            'new_recordings': [],
            'new_transcripts': [],
            'summary': {}
        }
        
    def authenticate(self):
        print("🔐 Authenticating with Microsoft Graph API...")
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
    
    def search_new_calendar_meetings(self):
        """Search for new calendar meetings since last check"""
        print("="*70)
        print("1. SEARCHING FOR NEW CALENDAR MEETINGS")
        print(f"   (since {LAST_CHECK_DATE})")
        print("="*70)
        
        filter_query = f"start/dateTime ge '{LAST_CHECK_DATE}T00:00:00Z'"
        url = f"https://graph.microsoft.com/v1.0/users/{HR_USER_ID}/calendar/events"
        params = {
            '$filter': filter_query,
            '$select': 'id,subject,start,end,organizer,isOnlineMeeting,onlineMeeting,attendees,createdDateTime,lastModifiedDateTime',
            '$top': 500,
            '$orderby': 'start/dateTime desc'
        }
        
        new_meetings = []
        page = 1
        
        while url:
            try:
                resp = requests.get(url, headers=self.headers, params=params if '?' not in url else None, timeout=30)
                if resp.status_code == 200:
                    data = resp.json()
                    events = data.get('value', [])
                    print(f"   Page {page}: {len(events)} events")
                    
                    for event in events:
                        subject = event.get('subject', '')
                        start = event.get('start', {}).get('dateTime', '')[:10]
                        start_time = event.get('start', {}).get('dateTime', '')[11:19]
                        online_meeting = event.get('onlineMeeting', {})
                        join_url = online_meeting.get('joinUrl', '') if online_meeting else ''
                        
                        # Get attendees
                        attendees = event.get('attendees', [])
                        attendee_names = [a.get('emailAddress', {}).get('name', '') for a in attendees]
                        
                        # Filter for check-in meetings only
                        if self.is_checkin_meeting(subject):
                            new_meetings.append({
                                'subject': subject,
                                'date': start,
                                'time': start_time,
                                'event_id': event.get('id', ''),
                                'is_online': event.get('isOnlineMeeting', False),
                                'join_url': join_url,
                                'attendees': attendee_names,
                                'created': event.get('createdDateTime', ''),
                                'modified': event.get('lastModifiedDateTime', ''),
                                'source': 'calendar'
                            })
                    
                    url = data.get('@odata.nextLink')
                    params = None
                    page += 1
                else:
                    print(f"   ❌ Error: {resp.status_code} - {resp.text[:200]}")
                    break
            except Exception as e:
                print(f"   ❌ Exception: {e}")
                break
        
        self.results['new_calendar_meetings'] = new_meetings
        
        # Group by date
        by_date = {}
        for m in new_meetings:
            date = m['date']
            if date not in by_date:
                by_date[date] = []
            by_date[date].append(m['subject'])
        
        print(f"\n   ✅ Found {len(new_meetings)} check-in meetings since {LAST_CHECK_DATE}")
        if by_date:
            print("\n   📅 Meetings by date:")
            for date in sorted(by_date.keys(), reverse=True)[:15]:
                print(f"      {date}: {len(by_date[date])} meetings")
        
        return new_meetings
    
    def search_new_recordings(self):
        """Search OneDrive for new recordings"""
        print("\n" + "="*70)
        print("2. SEARCHING FOR NEW RECORDINGS IN ONEDRIVE")
        print("="*70)
        
        recordings = []
        
        # Search in HR user's OneDrive recordings folder
        url = f"https://graph.microsoft.com/v1.0/users/{HR_USER_ID}/drive/root:/Recordings:/children"
        params = {'$top': 500, '$select': 'id,name,createdDateTime,lastModifiedDateTime,size,webUrl'}
        
        print("   Searching HR's OneDrive Recordings folder...")
        page = 1
        
        while url:
            try:
                resp = requests.get(url, headers=self.headers, params=params if '?' not in url else None, timeout=30)
                if resp.status_code == 200:
                    data = resp.json()
                    items = data.get('value', [])
                    print(f"   Page {page}: {len(items)} items")
                    
                    for item in items:
                        name = item.get('name', '')
                        created = item.get('createdDateTime', '')
                        modified = item.get('lastModifiedDateTime', '')
                        
                        # Check if created after last check
                        if created >= f"{LAST_CHECK_DATE}T00:00:00":
                            if self.is_checkin_meeting(name):
                                recordings.append({
                                    'name': name,
                                    'created': created,
                                    'modified': modified,
                                    'size': item.get('size', 0),
                                    'web_url': item.get('webUrl', ''),
                                    'id': item.get('id', ''),
                                    'source': 'onedrive'
                                })
                    
                    url = data.get('@odata.nextLink')
                    params = None
                    page += 1
                elif resp.status_code == 404:
                    print("   ℹ️  Recordings folder not found")
                    break
                else:
                    print(f"   ❌ Error: {resp.status_code}")
                    break
            except Exception as e:
                print(f"   ❌ Exception: {e}")
                break
        
        self.results['new_recordings'] = recordings
        print(f"\n   ✅ Found {len(recordings)} new check-in recordings since {LAST_CHECK_DATE}")
        
        if recordings:
            print("\n   📹 New recordings:")
            for r in recordings[:10]:
                print(f"      - {r['name'][:60]}... ({r['created'][:10]})")
            if len(recordings) > 10:
                print(f"      ... and {len(recordings) - 10} more")
        
        return recordings
    
    def search_new_transcripts(self):
        """Search for new transcripts in OneDrive"""
        print("\n" + "="*70)
        print("3. SEARCHING FOR NEW TRANSCRIPTS (AUTO-TRANSCRIPTION)")
        print("="*70)
        
        transcripts = []
        
        # Search in HR user's OneDrive for transcripts
        url = f"https://graph.microsoft.com/v1.0/users/{HR_USER_ID}/drive/root:/Recordings:/children"
        params = {'$top': 500}
        
        print("   Checking recording folders for transcripts...")
        
        try:
            resp = requests.get(url, headers=self.headers, params=params, timeout=30)
            if resp.status_code == 200:
                data = resp.json()
                folders = [item for item in data.get('value', []) if item.get('folder')]
                
                print(f"   Found {len(folders)} recording folders to check")
                
                for folder in folders:
                    folder_name = folder.get('name', '')
                    folder_id = folder.get('id', '')
                    
                    # Check if this is a new folder (created after last check)
                    created = folder.get('createdDateTime', '')
                    if created < f"{LAST_CHECK_DATE}T00:00:00":
                        continue
                    
                    # Search for transcript files in folder
                    folder_url = f"https://graph.microsoft.com/v1.0/users/{HR_USER_ID}/drive/items/{folder_id}/children"
                    try:
                        folder_resp = requests.get(folder_url, headers=self.headers, timeout=30)
                        if folder_resp.status_code == 200:
                            folder_items = folder_resp.json().get('value', [])
                            for item in folder_items:
                                name = item.get('name', '')
                                # Look for transcript files (.vtt, .docx, .txt with transcript in name)
                                if any(ext in name.lower() for ext in ['.vtt', 'transcript', '.docx']):
                                    transcripts.append({
                                        'name': name,
                                        'folder': folder_name,
                                        'created': item.get('createdDateTime', ''),
                                        'id': item.get('id', ''),
                                        'size': item.get('size', 0),
                                        'source': 'onedrive_transcript'
                                    })
                    except Exception as e:
                        continue
                        
            elif resp.status_code == 404:
                print("   ℹ️  Recordings folder not found")
        except Exception as e:
            print(f"   ❌ Exception: {e}")
        
        self.results['new_transcripts'] = transcripts
        print(f"\n   ✅ Found {len(transcripts)} new transcripts since {LAST_CHECK_DATE}")
        
        if transcripts:
            print("\n   📝 New transcripts:")
            for t in transcripts[:10]:
                print(f"      - {t['name'][:50]}... (in {t['folder'][:30]})")
            if len(transcripts) > 10:
                print(f"      ... and {len(transcripts) - 10} more")
        
        return transcripts
    
    def check_local_files(self):
        """Check for new local recordings and transcripts"""
        print("\n" + "="*70)
        print("4. CHECKING LOCAL FILES")
        print("="*70)
        
        new_local_recordings = []
        new_local_transcripts = []
        
        # Parse last check date
        last_check = datetime.strptime(LAST_CHECK_DATE, '%Y-%m-%d')
        
        # Check local recordings
        print(f"\n   Checking {RECORDINGS_DIR}/...")
        if RECORDINGS_DIR.exists():
            for item in RECORDINGS_DIR.iterdir():
                if item.is_dir():
                    # Parse date from folder name (format: YYYYMMDD_...)
                    folder_name = item.name
                    try:
                        date_str = folder_name[:8]
                        folder_date = datetime.strptime(date_str, '%Y%m%d')
                        if folder_date >= last_check:
                            if self.is_checkin_meeting(folder_name):
                                new_local_recordings.append(folder_name)
                    except:
                        continue
        
        print(f"   Found {len(new_local_recordings)} new local recording folders")
        
        # Check local transcripts
        print(f"\n   Checking {TRANSCRIPTS_DIR}/...")
        if TRANSCRIPTS_DIR.exists():
            for item in TRANSCRIPTS_DIR.rglob('*'):
                if item.is_file() and item.suffix in ['.txt', '.vtt', '.docx', '.json']:
                    # Check modification time
                    mod_time = datetime.fromtimestamp(item.stat().st_mtime)
                    if mod_time >= last_check:
                        parent_name = item.parent.name if item.parent != TRANSCRIPTS_DIR else ''
                        if self.is_checkin_meeting(item.name) or self.is_checkin_meeting(parent_name):
                            new_local_transcripts.append({
                                'name': item.name,
                                'path': str(item),
                                'modified': mod_time.isoformat(),
                                'size': item.stat().st_size
                            })
        
        print(f"   Found {len(new_local_transcripts)} new local transcripts")
        
        self.results['new_local_recordings'] = new_local_recordings
        self.results['new_local_transcripts'] = new_local_transcripts
        
        return new_local_recordings, new_local_transcripts
    
    def search_online_meeting_transcripts(self):
        """Search for online meeting transcripts via Graph API"""
        print("\n" + "="*70)
        print("5. SEARCHING ONLINE MEETING TRANSCRIPTS (GRAPH API)")
        print("="*70)
        
        # First get online meetings
        print("   Fetching online meetings with transcripts...")
        
        # Try the onlineMeetings endpoint
        url = f"https://graph.microsoft.com/v1.0/users/{HR_USER_ID}/onlineMeetings"
        params = {'$top': 50}
        
        meeting_transcripts = []
        
        try:
            resp = requests.get(url, headers=self.headers, params=params, timeout=30)
            if resp.status_code == 200:
                meetings = resp.json().get('value', [])
                print(f"   Found {len(meetings)} online meetings")
                
                for meeting in meetings:
                    meeting_id = meeting.get('id', '')
                    subject = meeting.get('subject', '')
                    
                    if not self.is_checkin_meeting(subject):
                        continue
                    
                    # Get transcripts for this meeting
                    transcript_url = f"https://graph.microsoft.com/v1.0/users/{HR_USER_ID}/onlineMeetings/{meeting_id}/transcripts"
                    try:
                        transcript_resp = requests.get(transcript_url, headers=self.headers, timeout=30)
                        if transcript_resp.status_code == 200:
                            transcripts = transcript_resp.json().get('value', [])
                            if transcripts:
                                for t in transcripts:
                                    created = t.get('createdDateTime', '')
                                    if created >= f"{LAST_CHECK_DATE}T00:00:00":
                                        meeting_transcripts.append({
                                            'meeting_id': meeting_id,
                                            'subject': subject,
                                            'transcript_id': t.get('id', ''),
                                            'created': created,
                                            'source': 'graph_api'
                                        })
                    except:
                        continue
                        
            elif resp.status_code == 403:
                print("   ⚠️  Access denied - need OnlineMeetings.Read.All permission")
            else:
                print(f"   ℹ️  Status: {resp.status_code}")
        except Exception as e:
            print(f"   ❌ Exception: {e}")
        
        self.results['online_meeting_transcripts'] = meeting_transcripts
        print(f"\n   ✅ Found {len(meeting_transcripts)} meeting transcripts via Graph API")
        
        return meeting_transcripts
    
    def generate_summary(self):
        """Generate summary of findings"""
        print("\n" + "="*70)
        print("SUMMARY")
        print("="*70)
        
        summary = {
            'calendar_meetings': len(self.results.get('new_calendar_meetings', [])),
            'onedrive_recordings': len(self.results.get('new_recordings', [])),
            'onedrive_transcripts': len(self.results.get('new_transcripts', [])),
            'local_recordings': len(self.results.get('new_local_recordings', [])),
            'local_transcripts': len(self.results.get('new_local_transcripts', [])),
            'graph_transcripts': len(self.results.get('online_meeting_transcripts', []))
        }
        
        self.results['summary'] = summary
        
        print(f"\n   📅 Calendar meetings since {LAST_CHECK_DATE}: {summary['calendar_meetings']}")
        print(f"   📹 OneDrive recordings: {summary['onedrive_recordings']}")
        print(f"   📝 OneDrive transcripts: {summary['onedrive_transcripts']}")
        print(f"   💾 Local recordings: {summary['local_recordings']}")
        print(f"   📄 Local transcripts: {summary['local_transcripts']}")
        print(f"   🌐 Graph API transcripts: {summary['graph_transcripts']}")
        
        # Calculate what needs processing
        need_download = []
        need_transcription = []
        
        # Check which OneDrive recordings need to be downloaded
        local_recording_names = set(self.results.get('new_local_recordings', []))
        for rec in self.results.get('new_recordings', []):
            name = rec['name']
            # Check if already downloaded
            if not any(name.replace('.mp4', '') in lr for lr in local_recording_names):
                need_download.append(rec)
        
        # Check which local recordings need transcription
        local_transcript_paths = set(t['path'] for t in self.results.get('new_local_transcripts', []))
        for rec in self.results.get('new_local_recordings', []):
            # Check if transcript exists
            has_transcript = False
            for tp in local_transcript_paths:
                if rec in tp or rec[:15] in tp:
                    has_transcript = True
                    break
            if not has_transcript:
                need_transcription.append(rec)
        
        self.results['need_download'] = need_download
        self.results['need_transcription'] = need_transcription
        
        print(f"\n   🔽 Need to download: {len(need_download)} recordings")
        print(f"   🎤 Need transcription: {len(need_transcription)} recordings")
        
        # List meetings needing action
        if need_download:
            print("\n   📥 Recordings to download:")
            for r in need_download[:10]:
                print(f"      - {r['name'][:60]}")
            if len(need_download) > 10:
                print(f"      ... and {len(need_download) - 10} more")
        
        if need_transcription:
            print("\n   🎤 Recordings to transcribe:")
            for r in need_transcription[:10]:
                print(f"      - {r[:60]}")
            if len(need_transcription) > 10:
                print(f"      ... and {len(need_transcription) - 10} more")
    
    def save_results(self):
        """Save results to JSON file"""
        OUTPUT_DIR.mkdir(exist_ok=True)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_file = OUTPUT_DIR / f'new_meetings_check_{timestamp}.json'
        
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(self.results, f, indent=2, ensure_ascii=False, default=str)
        
        print(f"\n💾 Results saved to: {output_file}")
        
        # Also save latest
        latest_file = OUTPUT_DIR / 'new_meetings_check_latest.json'
        with open(latest_file, 'w', encoding='utf-8') as f:
            json.dump(self.results, f, indent=2, ensure_ascii=False, default=str)
        
        return output_file
    
    def run(self):
        """Run all checks"""
        print("\n" + "="*70)
        print("🔍 CHECKING FOR NEW MEETINGS WITH AUTO-RECORDING & TRANSCRIPTION")
        print(f"   Last check: {LAST_CHECK_DATE}")
        print(f"   Current date: {datetime.now().strftime('%Y-%m-%d')}")
        print("="*70)
        
        self.authenticate()
        
        # Run all searches
        self.search_new_calendar_meetings()
        self.search_new_recordings()
        self.search_new_transcripts()
        self.check_local_files()
        self.search_online_meeting_transcripts()
        
        # Generate summary
        self.generate_summary()
        
        # Save results
        output_file = self.save_results()
        
        return self.results


if __name__ == '__main__':
    checker = NewMeetingChecker()
    results = checker.run()
