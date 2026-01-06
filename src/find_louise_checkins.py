"""
Find ALL Louise Check-in Meetings
=================================
Search everywhere for Louise check-in meetings since November 1, 2025:
1. Calendar events
2. OneDrive recordings (all users)
3. Local recordings folder
4. Local transcripts folder
5. Graph API transcripts

Goal: Find any recordings that can be transcribed
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

HR_USER_ID = '81835016-79d5-4a15-91b1-c104e2cd9adb'

# Paths
OUTPUT_DIR = Path('output')
RECORDINGS_DIR = Path('recordings')
TRANSCRIPTS_DIR = Path('transcripts')

START_DATE = '2025-11-01'


class LouiseCheckInFinder:
    def __init__(self):
        self.credential = ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
        self.token = None
        self.headers = None
        self.louise_meetings = []
        
    def authenticate(self):
        print("Authenticating with Microsoft Graph API...")
        self.token = self.credential.get_token('https://graph.microsoft.com/.default').token
        self.headers = {
            'Authorization': f'Bearer {self.token}',
            'Content-Type': 'application/json'
        }
        print("✅ Authenticated\n")
    
    def search_calendar(self):
        """Search HR calendar for Louise check-in meetings"""
        print("="*70)
        print("1. SEARCHING CALENDAR FOR LOUISE CHECK-INS")
        print("="*70)
        
        filter_query = f"start/dateTime ge '{START_DATE}T00:00:00Z'"
        url = f"https://graph.microsoft.com/v1.0/users/{HR_USER_ID}/calendar/events"
        params = {
            '$filter': filter_query,
            '$select': 'id,subject,start,end,organizer,isOnlineMeeting,onlineMeeting',
            '$top': 500,
            '$orderby': 'start/dateTime desc'
        }
        
        louise_calendar = []
        
        while url:
            try:
                resp = requests.get(url, headers=self.headers, params=params if '?' not in url else None, timeout=30)
                if resp.status_code == 200:
                    data = resp.json()
                    events = data.get('value', [])
                    
                    for event in events:
                        subject = event.get('subject', '')
                        subject_lower = subject.lower()
                        
                        # Check for Louise check-in meetings
                        if 'louise' in subject_lower and ('check-in' in subject_lower or 'checkin' in subject_lower):
                            start = event.get('start', {}).get('dateTime', '')[:10]
                            online_meeting = event.get('onlineMeeting', {})
                            join_url = online_meeting.get('joinUrl', '') if online_meeting else ''
                            
                            louise_calendar.append({
                                'subject': subject,
                                'date': start,
                                'event_id': event.get('id', ''),
                                'is_online': event.get('isOnlineMeeting', False),
                                'join_url': join_url
                            })
                    
                    url = data.get('@odata.nextLink')
                    params = None
                else:
                    print(f"  Error: {resp.status_code}")
                    break
            except Exception as e:
                print(f"  Error: {e}")
                break
        
        print(f"  Found {len(louise_calendar)} Louise check-in meetings in calendar")
        
        # Show them
        for m in louise_calendar[:20]:
            print(f"    {m['date']} - {m['subject'][:50]}")
        if len(louise_calendar) > 20:
            print(f"    ... and {len(louise_calendar) - 20} more")
        
        return louise_calendar
    
    def search_onedrive_all_users(self):
        """Search OneDrive Recordings folder for all users"""
        print("\n" + "="*70)
        print("2. SEARCHING ONEDRIVE RECORDINGS (ALL USERS)")
        print("="*70)
        
        # First get all users
        users_url = "https://graph.microsoft.com/v1.0/users?$select=id,displayName,mail&$top=100"
        
        try:
            resp = requests.get(users_url, headers=self.headers, timeout=30)
            if resp.status_code != 200:
                print(f"  Error getting users: {resp.status_code}")
                return []
            
            users = resp.json().get('value', [])
            print(f"  Found {len(users)} users to search")
        except Exception as e:
            print(f"  Error: {e}")
            return []
        
        louise_recordings = []
        start_date = datetime.strptime(START_DATE, '%Y-%m-%d')
        
        for user in users:
            user_id = user['id']
            user_name = user.get('displayName', 'Unknown')
            
            # Search Recordings folder
            url = f"https://graph.microsoft.com/v1.0/users/{user_id}/drive/root:/Recordings:/children?$top=500"
            
            try:
                resp = requests.get(url, headers=self.headers, timeout=30)
                if resp.status_code == 200:
                    items = resp.json().get('value', [])
                    
                    for item in items:
                        name = item.get('name', '')
                        name_lower = name.lower()
                        
                        # Check for Louise check-in
                        if 'louise' in name_lower and ('check-in' in name_lower or 'checkin' in name_lower):
                            created = item.get('createdDateTime', '')
                            
                            # Check date
                            if created:
                                try:
                                    created_date = datetime.fromisoformat(created.replace('Z', '+00:00'))
                                    if created_date.replace(tzinfo=None) >= start_date:
                                        louise_recordings.append({
                                            'name': name,
                                            'date': created[:10],
                                            'user_id': user_id,
                                            'user_name': user_name,
                                            'file_id': item.get('id', ''),
                                            'size_mb': item.get('size', 0) / 1024 / 1024,
                                            'web_url': item.get('webUrl', ''),
                                            'download_url': item.get('@microsoft.graph.downloadUrl', '')
                                        })
                                except:
                                    pass
                elif resp.status_code != 404:
                    pass  # No Recordings folder
            except:
                pass
        
        print(f"  Found {len(louise_recordings)} Louise recordings in OneDrive")
        
        for r in louise_recordings:
            print(f"    {r['date']} - {r['name'][:50]}... ({r['size_mb']:.1f} MB)")
        
        return louise_recordings
    
    def search_local_recordings(self):
        """Search local recordings folder"""
        print("\n" + "="*70)
        print("3. SEARCHING LOCAL RECORDINGS FOLDER")
        print("="*70)
        
        louise_local = []
        start_date = datetime.strptime(START_DATE, '%Y-%m-%d')
        
        for item in RECORDINGS_DIR.iterdir():
            name = item.name
            name_lower = name.lower()
            
            if 'louise' not in name_lower:
                continue
            if 'check-in' not in name_lower and 'checkin' not in name_lower:
                continue
            
            # Extract date
            match = re.match(r'(\d{8})_(\d{6})_(.+)', name)
            if match:
                date_str = match.group(1)
                try:
                    date = datetime.strptime(date_str, '%Y%m%d')
                    if date >= start_date:
                        # Check for transcript
                        has_vtt = any(TRANSCRIPTS_DIR.glob(f"{date_str}_{match.group(2)}*.vtt"))
                        
                        # Get file size
                        if item.is_file():
                            size = item.stat().st_size / 1024 / 1024
                        else:
                            mp4s = list(item.glob('*.mp4'))
                            size = mp4s[0].stat().st_size / 1024 / 1024 if mp4s else 0
                        
                        louise_local.append({
                            'name': name,
                            'date': date.strftime('%Y-%m-%d'),
                            'has_transcript': has_vtt,
                            'size_mb': size,
                            'path': str(item)
                        })
                except:
                    pass
        
        print(f"  Found {len(louise_local)} Louise recordings locally")
        
        for r in louise_local:
            status = "✓ Has VTT" if r['has_transcript'] else "✗ No VTT"
            print(f"    {r['date']} - {r['name'][:45]}... ({r['size_mb']:.1f} MB) {status}")
        
        return louise_local
    
    def search_local_transcripts(self):
        """Search local transcripts folder"""
        print("\n" + "="*70)
        print("4. SEARCHING LOCAL TRANSCRIPTS FOLDER")
        print("="*70)
        
        louise_transcripts = []
        start_date = datetime.strptime(START_DATE, '%Y-%m-%d')
        
        for vtt in TRANSCRIPTS_DIR.glob('*.vtt'):
            name = vtt.name
            name_lower = name.lower()
            
            if 'louise' not in name_lower:
                continue
            if 'check-in' not in name_lower and 'checkin' not in name_lower:
                continue
            
            # Extract date
            match = re.match(r'(\d{8})_(\d{6})_(.+)\.vtt', name)
            if match:
                date_str = match.group(1)
                try:
                    date = datetime.strptime(date_str, '%Y%m%d')
                    if date >= start_date:
                        louise_transcripts.append({
                            'name': name,
                            'date': date.strftime('%Y-%m-%d'),
                            'subject': match.group(3)
                        })
                except:
                    pass
            elif name.startswith('unknown_') and 'louise' in name_lower:
                louise_transcripts.append({
                    'name': name,
                    'date': 'Unknown',
                    'subject': name.replace('unknown_', '').replace('.vtt', '')
                })
        
        print(f"  Found {len(louise_transcripts)} Louise transcripts locally")
        
        for t in louise_transcripts:
            print(f"    {t['date']} - {t['subject'][:50]}")
        
        return louise_transcripts
    
    def search_graph_transcripts(self):
        """Search Graph API for Louise transcripts"""
        print("\n" + "="*70)
        print("5. SEARCHING GRAPH API TRANSCRIPTS")
        print("="*70)
        
        # Get all transcripts from HR user
        url = f"https://graph.microsoft.com/beta/users/{HR_USER_ID}/onlineMeetings/getAllTranscripts(meetingOrganizerUserId='{HR_USER_ID}')"
        
        louise_api_transcripts = []
        start_date = datetime.strptime(START_DATE, '%Y-%m-%d')
        
        while url:
            try:
                resp = requests.get(url, headers=self.headers, timeout=30)
                if resp.status_code == 200:
                    data = resp.json()
                    transcripts = data.get('value', [])
                    
                    for t in transcripts:
                        # We need to check meeting details
                        meeting_id = t.get('meetingId', '')
                        created = t.get('createdDateTime', '')
                        
                        if created:
                            try:
                                created_date = datetime.fromisoformat(created.replace('Z', '+00:00'))
                                if created_date.replace(tzinfo=None) >= start_date:
                                    # Get meeting details to check subject
                                    meeting_url = f"https://graph.microsoft.com/beta/users/{HR_USER_ID}/onlineMeetings/{meeting_id}"
                                    meeting_resp = requests.get(meeting_url, headers=self.headers, timeout=10)
                                    
                                    if meeting_resp.status_code == 200:
                                        meeting = meeting_resp.json()
                                        subject = meeting.get('subject', '')
                                        
                                        if 'louise' in subject.lower() and ('check-in' in subject.lower() or 'checkin' in subject.lower()):
                                            louise_api_transcripts.append({
                                                'subject': subject,
                                                'date': created[:10],
                                                'transcript_id': t.get('id', ''),
                                                'meeting_id': meeting_id
                                            })
                            except:
                                pass
                    
                    url = data.get('@odata.nextLink')
                else:
                    print(f"  Error: {resp.status_code}")
                    break
            except Exception as e:
                print(f"  Error: {e}")
                break
        
        print(f"  Found {len(louise_api_transcripts)} Louise transcripts via Graph API")
        
        for t in louise_api_transcripts:
            print(f"    {t['date']} - {t['subject'][:50]}")
        
        return louise_api_transcripts
    
    def consolidate_and_find_gaps(self, calendar, onedrive, local_rec, local_trans, api_trans):
        """Find what can be transcribed"""
        print("\n" + "="*70)
        print("CONSOLIDATION & GAP ANALYSIS")
        print("="*70)
        
        # Create master list by date + VA name
        all_meetings = {}
        
        def extract_va(subject):
            match = re.search(r'louise\s*x\s*([A-Za-z]+)', str(subject), re.IGNORECASE)
            return match.group(1) if match else 'Unknown'
        
        # Add calendar events
        for m in calendar:
            va = extract_va(m['subject'])
            key = (m['date'], va.lower())
            if key not in all_meetings:
                all_meetings[key] = {
                    'date': m['date'],
                    'va_name': va,
                    'subject': m['subject'],
                    'in_calendar': True,
                    'has_onedrive_recording': False,
                    'has_local_recording': False,
                    'has_transcript': False,
                    'can_transcribe': False,
                    'recording_path': None,
                    'recording_size_mb': 0
                }
            else:
                all_meetings[key]['in_calendar'] = True
        
        # Add OneDrive recordings
        for r in onedrive:
            va = extract_va(r['name'])
            key = (r['date'], va.lower())
            if key not in all_meetings:
                all_meetings[key] = {
                    'date': r['date'],
                    'va_name': va,
                    'subject': r['name'],
                    'in_calendar': False,
                    'has_onedrive_recording': True,
                    'has_local_recording': False,
                    'has_transcript': False,
                    'can_transcribe': True,
                    'recording_path': r.get('web_url', ''),
                    'recording_size_mb': r['size_mb'],
                    'onedrive_info': r
                }
            else:
                all_meetings[key]['has_onedrive_recording'] = True
                all_meetings[key]['can_transcribe'] = True
                all_meetings[key]['recording_size_mb'] = r['size_mb']
                all_meetings[key]['onedrive_info'] = r
        
        # Add local recordings
        for r in local_rec:
            va = extract_va(r['name'])
            key = (r['date'], va.lower())
            if key not in all_meetings:
                all_meetings[key] = {
                    'date': r['date'],
                    'va_name': va,
                    'subject': r['name'],
                    'in_calendar': False,
                    'has_onedrive_recording': False,
                    'has_local_recording': True,
                    'has_transcript': r['has_transcript'],
                    'can_transcribe': not r['has_transcript'],
                    'recording_path': r['path'],
                    'recording_size_mb': r['size_mb']
                }
            else:
                all_meetings[key]['has_local_recording'] = True
                all_meetings[key]['has_transcript'] = r['has_transcript']
                all_meetings[key]['can_transcribe'] = not r['has_transcript']
                all_meetings[key]['recording_path'] = r['path']
                all_meetings[key]['recording_size_mb'] = r['size_mb']
        
        # Add local transcripts
        for t in local_trans:
            va = extract_va(t['subject'])
            key = (t['date'], va.lower())
            if key not in all_meetings:
                all_meetings[key] = {
                    'date': t['date'],
                    'va_name': va,
                    'subject': t['subject'],
                    'in_calendar': False,
                    'has_onedrive_recording': False,
                    'has_local_recording': False,
                    'has_transcript': True,
                    'can_transcribe': False,
                    'recording_path': None,
                    'recording_size_mb': 0
                }
            else:
                all_meetings[key]['has_transcript'] = True
                all_meetings[key]['can_transcribe'] = False
        
        # Convert to list and sort
        meetings_list = list(all_meetings.values())
        meetings_list.sort(key=lambda x: x['date'])
        
        # Stats
        total = len(meetings_list)
        with_transcript = sum(1 for m in meetings_list if m['has_transcript'])
        can_transcribe = [m for m in meetings_list if m['can_transcribe']]
        
        print(f"\nTotal Louise check-in meetings found: {total}")
        print(f"  Already have transcript: {with_transcript}")
        print(f"  Can be transcribed: {len(can_transcribe)}")
        
        if can_transcribe:
            print(f"\n" + "="*70)
            print("RECORDINGS THAT CAN BE TRANSCRIBED")
            print("="*70)
            
            total_size = sum(m['recording_size_mb'] for m in can_transcribe)
            
            for m in can_transcribe:
                source = "OneDrive" if m['has_onedrive_recording'] else "Local"
                print(f"\n  {m['date']} - Louise x {m['va_name']}")
                print(f"    Source: {source}")
                print(f"    Size: {m['recording_size_mb']:.1f} MB")
                if m['recording_path']:
                    print(f"    Path: {m['recording_path'][:60]}...")
            
            print(f"\n  Total to transcribe: {len(can_transcribe)} recordings")
            print(f"  Total size: {total_size:.1f} MB")
            est_hours = total_size / 60
            est_cost = est_hours * 0.98
            print(f"  Estimated audio: ~{est_hours:.1f} hours")
            print(f"  Estimated cost: ~${est_cost:.2f}")
        
        # Save report
        report = {
            'generated': datetime.now().isoformat(),
            'search_period': f'{START_DATE} to present',
            'summary': {
                'total_meetings': total,
                'with_transcript': with_transcript,
                'can_transcribe': len(can_transcribe)
            },
            'all_meetings': meetings_list,
            'to_transcribe': can_transcribe
        }
        
        output_path = OUTPUT_DIR / f'louise_checkins_report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json'
        with open(output_path, 'w') as f:
            json.dump(report, f, indent=2, default=str)
        print(f"\n✅ Report saved to: {output_path.name}")
        
        return can_transcribe


def main():
    finder = LouiseCheckInFinder()
    finder.authenticate()
    
    # Search all sources
    calendar = finder.search_calendar()
    onedrive = finder.search_onedrive_all_users()
    local_rec = finder.search_local_recordings()
    local_trans = finder.search_local_transcripts()
    api_trans = finder.search_graph_transcripts()
    
    # Find what can be transcribed
    to_transcribe = finder.consolidate_and_find_gaps(calendar, onedrive, local_rec, local_trans, api_trans)
    
    return to_transcribe


if __name__ == '__main__':
    main()
