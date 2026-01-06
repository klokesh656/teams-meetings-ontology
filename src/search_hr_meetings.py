"""
Comprehensive search for HR, Shey, and Louise meetings in last 2 months.
Checks: Graph API transcripts, calendar events, communications, and recordings.
"""

import os
import re
import json
import time
import requests
from datetime import datetime, timedelta
from pathlib import Path
from dotenv import load_dotenv
from azure.identity import ClientSecretCredential

load_dotenv()

# Configuration
TENANT_ID = os.getenv('AZURE_TENANT_ID', '187b2af6-1bfb-490a-85dd-b720fe3d31bc')
CLIENT_ID = os.getenv('AZURE_CLIENT_ID')
CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET')

# HR User ID
HR_USER_ID = '81835016-79d5-4a15-91b1-c104e2cd9adb'

# Target users to search for
TARGET_NAMES = ['shey', 'louise', 'hr', 'check-in', 'checkin', 'check in', 'integration']

TRANSCRIPTS_DIR = Path('transcripts')
RECORDINGS_DIR = Path('recordings')
OUTPUT_DIR = Path('output')

TRANSCRIPTS_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


def get_auth_token():
    """Get authentication token for Graph API"""
    credential = ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
    token = credential.get_token('https://graph.microsoft.com/.default').token
    return token


def get_existing_transcripts():
    """Get list of existing transcript files"""
    existing = {}
    for f in TRANSCRIPTS_DIR.glob('*.vtt'):
        existing[f.stem.lower()] = f.name
    return existing


def matches_target(subject):
    """Check if meeting subject matches our target criteria"""
    if not subject:
        return False
    subject_lower = subject.lower()
    for target in TARGET_NAMES:
        if target in subject_lower:
            return True
    return False


def search_all_transcripts(headers, days_back=60):
    """Search all transcripts from Graph API"""
    print("\n" + "="*60)
    print("SEARCHING GRAPH API FOR ALL TRANSCRIPTS")
    print("="*60)
    
    all_transcripts = []
    cutoff_date = datetime.utcnow() - timedelta(days=days_back)
    
    url = f"https://graph.microsoft.com/beta/users/{HR_USER_ID}/onlineMeetings/getAllTranscripts(meetingOrganizerUserId='{HR_USER_ID}')"
    
    page_count = 0
    total_found = 0
    
    while url:
        try:
            resp = requests.get(url, headers=headers, timeout=60)
            if resp.status_code == 200:
                data = resp.json()
                transcripts = data.get('value', [])
                
                for t in transcripts:
                    created_date = t.get('createdDateTime', '')
                    if created_date:
                        try:
                            created = datetime.fromisoformat(created_date.replace('Z', '+00:00'))
                            if created.replace(tzinfo=None) >= cutoff_date:
                                t['user_id'] = HR_USER_ID
                                all_transcripts.append(t)
                                total_found += 1
                        except:
                            pass
                
                page_count += 1
                url = data.get('@odata.nextLink')
                
                if page_count % 5 == 0:
                    print(f"  📄 Processed {page_count} pages, found {total_found} transcripts in date range...")
                
                time.sleep(0.2)
            else:
                print(f"  ❌ Error: {resp.status_code}")
                break
        except requests.exceptions.Timeout:
            print(f"  ⚠️ Timeout, retrying...")
            time.sleep(2)
            continue
        except Exception as e:
            print(f"  ❌ Error: {str(e)[:50]}")
            break
    
    print(f"\n✅ Found {len(all_transcripts)} transcripts in last {days_back} days")
    return all_transcripts


def search_calendar_events(headers, days_back=60):
    """Search calendar events for meetings"""
    print("\n" + "="*60)
    print("SEARCHING CALENDAR EVENTS")
    print("="*60)
    
    start_date = (datetime.utcnow() - timedelta(days=days_back)).strftime('%Y-%m-%dT00:00:00Z')
    end_date = datetime.utcnow().strftime('%Y-%m-%dT23:59:59Z')
    
    url = f"https://graph.microsoft.com/v1.0/users/{HR_USER_ID}/calendar/calendarView"
    params = {
        'startDateTime': start_date,
        'endDateTime': end_date,
        '$top': 500,
        '$select': 'subject,start,end,isOnlineMeeting,onlineMeeting,organizer,attendees'
    }
    
    meetings = []
    try:
        resp = requests.get(url, headers=headers, params=params, timeout=60)
        if resp.status_code == 200:
            data = resp.json()
            events = data.get('value', [])
            
            for event in events:
                subject = event.get('subject', '')
                if matches_target(subject):
                    meetings.append({
                        'subject': subject,
                        'start': event.get('start', {}).get('dateTime', ''),
                        'end': event.get('end', {}).get('dateTime', ''),
                        'isOnlineMeeting': event.get('isOnlineMeeting', False),
                        'joinUrl': event.get('onlineMeeting', {}).get('joinUrl', '') if event.get('onlineMeeting') else ''
                    })
            
            print(f"✅ Found {len(meetings)} matching calendar events")
        else:
            print(f"  ❌ Calendar error: {resp.status_code} - {resp.text[:100]}")
    except Exception as e:
        print(f"  ❌ Error: {e}")
    
    return meetings


def search_online_meetings(headers, days_back=60):
    """Search online meetings directly"""
    print("\n" + "="*60)
    print("SEARCHING ONLINE MEETINGS")
    print("="*60)
    
    url = f"https://graph.microsoft.com/beta/users/{HR_USER_ID}/onlineMeetings"
    
    meetings = []
    page_count = 0
    
    while url:
        try:
            resp = requests.get(url, headers=headers, timeout=60)
            if resp.status_code == 200:
                data = resp.json()
                items = data.get('value', [])
                
                for meeting in items:
                    subject = meeting.get('subject', '')
                    if matches_target(subject):
                        meetings.append({
                            'id': meeting.get('id', ''),
                            'subject': subject,
                            'creationDateTime': meeting.get('creationDateTime', ''),
                            'startDateTime': meeting.get('startDateTime', ''),
                            'endDateTime': meeting.get('endDateTime', '')
                        })
                
                page_count += 1
                url = data.get('@odata.nextLink')
                
                if page_count % 10 == 0:
                    print(f"  📄 Processed {page_count} pages, found {len(meetings)} matching meetings...")
                
                time.sleep(0.2)
            else:
                print(f"  ❌ Error: {resp.status_code}")
                break
        except Exception as e:
            print(f"  ❌ Error: {str(e)[:50]}")
            break
    
    print(f"✅ Found {len(meetings)} matching online meetings")
    return meetings


def search_local_recordings():
    """Search local recordings folder"""
    print("\n" + "="*60)
    print("SEARCHING LOCAL RECORDINGS")
    print("="*60)
    
    recordings = []
    if RECORDINGS_DIR.exists():
        for item in RECORDINGS_DIR.iterdir():
            name = item.name.lower()
            for target in TARGET_NAMES:
                if target in name:
                    recordings.append({
                        'path': str(item),
                        'name': item.name,
                        'is_dir': item.is_dir()
                    })
                    break
    
    print(f"✅ Found {len(recordings)} matching recordings")
    return recordings


def get_meeting_details(headers, meeting_id):
    """Get meeting details including subject"""
    url = f"https://graph.microsoft.com/beta/users/{HR_USER_ID}/onlineMeetings/{meeting_id}"
    try:
        resp = requests.get(url, headers=headers, timeout=30)
        if resp.status_code == 200:
            return resp.json()
    except:
        pass
    return {}


def download_transcript(headers, meeting_id, transcript_id, meeting_subject, created_date):
    """Download a single transcript"""
    url = f"https://graph.microsoft.com/beta/users/{HR_USER_ID}/onlineMeetings/{meeting_id}/transcripts/{transcript_id}/content?$format=text/vtt"
    
    try:
        resp = requests.get(url, headers=headers, timeout=60)
        if resp.status_code == 200:
            safe_subject = re.sub(r'[<>:"/\\|?*]', '', meeting_subject or 'Unknown Meeting')[:50]
            date_str = created_date[:10].replace('-', '') if created_date else datetime.now().strftime('%Y%m%d')
            time_str = created_date[11:19].replace(':', '') if created_date and len(created_date) > 11 else ''
            filename = f"{date_str}_{time_str}_{safe_subject}.vtt"
            filepath = TRANSCRIPTS_DIR / filename
            
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(resp.text)
            
            return filepath, filename
        elif resp.status_code == 404:
            return None, "not_found"
    except Exception as e:
        pass
    
    return None, "error"


def main():
    print("="*70)
    print("COMPREHENSIVE HR/SHEY/LOUISE MEETING SEARCH")
    print("="*70)
    print(f"Date Range: Last 60 days (since {(datetime.now() - timedelta(days=60)).strftime('%Y-%m-%d')})")
    print(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Searching for: {', '.join(TARGET_NAMES)}")
    
    # Get existing transcripts
    existing = get_existing_transcripts()
    print(f"\n📁 Existing transcripts: {len(existing)}")
    
    # Authenticate
    print("\n🔗 Authenticating with Microsoft Graph...")
    token = get_auth_token()
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    print("✅ Authenticated")
    
    results = {
        'search_date': datetime.now().isoformat(),
        'date_range': 'Last 60 days',
        'transcripts': [],
        'calendar_events': [],
        'online_meetings': [],
        'local_recordings': [],
        'new_downloads': []
    }
    
    # 1. Search all transcripts
    all_transcripts = search_all_transcripts(headers, days_back=60)
    
    # 2. Search calendar events
    calendar_events = search_calendar_events(headers, days_back=60)
    results['calendar_events'] = calendar_events
    
    # 3. Search online meetings
    online_meetings = search_online_meetings(headers, days_back=60)
    results['online_meetings'] = online_meetings
    
    # 4. Search local recordings
    local_recordings = search_local_recordings()
    results['local_recordings'] = local_recordings
    
    # Process transcripts - download matching ones
    print("\n" + "="*60)
    print("DOWNLOADING MATCHING TRANSCRIPTS")
    print("="*60)
    
    new_downloads = []
    skipped = 0
    not_found = 0
    matching_count = 0
    
    for i, t in enumerate(all_transcripts):
        transcript_id = t.get('id', '')
        meeting_id = t.get('meetingId', '')
        created_date = t.get('createdDateTime', '')
        
        # Get meeting details to check subject
        meeting = get_meeting_details(headers, meeting_id)
        subject = meeting.get('subject', '')
        
        if not matches_target(subject):
            continue
        
        matching_count += 1
        
        # Check if already downloaded
        date_prefix = created_date[:10].replace('-', '') if created_date else ''
        already_have = False
        for ex_key in existing:
            if date_prefix and date_prefix in ex_key:
                safe_subj = re.sub(r'[<>:"/\\|?*]', '', subject or '')[:30].lower()
                if safe_subj and safe_subj[:15] in ex_key:
                    already_have = True
                    break
        
        if already_have:
            skipped += 1
            continue
        
        print(f"\n  [{matching_count}] {subject[:50]}...")
        print(f"      Date: {created_date[:10] if created_date else 'Unknown'}")
        
        # Download
        filepath, filename = download_transcript(headers, meeting_id, transcript_id, subject, created_date)
        
        if filepath:
            print(f"      ✅ Downloaded: {filename}")
            new_downloads.append({
                'filename': filename,
                'subject': subject,
                'date': created_date
            })
            existing[filename[:-4].lower()] = filename
        elif filename == "not_found":
            print(f"      ⏭️ Transcript expired/deleted")
            not_found += 1
        
        time.sleep(0.3)
    
    results['new_downloads'] = new_downloads
    
    # Summary
    print("\n" + "="*70)
    print("COMPREHENSIVE SEARCH SUMMARY")
    print("="*70)
    
    print(f"\n📊 GRAPH API RESULTS:")
    print(f"   Total transcripts in last 60 days: {len(all_transcripts)}")
    print(f"   Matching HR/Shey/Louise/Check-in: {matching_count}")
    print(f"   Already downloaded: {skipped}")
    print(f"   Expired/deleted: {not_found}")
    print(f"   🆕 Newly downloaded: {len(new_downloads)}")
    
    print(f"\n📅 CALENDAR EVENTS:")
    print(f"   Matching events found: {len(calendar_events)}")
    if calendar_events:
        print("   Recent events:")
        for e in calendar_events[:10]:
            print(f"     - {e['subject'][:50]} ({e['start'][:10] if e['start'] else 'Unknown'})")
    
    print(f"\n🎥 ONLINE MEETINGS:")
    print(f"   Matching meetings: {len(online_meetings)}")
    
    print(f"\n📂 LOCAL RECORDINGS:")
    print(f"   Matching recordings: {len(local_recordings)}")
    if local_recordings:
        print("   Recordings:")
        for r in local_recordings[:10]:
            print(f"     - {r['name'][:60]}")
    
    print(f"\n📁 TOTAL TRANSCRIPTS NOW: {len(existing)}")
    
    if new_downloads:
        print(f"\n✨ NEW TRANSCRIPTS DOWNLOADED:")
        for t in new_downloads:
            print(f"   - {t['subject'][:60]}")
            print(f"     Date: {t['date'][:10] if t['date'] else 'Unknown'}")
    
    # Save results
    with open(OUTPUT_DIR / 'meeting_search_results.json', 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=2, default=str)
    
    # Save with timestamp
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    with open(OUTPUT_DIR / f'meeting_search_results_{timestamp}.json', 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=2, default=str)
    
    print(f"\n💾 Results saved to: output/meeting_search_results_{timestamp}.json")
    
    return new_downloads


if __name__ == "__main__":
    new_transcripts = main()
    
    if new_transcripts:
        print("\n" + "="*60)
        print("NEXT STEP: Update Excel with new transcripts")
        print("="*60)
        print("Run: .venv\\Scripts\\python.exe src/update_excel_with_transcripts.py")
