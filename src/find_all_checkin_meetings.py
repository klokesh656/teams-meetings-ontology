"""
Comprehensive search for HR, Shey, and Louise meetings in the last 2 months.
Checks:
1. Microsoft Graph API - Online Meetings
2. Microsoft Graph API - Calendar Events
3. Microsoft Graph API - Call Records
4. Local recordings folder
5. Existing transcripts
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
from collections import defaultdict

load_dotenv()

# Configuration
TENANT_ID = os.getenv('AZURE_TENANT_ID', '187b2af6-1bfb-490a-85dd-b720fe3d31bc')
CLIENT_ID = os.getenv('AZURE_CLIENT_ID')
CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET')

# Target users to search
TARGET_USERS = {
    'hr': '81835016-79d5-4a15-91b1-c104e2cd9adb',  # HR account (known ID)
}

# User emails to try to resolve
USER_EMAILS = [
    'hr@our-assistants.com',
    'shey@our-assistants.com', 
    'louise@our-assistants.com',
    'kc@our-assistants.com',
]

# Keywords for check-in meetings
CHECKIN_KEYWORDS = [
    'check-in', 'checkin', 'check in',
    'integration team', 
    'daily', 'weekly',
    'catch up', 'catchup',
    'sync', '1:1', 'one on one',
    'status', 'update',
    'shey x', 'louise x', 'kc x'
]

TRANSCRIPTS_DIR = Path('transcripts')
RECORDINGS_DIR = Path('recordings')
OUTPUT_DIR = Path('output')


def get_auth_token():
    """Get authentication token for Graph API"""
    credential = ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
    token = credential.get_token('https://graph.microsoft.com/.default').token
    return token


def get_all_users(headers):
    """Get all users from the tenant"""
    print("\n🔍 Fetching all users from tenant...")
    users = []
    url = "https://graph.microsoft.com/v1.0/users?$select=id,displayName,mail,userPrincipalName&$top=999"
    
    try:
        resp = requests.get(url, headers=headers, timeout=60)
        if resp.status_code == 200:
            data = resp.json()
            users = data.get('value', [])
            print(f"   Found {len(users)} users")
            
            # Look for HR, Shey, Louise, KC
            target_found = []
            for user in users:
                name = (user.get('displayName') or '').lower()
                email = (user.get('mail') or user.get('userPrincipalName') or '').lower()
                
                if any(k in name or k in email for k in ['shey', 'louise', 'kc', 'hr@', 'human resource']):
                    target_found.append({
                        'id': user.get('id'),
                        'name': user.get('displayName'),
                        'email': user.get('mail') or user.get('userPrincipalName')
                    })
            
            if target_found:
                print(f"\n   🎯 Found target users:")
                for u in target_found:
                    print(f"      - {u['name']} ({u['email']}) - ID: {u['id']}")
                    TARGET_USERS[u['name'].lower()] = u['id']
        else:
            print(f"   ❌ Error: {resp.status_code} - {resp.text[:200]}")
    except Exception as e:
        print(f"   ❌ Error: {e}")
    
    return users


def get_user_calendar_events(headers, user_id, user_name, days_back=60):
    """Get calendar events for a user"""
    events = []
    start_date = (datetime.utcnow() - timedelta(days=days_back)).strftime('%Y-%m-%dT00:00:00Z')
    end_date = datetime.utcnow().strftime('%Y-%m-%dT23:59:59Z')
    
    print(f"\n📅 Fetching calendar events for {user_name}...")
    
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/calendar/calendarView?startDateTime={start_date}&endDateTime={end_date}&$top=500&$select=subject,start,end,organizer,isOnlineMeeting,onlineMeetingUrl"
    
    try:
        resp = requests.get(url, headers=headers, timeout=60)
        if resp.status_code == 200:
            data = resp.json()
            events = data.get('value', [])
            print(f"   Found {len(events)} calendar events")
            
            # Filter for check-in meetings
            checkin_events = []
            for event in events:
                subject = (event.get('subject') or '').lower()
                if any(kw in subject for kw in CHECKIN_KEYWORDS):
                    checkin_events.append(event)
            
            print(f"   🎯 Check-in meetings: {len(checkin_events)}")
            return events, checkin_events
        else:
            print(f"   ⚠️ Error: {resp.status_code}")
    except Exception as e:
        print(f"   ❌ Error: {e}")
    
    return [], []


def get_all_transcripts_from_api(headers, user_id, user_name, days_back=60):
    """Get all transcripts for a user from Graph API"""
    all_transcripts = []
    cutoff_date = datetime.utcnow() - timedelta(days=days_back)
    
    print(f"\n📝 Fetching transcripts for {user_name}...")
    url = f"https://graph.microsoft.com/beta/users/{user_id}/onlineMeetings/getAllTranscripts(meetingOrganizerUserId='{user_id}')"
    
    page_count = 0
    recent_count = 0
    
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
                                t['user_id'] = user_id
                                t['user_name'] = user_name
                                all_transcripts.append(t)
                                recent_count += 1
                        except:
                            pass
                
                page_count += 1
                url = data.get('@odata.nextLink')
                
                if page_count % 5 == 0:
                    print(f"   📄 Processed {page_count} pages, {recent_count} recent transcripts...")
                
                time.sleep(0.3)
            else:
                print(f"   ⚠️ Error: {resp.status_code}")
                break
        except requests.exceptions.Timeout:
            print(f"   ⚠️ Timeout, continuing...")
            break
        except Exception as e:
            print(f"   ❌ Error: {e}")
            break
    
    print(f"   ✅ Found {len(all_transcripts)} transcripts in last {days_back} days")
    return all_transcripts


def scan_local_recordings():
    """Scan local recordings folder for HR/Shey/Louise meetings"""
    print("\n📁 Scanning local recordings folder...")
    
    if not RECORDINGS_DIR.exists():
        print("   ⚠️ Recordings folder not found")
        return []
    
    recordings = []
    checkin_recordings = []
    
    for folder in RECORDINGS_DIR.iterdir():
        if folder.is_dir():
            name = folder.name.lower()
            recordings.append(folder.name)
            
            # Check if it's a check-in meeting
            if any(kw in name for kw in CHECKIN_KEYWORDS):
                checkin_recordings.append(folder.name)
    
    print(f"   Found {len(recordings)} total recordings")
    print(f"   🎯 Check-in recordings: {len(checkin_recordings)}")
    
    return recordings, checkin_recordings


def scan_local_transcripts():
    """Scan local transcripts folder"""
    print("\n📝 Scanning local transcripts folder...")
    
    if not TRANSCRIPTS_DIR.exists():
        print("   ⚠️ Transcripts folder not found")
        return [], []
    
    transcripts = []
    checkin_transcripts = []
    cutoff_date = datetime.now() - timedelta(days=60)
    
    for f in TRANSCRIPTS_DIR.glob('*.vtt'):
        name = f.stem.lower()
        transcripts.append(f.name)
        
        # Check date from filename
        try:
            date_str = f.stem[:8]
            file_date = datetime.strptime(date_str, '%Y%m%d')
            if file_date >= cutoff_date:
                if any(kw in name for kw in CHECKIN_KEYWORDS):
                    checkin_transcripts.append(f.name)
        except:
            pass
    
    print(f"   Found {len(transcripts)} total transcripts")
    print(f"   🎯 Recent check-in transcripts (last 60 days): {len(checkin_transcripts)}")
    
    return transcripts, checkin_transcripts


def get_meeting_details_and_download(headers, user_id, meeting_id, transcript_id, created_date):
    """Get meeting details and download transcript if available"""
    # Get meeting subject
    meeting_url = f"https://graph.microsoft.com/beta/users/{user_id}/onlineMeetings/{meeting_id}"
    subject = "Unknown Meeting"
    
    try:
        resp = requests.get(meeting_url, headers=headers, timeout=30)
        if resp.status_code == 200:
            subject = resp.json().get('subject', 'Unknown Meeting')
    except:
        pass
    
    # Download transcript
    transcript_url = f"https://graph.microsoft.com/beta/users/{user_id}/onlineMeetings/{meeting_id}/transcripts/{transcript_id}/content?$format=text/vtt"
    
    try:
        resp = requests.get(transcript_url, headers=headers, timeout=60)
        if resp.status_code == 200:
            # Save transcript
            safe_subject = re.sub(r'[<>:"/\\|?*]', '', subject)[:50]
            date_str = created_date[:10].replace('-', '') if created_date else datetime.now().strftime('%Y%m%d')
            time_str = created_date[11:19].replace(':', '') if created_date and len(created_date) > 11 else ''
            filename = f"{date_str}_{time_str}_{safe_subject}.vtt"
            filepath = TRANSCRIPTS_DIR / filename
            
            if not filepath.exists():
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write(resp.text)
                return filename, subject, True
            else:
                return filename, subject, False  # Already exists
        elif resp.status_code == 404:
            return None, subject, False
    except:
        pass
    
    return None, subject, False


def main():
    print("=" * 70)
    print("COMPREHENSIVE MEETING SEARCH - HR, SHEY, LOUISE")
    print("=" * 70)
    print(f"Date Range: Last 60 days (since {(datetime.now() - timedelta(days=60)).strftime('%Y-%m-%d')})")
    print(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Focus: Check-in meetings with Virtual Assistants")
    
    # Authenticate
    print("\n🔗 Authenticating with Microsoft Graph...")
    token = get_auth_token()
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    print("✅ Authenticated")
    
    results = {
        'searched_at': datetime.now().isoformat(),
        'date_range_days': 60,
        'users_found': [],
        'api_transcripts': [],
        'calendar_events': [],
        'checkin_meetings': [],
        'local_recordings': [],
        'local_transcripts': [],
        'new_downloads': [],
        'summary': {}
    }
    
    # Step 1: Get all users to find target users
    all_users = get_all_users(headers)
    results['users_found'] = list(TARGET_USERS.keys())
    
    # Step 2: Scan local files first
    local_recordings, checkin_recordings = scan_local_recordings()
    local_transcripts, checkin_local_transcripts = scan_local_transcripts()
    
    results['local_recordings'] = checkin_recordings
    results['local_transcripts'] = checkin_local_transcripts
    
    # Step 3: Get transcripts from API for each target user
    all_api_transcripts = []
    new_downloads = []
    
    for user_name, user_id in TARGET_USERS.items():
        print(f"\n{'='*50}")
        print(f"Processing: {user_name.upper()}")
        print(f"{'='*50}")
        
        # Get calendar events
        all_events, checkin_events = get_user_calendar_events(headers, user_id, user_name, days_back=60)
        
        for event in checkin_events[:20]:  # Log first 20
            results['checkin_meetings'].append({
                'subject': event.get('subject'),
                'start': event.get('start', {}).get('dateTime'),
                'organizer': event.get('organizer', {}).get('emailAddress', {}).get('name'),
                'user': user_name,
                'is_online': event.get('isOnlineMeeting')
            })
        
        # Get transcripts from API
        transcripts = get_all_transcripts_from_api(headers, user_id, user_name, days_back=60)
        all_api_transcripts.extend(transcripts)
        
        # Try to download new transcripts
        print(f"\n   🔄 Checking for downloadable transcripts...")
        downloaded = 0
        skipped = 0
        not_available = 0
        
        for t in transcripts[:100]:  # Check first 100
            transcript_id = t.get('id', '')
            meeting_id = t.get('meetingId', '')
            created_date = t.get('createdDateTime', '')
            
            filename, subject, is_new = get_meeting_details_and_download(
                headers, user_id, meeting_id, transcript_id, created_date
            )
            
            if filename and is_new:
                downloaded += 1
                new_downloads.append({
                    'filename': filename,
                    'subject': subject,
                    'date': created_date,
                    'user': user_name
                })
                print(f"      ✅ Downloaded: {subject[:40]}...")
            elif filename:
                skipped += 1
            else:
                not_available += 1
            
            time.sleep(0.3)
        
        print(f"   📊 Downloaded: {downloaded}, Skipped (exists): {skipped}, Not available: {not_available}")
    
    results['new_downloads'] = new_downloads
    
    # Summary
    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    
    print(f"\n📊 Users searched: {', '.join(TARGET_USERS.keys())}")
    print(f"\n📅 Calendar check-in meetings found: {len(results['checkin_meetings'])}")
    print(f"📝 API transcripts (last 60 days): {len(all_api_transcripts)}")
    print(f"📁 Local check-in recordings: {len(checkin_recordings)}")
    print(f"📄 Local check-in transcripts: {len(checkin_local_transcripts)}")
    print(f"🆕 Newly downloaded: {len(new_downloads)}")
    
    # List recent check-in meetings
    print(f"\n📋 Recent Check-in Meetings (from calendar):")
    for meeting in sorted(results['checkin_meetings'], key=lambda x: x.get('start', ''), reverse=True)[:20]:
        date = meeting.get('start', '')[:10] if meeting.get('start') else 'Unknown'
        print(f"   - [{date}] {meeting.get('subject', 'No subject')[:60]}")
    
    # List newly downloaded
    if new_downloads:
        print(f"\n✨ Newly Downloaded Transcripts:")
        for t in new_downloads:
            date = t.get('date', '')[:10] if t.get('date') else 'Unknown'
            print(f"   - [{date}] {t.get('subject', 'Unknown')[:60]}")
    
    # List local check-in recordings without transcripts
    print(f"\n📁 Check-in Recordings (local folder) - Sample:")
    for rec in sorted(checkin_recordings, reverse=True)[:15]:
        print(f"   - {rec[:70]}")
    
    results['summary'] = {
        'users_searched': list(TARGET_USERS.keys()),
        'calendar_checkins': len(results['checkin_meetings']),
        'api_transcripts': len(all_api_transcripts),
        'local_checkin_recordings': len(checkin_recordings),
        'local_checkin_transcripts': len(checkin_local_transcripts),
        'new_downloads': len(new_downloads)
    }
    
    # Save results
    output_file = OUTPUT_DIR / f'meeting_search_results_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json'
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=2, default=str)
    
    print(f"\n💾 Full results saved to: {output_file}")
    
    # Final count
    total_transcripts = len(list(TRANSCRIPTS_DIR.glob('*.vtt')))
    print(f"\n📊 Total transcripts now: {total_transcripts}")
    
    return results


if __name__ == "__main__":
    main()
