"""
Check all check-in meetings for HR@our-assistants.com
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

TENANT_ID = os.getenv('AZURE_TENANT_ID')
CLIENT_ID = os.getenv('AZURE_CLIENT_ID')
CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET')

TRANSCRIPTS_DIR = Path('transcripts')
OUTPUT_DIR = Path('output')

def get_auth():
    credential = ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
    token = credential.get_token('https://graph.microsoft.com/.default').token
    return {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}

def main():
    print("=" * 70)
    print("CHECK-IN MEETINGS FOR HR@our-assistants.com")
    print("=" * 70)
    print(f"Date Range: Last 60 days")
    print(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    headers = get_auth()
    
    # Look up HR user
    print("\n🔍 Looking up HR@our-assistants.com...")
    resp = requests.get('https://graph.microsoft.com/v1.0/users/HR@our-assistants.com', headers=headers)
    
    if resp.status_code == 200:
        user = resp.json()
        user_id = user.get('id')
        print(f"✅ Found: {user.get('displayName')} - ID: {user_id}")
    else:
        print(f"❌ User lookup error: {resp.status_code}")
        print(f"   {resp.text[:200]}")
        return
    
    # Get calendar events
    print(f"\n📅 Searching calendar events (last 60 days)...")
    start_date = (datetime.utcnow() - timedelta(days=60)).strftime('%Y-%m-%dT00:00:00Z')
    end_date = datetime.utcnow().strftime('%Y-%m-%dT23:59:59Z')
    
    all_events = []
    url = f'https://graph.microsoft.com/v1.0/users/{user_id}/calendar/calendarView'
    params = {'startDateTime': start_date, 'endDateTime': end_date, '$top': 500}
    
    while url:
        resp = requests.get(url, headers=headers, params=params if '?' not in url else None)
        if resp.status_code == 200:
            data = resp.json()
            all_events.extend(data.get('value', []))
            url = data.get('@odata.nextLink')
            params = None
        else:
            print(f"❌ Calendar error: {resp.status_code}")
            break
    
    print(f"✅ Total calendar events: {len(all_events)}")
    
    # Filter check-in meetings
    checkins = []
    for e in all_events:
        subject = e.get('subject', '').lower()
        if 'check-in' in subject or 'check in' in subject or 'checkin' in subject or 'integration' in subject:
            checkins.append({
                'subject': e.get('subject', 'No Subject'),
                'start': e.get('start', {}).get('dateTime', ''),
                'end': e.get('end', {}).get('dateTime', ''),
                'isOnlineMeeting': e.get('isOnlineMeeting', False),
                'isCancelled': e.get('isCancelled', False),
                'organizer': e.get('organizer', {}).get('emailAddress', {}).get('address', ''),
                'attendees': [a.get('emailAddress', {}).get('address', '') for a in e.get('attendees', [])]
            })
    
    print(f"✅ Check-in/Integration meetings found: {len(checkins)}")
    
    # Sort by date (most recent first)
    checkins.sort(key=lambda x: x['start'], reverse=True)
    
    # Display meetings grouped by type
    print("\n" + "=" * 70)
    print("CHECK-IN MEETINGS SUMMARY")
    print("=" * 70)
    
    # Group by person
    by_person = {}
    for m in checkins:
        subject = m['subject']
        # Extract person name from subject
        if ' x ' in subject.lower():
            parts = subject.lower().split(' x ')
            if len(parts) >= 2:
                person = parts[-1].split('-')[0].split('(')[0].strip()
                person = person.title()
            else:
                person = 'Other'
        else:
            person = 'Other'
        
        if person not in by_person:
            by_person[person] = []
        by_person[person].append(m)
    
    print(f"\n📊 Meetings by Person:")
    for person in sorted(by_person.keys()):
        meetings = by_person[person]
        print(f"\n  👤 {person}: {len(meetings)} meetings")
        for m in meetings[:5]:  # Show last 5
            date = m['start'][:10] if m['start'] else 'Unknown'
            cancelled = " [CANCELLED]" if m['isCancelled'] else ""
            print(f"     {date} - {m['subject'][:50]}{cancelled}")
        if len(meetings) > 5:
            print(f"     ... and {len(meetings) - 5} more")
    
    # Now check for transcripts
    print("\n" + "=" * 70)
    print("CHECKING FOR TRANSCRIPTS")
    print("=" * 70)
    
    # Get all transcripts from Graph API
    print("\n🔍 Fetching transcripts from Graph API...")
    
    all_transcripts = []
    url = f"https://graph.microsoft.com/beta/users/{user_id}/onlineMeetings/getAllTranscripts(meetingOrganizerUserId='{user_id}')"
    
    page_count = 0
    while url:
        try:
            resp = requests.get(url, headers=headers, timeout=60)
            if resp.status_code == 200:
                data = resp.json()
                transcripts = data.get('value', [])
                
                cutoff = datetime.utcnow() - timedelta(days=60)
                for t in transcripts:
                    created = t.get('createdDateTime', '')
                    if created:
                        try:
                            dt = datetime.fromisoformat(created.replace('Z', '+00:00'))
                            if dt.replace(tzinfo=None) >= cutoff:
                                all_transcripts.append(t)
                        except:
                            pass
                
                page_count += 1
                url = data.get('@odata.nextLink')
                time.sleep(0.2)
            else:
                print(f"❌ Transcript API error: {resp.status_code}")
                break
        except Exception as e:
            print(f"❌ Error: {e}")
            break
    
    print(f"✅ Transcripts in last 60 days: {len(all_transcripts)}")
    
    # Get meeting details for each transcript
    print("\n🔍 Getting meeting details for transcripts...")
    transcript_meetings = []
    
    for i, t in enumerate(all_transcripts):
        meeting_id = t.get('meetingId', '')
        transcript_id = t.get('id', '')
        created = t.get('createdDateTime', '')
        
        # Get meeting details
        url = f"https://graph.microsoft.com/beta/users/{user_id}/onlineMeetings/{meeting_id}"
        try:
            resp = requests.get(url, headers=headers, timeout=30)
            if resp.status_code == 200:
                meeting = resp.json()
                subject = meeting.get('subject', 'Unknown')
                
                if 'check' in subject.lower() or 'integration' in subject.lower():
                    transcript_meetings.append({
                        'subject': subject,
                        'date': created[:10] if created else 'Unknown',
                        'meeting_id': meeting_id,
                        'transcript_id': transcript_id,
                        'created': created
                    })
                    print(f"  ✅ [{i+1}/{len(all_transcripts)}] {subject[:50]} ({created[:10]})")
        except:
            pass
        
        time.sleep(0.3)
    
    print(f"\n✅ Check-in meetings with transcripts: {len(transcript_meetings)}")
    
    # Check which ones we already have
    existing = set()
    for f in TRANSCRIPTS_DIR.glob('*.vtt'):
        existing.add(f.stem.lower())
    
    # Download missing transcripts
    print("\n" + "=" * 70)
    print("DOWNLOADING MISSING TRANSCRIPTS")
    print("=" * 70)
    
    new_downloads = []
    for tm in transcript_meetings:
        # Check if we have it
        date_prefix = tm['date'].replace('-', '')
        safe_subj = re.sub(r'[<>:"/\\|?*]', '', tm['subject'])[:30].lower()
        
        have_it = False
        for ex in existing:
            if date_prefix in ex and safe_subj[:15] in ex:
                have_it = True
                break
        
        if have_it:
            continue
        
        # Download
        print(f"\n  📥 Downloading: {tm['subject'][:50]}...")
        url = f"https://graph.microsoft.com/beta/users/{user_id}/onlineMeetings/{tm['meeting_id']}/transcripts/{tm['transcript_id']}/content?$format=text/vtt"
        
        try:
            resp = requests.get(url, headers=headers, timeout=60)
            if resp.status_code == 200:
                safe_subject = re.sub(r'[<>:"/\\|?*]', '', tm['subject'])[:50]
                date_str = tm['date'].replace('-', '')
                time_str = tm['created'][11:19].replace(':', '') if len(tm['created']) > 11 else ''
                filename = f"{date_str}_{time_str}_{safe_subject}.vtt"
                filepath = TRANSCRIPTS_DIR / filename
                
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write(resp.text)
                
                print(f"     ✅ Saved: {filename}")
                new_downloads.append({
                    'filename': filename,
                    'subject': tm['subject'],
                    'date': tm['date']
                })
            elif resp.status_code == 404:
                print(f"     ⏭️ Transcript expired/deleted")
            else:
                print(f"     ❌ Error: {resp.status_code}")
        except Exception as e:
            print(f"     ❌ Error: {e}")
        
        time.sleep(0.5)
    
    # Summary
    print("\n" + "=" * 70)
    print("FINAL SUMMARY")
    print("=" * 70)
    print(f"\n📅 Calendar Check-in Meetings (last 60 days): {len(checkins)}")
    print(f"📄 Transcripts Available: {len(transcript_meetings)}")
    print(f"🆕 Newly Downloaded: {len(new_downloads)}")
    print(f"📁 Total Transcripts in Folder: {len(list(TRANSCRIPTS_DIR.glob('*.vtt')))}")
    
    if new_downloads:
        print(f"\n✨ New Transcripts:")
        for t in new_downloads:
            print(f"   - {t['date']} - {t['subject'][:50]}")
    
    # Save results
    results = {
        'search_date': datetime.now().isoformat(),
        'user': 'HR@our-assistants.com',
        'calendar_checkins': len(checkins),
        'transcripts_available': len(transcript_meetings),
        'new_downloads': new_downloads,
        'meetings_by_person': {k: len(v) for k, v in by_person.items()}
    }
    
    with open(OUTPUT_DIR / 'hr_checkin_meetings.json', 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=2, default=str)
    
    print(f"\n💾 Results saved to: output/hr_checkin_meetings.json")
    
    return new_downloads

if __name__ == "__main__":
    new = main()
    if new:
        print("\n" + "=" * 60)
        print("NEXT: Run Excel update to analyze new transcripts")
        print("=" * 60)
        print("Run: .venv\\Scripts\\python.exe src/update_excel_with_transcripts.py")
