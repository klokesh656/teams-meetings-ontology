"""
Check for recent meetings and transcripts from Microsoft Graph API.
Uses the getAllTranscripts endpoint like daily_sync.py
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

# User IDs for transcript access
USER_IDS = [
    '81835016-79d5-4a15-91b1-c104e2cd9adb',  # HR account
]

TRANSCRIPTS_DIR = Path('transcripts')
OUTPUT_DIR = Path('output')

# Ensure directories exist
TRANSCRIPTS_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


def get_auth_token():
    """Get authentication token for Graph API"""
    credential = ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
    token = credential.get_token('https://graph.microsoft.com/.default').token
    return token


def get_existing_transcripts():
    """Get list of existing transcript files"""
    existing = set()
    for f in TRANSCRIPTS_DIR.glob('*.vtt'):
        existing.add(f.stem)
    return existing


def get_all_transcripts(headers, days_back=14):
    """Fetch transcripts from the last N days"""
    all_transcripts = []
    cutoff_date = datetime.utcnow() - timedelta(days=days_back)
    
    for user_id in USER_IDS:
        print(f"\n👤 Fetching transcripts for user: {user_id}")
        url = f"https://graph.microsoft.com/beta/users/{user_id}/onlineMeetings/getAllTranscripts(meetingOrganizerUserId='{user_id}')"
        
        page_count = 0
        while url:
            try:
                resp = requests.get(url, headers=headers, timeout=60)
                if resp.status_code == 200:
                    data = resp.json()
                    transcripts = data.get('value', [])
                    
                    # Filter by date
                    for t in transcripts:
                        created_date = t.get('createdDateTime', '')
                        if created_date:
                            try:
                                created = datetime.fromisoformat(created_date.replace('Z', '+00:00'))
                                if created.replace(tzinfo=None) >= cutoff_date:
                                    t['user_id'] = user_id
                                    all_transcripts.append(t)
                            except:
                                t['user_id'] = user_id
                                all_transcripts.append(t)
                    
                    page_count += 1
                    url = data.get('@odata.nextLink')
                    
                    if page_count % 10 == 0:
                        print(f"  📄 Processed {page_count} pages, found {len(all_transcripts)} recent transcripts...")
                    
                    time.sleep(0.3)  # Rate limiting
                else:
                    print(f"  ❌ Error: {resp.status_code} - {resp.text[:100]}")
                    break
            except requests.exceptions.Timeout:
                print(f"  ⚠️ Timeout, retrying...")
                time.sleep(2)
                continue
            except Exception as e:
                print(f"  ❌ Error: {e}")
                break
    
    return all_transcripts


def get_meeting_details(headers, user_id, meeting_id):
    """Get meeting details including subject"""
    url = f"https://graph.microsoft.com/beta/users/{user_id}/onlineMeetings/{meeting_id}"
    try:
        resp = requests.get(url, headers=headers, timeout=30)
        if resp.status_code == 200:
            return resp.json()
    except:
        pass
    return {}


def download_transcript(headers, user_id, meeting_id, transcript_id, meeting_subject, created_date):
    """Download a single transcript"""
    url = f"https://graph.microsoft.com/beta/users/{user_id}/onlineMeetings/{meeting_id}/transcripts/{transcript_id}/content?$format=text/vtt"
    
    try:
        resp = requests.get(url, headers=headers, timeout=60)
        if resp.status_code == 200:
            # Create filename
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
        print(f"    ❌ Error: {e}")
    
    return None, "error"


def main():
    print("=" * 60)
    print("CHECKING FOR RECENT MEETINGS & TRANSCRIPTS")
    print("=" * 60)
    print(f"Date Range: Last 14 days (since {(datetime.now() - timedelta(days=14)).strftime('%Y-%m-%d')})")
    print(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
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
    
    # Get all transcripts from last 14 days
    transcripts = get_all_transcripts(headers, days_back=14)
    print(f"\n📊 Found {len(transcripts)} transcripts in the last 14 days")
    
    if not transcripts:
        print("\n⚠️ No recent transcripts found. This could mean:")
        print("   - No meetings with transcripts in the last 14 days")
        print("   - Transcripts have expired (they expire after a period)")
        print("   - Permission issues")
        return
    
    # Process each transcript
    new_downloaded = []
    skipped = 0
    not_found = 0
    
    print("\n🔄 Processing transcripts...")
    for i, t in enumerate(transcripts):
        transcript_id = t.get('id', '')
        meeting_id = t.get('meetingId', '')
        created_date = t.get('createdDateTime', '')
        user_id = t.get('user_id', USER_IDS[0])
        
        # Generate a check name based on available info
        check_name = f"{created_date[:10].replace('-', '')}_{transcript_id[:8]}" if created_date else transcript_id[:20]
        
        # Check if we might already have this
        already_have = False
        for ex in existing:
            if check_name[:8] in ex or transcript_id[:8] in ex:
                already_have = True
                break
        
        if already_have:
            skipped += 1
            continue
        
        # Get meeting details
        meeting = get_meeting_details(headers, user_id, meeting_id)
        subject = meeting.get('subject', 'Unknown Meeting')
        
        print(f"\n  [{i+1}/{len(transcripts)}] {subject[:50]}...")
        print(f"      Date: {created_date[:10] if created_date else 'Unknown'}")
        
        # Download transcript
        filepath, filename = download_transcript(headers, user_id, meeting_id, transcript_id, subject, created_date)
        
        if filepath:
            print(f"      ✅ Downloaded: {filename}")
            new_downloaded.append({
                'filename': filename,
                'subject': subject,
                'date': created_date,
                'meeting_id': meeting_id
            })
            existing.add(filename[:-4])  # Add without .vtt
        elif filename == "not_found":
            print(f"      ⏭️ Transcript not available (expired/deleted)")
            not_found += 1
        
        time.sleep(0.5)  # Rate limiting
    
    # Summary
    print("\n" + "=" * 60)
    print("SUMMARY")
    print("=" * 60)
    print(f"📊 Transcripts in last 14 days: {len(transcripts)}")
    print(f"⏭️ Already downloaded (skipped): {skipped}")
    print(f"❌ Not available (expired): {not_found}")
    print(f"🆕 Newly downloaded: {len(new_downloaded)}")
    print(f"📁 Total transcripts now: {len(existing)}")
    
    if new_downloaded:
        print("\n✨ New transcripts:")
        for t in new_downloaded:
            print(f"   - {t['subject'][:60]}")
            print(f"     Date: {t['date'][:10] if t['date'] else 'Unknown'}")
    
    # Save results
    results = {
        'checked_at': datetime.now().isoformat(),
        'date_range': f"Last 14 days (since {(datetime.now() - timedelta(days=14)).strftime('%Y-%m-%d')})",
        'total_found': len(transcripts),
        'skipped': skipped,
        'not_available': not_found,
        'new_downloaded': len(new_downloaded),
        'new_transcripts': new_downloaded
    }
    
    with open(OUTPUT_DIR / 'recent_check_results.json', 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=2, default=str)
    
    print(f"\n💾 Results saved to: output/recent_check_results.json")
    
    return new_downloaded


if __name__ == "__main__":
    new_transcripts = main()
    
    if new_transcripts:
        print("\n" + "=" * 60)
        print("NEXT STEP: Update Excel with new transcripts")
        print("=" * 60)
        print("Run: .venv\\Scripts\\python.exe src/update_excel_with_transcripts.py")
