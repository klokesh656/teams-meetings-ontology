"""
Download all transcripts from Microsoft Teams meetings.
This script accesses transcripts via the Graph API Communications endpoint.
Includes retry logic, rate limiting, and resume capability.
"""
import os
import json
import time
import requests
from datetime import datetime
from dotenv import load_dotenv
from azure.identity import ClientSecretCredential
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

load_dotenv()

CLIENT_ID = os.getenv('AZURE_CLIENT_ID')
CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET')
TENANT_ID = os.getenv('AZURE_TENANT_ID')

credential = ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
token = credential.get_token('https://graph.microsoft.com/.default').token
headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}

# Create session with retry logic
session = requests.Session()
retry_strategy = Retry(
    total=3,
    backoff_factor=2,
    status_forcelist=[429, 500, 502, 503, 504],
)
adapter = HTTPAdapter(max_retries=retry_strategy)
session.mount("https://", adapter)
session.headers.update(headers)

# HR user ID (organizer of meetings)
USER_ID = '81835016-79d5-4a15-91b1-c104e2cd9adb'

# Output directory
OUTPUT_DIR = 'transcripts'
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Rate limiting - delay between requests
REQUEST_DELAY = 0.5  # seconds

def get_all_transcripts():
    """Get all transcripts for the user"""
    url = f'https://graph.microsoft.com/beta/users/{USER_ID}/onlineMeetings/getAllTranscripts(meetingOrganizerUserId=\'{USER_ID}\')'
    
    all_transcripts = []
    while url:
        try:
            resp = session.get(url, timeout=30)
            if resp.status_code != 200:
                print(f"Error: {resp.status_code} - {resp.text[:200]}")
                break
            
            data = resp.json()
            all_transcripts.extend(data.get('value', []))
            url = data.get('@odata.nextLink')
            time.sleep(REQUEST_DELAY)
        except Exception as e:
            print(f"Error fetching transcripts: {e}")
            time.sleep(5)
            continue
    
    return all_transcripts

def download_transcript_content(meeting_id, transcript_id):
    """Download the actual transcript content in VTT format"""
    url = f'https://graph.microsoft.com/beta/users/{USER_ID}/onlineMeetings/{meeting_id}/transcripts/{transcript_id}/content?$format=text/vtt'
    
    try:
        resp = session.get(url, timeout=60)
        time.sleep(REQUEST_DELAY)
        if resp.status_code == 200:
            return resp.text
        else:
            return None
    except Exception as e:
        print(f"   Download error: {str(e)[:50]}")
        return None

def get_meeting_details(meeting_id):
    """Get meeting details like subject and participants"""
    url = f'https://graph.microsoft.com/beta/users/{USER_ID}/onlineMeetings/{meeting_id}'
    try:
        resp = session.get(url, timeout=30)
        time.sleep(REQUEST_DELAY)
        if resp.status_code == 200:
            return resp.json()
    except:
        pass
    return {}

def main():
    print("=" * 70)
    print("DOWNLOADING TEAMS MEETING TRANSCRIPTS")
    print("=" * 70)
    
    # Get existing files to skip
    existing_files = set(os.listdir(OUTPUT_DIR))
    print(f"Found {len(existing_files)} existing files to skip")
    
    # Get all transcripts
    print("\nFetching transcript list...")
    transcripts = get_all_transcripts()
    print(f"Found {len(transcripts)} transcripts")
    
    # Download each transcript
    downloaded = 0
    failed = 0
    skipped = 0
    
    for i, t in enumerate(transcripts):
        meeting_id = t.get('meetingId', '')
        transcript_id = t.get('id', '')
        created = t.get('createdDateTime', '')
        
        # Parse date for filename
        try:
            dt = datetime.fromisoformat(created.replace('Z', '+00:00'))
            date_str = dt.strftime('%Y%m%d_%H%M%S')
        except:
            date_str = 'unknown'
        
        # Get meeting details for subject
        meeting = get_meeting_details(meeting_id)
        subject = meeting.get('subject', 'No Subject')
        # Clean subject for filename
        safe_subject = "".join(c for c in subject if c.isalnum() or c in ' -_')[:50]
        
        # Create filename
        filename = f"{date_str}_{safe_subject}.vtt"
        
        # Check if already downloaded
        if filename in existing_files:
            skipped += 1
            continue
        
        print(f"\n[{i+1}/{len(transcripts)}] {created[:10]} - {subject[:40]}")
        
        # Skip meetings with no subject (usually have no content)
        if subject == 'No Subject':
            failed += 1
            continue
        
        # Download transcript content
        content = download_transcript_content(meeting_id, transcript_id)
        
        if content:
            filepath = os.path.join(OUTPUT_DIR, filename)
            
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(content)
            
            print(f"   ✅ Saved: {filename}")
            downloaded += 1
        else:
            failed += 1
    
    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    print(f"Total transcripts: {len(transcripts)}")
    print(f"Downloaded: {downloaded}")
    print(f"Skipped (already exists): {skipped}")
    print(f"Failed/No content: {failed}")
    print(f"\nTranscripts saved to: {os.path.abspath(OUTPUT_DIR)}")

if __name__ == '__main__':
    main()
