"""Check for recent transcripts and recordings in Microsoft Graph API."""
import requests
from azure.identity import ClientSecretCredential
import os
from dotenv import load_dotenv
from datetime import datetime

load_dotenv()

TENANT_ID = os.getenv('AZURE_TENANT_ID', '187b2af6-1bfb-490a-85dd-b720fe3d31bc')
CLIENT_ID = os.getenv('AZURE_CLIENT_ID')
CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET')
HR_USER_ID = '81835016-79d5-4a15-91b1-c104e2cd9adb'

def main():
    credential = ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
    token = credential.get_token('https://graph.microsoft.com/.default').token
    print('✅ Authentication successful!')
    print()
    
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    
    # ========================================
    # 1. Check Transcripts
    # ========================================
    print('=' * 70)
    print('CHECKING FOR TRANSCRIPTS (newest first)')
    print('=' * 70)
    
    url = f"https://graph.microsoft.com/beta/users/{HR_USER_ID}/onlineMeetings/getAllTranscripts(meetingOrganizerUserId='{HR_USER_ID}')"
    resp = requests.get(url, headers=headers)
    print(f'Transcripts API Status: {resp.status_code}')
    
    if resp.status_code == 200:
        data = resp.json()
        transcripts = data.get('value', [])
        print(f'Total transcripts found: {len(transcripts)}')
        print()
        
        # Sort by date (newest first)
        transcripts_sorted = sorted(transcripts, key=lambda x: x.get('createdDateTime', ''), reverse=True)
        
        print('NEWEST 10 TRANSCRIPTS:')
        print('-' * 50)
        for i, t in enumerate(transcripts_sorted[:10], 1):
            created = t.get('createdDateTime', 'unknown')
            transcript_id = t.get('id', 'unknown')
            print(f'[{i}] Created: {created}')
            print(f'    Transcript ID: {transcript_id}')
            print()
    else:
        print(f'Error: {resp.text[:500]}')
    
    # ========================================
    # 2. Check Recordings
    # ========================================
    print()
    print('=' * 70)
    print('CHECKING FOR RECORDINGS (newest first)')
    print('=' * 70)
    
    url = f"https://graph.microsoft.com/beta/users/{HR_USER_ID}/onlineMeetings/getAllRecordings(meetingOrganizerUserId='{HR_USER_ID}')"
    resp = requests.get(url, headers=headers)
    print(f'Recordings API Status: {resp.status_code}')
    
    if resp.status_code == 200:
        data = resp.json()
        recordings = data.get('value', [])
        print(f'Total recordings found: {len(recordings)}')
        print()
        
        # Sort by date (newest first)
        recordings_sorted = sorted(recordings, key=lambda x: x.get('createdDateTime', ''), reverse=True)
        
        print('NEWEST 10 RECORDINGS:')
        print('-' * 50)
        for i, r in enumerate(recordings_sorted[:10], 1):
            created = r.get('createdDateTime', 'unknown')
            recording_id = r.get('id', 'unknown')
            print(f'[{i}] Created: {created}')
            print(f'    Recording ID: {recording_id}')
            print()
    else:
        print(f'Error: {resp.text[:500]}')
    
    # ========================================
    # 3. Check OneDrive Recordings folder
    # ========================================
    print()
    print('=' * 70)
    print('CHECKING ONEDRIVE RECORDINGS FOLDER')
    print('=' * 70)
    
    # Check HR user's OneDrive for Recordings folder
    url = f"https://graph.microsoft.com/v1.0/users/{HR_USER_ID}/drive/root:/Recordings:/children"
    resp = requests.get(url, headers=headers)
    print(f'OneDrive Recordings API Status: {resp.status_code}')
    
    if resp.status_code == 200:
        data = resp.json()
        files = data.get('value', [])
        print(f'Files in Recordings folder: {len(files)}')
        print()
        
        # Sort by date (newest first)
        files_sorted = sorted(files, key=lambda x: x.get('createdDateTime', ''), reverse=True)
        
        print('NEWEST 10 FILES:')
        print('-' * 50)
        for i, f in enumerate(files_sorted[:10], 1):
            created = f.get('createdDateTime', 'unknown')
            name = f.get('name', 'unknown')
            size = f.get('size', 0)
            print(f'[{i}] {name}')
            print(f'    Created: {created}')
            print(f'    Size: {size / 1024 / 1024:.1f} MB')
            print()
    elif resp.status_code == 404:
        print('Recordings folder not found in OneDrive')
    else:
        print(f'Error: {resp.text[:300]}')
    
    # ========================================
    # 4. Summary
    # ========================================
    print()
    print('=' * 70)
    print('WHERE TO FIND RECORDINGS & TRANSCRIPTS')
    print('=' * 70)
    print('''
1. MICROSOFT STREAM (stream.microsoft.com)
   - Go to "My content" → "Meetings"
   - All meeting recordings with transcripts appear here
   - Best place to check first!

2. MICROSOFT TEAMS
   - Open Teams → Calendar → Find the meeting
   - Click the meeting → "Recordings & Transcripts" tab
   - Or check the meeting chat for the recording link

3. ONEDRIVE
   - Recordings folder: OneDrive/Recordings
   - Only appears after first recording is made
   
4. SHAREPOINT (for channel meetings)
   - Team → Channel → Files → Recordings

NOTE: It can take 5-15 minutes for recordings/transcripts 
to appear after a meeting ends!
''')

if __name__ == '__main__':
    main()
