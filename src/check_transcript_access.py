"""Check Graph API access to transcripts via Communications API."""
import requests
import os
from dotenv import load_dotenv
from azure.identity import ClientSecretCredential

load_dotenv()

credential = ClientSecretCredential(
    tenant_id=os.getenv('AZURE_TENANT_ID'),
    client_id=os.getenv('AZURE_CLIENT_ID'),
    client_secret=os.getenv('AZURE_CLIENT_SECRET')
)

token = credential.get_token('https://graph.microsoft.com/.default')
headers = {'Authorization': f'Bearer {token.token}'}

# Get HR user ID first
print("Finding HR user...")
resp = requests.get('https://graph.microsoft.com/v1.0/users', headers=headers)
users = resp.json().get('value', [])
hr_user = None
for u in users:
    email = u.get('mail', '') or u.get('userPrincipalName', '') or ''
    if 'hr@our-assistants' in email.lower():
        hr_user = u
        print(f"HR User: {u['displayName']} ({email})")
        print(f"User ID: {u['id']}")
        break

if not hr_user:
    print("HR user not found!")
    exit(1)

user_id = hr_user['id']

# Try to get online meetings for this user
print('\n' + '='*60)
print('Checking Online Meetings API...')
print('='*60)
meetings_url = f'https://graph.microsoft.com/v1.0/users/{user_id}/onlineMeetings'
resp = requests.get(meetings_url, headers=headers)
print(f'Status: {resp.status_code}')
if resp.status_code == 200:
    meetings = resp.json().get('value', [])
    print(f'Found {len(meetings)} meetings')
    for m in meetings[:5]:
        print(f"  - {m.get('subject', 'No subject')}")
        meeting_id = m.get('id')
        if meeting_id:
            # Try to get transcripts for this meeting
            transcript_url = f'https://graph.microsoft.com/v1.0/users/{user_id}/onlineMeetings/{meeting_id}/transcripts'
            t_resp = requests.get(transcript_url, headers=headers)
            print(f"    Transcripts: {t_resp.status_code}")
            if t_resp.status_code == 200:
                transcripts = t_resp.json().get('value', [])
                print(f"    Found {len(transcripts)} transcripts")
else:
    print(f'Error: {resp.text[:300]}')

# Check communications/callRecords
print('\n' + '='*60)
print('Checking Call Records API...')
print('='*60)
call_records_url = 'https://graph.microsoft.com/v1.0/communications/callRecords'
resp = requests.get(call_records_url, headers=headers)
print(f'Status: {resp.status_code}')
if resp.status_code == 200:
    records = resp.json().get('value', [])
    print(f'Found {len(records)} call records')
else:
    print(f'Error: {resp.text[:300]}')

# Check calendar events with online meetings
print('\n' + '='*60)
print('Checking Calendar Events with Meetings...')
print('='*60)
cal_url = f"https://graph.microsoft.com/v1.0/users/{user_id}/events?$top=10&$orderby=start/dateTime desc&$filter=isOnlineMeeting eq true"
resp = requests.get(cal_url, headers=headers)
print(f'Status: {resp.status_code}')
if resp.status_code == 200:
    events = resp.json().get('value', [])
    print(f'Found {len(events)} online meeting events')
    for e in events[:5]:
        subject = e.get('subject', 'No subject')
        start = e.get('start', {}).get('dateTime', '')[:10]
        print(f"  - [{start}] {subject}")
        
        # Check if there's a join URL we can use to get meeting ID
        online_meeting = e.get('onlineMeeting', {})
        if online_meeting:
            join_url = online_meeting.get('joinUrl', '')
            if join_url:
                print(f"    Has join URL")
else:
    print(f'Error: {resp.text[:300]}')

# Check Drive for .vtt files specifically
print('\n' + '='*60)
print('Searching Drive for VTT files...')
print('='*60)
search_url = f"https://graph.microsoft.com/v1.0/users/{user_id}/drive/root/search(q='.vtt')"
resp = requests.get(search_url, headers=headers)
print(f'Status: {resp.status_code}')
if resp.status_code == 200:
    files = resp.json().get('value', [])
    print(f'Found {len(files)} VTT files')
    for f in files[:10]:
        print(f"  - {f.get('name')} ({f.get('size', 0)} bytes)")
        print(f"    Path: {f.get('parentReference', {}).get('path', 'N/A')}")
else:
    print(f'Error: {resp.text[:300]}')

# Also check in Microsoft Stream or other locations
print('\n' + '='*60)
print('Checking SharePoint sites for transcripts...')
print('='*60)
sites_url = "https://graph.microsoft.com/v1.0/sites?search=*"
resp = requests.get(sites_url, headers=headers)
print(f'Status: {resp.status_code}')
if resp.status_code == 200:
    sites = resp.json().get('value', [])
    print(f'Found {len(sites)} SharePoint sites')
    for s in sites[:5]:
        print(f"  - {s.get('displayName', s.get('name', 'Unknown'))}")
