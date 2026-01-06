"""
Test transcript access using correct API format.
The Application Access Policy takes 15-60 minutes to propagate.
"""
import os
import requests
from dotenv import load_dotenv
from azure.identity import ClientSecretCredential

load_dotenv()

CLIENT_ID = os.getenv('AZURE_CLIENT_ID')
CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET')
TENANT_ID = os.getenv('AZURE_TENANT_ID')

credential = ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
token = credential.get_token('https://graph.microsoft.com/.default').token
headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}

# HR user ID
USER_ID = '81835016-79d5-4a15-91b1-c104e2cd9adb'

print("=" * 70)
print("TESTING TRANSCRIPT API ACCESS")
print("=" * 70)

# Method 1: Get all transcripts for a user (beta API)
print("\n1. Getting all transcripts (beta getAllTranscripts)...")
url = f'https://graph.microsoft.com/beta/users/{USER_ID}/onlineMeetings/getAllTranscripts(meetingOrganizerUserId=\'{USER_ID}\')'
resp = requests.get(url, headers=headers)
print(f"   Status: {resp.status_code}")
if resp.status_code == 200:
    data = resp.json()
    transcripts = data.get('value', [])
    print(f"   ✅ SUCCESS! Found {len(transcripts)} transcripts")
    for t in transcripts[:5]:
        print(f"      - Meeting: {t.get('meetingId', '')[:30]}...")
        print(f"        Transcript ID: {t.get('id', '')}")
        print(f"        Created: {t.get('createdDateTime', '')}")
elif resp.status_code == 403:
    print("   ❌ Access Denied - Policy may still be propagating (wait 15-60 min)")
else:
    print(f"   Response: {resp.text[:300]}")

# Method 2: Get all recordings for a user
print("\n2. Getting all recordings (beta getAllRecordings)...")
url = f'https://graph.microsoft.com/beta/users/{USER_ID}/onlineMeetings/getAllRecordings(meetingOrganizerUserId=\'{USER_ID}\')'
resp = requests.get(url, headers=headers)
print(f"   Status: {resp.status_code}")
if resp.status_code == 200:
    data = resp.json()
    recordings = data.get('value', [])
    print(f"   ✅ SUCCESS! Found {len(recordings)} recordings")
    for r in recordings[:5]:
        print(f"      - Meeting: {r.get('meetingId', '')[:30]}...")
        print(f"        Recording ID: {r.get('id', '')}")
elif resp.status_code == 403:
    print("   ❌ Access Denied - Policy may still be propagating")
else:
    print(f"   Response: {resp.text[:300]}")

# Method 3: Try to list online meetings with filter
print("\n3. Listing online meetings...")
# The filter needs a specific format for application permissions
url = f'https://graph.microsoft.com/v1.0/users/{USER_ID}/onlineMeetings?$filter=startDateTime ge 2024-01-01T00:00:00Z'
resp = requests.get(url, headers=headers)
print(f"   Status: {resp.status_code}")
if resp.status_code == 200:
    data = resp.json()
    meetings = data.get('value', [])
    print(f"   ✅ SUCCESS! Found {len(meetings)} meetings")
else:
    print(f"   Response: {resp.text[:300]}")

# Method 4: Use beta me endpoint variant
print("\n4. Getting AI insights for meetings...")
url = f'https://graph.microsoft.com/beta/users/{USER_ID}/onlineMeetings/getAllMeetingAiInsights(meetingOrganizerUserId=\'{USER_ID}\')'
resp = requests.get(url, headers=headers)
print(f"   Status: {resp.status_code}")
if resp.status_code == 200:
    data = resp.json()
    insights = data.get('value', [])
    print(f"   ✅ SUCCESS! Found {len(insights)} AI insights")
else:
    print(f"   Response: {resp.text[:200]}")

print("\n" + "=" * 70)
print("STATUS")
print("=" * 70)
if resp.status_code == 403:
    print("""
⏳ The Application Access Policy is still propagating.
   - Created at: Just now
   - Expected propagation: 15-60 minutes
   
   Please wait and try again in 15 minutes.
   
   Run: python src/test_transcript_api2.py
""")
else:
    print("Check the results above. If you see transcripts, the API is working!")
