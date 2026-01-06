"""
Test transcript access via Communications API after Application Access Policy is set.
The policy may take 15-60 minutes to propagate.
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
print("Policy propagation can take 15-60 minutes after creation")
print("=" * 70)

# Test 1: List user's online meetings (requires OnlineMeetings.Read.All)
print("\n1. Testing OnlineMeetings API...")
# Need to use getAllTranscripts or get specific meeting by joinWebUrl
url = f'https://graph.microsoft.com/v1.0/users/{USER_ID}/onlineMeetings'
resp = requests.get(url, headers=headers)
print(f"   Status: {resp.status_code}")
if resp.status_code == 200:
    meetings = resp.json().get('value', [])
    print(f"   ✅ Success! Found {len(meetings)} meetings")
    for m in meetings[:3]:
        print(f"      - {m.get('subject', 'No subject')} ({m.get('id', '')[:20]}...)")
else:
    print(f"   ❌ Error: {resp.text[:200]}")

# Test 2: Try beta endpoint for transcripts
print("\n2. Testing Beta Transcripts API...")
url = f'https://graph.microsoft.com/beta/users/{USER_ID}/onlineMeetings/getAllTranscripts'
resp = requests.get(url, headers=headers)
print(f"   Status: {resp.status_code}")
if resp.status_code == 200:
    transcripts = resp.json().get('value', [])
    print(f"   ✅ Success! Found {len(transcripts)} transcripts")
else:
    print(f"   ❌ Error: {resp.text[:200]}")

# Test 3: Check Communications API permissions
print("\n3. Testing Communications.Read.All (Call Records)...")
url = 'https://graph.microsoft.com/v1.0/communications/callRecords'
resp = requests.get(url, headers=headers)
print(f"   Status: {resp.status_code}")
if resp.status_code == 200:
    print("   ✅ Success!")
else:
    print(f"   ❌ Error: {resp.text[:200]}")

# Test 4: Get a specific meeting by join URL (if we have one from recordings)
print("\n4. Getting sample meeting from recordings to test transcript access...")
rec_url = f'https://graph.microsoft.com/v1.0/users/{USER_ID}/drive/root:/Recordings:/children?$top=5'
rec_resp = requests.get(rec_url, headers=headers)
if rec_resp.status_code == 200:
    recordings = rec_resp.json().get('value', [])
    print(f"   Found {len(recordings)} recordings")
    for rec in recordings[:2]:
        name = rec.get('name', '')
        print(f"   - {name}")
else:
    print(f"   Could not access recordings: {rec_resp.status_code}")

print("\n" + "=" * 70)
print("SUMMARY")
print("=" * 70)
print("""
If you're still getting 403 errors:
1. Wait 15-60 minutes for policy propagation
2. Verify the app has these API permissions in Azure AD:
   - OnlineMeetings.Read.All (Application)
   - OnlineMeetingTranscript.Read.All (Application)
   - CallRecords.Read.All (Application)
   
3. Ensure Admin Consent was granted for these permissions

To check/add permissions:
- Go to portal.azure.com
- Azure Active Directory > App registrations
- Find your app (7b98108a-a799-45c1-aad6-93af90c1134c)
- API permissions > Add permission > Microsoft Graph > Application permissions
- Search for "OnlineMeetingTranscript" and add it
- Click "Grant admin consent"
""")
