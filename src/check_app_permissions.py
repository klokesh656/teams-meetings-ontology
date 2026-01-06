"""Check what permissions the app token actually has"""
import os
import base64
import json
from dotenv import load_dotenv
from azure.identity import ClientSecretCredential

load_dotenv()

CLIENT_ID = os.getenv('AZURE_CLIENT_ID')
CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET')
TENANT_ID = os.getenv('AZURE_TENANT_ID')

credential = ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
token = credential.get_token('https://graph.microsoft.com/.default').token

# Decode the JWT token to see its claims
parts = token.split('.')
# Add padding if needed
payload = parts[1] + '=' * (4 - len(parts[1]) % 4)
decoded = base64.urlsafe_b64decode(payload)
claims = json.loads(decoded)

print("=" * 70)
print("APP TOKEN PERMISSIONS (ROLES)")
print("=" * 70)
print(f"App ID: {claims.get('appid', 'N/A')}")
print(f"Tenant: {claims.get('tid', 'N/A')}")
print()

roles = claims.get('roles', [])
print(f"Granted Permissions ({len(roles)} total):")
print("-" * 50)

# Check for transcript-related permissions
transcript_perms = ['OnlineMeetings', 'OnlineMeetingTranscript', 'CallRecords', 'Communications']
has_transcript_perms = []

for role in sorted(roles):
    marker = ""
    for perm in transcript_perms:
        if perm in role:
            marker = " ⭐"
            has_transcript_perms.append(role)
    print(f"  ✓ {role}{marker}")

print()
print("=" * 70)
print("TRANSCRIPT-RELATED PERMISSIONS CHECK")
print("=" * 70)

required = [
    'OnlineMeetings.Read.All',
    'OnlineMeetingTranscript.Read.All',
    'CallRecords.Read.All'
]

for perm in required:
    if perm in roles:
        print(f"  ✅ {perm}")
    else:
        print(f"  ❌ {perm} - MISSING!")

if all(p in roles for p in required):
    print("\n✅ All required permissions are present!")
    print("   If still getting errors, wait for policy propagation (15-60 min)")
else:
    print("\n❌ Missing permissions! Add them in Azure AD > App registrations > API permissions")
