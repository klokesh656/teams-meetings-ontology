"""Search all users for VTT transcript files."""
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

# Get all users
resp = requests.get('https://graph.microsoft.com/v1.0/users?$top=100', headers=headers)
users = resp.json().get('value', [])

print("="*70)
print("SCANNING ALL USERS FOR VTT TRANSCRIPT FILES")
print("="*70)

total_vtt = 0
users_with_vtt = []

for u in users:
    email = u.get('mail', '') or u.get('userPrincipalName', '') or ''
    name = u.get('displayName', '')
    user_id = u['id']
    
    try:
        # Check Recordings folder
        url = f'https://graph.microsoft.com/v1.0/users/{user_id}/drive/root:/Recordings:/children?$top=500'
        resp = requests.get(url, headers=headers, timeout=30)
        
        if resp.status_code == 200:
            items = resp.json().get('value', [])
            vtt_files = [i for i in items if i.get('name', '').lower().endswith('.vtt')]
            docx_transcripts = [i for i in items if i.get('name', '').lower().endswith('.docx') and 'transcript' in i.get('name', '').lower()]
            mp4_files = [i for i in items if i.get('name', '').lower().endswith('.mp4')]
            
            if vtt_files or docx_transcripts:
                total_vtt += len(vtt_files) + len(docx_transcripts)
                users_with_vtt.append({
                    'name': name,
                    'email': email,
                    'vtt': len(vtt_files),
                    'docx': len(docx_transcripts),
                    'mp4': len(mp4_files),
                    'files': vtt_files + docx_transcripts
                })
                print(f"\n✅ {name} ({email})")
                print(f"   VTT: {len(vtt_files)}, DOCX transcripts: {len(docx_transcripts)}, MP4: {len(mp4_files)}")
                for f in (vtt_files + docx_transcripts)[:5]:
                    print(f"   📝 {f.get('name')}")
            elif mp4_files:
                # User has recordings but no transcripts
                print(f"⚠️  {name}: {len(mp4_files)} MP4s, 0 transcripts")
    except Exception as e:
        print(f"❌ Error scanning {name}: {str(e)[:50]}")

print("\n" + "="*70)
print("SUMMARY")
print("="*70)
print(f"Total VTT/DOCX transcript files found: {total_vtt}")
print(f"Users with transcripts: {len(users_with_vtt)}")

if not users_with_vtt:
    print("\n❌ NO TRANSCRIPT FILES FOUND IN ANY USER'S RECORDINGS FOLDER")
    print("\nThis means transcripts are either:")
    print("  1. Not being saved to OneDrive (need to enable in Teams Admin)")
    print("  2. Stored in a different location (SharePoint, Stream)")
    print("  3. Only accessible via Communications API (need Application Access Policy)")
