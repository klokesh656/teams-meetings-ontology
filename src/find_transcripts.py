"""Search for DOCX and VTT transcript files matching MP4 recordings."""
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

# Get HR user
print("Finding HR user...")
resp = requests.get('https://graph.microsoft.com/v1.0/users', headers=headers)
users = resp.json().get('value', [])
hr_user = None
for u in users:
    email = u.get('mail', '') or u.get('userPrincipalName', '') or ''
    if 'hr@our-assistants' in email.lower():
        hr_user = u
        print(f"HR User: {u['displayName']} ({email})")
        break

if not hr_user:
    print("HR user not found!")
    exit(1)

user_id = hr_user['id']

# List ALL files in Recordings folder
print('\n' + '='*60)
print('Listing ALL files in Recordings folder...')
print('='*60)

url = f'https://graph.microsoft.com/v1.0/users/{user_id}/drive/root:/Recordings:/children?$top=500'
resp = requests.get(url, headers=headers)

if resp.status_code == 200:
    items = resp.json().get('value', [])
    
    mp4_files = []
    vtt_files = []
    docx_files = []
    other_files = []
    
    for item in items:
        name = item.get('name', '')
        size = item.get('size', 0)
        size_mb = size / (1024 * 1024)
        
        if name.lower().endswith('.mp4'):
            mp4_files.append((name, size_mb))
        elif name.lower().endswith('.vtt'):
            vtt_files.append((name, size_mb))
        elif name.lower().endswith('.docx'):
            docx_files.append((name, size_mb))
        else:
            other_files.append((name, size_mb))
    
    print(f'\nTotal items: {len(items)}')
    print(f'  MP4 recordings: {len(mp4_files)}')
    print(f'  VTT transcripts: {len(vtt_files)}')
    print(f'  DOCX files: {len(docx_files)}')
    print(f'  Other files: {len(other_files)}')
    
    if vtt_files:
        print('\n' + '='*60)
        print('VTT TRANSCRIPTS FOUND:')
        print('='*60)
        for name, size in vtt_files[:20]:
            print(f'  📝 {name} ({size:.2f} MB)')
        if len(vtt_files) > 20:
            print(f'  ... and {len(vtt_files) - 20} more')
    
    if docx_files:
        print('\n' + '='*60)
        print('DOCX FILES FOUND:')
        print('='*60)
        for name, size in docx_files[:20]:
            print(f'  📄 {name} ({size:.2f} MB)')
        if len(docx_files) > 20:
            print(f'  ... and {len(docx_files) - 20} more')
    
    if other_files:
        print('\n' + '='*60)
        print('OTHER FILES:')
        print('='*60)
        for name, size in other_files[:10]:
            print(f'  📁 {name}')
    
    # Check for matching pairs
    if vtt_files or docx_files:
        print('\n' + '='*60)
        print('CHECKING FOR TRANSCRIPT MATCHES:')
        print('='*60)
        
        matches = 0
        for mp4_name, _ in mp4_files[:20]:
            # Get base name without extension and common suffixes
            base_name = mp4_name.replace('.mp4', '').replace('-Meeting Recording', '')
            
            # Look for matching VTT
            for vtt_name, _ in vtt_files:
                vtt_base = vtt_name.replace('.vtt', '')
                if base_name in vtt_base or vtt_base in base_name:
                    print(f'  ✅ MATCH:')
                    print(f'     MP4:  {mp4_name[:60]}...')
                    print(f'     VTT:  {vtt_name}')
                    matches += 1
                    break
            
            # Look for matching DOCX
            for docx_name, _ in docx_files:
                docx_base = docx_name.replace('.docx', '')
                if base_name in docx_base or docx_base in base_name:
                    print(f'  ✅ MATCH:')
                    print(f'     MP4:  {mp4_name[:60]}...')
                    print(f'     DOCX: {docx_name}')
                    matches += 1
                    break
        
        print(f'\nTotal matches found: {matches}')
    else:
        print('\n❌ No VTT or DOCX transcript files found in Recordings folder.')
        
else:
    print(f'Error: {resp.status_code}')
    print(resp.text[:500])

# Also search in Documents folder root
print('\n' + '='*60)
print('Searching entire OneDrive for transcript files...')
print('='*60)

# Search for VTT files
search_url = f"https://graph.microsoft.com/v1.0/users/{user_id}/drive/root/search(q='transcript')"
resp = requests.get(search_url, headers=headers)
if resp.status_code == 200:
    results = resp.json().get('value', [])
    if results:
        print(f'\nFound {len(results)} files matching "transcript":')
        for item in results[:15]:
            name = item.get('name', '')
            path = item.get('parentReference', {}).get('path', '')
            print(f'  • {name}')
            print(f'    Path: {path}')
    else:
        print('No files found matching "transcript"')
else:
    print(f'Search error: {resp.status_code}')

# Search for .vtt extension
search_url = f"https://graph.microsoft.com/v1.0/users/{user_id}/drive/root/search(q='.vtt')"
resp = requests.get(search_url, headers=headers)
if resp.status_code == 200:
    results = resp.json().get('value', [])
    if results:
        print(f'\nFound {len(results)} .vtt files:')
        for item in results[:15]:
            name = item.get('name', '')
            path = item.get('parentReference', {}).get('path', '')
            size = item.get('size', 0)
            print(f'  • {name} ({size/1024:.1f} KB)')
            print(f'    Path: {path}')
    else:
        print('No .vtt files found')
