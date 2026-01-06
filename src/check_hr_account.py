"""Check HR@Our-Assistants.com account for transcripts."""
import asyncio
import os
import sys
import requests
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.transcript_extractor import TranscriptExtractor
from dotenv import load_dotenv

load_dotenv()

async def check_hr_account():
    extractor = TranscriptExtractor(
        os.getenv('AZURE_CLIENT_ID'),
        os.getenv('AZURE_CLIENT_SECRET'),
        os.getenv('AZURE_TENANT_ID')
    )
    
    # Find HR account using Graph API directly
    print("Searching for HR@Our-Assistants.com account...")
    
    try:
        result = await extractor.graph_client.users.get()
        users = result.value if result else []
        
        hr_user = None
        all_users = []
        
        for u in users:
            email = getattr(u, 'mail', '') or getattr(u, 'user_principal_name', '') or ''
            name = getattr(u, 'display_name', '') or ''
            user_id = getattr(u, 'id', '')
            all_users.append({'id': user_id, 'name': name, 'email': email})
            
            if 'hr@our-assistants' in email.lower() or 'hr_our-assistants' in email.lower():
                hr_user = {'id': user_id, 'name': name, 'email': email}
                print(f"✅ Found HR account: {name} ({email})")
                break
        
        if not hr_user:
            print("\n❌ HR account not found directly. Searching for similar accounts...")
            for user in all_users:
                if 'hr' in user['email'].lower() or 'our-assistants' in user['email'].lower():
                    print(f"  - {user['name']} ({user['email']})")
            
            print("\nAll available accounts:")
            for user in all_users[:20]:
                print(f"  - {user['name']} ({user['email']})")
            return
        
        # List files in Recordings folder using REST API
        user_id = hr_user['id']
        print(f"\n📁 Scanning Recordings folder for {hr_user['name']}...")
        
        # Get token and use direct REST call
        token = extractor.credential.get_token("https://graph.microsoft.com/.default")
        headers = {"Authorization": f"Bearer {token.token}"}
        
        # Try to list Recordings folder
        url = f"https://graph.microsoft.com/v1.0/users/{user_id}/drive/root:/Recordings:/children"
        resp = requests.get(url, headers=headers)
        
        if resp.status_code == 200:
            data = resp.json()
            items = data.get('value', [])
            print(f"Found {len(items)} items:\n")
            vtt_files = []
            mp4_files = []
            other_files = []
            
            for item in items:
                name = item.get('name', '')
                size = item.get('size', 0)
                size_mb = size / (1024 * 1024)
                
                if name.lower().endswith('.vtt'):
                    vtt_files.append((name, size_mb))
                elif name.lower().endswith('.mp4'):
                    mp4_files.append((name, size_mb))
                else:
                    other_files.append((name, size_mb))
            
            # Show transcripts first
            if vtt_files:
                print(f"📝 TRANSCRIPTS ({len(vtt_files)} files):")
                for name, size in vtt_files[:10]:
                    print(f"   • {name} ({size:.2f} MB)")
                if len(vtt_files) > 10:
                    print(f"   ... and {len(vtt_files) - 10} more")
            else:
                print("📝 TRANSCRIPTS: None found")
            
            print()
            
            # Show recordings
            if mp4_files:
                print(f"🎥 RECORDINGS ({len(mp4_files)} files):")
                for name, size in mp4_files[:10]:
                    print(f"   • {name[:60]}... ({size:.1f} MB)")
                if len(mp4_files) > 10:
                    print(f"   ... and {len(mp4_files) - 10} more")
            
            # Show other files
            if other_files:
                print(f"\n📄 OTHER FILES ({len(other_files)} files):")
                for name, size in other_files[:10]:
                    print(f"   • {name}")
            
            print(f"\n{'='*60}")
            print(f"SUMMARY:")
            print(f"  Transcripts (VTT): {len(vtt_files)}")
            print(f"  Recordings (MP4):  {len(mp4_files)}")
            print(f"  Other files:       {len(other_files)}")
            print(f"{'='*60}")
        else:
            print(f"Error accessing Recordings folder: {resp.status_code}")
            print(resp.text[:500] if resp.text else "No error details")
            
    except Exception as e:
        print(f"Error: {e}")
        
        # Try to list root folder
        print("\nTrying to list root folder instead...")
        try:
            extractor2 = TranscriptExtractor(
                os.getenv('AZURE_CLIENT_ID'),
                os.getenv('AZURE_CLIENT_SECRET'),
                os.getenv('AZURE_TENANT_ID')
            )
            result = await extractor2.graph_client.users.get()
            users = result.value if result else []
            for u in users[:20]:
                email = getattr(u, 'mail', '') or getattr(u, 'user_principal_name', '') or ''
                name = getattr(u, 'display_name', '') or ''
                print(f"  - {name} ({email})")
        except Exception as e2:
            print(f"Error listing users: {e2}")

if __name__ == "__main__":
    asyncio.run(check_hr_account())
