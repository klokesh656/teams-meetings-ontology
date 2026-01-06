"""
Microsoft Teams Transcript Extractor
Extracts meeting transcripts using Microsoft Graph API
"""

import os
import json
import re
from datetime import datetime
from typing import List, Dict, Optional
from pathlib import Path
from dotenv import load_dotenv
import requests

from msal import ConfidentialClientApplication
from msgraph import GraphServiceClient
from msgraph.generated.users.item.online_meetings.online_meetings_request_builder import OnlineMeetingsRequestBuilder
from azure.identity import ClientSecretCredential
from kiota_abstractions.base_request_configuration import RequestConfiguration

# Optional imports for Excel and Blob storage
try:
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False

try:
    from azure.storage.blob import BlobServiceClient, generate_blob_sas, BlobSasPermissions
    from datetime import timedelta
    BLOB_AVAILABLE = True
except ImportError:
    BLOB_AVAILABLE = False


class TranscriptExtractor:
    """Extracts Microsoft Teams meeting transcripts and saves them locally."""
    
    def __init__(self, client_id: str, client_secret: str, tenant_id: str, output_dir: str = "transcripts"):
        """
        Initialize the transcript extractor.
        
        Args:
            client_id: Azure AD application client ID
            client_secret: Azure AD application client secret
            tenant_id: Azure AD tenant ID
            output_dir: Directory to save transcripts (default: "transcripts")
        """
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
        
        # Initialize Graph client
        self.credential = ClientSecretCredential(
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret
        )
        self.graph_client = GraphServiceClient(credentials=self.credential)
    
    async def verify_user_exists(self, user_id: str) -> bool:
        """
        Verify that a user exists in the organization.
        
        Args:
            user_id: User principal name or user ID
            
        Returns:
            True if user exists, False otherwise
        """
        try:
            user = await self.graph_client.users.by_user_id(user_id).get()
            if user:
                print(f"✓ User found: {user.display_name} ({user.mail or user.user_principal_name})")
                return True
            return False
        except Exception as e:
            print(f"✗ User not found: {user_id}")
            print(f"  Error: {e}")
            return False
    
    async def list_users(self, top: int = 10) -> List:
        """
        List users in the organization.
        
        Args:
            top: Number of users to retrieve
            
        Returns:
            List of users
        """
        try:
            from msgraph.generated.users.users_request_builder import UsersRequestBuilder
            query_params = UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
                top=top,
                select=["displayName", "mail", "userPrincipalName"]
            )
            request_config = UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(
                query_parameters=query_params
            )
            users = await self.graph_client.users.get(request_configuration=request_config)
            return users.value if users else []
        except Exception as e:
            print(f"Error listing users: {e}")
            return []

    async def get_communications_meetings(self, max_items: int = 500) -> List[Dict]:
        """
        Call the /communications/onlineMeetings endpoint directly using an access token.

        The Graph communications endpoint does not accept $top in this tenant, so
        this method paginates using @odata.nextLink and returns up to `max_items`.

        Returns a list of meeting objects (raw JSON) or empty list on failure.
        """
        try:
            # Acquire a token for Microsoft Graph
            token = self.credential.get_token("https://graph.microsoft.com/.default")
            headers = {
                "Authorization": f"Bearer {token.token}",
                "Accept": "application/json"
            }

            url = "https://graph.microsoft.com/v1.0/communications/onlineMeetings"
            results: List[Dict] = []

            while url and len(results) < max_items:
                resp = requests.get(url, headers=headers, timeout=30)
                if resp.status_code != 200:
                    print(f"communications endpoint returned {resp.status_code}: {resp.text}")
                    break

                data = resp.json()
                items = data.get("value", [])
                results.extend(items)

                # Pagination
                next_link = data.get("@odata.nextLink")
                url = next_link if next_link else None

                # safety: stop if no next_link
                if not next_link:
                    break

            # Trim to max_items
            if len(results) > max_items:
                results = results[:max_items]

            return results
        except Exception as e:
            print(f"Error calling communications endpoint: {e}")
            return []

    async def post_get_all_online_meetings(self, start_datetime: str, end_datetime: str) -> List[Dict]:
        """
        Call the POST action to retrieve online meetings in a time range.

        This uses the action POST /communications/getAllOnlineMeetings with a JSON body
        containing `startDateTime` and `endDateTime` in ISO 8601 format.
        Returns a list of meetings or empty list on failure.
        """
        try:
            token = self.credential.get_token("https://graph.microsoft.com/.default")
            headers = {
                "Authorization": f"Bearer {token.token}",
                "Accept": "application/json",
                "Content-Type": "application/json"
            }

            url = "https://graph.microsoft.com/v1.0/communications/getAllOnlineMeetings"
            body = {
                "startDateTime": start_datetime,
                "endDateTime": end_datetime
            }

            resp = requests.post(url, headers=headers, json=body, timeout=60)
            if resp.status_code == 200:
                data = resp.json()
                return data.get("value", []) or data
            else:
                print(f"POST communications action returned {resp.status_code}: {resp.text}")
                return []
        except Exception as e:
            print(f"Error calling getAllOnlineMeetings action: {e}")
            return []

    async def get_communications_meetings_beta(self, max_items: int = 500) -> List[Dict]:
        """
        Try the beta communications endpoint: /beta/communications/onlineMeetings

        Some tenants expose additional communications actions in beta. This method
        paginatedly fetches items from the beta endpoint and returns a list.
        """
        try:
            token = self.credential.get_token("https://graph.microsoft.com/.default")
            headers = {
                "Authorization": f"Bearer {token.token}",
                "Accept": "application/json"
            }

            url = "https://graph.microsoft.com/beta/communications/onlineMeetings"
            results: List[Dict] = []

            while url and len(results) < max_items:
                resp = requests.get(url, headers=headers, timeout=30)
                if resp.status_code != 200:
                    print(f"beta communications endpoint returned {resp.status_code}: {resp.text}")
                    break

                data = resp.json()
                items = data.get("value", [])
                results.extend(items)

                next_link = data.get("@odata.nextLink")
                url = next_link if next_link else None
                if not next_link:
                    break

            if len(results) > max_items:
                results = results[:max_items]
            return results
        except Exception as e:
            print(f"Error calling beta communications endpoint: {e}")
            return []

    async def search_user_drive(self, user_id: str, query: str = "transcript") -> List[Dict]:
        """
        Search a user's OneDrive for files matching the query.
        
        Teams cloud recordings and transcripts are typically saved to the organizer's
        OneDrive in a folder like "Recordings" with .vtt transcript files.
        
        Args:
            user_id: User principal name or user ID
            query: Search query (default: "transcript")
            
        Returns:
            List of matching file items (dict with id, name, webUrl, etc.)
        """
        try:
            token = self.credential.get_token("https://graph.microsoft.com/.default")
            headers = {
                "Authorization": f"Bearer {token.token}",
                "Accept": "application/json"
            }
            
            # Search the user's drive
            url = f"https://graph.microsoft.com/v1.0/users/{user_id}/drive/root/search(q='{query}')"
            results: List[Dict] = []
            
            while url:
                resp = requests.get(url, headers=headers, timeout=30)
                if resp.status_code == 200:
                    data = resp.json()
                    items = data.get("value", [])
                    results.extend(items)
                    url = data.get("@odata.nextLink")
                else:
                    print(f"  Drive search returned {resp.status_code}: {resp.text[:500]}")
                    break
            
            return results
        except Exception as e:
            print(f"Error searching drive for {user_id}: {e}")
            return []

    async def search_drives_for_transcripts(self, user_emails: List[str], download: bool = False) -> Dict[str, any]:
        """
        Search multiple users' OneDrive for transcript/recording files.
        
        Args:
            user_emails: List of user emails to search
            download: Whether to download found files (default: False)
            
        Returns:
            Dictionary with search results
        """
        results = {
            "users_searched": 0,
            "files_found": [],
            "files_downloaded": 0,
            "errors": []
        }
        
        # Common search terms for Teams recordings/transcripts
        search_terms = ["transcript", "recording", ".vtt"]
        
        for email in user_emails:
            results["users_searched"] += 1
            print(f"\n🔍 Searching OneDrive for: {email}")
            
            for term in search_terms:
                try:
                    items = await self.search_user_drive(email, term)
                    for item in items:
                        file_info = {
                            "user": email,
                            "name": item.get("name", ""),
                            "id": item.get("id", ""),
                            "webUrl": item.get("webUrl", ""),
                            "size": item.get("size", 0),
                            "lastModified": item.get("lastModifiedDateTime", ""),
                            "downloadUrl": item.get("@microsoft.graph.downloadUrl", "")
                        }
                        
                        # Check if it's a transcript-like file
                        name_lower = file_info["name"].lower()
                        if any(ext in name_lower for ext in [".vtt", ".txt", "transcript", "recording"]):
                            # Avoid duplicates
                            if not any(f["id"] == file_info["id"] for f in results["files_found"]):
                                results["files_found"].append(file_info)
                                print(f"  📄 {file_info['name']} ({file_info['size']} bytes)")
                                print(f"     Modified: {file_info['lastModified']}")
                                print(f"     URL: {file_info['webUrl']}")
                                
                                if download and file_info["downloadUrl"]:
                                    try:
                                        self._download_file(file_info)
                                        results["files_downloaded"] += 1
                                    except Exception as e:
                                        results["errors"].append(f"Download error for {file_info['name']}: {e}")
                except Exception as e:
                    results["errors"].append(f"Search error for {email}/{term}: {e}")
        
        return results

    def _download_file(self, file_info: Dict) -> str:
        """
        Download a file from OneDrive to the output directory.
        
        Args:
            file_info: Dictionary with file information including downloadUrl
            
        Returns:
            Path to downloaded file
        """
        token = self.credential.get_token("https://graph.microsoft.com/.default")
        headers = {"Authorization": f"Bearer {token.token}"}
        
        # Create user subdirectory
        user_dir = self.output_dir / file_info["user"].replace("@", "_at_")
        user_dir.mkdir(exist_ok=True)
        
        # Download the file
        download_url = file_info["downloadUrl"]
        resp = requests.get(download_url, headers=headers, timeout=60)
        
        if resp.status_code == 200:
            filepath = user_dir / file_info["name"]
            with open(filepath, 'wb') as f:
                f.write(resp.content)
            print(f"  ✅ Downloaded: {filepath}")
            return str(filepath)
        else:
            raise Exception(f"Download failed with status {resp.status_code}")

    async def list_recordings_folder(self, user_id: str) -> List[Dict]:
        """
        List files in a user's OneDrive /Documents/Recordings folder.
        
        Teams cloud recordings are typically saved here.
        
        Args:
            user_id: User principal name or user ID
            
        Returns:
            List of file items in the Recordings folder
        """
        try:
            token = self.credential.get_token("https://graph.microsoft.com/.default")
            headers = {
                "Authorization": f"Bearer {token.token}",
                "Accept": "application/json"
            }
            
            # Try common paths for recordings
            paths_to_try = [
                "/Documents/Recordings",
                "/Recordings",
                "/Documents/Microsoft Teams Chat Files",
                "/Documents"
            ]
            
            results: List[Dict] = []
            
            for path in paths_to_try:
                url = f"https://graph.microsoft.com/v1.0/users/{user_id}/drive/root:{path}:/children"
                resp = requests.get(url, headers=headers, timeout=30)
                
                if resp.status_code == 200:
                    data = resp.json()
                    items = data.get("value", [])
                    if items:
                        print(f"  📁 Found {len(items)} items in {path}")
                        for item in items:
                            item["_folder_path"] = path
                        results.extend(items)
                        
                        # Paginate if needed
                        next_link = data.get("@odata.nextLink")
                        while next_link:
                            resp = requests.get(next_link, headers=headers, timeout=30)
                            if resp.status_code == 200:
                                data = resp.json()
                                for item in data.get("value", []):
                                    item["_folder_path"] = path
                                results.extend(data.get("value", []))
                                next_link = data.get("@odata.nextLink")
                            else:
                                break
                elif resp.status_code == 404:
                    pass  # Folder doesn't exist, try next
                else:
                    print(f"  ⚠ Error accessing {path}: {resp.status_code}")
            
            return results
        except Exception as e:
            print(f"Error listing recordings folder for {user_id}: {e}")
            return []

    async def search_all_users_recordings(self, max_users: int = 20, download: bool = False) -> Dict[str, any]:
        """
        Search all users' Recordings folders for meeting recordings and transcripts.
        
        Args:
            max_users: Maximum number of users to scan
            download: Whether to download found files
            
        Returns:
            Dictionary with search results
        """
        results = {
            "users_searched": 0,
            "users_with_recordings": [],
            "files_found": [],
            "transcript_files": [],
            "recording_files": [],
            "files_downloaded": 0,
            "errors": []
        }
        
        print(f"Scanning up to {max_users} users for Recordings folders...")
        print("=" * 60)
        
        users = await self.list_users(max_users)
        
        for user in users:
            email = user.mail or user.user_principal_name
            if not email:
                continue
                
            results["users_searched"] += 1
            print(f"\n👤 {user.display_name} ({email})")
            
            try:
                items = await self.list_recordings_folder(email)
                
                if items:
                    user_files = []
                    for item in items:
                        name = item.get("name", "")
                        name_lower = name.lower()
                        
                        file_info = {
                            "user": email,
                            "user_name": user.display_name,
                            "name": name,
                            "id": item.get("id", ""),
                            "webUrl": item.get("webUrl", ""),
                            "size": item.get("size", 0),
                            "lastModified": item.get("lastModifiedDateTime", ""),
                            "downloadUrl": item.get("@microsoft.graph.downloadUrl", ""),
                            "folder_path": item.get("_folder_path", "")
                        }
                        
                        # Categorize files
                        is_transcript = any(ext in name_lower for ext in [".vtt", ".docx", "transcript"])
                        is_recording = any(ext in name_lower for ext in [".mp4", ".mp3", ".m4a", "recording"])
                        
                        if is_transcript:
                            results["transcript_files"].append(file_info)
                            print(f"  📝 TRANSCRIPT: {name}")
                        elif is_recording:
                            results["recording_files"].append(file_info)
                            print(f"  🎥 RECORDING: {name}")
                        
                        if is_transcript or is_recording:
                            results["files_found"].append(file_info)
                            user_files.append(file_info)
                            
                            if download and is_transcript:
                                try:
                                    # Get download URL if not present
                                    if not file_info["downloadUrl"]:
                                        file_info["downloadUrl"] = await self._get_file_download_url(email, file_info["id"])
                                    if file_info["downloadUrl"]:
                                        self._download_file(file_info)
                                        results["files_downloaded"] += 1
                                except Exception as e:
                                    results["errors"].append(f"Download error: {e}")
                    
                    if user_files:
                        results["users_with_recordings"].append({
                            "name": user.display_name,
                            "email": email,
                            "file_count": len(user_files)
                        })
            except Exception as e:
                results["errors"].append(f"Error for {email}: {e}")
        
        return results

    async def _get_file_download_url(self, user_id: str, file_id: str) -> Optional[str]:
        """Get the download URL for a specific file."""
        try:
            token = self.credential.get_token("https://graph.microsoft.com/.default")
            headers = {
                "Authorization": f"Bearer {token.token}",
                "Accept": "application/json"
            }
            
            url = f"https://graph.microsoft.com/v1.0/users/{user_id}/drive/items/{file_id}"
            resp = requests.get(url, headers=headers, timeout=30)
            
            if resp.status_code == 200:
                data = resp.json()
                return data.get("@microsoft.graph.downloadUrl")
            return None
        except:
            return None
        
    async def get_user_meetings(self, user_id: str) -> List:
        """
        Get all online meetings for a specific user.
        
        Args:
            user_id: User principal name or user ID
            
        Returns:
            List of online meetings
        """
        try:
            meetings = await self.graph_client.users.by_user_id(user_id).online_meetings.get()
            return meetings.value if meetings else []
        except Exception as e:
            error_msg = str(e)
            print(f"\n⚠ Unable to access meetings for {user_id}")
            
            if "403" in error_msg or "Authorization" in error_msg or "Forbidden" in error_msg:
                print("\n🚫 Permission Denied Error")
                print("  Your application is missing required API permissions.")
                print("\n  Required Microsoft Graph Application Permissions:")
                print("  ✓ User.Read.All (you have this)")
                print("  ✗ OnlineMeetings.Read.All (MISSING - add this!)")
                print("  ✗ OnlineMeetings.ReadWrite.All (alternative)")
                print("\n  Steps to fix:")
                print("  1. Go to: https://portal.azure.com")
                print("  2. Azure AD → App registrations → Your app")
                print("  3. API permissions → Add permission → Microsoft Graph")
                print("  4. Application permissions → OnlineMeetings.Read.All")
                print("  5. Grant admin consent")
            elif "404" in error_msg or "UnknownError" in error_msg:
                print("  Possible reasons:")
                print("  1. User has no online meetings")
                print("  2. Meetings may not be accessible through this API endpoint")
            else:
                print(f"  Error details: {e}")
            
            return []
    
    async def get_meeting_transcripts(self, user_id: str, meeting_id: str) -> List:
        """
        Get all transcripts for a specific meeting.
        
        Args:
            user_id: User principal name or user ID
            meeting_id: Online meeting ID
            
        Returns:
            List of transcripts
        """
        try:
            transcripts = await self.graph_client.users.by_user_id(user_id).online_meetings.by_online_meeting_id(meeting_id).transcripts.get()
            return transcripts.value if transcripts else []
        except Exception as e:
            print(f"Error fetching transcripts for meeting {meeting_id}: {e}")
            return []
    
    async def get_transcript_content(self, user_id: str, meeting_id: str, transcript_id: str) -> Optional[str]:
        """
        Get the content of a specific transcript.
        
        Args:
            user_id: User principal name or user ID
            meeting_id: Online meeting ID
            transcript_id: Transcript ID
            
        Returns:
            Transcript content as string
        """
        try:
            content = await self.graph_client.users.by_user_id(user_id).online_meetings.by_online_meeting_id(meeting_id).transcripts.by_call_transcript_id(transcript_id).content.get()
            return content.decode('utf-8') if content else None
        except Exception as e:
            print(f"Error fetching transcript content {transcript_id}: {e}")
            return None
    
    def save_transcript(self, content: str, meeting_id: str, transcript_id: str, created_date: datetime) -> str:
        """
        Save transcript content to a file.
        
        Args:
            content: Transcript content
            meeting_id: Meeting ID
            transcript_id: Transcript ID
            created_date: Transcript creation date
            
        Returns:
            Path to saved file
        """
        # Create subdirectory for the meeting
        date_str = created_date.strftime("%Y-%m-%d")
        meeting_dir = self.output_dir / f"{date_str}_{meeting_id[:8]}"
        meeting_dir.mkdir(exist_ok=True)
        
        # Save transcript file
        filename = f"transcript_{transcript_id[:8]}.vtt"
        filepath = meeting_dir / filename
        
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(content)
        
        # Save metadata
        metadata = {
            "meeting_id": meeting_id,
            "transcript_id": transcript_id,
            "created_date": created_date.isoformat(),
            "saved_date": datetime.now().isoformat()
        }
        
        metadata_file = meeting_dir / f"metadata_{transcript_id[:8]}.json"
        with open(metadata_file, 'w', encoding='utf-8') as f:
            json.dump(metadata, f, indent=2)
        
        return str(filepath)
    
    async def scan_all_users_for_meetings(self, max_users: int = 50) -> Dict[str, any]:
        """
        Scan all users in the organization to find who has meetings with transcripts.
        
        Args:
            max_users: Maximum number of users to scan
            
        Returns:
            Dictionary with scan results
        """
        print(f"Scanning up to {max_users} users for meetings...")
        print("=" * 60)
        
        users = await self.list_users(max_users)
        results = {
            "users_scanned": 0,
            "users_with_meetings": [],
            "total_meetings": 0,
            "total_transcripts": 0
        }
        
        for user in users:
            email = user.mail or user.user_principal_name
            results["users_scanned"] += 1
            
            try:
                meetings = await self.graph_client.users.by_user_id(user.id).online_meetings.get()
                if meetings and meetings.value and len(meetings.value) > 0:
                    meeting_count = len(meetings.value)
                    results["total_meetings"] += meeting_count
                    
                    # Check for transcripts
                    transcript_count = 0
                    for meeting in meetings.value:
                        try:
                            transcripts = await self.graph_client.users.by_user_id(user.id).online_meetings.by_online_meeting_id(meeting.id).transcripts.get()
                            if transcripts and transcripts.value:
                                transcript_count += len(transcripts.value)
                        except:
                            pass
                    
                    results["total_transcripts"] += transcript_count
                    results["users_with_meetings"].append({
                        "name": user.display_name,
                        "email": email,
                        "meetings": meeting_count,
                        "transcripts": transcript_count
                    })
                    
                    icon = "📝" if transcript_count > 0 else "📅"
                    print(f"{icon} {user.display_name} ({email})")
                    print(f"   Meetings: {meeting_count} | Transcripts: {transcript_count}")
            except:
                pass
        
        return results
    
    async def extract_all_transcripts(self, user_id: str, verify_user: bool = True) -> Dict[str, int]:
        """
        Extract all transcripts for a user's meetings.
        
        Args:
            user_id: User principal name or user ID
            verify_user: Whether to verify user exists first (default: True)
            
        Returns:
            Dictionary with extraction statistics
        """
        stats = {
            "meetings_found": 0,
            "transcripts_found": 0,
            "transcripts_saved": 0,
            "errors": 0
        }
        
        if verify_user:
            print(f"Verifying user: {user_id}")
            if not await self.verify_user_exists(user_id):
                print("\n❌ Cannot proceed: User verification failed")
                print("\nTip: Run with --list-users to see available users")
                return stats
        
        print(f"\nFetching meetings for user: {user_id}")
        meetings = await self.get_user_meetings(user_id)
        stats["meetings_found"] = len(meetings)
        
        if len(meetings) == 0:
            print("\n⚠ No meetings found for this user.")
            print("\nPossible reasons:")
            print("  • User hasn't hosted or attended any Teams meetings")
            print("  • Meetings exist but aren't accessible via this API")
            print("  • Try running: python src/transcript_extractor.py --scan-all")
            print("    to find users who have meetings with transcripts")
        else:
            print(f"Found {len(meetings)} meetings")
        
        for meeting in meetings:
            meeting_id = meeting.id
            print(f"\nProcessing meeting: {meeting_id}")
            
            transcripts = await self.get_meeting_transcripts(user_id, meeting_id)
            stats["transcripts_found"] += len(transcripts)
            
            for transcript in transcripts:
                transcript_id = transcript.id
                created_date = transcript.created_date_time
                
                print(f"  Downloading transcript: {transcript_id[:8]}...")
                content = await self.get_transcript_content(user_id, meeting_id, transcript_id)
                
                if content:
                    try:
                        filepath = self.save_transcript(content, meeting_id, transcript_id, created_date)
                        stats["transcripts_saved"] += 1
                        print(f"  ✓ Saved to: {filepath}")
                    except Exception as e:
                        print(f"  ✗ Error saving transcript: {e}")
                        stats["errors"] += 1
                else:
                    print(f"  ✗ Failed to download transcript content")
                    stats["errors"] += 1
        
        return stats


    def parse_vtt_metadata(self, vtt_content: str, filename: str = "") -> Dict:
        """
        Parse VTT transcript content to extract metadata.
        
        Args:
            vtt_content: The raw VTT file content
            filename: Original filename for additional context
            
        Returns:
            Dictionary with extracted metadata
        """
        metadata = {
            "duration_seconds": 0,
            "participant_count": 0,
            "participants": [],
            "word_count": 0,
            "line_count": 0,
            "first_timestamp": None,
            "last_timestamp": None,
            "meeting_subject": ""
        }
        
        try:
            # Extract meeting subject from filename
            # Format: "Meeting Subject-YYYYMMDD_HHMMSS-Meeting Recording.vtt"
            if filename:
                # Remove common suffixes
                clean_name = filename.replace("-Meeting Recording.vtt", "")
                clean_name = clean_name.replace("-Transcript.vtt", "")
                clean_name = clean_name.replace(".vtt", "")
                
                # Try to extract date and subject
                date_match = re.search(r'-(\d{8}_\d{6})', clean_name)
                if date_match:
                    metadata["meeting_subject"] = clean_name[:date_match.start()]
                else:
                    metadata["meeting_subject"] = clean_name
            
            # Parse VTT content
            lines = vtt_content.split('\n')
            metadata["line_count"] = len(lines)
            
            participants = set()
            timestamps = []
            text_content = []
            
            # VTT timestamp pattern: 00:00:00.000 --> 00:00:05.000
            timestamp_pattern = re.compile(r'(\d{2}:\d{2}:\d{2}\.\d{3})\s*-->\s*(\d{2}:\d{2}:\d{2}\.\d{3})')
            # Speaker pattern: <v Speaker Name>text</v> or Speaker Name: text
            speaker_pattern1 = re.compile(r'<v\s+([^>]+)>')
            speaker_pattern2 = re.compile(r'^([A-Za-z\s\.\-\']+):\s*(.+)$')
            
            for line in lines:
                line = line.strip()
                
                # Extract timestamps
                ts_match = timestamp_pattern.search(line)
                if ts_match:
                    start_ts = ts_match.group(1)
                    end_ts = ts_match.group(2)
                    timestamps.append((start_ts, end_ts))
                
                # Extract speakers from VTT voice tags
                speaker_matches = speaker_pattern1.findall(line)
                for speaker in speaker_matches:
                    participants.add(speaker.strip())
                
                # Extract speakers from "Name: text" format
                if ':' in line and not '-->' in line:
                    speaker_match = speaker_pattern2.match(line)
                    if speaker_match:
                        speaker = speaker_match.group(1).strip()
                        # Filter out common non-speaker patterns
                        if len(speaker) > 2 and not speaker.isdigit():
                            participants.add(speaker)
                
                # Count words (exclude timestamps and metadata)
                if not ts_match and not line.startswith('WEBVTT') and line:
                    # Remove VTT tags for word counting
                    clean_text = re.sub(r'<[^>]+>', '', line)
                    text_content.append(clean_text)
            
            # Calculate metadata
            metadata["participants"] = sorted(list(participants))
            metadata["participant_count"] = len(participants)
            metadata["word_count"] = sum(len(t.split()) for t in text_content)
            
            # Calculate duration from timestamps
            if timestamps:
                metadata["first_timestamp"] = timestamps[0][0]
                metadata["last_timestamp"] = timestamps[-1][1]
                
                # Parse last timestamp to get duration
                last_ts = timestamps[-1][1]
                parts = last_ts.split(':')
                if len(parts) == 3:
                    hours = int(parts[0])
                    minutes = int(parts[1])
                    seconds = float(parts[2])
                    metadata["duration_seconds"] = int(hours * 3600 + minutes * 60 + seconds)
            
        except Exception as e:
            print(f"  ⚠ Error parsing VTT metadata: {e}")
        
        return metadata

    def parse_filename_metadata(self, filename: str, last_modified: str = "") -> Dict:
        """
        Parse meeting metadata from filename and modification date.
        
        Args:
            filename: The file name (e.g., "Team Meeting-20231201_140000-Meeting Recording.vtt")
            last_modified: ISO format datetime string
            
        Returns:
            Dictionary with parsed metadata
        """
        metadata = {
            "meeting_subject": "",
            "meeting_date": None,
            "meeting_time": None,
            "file_type": "",
            "is_transcript": False,
            "is_recording": False
        }
        
        try:
            # Determine file type
            filename_lower = filename.lower()
            if '.vtt' in filename_lower or 'transcript' in filename_lower:
                metadata["file_type"] = "transcript"
                metadata["is_transcript"] = True
            elif any(ext in filename_lower for ext in ['.mp4', '.mp3', '.m4a']):
                metadata["file_type"] = "recording"
                metadata["is_recording"] = True
            
            # Extract date from filename pattern: -YYYYMMDD_HHMMSS-
            date_match = re.search(r'-(\d{4})(\d{2})(\d{2})_(\d{2})(\d{2})(\d{2})-', filename)
            if date_match:
                year, month, day, hour, minute, second = date_match.groups()
                metadata["meeting_date"] = f"{year}-{month}-{day}"
                metadata["meeting_time"] = f"{hour}:{minute}:{second}"
                
                # Extract subject (everything before the date)
                subject_end = filename.find(f"-{year}{month}{day}")
                if subject_end > 0:
                    metadata["meeting_subject"] = filename[:subject_end]
            elif last_modified:
                # Use last modified date if no date in filename
                try:
                    dt = datetime.fromisoformat(last_modified.replace('Z', '+00:00'))
                    metadata["meeting_date"] = dt.strftime("%Y-%m-%d")
                    metadata["meeting_time"] = dt.strftime("%H:%M:%S")
                except:
                    pass
            
            # Clean up subject
            if not metadata["meeting_subject"]:
                # Remove common suffixes and extensions
                clean_name = filename
                for suffix in ['-Meeting Recording.vtt', '-Transcript.vtt', '.vtt', '.mp4', '.mp3', '.m4a']:
                    clean_name = clean_name.replace(suffix, '')
                metadata["meeting_subject"] = clean_name
                
        except Exception as e:
            print(f"  ⚠ Error parsing filename metadata: {e}")
        
        return metadata

    async def download_and_parse_transcript(self, user_id: str, file_info: Dict) -> Dict:
        """
        Download a transcript file and parse its content for metadata.
        
        Args:
            user_id: User principal name
            file_info: Dictionary with file information
            
        Returns:
            Dictionary with full metadata including parsed content
        """
        metadata = {
            "file_id": file_info.get("id", ""),
            "file_name": file_info.get("name", ""),
            "web_url": file_info.get("webUrl", ""),
            "size_bytes": file_info.get("size", 0),
            "last_modified": file_info.get("lastModified", ""),
            "user_email": file_info.get("user", user_id),
            "user_name": file_info.get("user_name", ""),
            "content_parsed": False,
            "error": None
        }
        
        # Parse filename metadata
        filename_meta = self.parse_filename_metadata(
            file_info.get("name", ""), 
            file_info.get("lastModified", "")
        )
        metadata.update(filename_meta)
        
        # Try to download and parse VTT content
        try:
            download_url = file_info.get("downloadUrl")
            if not download_url:
                download_url = await self._get_file_download_url(user_id, file_info.get("id"))
            
            if download_url:
                token = self.credential.get_token("https://graph.microsoft.com/.default")
                headers = {"Authorization": f"Bearer {token.token}"}
                
                resp = requests.get(download_url, headers=headers, timeout=60)
                if resp.status_code == 200:
                    content = resp.content.decode('utf-8', errors='ignore')
                    vtt_meta = self.parse_vtt_metadata(content, file_info.get("name", ""))
                    metadata.update(vtt_meta)
                    metadata["content_parsed"] = True
                else:
                    metadata["error"] = f"Download failed: {resp.status_code}"
            else:
                metadata["error"] = "No download URL available"
        except Exception as e:
            metadata["error"] = str(e)
        
        return metadata

    async def extract_metadata_for_export(self, max_users: int = 30, download_content: bool = True, 
                                         include_recordings: bool = True) -> List[Dict]:
        """
        Extract metadata from all transcript and recording files for Excel export.
        
        Args:
            max_users: Maximum number of users to scan
            download_content: Whether to download and parse VTT content
            include_recordings: Whether to include recording files (mp4) in export
            
        Returns:
            List of metadata dictionaries
        """
        print(f"📊 Extracting metadata for Excel export...")
        print(f"   Max users: {max_users}")
        print(f"   Parse content: {download_content}")
        print(f"   Include recordings: {include_recordings}")
        print("=" * 60)
        
        all_metadata = []
        
        # Scan for transcript and recording files
        results = await self.search_all_users_recordings(max_users=max_users, download=False)
        
        # Determine which files to process
        files_to_process = results['transcript_files'].copy()
        if include_recordings:
            files_to_process.extend(results['recording_files'])
        
        print(f"\n📝 Processing {len(files_to_process)} files...")
        print(f"   Transcript files: {len(results['transcript_files'])}")
        print(f"   Recording files: {len(results['recording_files']) if include_recordings else 0}")
        
        for i, file_info in enumerate(files_to_process):
            filename = file_info.get('name', '')[:50]
            print(f"  [{i+1}/{len(files_to_process)}] {filename}...")
            
            # Check if it's a VTT file that we can parse
            is_vtt = file_info.get('name', '').lower().endswith('.vtt')
            
            if download_content and is_vtt:
                metadata = await self.download_and_parse_transcript(
                    file_info.get('user', ''),
                    file_info
                )
            else:
                # Just use filename-based metadata
                metadata = {
                    "file_id": file_info.get("id", ""),
                    "file_name": file_info.get("name", ""),
                    "web_url": file_info.get("webUrl", ""),
                    "size_bytes": file_info.get("size", 0),
                    "last_modified": file_info.get("lastModified", ""),
                    "user_email": file_info.get("user", ""),
                    "user_name": file_info.get("user_name", ""),
                    "content_parsed": False
                }
                filename_meta = self.parse_filename_metadata(
                    file_info.get("name", ""),
                    file_info.get("lastModified", "")
                )
                metadata.update(filename_meta)
            
            # Add unique meeting ID
            metadata["meeting_id"] = f"{metadata.get('meeting_date', 'unknown')}_{metadata.get('file_id', '')[:8]}"
            
            all_metadata.append(metadata)
        
        return all_metadata

    def export_to_excel(self, metadata_list: List[Dict], output_file: str = "meeting_metadata.xlsx") -> str:
        """
        Export metadata to an Excel file with the schema for Power Automate.
        
        Args:
            metadata_list: List of metadata dictionaries
            output_file: Output Excel file path
            
        Returns:
            Path to the created Excel file
        """
        if not PANDAS_AVAILABLE:
            raise ImportError("pandas and openpyxl are required. Install with: pip install pandas openpyxl")
        
        print(f"\n📑 Exporting {len(metadata_list)} records to Excel...")
        
        # Prepare data with the schema from ARCHITECTURE.md
        rows = []
        for meta in metadata_list:
            row = {
                # Core meeting metadata
                "meeting_id": meta.get("meeting_id", ""),
                "meeting_date": meta.get("meeting_date", ""),
                "meeting_time": meta.get("meeting_time", ""),
                "meeting_subject": meta.get("meeting_subject", ""),
                "organizer_email": meta.get("user_email", ""),
                "organizer_name": meta.get("user_name", ""),
                "duration_seconds": meta.get("duration_seconds", 0),
                "participant_count": meta.get("participant_count", 0),
                "participants": ", ".join(meta.get("participants", [])) if meta.get("participants") else "",
                
                # File information
                "file_name": meta.get("file_name", ""),
                "file_size_bytes": meta.get("size_bytes", 0),
                "sharepoint_url": meta.get("web_url", ""),
                "last_modified": meta.get("last_modified", ""),
                
                # Blob storage (to be filled after upload)
                "transcript_blob_url": "",
                "blob_uploaded": False,
                
                # AI Analysis scores (to be filled by Power Automate)
                "sentiment_score": None,
                "churn_risk_score": None,
                "upsell_potential_score": None,
                "execution_reliability_score": None,
                "operational_complexity_score": None,
                
                # AI Analysis events (to be filled by Power Automate)
                "events_detected": "",
                "key_topics": "",
                "action_items": "",
                "ai_summary": "",
                
                # Processing status
                "ai_processed": False,
                "ai_processed_date": "",
                "processing_error": meta.get("error", ""),
                
                # Additional metadata
                "word_count": meta.get("word_count", 0),
                "content_parsed": meta.get("content_parsed", False),
                "extraction_date": datetime.now().isoformat()
            }
            rows.append(row)
        
        # Create DataFrame
        df = pd.DataFrame(rows)
        
        # Ensure output directory exists
        output_path = Path(output_file)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Write to Excel with formatting
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Meeting Metadata', index=False)
            
            # Get workbook and worksheet
            workbook = writer.book
            worksheet = writer.sheets['Meeting Metadata']
            
            # Format header row
            header_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
            header_font = Font(color='FFFFFF', bold=True)
            
            for cell in worksheet[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', wrap_text=True)
            
            # Adjust column widths
            column_widths = {
                'A': 20,  # meeting_id
                'B': 12,  # meeting_date
                'C': 10,  # meeting_time
                'D': 30,  # meeting_subject
                'E': 30,  # organizer_email
                'F': 20,  # organizer_name
                'G': 12,  # duration_seconds
                'H': 12,  # participant_count
                'I': 40,  # participants
                'J': 40,  # file_name
                'K': 12,  # file_size_bytes
                'L': 50,  # sharepoint_url
                'M': 20,  # last_modified
                'N': 50,  # transcript_blob_url
                'O': 12,  # blob_uploaded
                'P': 12,  # sentiment_score
                'Q': 12,  # churn_risk_score
                'R': 12,  # upsell_potential_score
                'S': 12,  # execution_reliability_score
                'T': 12,  # operational_complexity_score
                'U': 30,  # events_detected
                'V': 30,  # key_topics
                'W': 40,  # action_items
                'X': 50,  # ai_summary
                'Y': 12,  # ai_processed
                'Z': 20,  # ai_processed_date
                'AA': 30, # processing_error
                'AB': 10, # word_count
                'AC': 12, # content_parsed
                'AD': 20  # extraction_date
            }
            
            for col, width in column_widths.items():
                worksheet.column_dimensions[col].width = width
            
            # Freeze header row
            worksheet.freeze_panes = 'A2'
        
        print(f"✅ Excel file created: {output_path}")
        return str(output_path)

    async def upload_transcripts_to_blob(self, metadata_list: List[Dict], 
                                          connection_string: str = None,
                                          container_name: str = "transcripts",
                                          transcripts_only: bool = True,
                                          max_file_size_mb: int = 100) -> List[Dict]:
        """
        Upload transcript/recording files to Azure Blob Storage organized by user folders.
        
        Args:
            metadata_list: List of metadata dictionaries with file info
            connection_string: Azure Blob Storage connection string
            container_name: Blob container name
            transcripts_only: If True, only upload VTT/transcript files (skip large MP4s)
            max_file_size_mb: Maximum file size in MB to upload (default 100MB)
            
        Returns:
            Updated metadata list with blob URLs
        """
        if not BLOB_AVAILABLE:
            raise ImportError("azure-storage-blob is required. Install with: pip install azure-storage-blob")
        
        if not connection_string:
            connection_string = os.getenv("AZURE_STORAGE_CONNECTION_STRING")
        
        if not connection_string:
            print("⚠ No Azure Storage connection string provided.")
            print("  Set AZURE_STORAGE_CONNECTION_STRING in .env or pass it as parameter.")
            return metadata_list
        
        print(f"\n☁️ Uploading files to Azure Blob Storage...")
        print(f"   Container: {container_name}")
        if transcripts_only:
            print(f"   Mode: Transcripts only (VTT files)")
        else:
            print(f"   Mode: All files (max {max_file_size_mb}MB each)")
        
        try:
            blob_service_client = BlobServiceClient.from_connection_string(connection_string)
            
            # Create container if it doesn't exist
            container_client = blob_service_client.get_container_client(container_name)
            try:
                container_client.create_container()
                print(f"   Created container: {container_name}")
            except Exception:
                pass  # Container already exists
            
            uploaded_count = 0
            skipped_count = 0
            user_folders_created = set()
            
            # Filter files to upload
            files_to_upload = []
            for meta in metadata_list:
                if not meta.get("web_url"):
                    continue
                
                file_type = meta.get("file_type", "")
                file_name = meta.get("file_name", "")
                
                # Check if we should upload this file
                if transcripts_only:
                    # Only upload VTT transcript files
                    if file_type == "transcript" or file_name.lower().endswith('.vtt'):
                        files_to_upload.append(meta)
                    else:
                        skipped_count += 1
                        # Store SharePoint URL as the link instead
                        meta["transcript_blob_url"] = meta.get("web_url", "")
                        meta["blob_uploaded"] = False
                        meta["processing_notes"] = "MP4 recording - using SharePoint link"
                else:
                    files_to_upload.append(meta)
            
            if transcripts_only and skipped_count > 0:
                print(f"   ⏭️ Skipping {skipped_count} MP4 recordings (using SharePoint links)")
            
            if not files_to_upload:
                print(f"\n📝 No transcript files (VTT) found to upload.")
                print(f"   SharePoint URLs will be used for {skipped_count} recording files.")
                return metadata_list
            
            print(f"\n📤 Uploading {len(files_to_upload)} files...")
            
            for i, meta in enumerate(files_to_upload):
                try:
                    # Download the file content from SharePoint/OneDrive
                    download_url = await self._get_file_download_url(
                        meta.get("user_email", ""),
                        meta.get("file_id", "")
                    )
                    
                    if download_url:
                        token = self.credential.get_token("https://graph.microsoft.com/.default")
                        headers = {"Authorization": f"Bearer {token.token}"}
                        
                        # Stream download for large files with longer timeout
                        resp = requests.get(download_url, headers=headers, timeout=300, stream=True)
                        if resp.status_code == 200:
                            # Check content length
                            content_length = int(resp.headers.get('content-length', 0))
                            file_size_mb = content_length / (1024 * 1024)
                            
                            if file_size_mb > max_file_size_mb:
                                print(f"     ⏭️ Skipping {meta.get('file_name', 'unknown')[:30]}... ({file_size_mb:.1f}MB > {max_file_size_mb}MB limit)")
                                meta["transcript_blob_url"] = meta.get("web_url", "")
                                meta["processing_notes"] = f"File too large ({file_size_mb:.1f}MB)"
                                continue
                            
                            # Create user folder name (sanitize email for folder name)
                            user_email = meta.get("user_email", "unknown")
                            user_folder = user_email.replace("@", "_at_").replace(".", "_")
                            
                            # Track user folders
                            if user_folder not in user_folders_created:
                                user_folders_created.add(user_folder)
                                print(f"\n  📁 User folder: {user_folder}")
                            
                            # Create blob path: user_folder/filename
                            file_name = meta.get('file_name', 'file')
                            # Sanitize filename for blob storage
                            safe_filename = re.sub(r'[<>:"/\\|?*]', '_', file_name)
                            blob_name = f"{user_folder}/{safe_filename}"
                            
                            # Upload to blob with chunked upload for reliability
                            blob_client = container_client.get_blob_client(blob_name)
                            
                            # Read content in chunks
                            content = b''
                            for chunk in resp.iter_content(chunk_size=8192):
                                content += chunk
                            
                            blob_client.upload_blob(content, overwrite=True)
                            
                            # Generate SAS URL (valid for 1 year)
                            sas_token = generate_blob_sas(
                                account_name=blob_service_client.account_name,
                                container_name=container_name,
                                blob_name=blob_name,
                                account_key=blob_service_client.credential.account_key,
                                permission=BlobSasPermissions(read=True),
                                expiry=datetime.utcnow() + timedelta(days=365)
                            )
                            
                            meta["transcript_blob_url"] = f"{blob_client.url}?{sas_token}"
                            meta["blob_uploaded"] = True
                            uploaded_count += 1
                            
                            # Progress indicator
                            size_str = f"({file_size_mb:.1f}MB)" if file_size_mb > 1 else ""
                            print(f"     ✅ [{i+1}/{len(files_to_upload)}] {safe_filename[:40]}... {size_str}")
                        else:
                            meta["processing_error"] = f"Download failed: {resp.status_code}"
                            print(f"     ❌ [{i+1}/{len(files_to_upload)}] Download failed: {resp.status_code}")
                    else:
                        meta["processing_error"] = "No download URL available"
                        
                except Exception as e:
                    meta["processing_error"] = str(e)[:100]
                    # Use SharePoint URL as fallback
                    meta["transcript_blob_url"] = meta.get("web_url", "")
                    print(f"     ❌ Error: {meta.get('file_name', 'unknown')[:30]}... - {str(e)[:40]}")
            
            print(f"\n☁️ Upload Summary:")
            print(f"   Files uploaded to blob: {uploaded_count}")
            print(f"   Files using SharePoint links: {skipped_count}")
            print(f"   User folders created: {len(user_folders_created)}")
            
            if user_folders_created:
                print(f"\n📁 User folders in blob storage:")
                for folder in sorted(user_folders_created):
                    print(f"   • {folder}")
            
        except Exception as e:
            print(f"❌ Blob storage error: {e}")
        
        return metadata_list


async def main():
    """Main function to run the transcript extractor."""
    import sys
    
    # Load environment variables from .env file
    load_dotenv()
    
    # Load configuration from environment variables
    client_id = os.getenv("AZURE_CLIENT_ID")
    client_secret = os.getenv("AZURE_CLIENT_SECRET")
    tenant_id = os.getenv("AZURE_TENANT_ID")
    user_id = os.getenv("TEAMS_USER_ID")
    
    if not all([client_id, client_secret, tenant_id]):
        print("Error: Missing required environment variables")
        print("Please set: AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, AZURE_TENANT_ID")
        return
    
    # Initialize extractor
    output_dir = os.getenv("OUTPUT_DIR", "transcripts")
    extractor = TranscriptExtractor(client_id, client_secret, tenant_id, output_dir)
    
    # Check for command line arguments
    if len(sys.argv) > 1:
        if sys.argv[1] == "--list-users":
            print("Fetching users from your organization...")
            print("=" * 50)
            users = await extractor.list_users(20)
            if users:
                print(f"\nFound {len(users)} users:")
                for user in users:
                    email = user.mail or user.user_principal_name
                    print(f"  • {user.display_name} - {email}")
                print("\nUpdate your .env file with the correct TEAMS_USER_ID")
            else:
                print("No users found or insufficient permissions")
            return
        
        elif sys.argv[1] == "--scan-all":
            print("Scanning all users for meetings and transcripts...")
            print("This may take a few minutes...\n")
            results = await extractor.scan_all_users_for_meetings(50)
            
            print("\n" + "=" * 60)
            print("SCAN SUMMARY:")
            print("=" * 60)
            print(f"Users scanned: {results['users_scanned']}")
            print(f"Users with meetings: {len(results['users_with_meetings'])}")
            print(f"Total meetings found: {results['total_meetings']}")
            print(f"Total transcripts found: {results['total_transcripts']}")
            
            if results['users_with_meetings']:
                print("\nUsers with meetings:")
                for user_info in results['users_with_meetings']:
                    if user_info['transcripts'] > 0:
                        print(f"  ✨ {user_info['name']} ({user_info['email']})")
                        print(f"     {user_info['meetings']} meetings, {user_info['transcripts']} transcripts")
                        print(f"     Run: TEAMS_USER_ID={user_info['email']} in .env")
            else:
                print("\n⚠ No users found with meetings")
                print("  This could mean:")
                print("  • No Teams meetings have been held yet")
                print("  • Transcription wasn't enabled for meetings")
                print("  • Meetings are not accessible via this API")
            return
        
        elif sys.argv[1] == "--use-communications":
            # Optionally accept an organizer email as the second argument
            organizer_email = None
            if len(sys.argv) > 2:
                organizer_email = sys.argv[2]
            else:
                organizer_email = user_id

            print(f"Calling communications endpoint and filtering for organizer: {organizer_email}")
            meetings = await extractor.get_communications_meetings(200)
            if not meetings:
                print("No meetings returned by the communications endpoint.")
                return

            # Filter meetings where organizer email matches (best-effort)
            matched = []
            for m in meetings:
                try:
                    org = m.get("organizer") or {}
                    # organizer structure can vary; try common patterns
                    org_email = None
                    if isinstance(org, dict):
                        org_email = org.get("email") or org.get("identity", {}).get("user", {}).get("email") or org.get("user", {}).get("email")

                    # fallback: check joinInformation or participants
                    if not org_email:
                        participants = m.get("participants") or {}
                        # participants may contain "organizer" key
                        if isinstance(participants, dict) and participants.get("organizer"):
                            p_org = participants.get("organizer")
                            if isinstance(p_org, dict):
                                org_email = p_org.get("identity", {}).get("user", {}).get("email") or p_org.get("email")

                    if org_email and organizer_email and organizer_email.lower() in org_email.lower():
                        matched.append(m)
                except Exception:
                    continue

            print(f"Total communications meetings returned: {len(meetings)}, matched for organizer: {len(matched)}")
            if matched:
                for mm in matched[:50]:
                    subj = mm.get("subject") or "(no subject)"
                    mid = mm.get("id") or "(no id)"
                    print(f"- id: {mid} | subject: {subj}")
                print("\nYou can set TEAMS_USER_ID in your .env to one of these organizers and re-run the extractor.")
            else:
                print("No meetings matched the organizer filter. The communications endpoint may require different filtering or the meetings are stored elsewhere (cloud recordings, callRecords, or retention policies).")
            return
        
        elif sys.argv[1] == "--use-communications-post":
            # Expect two additional args: start and end datetimes (ISO 8601)
            if len(sys.argv) < 4:
                print("Usage: python src/transcript_extractor.py --use-communications-post <startIso> <endIso>")
                print("Example: python src/transcript_extractor.py --use-communications-post 2025-11-01T00:00:00Z 2025-12-06T23:59:59Z")
                return

            start_dt = sys.argv[2]
            end_dt = sys.argv[3]
            print(f"Calling getAllOnlineMeetings from {start_dt} to {end_dt}...")
            items = await extractor.post_get_all_online_meetings(start_dt, end_dt)
            if not items:
                print("No meetings returned by the getAllOnlineMeetings action.")
                return

            print(f"Total meetings returned: {len(items)}")
            # Show a few items
            for it in items[:50]:
                mid = it.get("id") or it.get("meetingId") or "(no id)"
                subj = it.get("subject") or it.get("subjectText") or "(no subject)"
                org = ""
                try:
                    org_obj = it.get("organizer") or {}
                    org = org_obj.get("email") or org_obj.get("identity", {}).get("user", {}).get("email") or org_obj.get("user", {}).get("email") or "(unknown)"
                except:
                    org = "(unknown)"
                print(f"- id: {mid} | org: {org} | subject: {subj}")
            return

        elif sys.argv[1] == "--use-communications-beta":
            print("Calling beta communications endpoint (may expose additional actions)...")
            meetings = await extractor.get_communications_meetings_beta(200)
            if not meetings:
                print("No meetings returned by beta communications endpoint.")
                return
            print(f"Total beta communications meetings returned: {len(meetings)}")
            for m in meetings[:50]:
                mid = m.get("id") or "(no id)"
                subj = m.get("subject") or "(no subject)"
                print(f"- id: {mid} | subject: {subj}")
            print("\nIf you want raw output for diagnosis, run with: --dump-communications-raw <N>")
            return

        elif sys.argv[1] == "--dump-communications-raw":
            # Dump raw communications meetings for inspection
            n = 10
            if len(sys.argv) > 2:
                try:
                    n = int(sys.argv[2])
                except:
                    pass
            print(f"Fetching raw communications meetings (up to {n})...")
            meetings = await extractor.get_communications_meetings(200)
            if not meetings:
                print("No meetings returned by communications endpoint. Trying beta...")
                meetings = await extractor.get_communications_meetings_beta(200)
            if not meetings:
                print("No meetings available to dump.")
                return
            import pprint, json
            pp = pprint.PrettyPrinter(depth=3)
            for i, m in enumerate(meetings[:n]):
                print(f"\n--- ITEM {i+1} ---")
                print(json.dumps(m, indent=2, default=str)[:4000])
            return

        elif sys.argv[1] == "--search-drives":
            # Search OneDrive for transcript files
            # Usage: --search-drives [email1,email2,...] [--download]
            download = "--download" in sys.argv
            
            # Get list of users to search
            user_emails = []
            if len(sys.argv) > 2 and not sys.argv[2].startswith("--"):
                # Comma-separated list of emails
                user_emails = [e.strip() for e in sys.argv[2].split(",") if e.strip()]
            
            if not user_emails:
                # Default: use TEAMS_USER_ID or scan top users
                if user_id:
                    user_emails = [user_id]
                else:
                    print("Fetching top 10 users to search their drives...")
                    users = await extractor.list_users(10)
                    user_emails = [u.mail or u.user_principal_name for u in users if u.mail or u.user_principal_name]
            
            if not user_emails:
                print("No users to search. Set TEAMS_USER_ID in .env or provide emails.")
                return
            
            print(f"🔎 Searching OneDrive for transcript files...")
            print(f"   Users to search: {', '.join(user_emails)}")
            print(f"   Download files: {download}")
            print("=" * 60)
            
            results = await extractor.search_drives_for_transcripts(user_emails, download=download)
            
            print("\n" + "=" * 60)
            print("DRIVE SEARCH SUMMARY:")
            print("=" * 60)
            print(f"  Users searched: {results['users_searched']}")
            print(f"  Transcript files found: {len(results['files_found'])}")
            print(f"  Files downloaded: {results['files_downloaded']}")
            
            if results['files_found']:
                print("\n📁 Found files:")
                for f in results['files_found']:
                    print(f"  • {f['name']} ({f['user']})")
                    print(f"    {f['webUrl']}")
                
                if not download:
                    print("\n💡 To download these files, run:")
                    print(f"   python src/transcript_extractor.py --search-drives \"{','.join(user_emails)}\" --download")
            else:
                print("\n⚠ No transcript files found in OneDrive.")
                print("  This could mean:")
                print("  • Recordings are saved to SharePoint instead of OneDrive")
                print("  • Transcription wasn't enabled for meetings")
                print("  • Files are in a different location or have different names")
                print("\n  Try searching SharePoint sites directly or check Teams admin settings.")
            
            if results['errors']:
                print("\n⚠ Errors encountered:")
                for err in results['errors'][:5]:
                    print(f"  • {err}")
            return

        elif sys.argv[1] == "--scan-recordings":
            # Scan all users' Recordings folders
            download = "--download" in sys.argv
            max_users = 30
            
            # Check for max users arg
            for i, arg in enumerate(sys.argv):
                if arg == "--max" and i + 1 < len(sys.argv):
                    try:
                        max_users = int(sys.argv[i + 1])
                    except:
                        pass
            
            print(f"🎬 Scanning users' Recordings folders for meeting files...")
            print(f"   Max users to scan: {max_users}")
            print(f"   Download transcripts: {download}")
            print("=" * 60)
            
            results = await extractor.search_all_users_recordings(max_users=max_users, download=download)
            
            print("\n" + "=" * 60)
            print("RECORDINGS SCAN SUMMARY:")
            print("=" * 60)
            print(f"  Users scanned: {results['users_searched']}")
            print(f"  Users with recordings: {len(results['users_with_recordings'])}")
            print(f"  Total files found: {len(results['files_found'])}")
            print(f"  📝 Transcript files: {len(results['transcript_files'])}")
            print(f"  🎥 Recording files: {len(results['recording_files'])}")
            print(f"  Files downloaded: {results['files_downloaded']}")
            
            if results['users_with_recordings']:
                print("\n👥 Users with recordings:")
                for u in results['users_with_recordings']:
                    print(f"  • {u['name']} ({u['email']}) - {u['file_count']} files")
            
            if results['transcript_files']:
                print("\n📝 Transcript files found:")
                for f in results['transcript_files'][:20]:
                    print(f"  • {f['name']}")
                    print(f"    User: {f['user_name']} | Modified: {f['lastModified'][:10] if f['lastModified'] else 'N/A'}")
                    print(f"    URL: {f['webUrl']}")
                
                if not download:
                    print("\n💡 To download transcript files, run:")
                    print("   python src/transcript_extractor.py --scan-recordings --download")
            else:
                print("\n⚠ No transcript files found.")
                print("  Recordings may exist but transcription wasn't enabled.")
            
            if results['errors']:
                print("\n⚠ Errors:")
                for err in results['errors'][:5]:
                    print(f"  • {err}")
            return

        elif sys.argv[1] == "--export-metadata":
            # Export metadata to Excel
            max_users = 30
            parse_content = "--parse" in sys.argv or "--parse-content" in sys.argv
            include_recordings = "--include-recordings" in sys.argv or "--recordings" in sys.argv
            output_file = "meeting_metadata.xlsx"
            
            # Parse arguments
            for i, arg in enumerate(sys.argv):
                if arg == "--max" and i + 1 < len(sys.argv):
                    try:
                        max_users = int(sys.argv[i + 1])
                    except:
                        pass
                if arg == "--output" and i + 1 < len(sys.argv):
                    output_file = sys.argv[i + 1]
            
            print(f"📊 Extracting meeting metadata for Excel export...")
            print(f"   Max users: {max_users}")
            print(f"   Parse content: {parse_content}")
            print(f"   Include recordings: {include_recordings}")
            print(f"   Output file: {output_file}")
            print("=" * 60)
            
            try:
                metadata_list = await extractor.extract_metadata_for_export(
                    max_users=max_users,
                    download_content=parse_content,
                    include_recordings=include_recordings
                )
                
                if metadata_list:
                    excel_path = extractor.export_to_excel(metadata_list, output_file)
                    
                    print("\n" + "=" * 60)
                    print("EXPORT SUMMARY:")
                    print("=" * 60)
                    print(f"  Transcripts processed: {len(metadata_list)}")
                    print(f"  Excel file: {excel_path}")
                    print("\n💡 Next steps:")
                    print("  1. Upload transcripts to Azure Blob Storage:")
                    print("     python src/transcript_extractor.py --upload-blobs")
                    print("  2. Configure Power Automate to read the Excel file")
                    print("  3. Set up AI analysis flow for each transcript")
                else:
                    print("\n⚠ No transcript files found to export.")
            except ImportError as e:
                print(f"\n❌ Missing dependencies: {e}")
                print("   Install with: pip install pandas openpyxl")
            return

        elif sys.argv[1] == "--upload-blobs":
            # Upload transcripts to Azure Blob Storage
            max_users = 30
            container_name = "transcripts"
            output_file = "meeting_metadata.xlsx"
            transcripts_only = True  # Default: only upload VTT transcripts, not MP4s
            max_file_size_mb = 100
            
            # Parse arguments
            for i, arg in enumerate(sys.argv):
                if arg == "--max" and i + 1 < len(sys.argv):
                    try:
                        max_users = int(sys.argv[i + 1])
                    except:
                        pass
                if arg == "--container" and i + 1 < len(sys.argv):
                    container_name = sys.argv[i + 1]
                if arg == "--output" and i + 1 < len(sys.argv):
                    output_file = sys.argv[i + 1]
                if arg == "--include-recordings":
                    transcripts_only = False
                if arg == "--max-size" and i + 1 < len(sys.argv):
                    try:
                        max_file_size_mb = int(sys.argv[i + 1])
                    except:
                        pass
            
            print(f"☁️ Uploading transcripts to Azure Blob Storage...")
            print(f"   Max users: {max_users}")
            print(f"   Container: {container_name}")
            print(f"   Output file: {output_file}")
            if transcripts_only:
                print(f"   Mode: Transcripts only (VTT files)")
                print(f"   Note: MP4 recordings will use SharePoint links")
            else:
                print(f"   Mode: All files (max {max_file_size_mb}MB each)")
            print("=" * 60)
            
            try:
                # First extract metadata
                metadata_list = await extractor.extract_metadata_for_export(
                    max_users=max_users,
                    download_content=True
                )
                
                if metadata_list:
                    # Upload to blob storage
                    updated_metadata = await extractor.upload_transcripts_to_blob(
                        metadata_list,
                        container_name=container_name,
                        transcripts_only=transcripts_only,
                        max_file_size_mb=max_file_size_mb
                    )
                    
                    # Export to Excel with blob URLs
                    excel_path = extractor.export_to_excel(updated_metadata, output_file)
                    
                    uploaded_count = sum(1 for m in updated_metadata if m.get("blob_uploaded"))
                    sharepoint_count = sum(1 for m in updated_metadata if not m.get("blob_uploaded") and m.get("transcript_blob_url"))
                    
                    print("\n" + "=" * 60)
                    print("UPLOAD SUMMARY:")
                    print("=" * 60)
                    print(f"  Total files processed: {len(updated_metadata)}")
                    print(f"  Uploaded to blob: {uploaded_count}")
                    print(f"  Using SharePoint links: {sharepoint_count}")
                    print(f"  Excel file: {excel_path}")
                    print("\n💡 The Excel file now contains URLs for Power Automate.")
                    if transcripts_only and uploaded_count == 0:
                        print("\n⚠️ Note: No VTT transcript files found.")
                        print("   Enable meeting transcription in Teams Admin Center to generate transcripts.")
                        print("   MP4 recordings can be accessed via SharePoint links in the Excel file.")
                else:
                    print("\n⚠ No transcript files found to upload.")
            except ImportError as e:
                print(f"\n❌ Missing dependencies: {e}")
                print("   Install with: pip install pandas openpyxl azure-storage-blob")
            except Exception as e:
                print(f"\n❌ Error: {e}")
            return
            return
    
    if not user_id:
        print("Error: TEAMS_USER_ID not set in .env file")
        print("Run 'python src/transcript_extractor.py --list-users' to see available users")
        return
    
    # Extract transcripts
    print("Starting transcript extraction...")
    print("=" * 50)
    stats = await extractor.extract_all_transcripts(user_id, verify_user=True)
    
    # Print summary
    print("\n" + "=" * 50)
    print("EXTRACTION SUMMARY:")
    print(f"  Meetings found: {stats['meetings_found']}")
    print(f"  Transcripts found: {stats['transcripts_found']}")
    print(f"  Transcripts saved: {stats['transcripts_saved']}")
    print(f"  Errors: {stats['errors']}")
    print(f"\nTranscripts saved to: {output_dir}")


if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
