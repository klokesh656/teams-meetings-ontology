"""
Move Recordings to SharePoint Folder
=====================================
Moves recordings from KC and Louise to HR SharePoint in a folder called "team 1 AI test"
"""

import os
import sys
import requests
import time
from datetime import datetime
from azure.identity import ClientSecretCredential
from dotenv import load_dotenv
import logging

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/move_recordings.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# Load environment variables
load_dotenv()

# Configuration
TENANT_ID = os.getenv('AZURE_TENANT_ID')
CLIENT_ID = os.getenv('AZURE_CLIENT_ID')
CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET')

# Target users to copy from
SOURCE_USERS = {
    'KC argente': 'KC.argente@our-assistants.com',
    'Louise Alatraca': 'Louise.Alatraca@our-assistants.com'
}

# Destination
HR_EMAIL = 'HR@our-assistants.com'
DESTINATION_FOLDER = 'team 1 AI test'


class RecordingMover:
    def __init__(self):
        self.credential = ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
        self.token = None
        self.headers = None
        self.hr_user_id = None
        self.destination_folder_id = None
        self.stats = {
            'files_found': 0,
            'files_copied': 0,
            'errors': 0
        }
    
    def authenticate(self):
        """Authenticate with Microsoft Graph API"""
        logger.info("Authenticating with Microsoft Graph API...")
        self.token = self.credential.get_token('https://graph.microsoft.com/.default').token
        self.headers = {
            'Authorization': f'Bearer {self.token}',
            'Content-Type': 'application/json'
        }
        logger.info("Authentication successful!")
    
    def get_user_id(self, email):
        """Get user ID from email"""
        url = f"https://graph.microsoft.com/v1.0/users/{email}"
        resp = requests.get(url, headers=self.headers, timeout=30)
        if resp.status_code == 200:
            return resp.json().get('id')
        logger.error(f"Failed to get user ID for {email}: {resp.status_code}")
        return None
    
    def create_destination_folder(self):
        """Create the destination folder in HR's OneDrive"""
        logger.info(f"Creating destination folder: {DESTINATION_FOLDER}")
        
        # Get HR user ID
        self.hr_user_id = self.get_user_id(HR_EMAIL)
        if not self.hr_user_id:
            logger.error("Failed to get HR user ID!")
            return False
        
        # Check if folder exists in Recordings
        url = f"https://graph.microsoft.com/v1.0/users/{self.hr_user_id}/drive/root:/Recordings/{DESTINATION_FOLDER}"
        resp = requests.get(url, headers=self.headers, timeout=30)
        
        if resp.status_code == 200:
            self.destination_folder_id = resp.json().get('id')
            logger.info(f"Folder already exists: {DESTINATION_FOLDER}")
            return True
        
        # Create the folder
        url = f"https://graph.microsoft.com/v1.0/users/{self.hr_user_id}/drive/root:/Recordings:/children"
        body = {
            "name": DESTINATION_FOLDER,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "rename"
        }
        
        resp = requests.post(url, headers=self.headers, json=body, timeout=30)
        if resp.status_code in [200, 201]:
            self.destination_folder_id = resp.json().get('id')
            logger.info(f"Created folder: {DESTINATION_FOLDER}")
            return True
        else:
            logger.error(f"Failed to create folder: {resp.status_code} - {resp.text}")
            return False
    
    def get_user_recordings(self, user_email, user_name):
        """Get all recordings from a user's OneDrive"""
        logger.info(f"Getting recordings for {user_name}...")
        
        user_id = self.get_user_id(user_email)
        if not user_id:
            return []
        
        recordings = []
        url = f"https://graph.microsoft.com/v1.0/users/{user_id}/drive/root:/Recordings:/children?$top=500"
        
        while url:
            resp = requests.get(url, headers=self.headers, timeout=30)
            if resp.status_code == 200:
                data = resp.json()
                items = data.get('value', [])
                mp4_files = [i for i in items if i.get('name', '').lower().endswith('.mp4')]
                for mp4 in mp4_files:
                    mp4['source_user_id'] = user_id
                    mp4['source_user_name'] = user_name
                recordings.extend(mp4_files)
                url = data.get('@odata.nextLink')
            else:
                logger.error(f"Failed to get recordings: {resp.status_code}")
                break
        
        logger.info(f"  Found {len(recordings)} recordings for {user_name}")
        return recordings
    
    def copy_file_to_destination(self, file_info):
        """Copy a file to the destination folder"""
        file_name = file_info.get('name')
        file_id = file_info.get('id')
        source_user_id = file_info.get('source_user_id')
        source_user_name = file_info.get('source_user_name')
        
        # Add source user prefix to filename
        new_name = f"{source_user_name} - {file_name}"
        
        # Use copy endpoint
        url = f"https://graph.microsoft.com/v1.0/users/{source_user_id}/drive/items/{file_id}/copy"
        
        body = {
            "parentReference": {
                "driveId": None,  # Will be set below
                "id": self.destination_folder_id
            },
            "name": new_name
        }
        
        # Get destination drive ID
        drive_url = f"https://graph.microsoft.com/v1.0/users/{self.hr_user_id}/drive"
        drive_resp = requests.get(drive_url, headers=self.headers, timeout=30)
        if drive_resp.status_code == 200:
            body["parentReference"]["driveId"] = drive_resp.json().get('id')
        else:
            logger.error(f"Failed to get destination drive: {drive_resp.status_code}")
            return False
        
        # Execute copy
        resp = requests.post(url, headers=self.headers, json=body, timeout=60)
        
        if resp.status_code == 202:  # Accepted - copy in progress
            logger.info(f"  [COPYING] {file_name} -> {DESTINATION_FOLDER}")
            return True
        elif resp.status_code == 201:  # Created
            logger.info(f"  [COPIED] {file_name} -> {DESTINATION_FOLDER}")
            return True
        else:
            logger.error(f"  [ERROR] Failed to copy {file_name}: {resp.status_code} - {resp.text[:200]}")
            return False
    
    def run(self):
        """Run the move operation"""
        start_time = datetime.now()
        logger.info(f"Starting recording move - {start_time}")
        logger.info(f"Moving recordings from KC and Louise to HR/{DESTINATION_FOLDER}")
        
        try:
            # Authenticate
            self.authenticate()
            
            # Create destination folder
            if not self.create_destination_folder():
                logger.error("Failed to create destination folder!")
                return
            
            # Get all recordings from source users
            all_recordings = []
            for user_name, user_email in SOURCE_USERS.items():
                recordings = self.get_user_recordings(user_email, user_name)
                all_recordings.extend(recordings)
            
            self.stats['files_found'] = len(all_recordings)
            logger.info(f"\nTotal recordings to copy: {len(all_recordings)}")
            
            # Copy each file
            logger.info("\n" + "="*60)
            logger.info("COPYING FILES")
            logger.info("="*60)
            
            for i, file_info in enumerate(all_recordings, 1):
                logger.info(f"\n[{i}/{len(all_recordings)}] {file_info.get('name')}")
                
                if self.copy_file_to_destination(file_info):
                    self.stats['files_copied'] += 1
                else:
                    self.stats['errors'] += 1
                
                time.sleep(0.5)  # Rate limiting
            
            # Summary
            logger.info("\n" + "="*60)
            logger.info("SUMMARY")
            logger.info("="*60)
            logger.info(f"Files Found:    {self.stats['files_found']}")
            logger.info(f"Files Copied:   {self.stats['files_copied']}")
            logger.info(f"Errors:         {self.stats['errors']}")
            
            end_time = datetime.now()
            duration = (end_time - start_time).total_seconds()
            logger.info(f"\nCompleted in {duration:.1f} seconds")
            
        except Exception as e:
            logger.error(f"Process failed: {e}")
            raise


def main():
    mover = RecordingMover()
    mover.run()


if __name__ == '__main__':
    main()
