"""
Find All Recordings Without Transcripts and Transcribe Them
============================================================
This script:
1. Scans all users' OneDrive for MP4 recordings
2. Downloads recordings that don't have matching transcripts
3. Transcribes using Azure Speech-to-Text
4. Saves transcripts in VTT format

Requirements:
- FFmpeg installed (for audio extraction)
- Azure Speech Service configured (AZURE_SPEECH_KEY in .env)
"""

import os
import sys
import requests
import time
from datetime import datetime
from pathlib import Path
from azure.identity import ClientSecretCredential
from dotenv import load_dotenv
import subprocess
import logging
import re

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/find_and_transcribe.log', encoding='utf-8'),
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

# Azure Speech Service
SPEECH_KEY = os.getenv('AZURE_SPEECH_KEY')
SPEECH_REGION = os.getenv('AZURE_SPEECH_REGION', 'eastus')

# Paths
RECORDINGS_DIR = Path('recordings')
TRANSCRIPTS_DIR = Path('transcripts')
LOGS_DIR = Path('logs')

# Ensure directories exist
RECORDINGS_DIR.mkdir(exist_ok=True)
TRANSCRIPTS_DIR.mkdir(exist_ok=True)
LOGS_DIR.mkdir(exist_ok=True)


class OneDriveRecordingTranscriber:
    def __init__(self):
        self.credential = ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
        self.token = None
        self.token_expires_at = None
        self.headers = None
        self.all_recordings = []
        self.recordings_to_transcribe = []
        self.stats = {
            'users_scanned': 0,
            'total_recordings': 0,
            'recordings_without_transcript': 0,
            'downloaded': 0,
            'transcribed': 0,
            'errors': 0
        }
    
    def authenticate(self):
        """Authenticate with Microsoft Graph API"""
        logger.info("Authenticating with Microsoft Graph API...")
        token_response = self.credential.get_token('https://graph.microsoft.com/.default')
        self.token = token_response.token
        # Token typically expires in 3600 seconds (1 hour), refresh at 45 min
        self.token_expires_at = time.time() + 2700  # 45 minutes
        self.headers = {
            'Authorization': f'Bearer {self.token}',
            'Content-Type': 'application/json'
        }
        logger.info("Authentication successful!")
    
    def refresh_token_if_needed(self):
        """Refresh token if it's about to expire"""
        if self.token_expires_at and time.time() > self.token_expires_at:
            logger.info("Token expiring soon, refreshing...")
            self.authenticate()
    
    def get_all_users(self):
        """Get all users in the organization"""
        logger.info("Fetching all users...")
        url = 'https://graph.microsoft.com/v1.0/users?$top=100'
        users = []
        
        while url:
            resp = requests.get(url, headers=self.headers, timeout=30)
            if resp.status_code == 200:
                data = resp.json()
                users.extend(data.get('value', []))
                url = data.get('@odata.nextLink')
            else:
                logger.error(f"Failed to get users: {resp.status_code}")
                break
        
        logger.info(f"Found {len(users)} users")
        return users
    
    def get_existing_transcripts(self):
        """Get list of existing transcript filenames"""
        existing = set()
        for f in TRANSCRIPTS_DIR.glob('*.vtt'):
            # Extract date pattern from filename
            match = re.search(r'(\d{8}_\d{6})', f.name)
            if match:
                existing.add(match.group(1))
        return existing
    
    def scan_user_recordings(self, user_id, user_name, user_email):
        """Scan a user's OneDrive for MP4 recordings"""
        recordings = []
        
        try:
            url = f'https://graph.microsoft.com/v1.0/users/{user_id}/drive/root:/Recordings:/children?$top=500'
            resp = requests.get(url, headers=self.headers, timeout=30)
            
            if resp.status_code == 200:
                items = resp.json().get('value', [])
                mp4_files = [i for i in items if i.get('name', '').lower().endswith('.mp4')]
                
                for mp4 in mp4_files:
                    recordings.append({
                        'user_id': user_id,
                        'user_name': user_name,
                        'user_email': user_email,
                        'file_id': mp4.get('id'),
                        'file_name': mp4.get('name'),
                        'size': mp4.get('size', 0),
                        'created': mp4.get('createdDateTime', ''),
                        'download_url': mp4.get('@microsoft.graph.downloadUrl', '')
                    })
                
                if mp4_files:
                    logger.info(f"  {user_name}: {len(mp4_files)} recordings found")
            elif resp.status_code == 404:
                # No Recordings folder
                pass
            else:
                logger.warning(f"  {user_name}: Error {resp.status_code}")
        except Exception as e:
            logger.error(f"  Error scanning {user_name}: {e}")
        
        return recordings
    
    def scan_all_users(self):
        """Scan all users for recordings"""
        logger.info("="*60)
        logger.info("SCANNING ALL USERS FOR RECORDINGS")
        logger.info("="*60)
        
        users = self.get_all_users()
        existing_transcripts = self.get_existing_transcripts()
        
        for user in users:
            user_id = user['id']
            user_name = user.get('displayName', 'Unknown')
            user_email = user.get('mail', '') or user.get('userPrincipalName', '')
            
            recordings = self.scan_user_recordings(user_id, user_name, user_email)
            self.all_recordings.extend(recordings)
            self.stats['users_scanned'] += 1
        
        self.stats['total_recordings'] = len(self.all_recordings)
        logger.info(f"\nTotal recordings found: {self.stats['total_recordings']}")
        
        # Check which recordings don't have transcripts
        for rec in self.all_recordings:
            created = rec.get('created', '')
            if created:
                # Create date pattern to match against existing transcripts
                date_str = created[:10].replace('-', '')
                time_str = created[11:19].replace(':', '') if len(created) > 11 else ''
                pattern = f"{date_str}_{time_str}"
                
                if pattern not in existing_transcripts:
                    self.recordings_to_transcribe.append(rec)
        
        self.stats['recordings_without_transcript'] = len(self.recordings_to_transcribe)
        logger.info(f"Recordings WITHOUT transcripts: {self.stats['recordings_without_transcript']}")
        
        return self.recordings_to_transcribe
    
    def download_recording(self, recording):
        """Download a recording from OneDrive"""
        # Refresh token if needed before each download
        self.refresh_token_if_needed()
        
        user_id = recording['user_id']
        file_id = recording['file_id']
        file_name = recording['file_name']
        created = recording['created']
        
        # Get download URL
        url = f'https://graph.microsoft.com/v1.0/users/{user_id}/drive/items/{file_id}/content'
        
        try:
            resp = requests.get(url, headers=self.headers, timeout=120, stream=True, allow_redirects=True)
            
            if resp.status_code == 200:
                # Create safe filename
                date_str = created[:10].replace('-', '') if created else 'unknown'
                time_str = created[11:19].replace(':', '') if created and len(created) > 11 else ''
                safe_name = re.sub(r'[<>:"/\\|?*]', '', file_name)[:50]
                filename = f"{date_str}_{time_str}_{safe_name}"
                if not filename.lower().endswith('.mp4'):
                    filename += '.mp4'
                
                filepath = RECORDINGS_DIR / filename
                
                # Download file
                total_size = int(resp.headers.get('content-length', 0))
                with open(filepath, 'wb') as f:
                    downloaded = 0
                    for chunk in resp.iter_content(chunk_size=8192):
                        f.write(chunk)
                        downloaded += len(chunk)
                
                size_mb = filepath.stat().st_size / 1024 / 1024
                logger.info(f"    Downloaded: {filename} ({size_mb:.1f} MB)")
                return filepath
            else:
                logger.error(f"    Failed to download: {resp.status_code}")
                return None
        except Exception as e:
            logger.error(f"    Error downloading: {e}")
            return None
    
    def extract_audio(self, video_path):
        """Extract audio from MP4 using ffmpeg"""
        audio_path = video_path.with_suffix('.wav')
        
        try:
            # Check if ffmpeg is available
            result = subprocess.run(['ffmpeg', '-version'], capture_output=True)
            if result.returncode != 0:
                logger.error("    FFmpeg not found!")
                return None
            
            # Extract audio (16kHz mono WAV for best speech recognition)
            cmd = [
                'ffmpeg', '-i', str(video_path),
                '-ar', '16000',  # 16kHz sample rate
                '-ac', '1',      # Mono
                '-y',            # Overwrite
                str(audio_path)
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True)
            if result.returncode == 0:
                logger.info(f"    Extracted audio: {audio_path.name}")
                return audio_path
            else:
                logger.error(f"    FFmpeg error: {result.stderr[:100]}")
                return None
        except FileNotFoundError:
            logger.error("    FFmpeg not installed! Run: choco install ffmpeg")
            return None
        except Exception as e:
            logger.error(f"    Error extracting audio: {e}")
            return None
    
    def transcribe_audio(self, audio_path):
        """Transcribe audio using Azure Speech-to-Text"""
        if not SPEECH_KEY:
            logger.error("    Azure Speech API key not configured!")
            logger.error("    Add AZURE_SPEECH_KEY to your .env file")
            return None
        
        try:
            import azure.cognitiveservices.speech as speechsdk
            
            speech_config = speechsdk.SpeechConfig(subscription=SPEECH_KEY, region=SPEECH_REGION)
            speech_config.speech_recognition_language = "en-US"
            
            audio_config = speechsdk.audio.AudioConfig(filename=str(audio_path))
            speech_recognizer = speechsdk.SpeechRecognizer(speech_config=speech_config, audio_config=audio_config)
            
            logger.info(f"    Transcribing... (this may take several minutes)")
            
            done = False
            transcript_parts = []
            
            def stop_cb(evt):
                nonlocal done
                done = True
            
            def recognized_cb(evt):
                if evt.result.reason == speechsdk.ResultReason.RecognizedSpeech:
                    transcript_parts.append(evt.result.text)
            
            speech_recognizer.recognized.connect(recognized_cb)
            speech_recognizer.session_stopped.connect(stop_cb)
            speech_recognizer.canceled.connect(stop_cb)
            
            speech_recognizer.start_continuous_recognition()
            
            # Wait for transcription (max 20 minutes)
            timeout = 1200
            start_time = time.time()
            while not done and (time.time() - start_time) < timeout:
                time.sleep(1)
            
            speech_recognizer.stop_continuous_recognition()
            
            full_transcript = ' '.join(transcript_parts)
            logger.info(f"    Transcription complete! ({len(full_transcript)} characters)")
            return full_transcript
        
        except ImportError:
            logger.error("    Azure Speech SDK not installed!")
            logger.error("    Run: pip install azure-cognitiveservices-speech")
            return None
        except Exception as e:
            logger.error(f"    Transcription error: {e}")
            return None
    
    def save_transcript(self, transcript_text, recording, audio_path):
        """Save transcript as VTT file"""
        created = recording.get('created', '')
        user_name = recording.get('user_name', 'Unknown')
        file_name = recording.get('file_name', 'recording')
        
        # Create filename
        date_str = created[:10].replace('-', '') if created else 'unknown'
        time_str = created[11:19].replace(':', '') if created and len(created) > 11 else ''
        safe_name = re.sub(r'[<>:"/\\|?*]', '', file_name.replace('.mp4', ''))[:40]
        vtt_filename = f"{date_str}_{time_str}_{safe_name}.vtt"
        vtt_path = TRANSCRIPTS_DIR / vtt_filename
        
        # Create VTT content
        vtt_content = f"""WEBVTT

NOTE Transcribed from OneDrive recording using Azure Speech-to-Text
NOTE Source: {user_name} - {file_name}
NOTE Date: {created}

00:00:00.000 --> 99:59:59.999
{transcript_text}
"""
        
        with open(vtt_path, 'w', encoding='utf-8') as f:
            f.write(vtt_content)
        
        logger.info(f"    Saved: {vtt_filename}")
        return vtt_path
    
    def process_recordings(self, max_recordings=5):
        """Process recordings - download, extract audio, transcribe"""
        logger.info("="*60)
        logger.info(f"PROCESSING RECORDINGS (max: {max_recordings})")
        logger.info("="*60)
        
        if not self.recordings_to_transcribe:
            logger.info("No recordings to process!")
            return
        
        # Sort by date (newest first)
        sorted_recordings = sorted(
            self.recordings_to_transcribe,
            key=lambda x: x.get('created', ''),
            reverse=True
        )
        
        # Process limited number
        to_process = sorted_recordings[:max_recordings]
        
        for i, rec in enumerate(to_process, 1):
            logger.info(f"\n[{i}/{len(to_process)}] {rec['user_name']} - {rec['file_name']}")
            logger.info(f"    Date: {rec['created']}")
            logger.info(f"    Size: {rec['size'] / 1024 / 1024:.1f} MB")
            
            # Step 1: Download recording
            video_path = self.download_recording(rec)
            if not video_path:
                self.stats['errors'] += 1
                continue
            self.stats['downloaded'] += 1
            
            # Step 2: Extract audio
            audio_path = self.extract_audio(video_path)
            if not audio_path:
                self.stats['errors'] += 1
                continue
            
            # Step 3: Transcribe
            transcript = self.transcribe_audio(audio_path)
            if not transcript:
                self.stats['errors'] += 1
                # Cleanup
                if audio_path.exists():
                    audio_path.unlink()
                continue
            
            # Step 4: Save transcript
            vtt_path = self.save_transcript(transcript, rec, audio_path)
            self.stats['transcribed'] += 1
            
            # Cleanup audio file (keep video for reference)
            if audio_path.exists():
                audio_path.unlink()
            
            time.sleep(2)  # Rate limiting
    
    def generate_report(self):
        """Generate summary report"""
        logger.info("\n" + "="*60)
        logger.info("TRANSCRIPTION SUMMARY")
        logger.info("="*60)
        
        # Group recordings by user
        by_user = {}
        for rec in self.all_recordings:
            user = rec['user_name']
            if user not in by_user:
                by_user[user] = {'total': 0, 'without_transcript': 0}
            by_user[user]['total'] += 1
        
        for rec in self.recordings_to_transcribe:
            user = rec['user_name']
            if user in by_user:
                by_user[user]['without_transcript'] += 1
        
        logger.info("\nRecordings by User:")
        for user, counts in sorted(by_user.items(), key=lambda x: x[1]['total'], reverse=True):
            logger.info(f"  {user}: {counts['total']} total, {counts['without_transcript']} without transcript")
        
        logger.info(f"""
Summary:
--------
Users Scanned:              {self.stats['users_scanned']}
Total Recordings:           {self.stats['total_recordings']}
Without Transcript:         {self.stats['recordings_without_transcript']}
Downloaded:                 {self.stats['downloaded']}
Successfully Transcribed:   {self.stats['transcribed']}
Errors:                     {self.stats['errors']}
""")
    
    def run(self, max_recordings=5):
        """Run the full process"""
        start_time = datetime.now()
        logger.info(f"Starting OneDrive recording transcription - {start_time}")
        
        try:
            # Step 1: Authenticate
            self.authenticate()
            
            # Step 2: Scan all users
            self.scan_all_users()
            
            # Step 3: Process recordings
            self.process_recordings(max_recordings=max_recordings)
            
            # Step 4: Generate report
            self.generate_report()
            
            end_time = datetime.now()
            duration = (end_time - start_time).total_seconds()
            logger.info(f"\nCompleted in {duration:.1f} seconds")
            
        except Exception as e:
            logger.error(f"Process failed: {e}")
            raise


def main():
    """Main entry point"""
    import argparse
    
    parser = argparse.ArgumentParser(description='Transcribe OneDrive recordings without transcripts')
    parser.add_argument('--max', type=int, default=5, help='Maximum recordings to process (default: 5)')
    parser.add_argument('--scan-only', action='store_true', help='Only scan, do not transcribe')
    args = parser.parse_args()
    
    transcriber = OneDriveRecordingTranscriber()
    
    if args.scan_only:
        transcriber.authenticate()
        transcriber.scan_all_users()
        transcriber.generate_report()
    else:
        transcriber.run(max_recordings=args.max)


if __name__ == '__main__':
    main()
