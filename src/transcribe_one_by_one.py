"""
Transcribe Recordings One-by-One with Progress Tracking
========================================================
This script processes ONE recording at a time and saves progress immediately.
Much more resilient to crashes and network issues.

Usage:
    python src/transcribe_one_by_one.py              # Process next recording
    python src/transcribe_one_by_one.py --count 10   # Process 10 recordings
    python src/transcribe_one_by_one.py --list       # List pending recordings
    python src/transcribe_one_by_one.py --status     # Show progress status
"""

import os
import sys
import json
import requests
import time
from datetime import datetime
from pathlib import Path
from azure.identity import ClientSecretCredential
from dotenv import load_dotenv
import subprocess
import logging
import re
import argparse

# Setup logging  
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/transcribe_progress.log', encoding='utf-8'),
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
SPEECH_KEY = os.getenv('AZURE_SPEECH_KEY')
SPEECH_REGION = os.getenv('AZURE_SPEECH_REGION', 'eastus')

# Paths
BASE_DIR = Path(__file__).parent.parent
TRANSCRIPTS_DIR = BASE_DIR / 'transcripts'
RECORDINGS_DIR = BASE_DIR / 'recordings'
PROGRESS_FILE = BASE_DIR / 'transcription_progress.json'

TRANSCRIPTS_DIR.mkdir(exist_ok=True)
RECORDINGS_DIR.mkdir(exist_ok=True)

class TranscriptionManager:
    def __init__(self):
        self.credential = None
        self.access_token = None
        self.token_expires = 0
        self.progress = self.load_progress()
        
    def load_progress(self):
        """Load progress from file"""
        if PROGRESS_FILE.exists():
            with open(PROGRESS_FILE, 'r') as f:
                return json.load(f)
        return {
            'completed': [],  # List of completed recording IDs
            'failed': [],     # List of failed recording IDs  
            'last_scan': None,
            'pending_recordings': [],
            'stats': {'total_completed': 0, 'total_failed': 0}
        }
    
    def save_progress(self):
        """Save progress to file"""
        with open(PROGRESS_FILE, 'w') as f:
            json.dump(self.progress, f, indent=2, default=str)
        logger.info("Progress saved!")
    
    def get_access_token(self):
        """Get fresh access token"""
        if time.time() < self.token_expires - 300:  # 5 min buffer
            return self.access_token
            
        logger.info("Getting fresh access token...")
        self.credential = ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
        token = self.credential.get_token("https://graph.microsoft.com/.default")
        self.access_token = token.token
        self.token_expires = token.expires_on
        return self.access_token
    
    def graph_request(self, url):
        """Make Graph API request with auto-refresh"""
        headers = {'Authorization': f'Bearer {self.get_access_token()}'}
        response = requests.get(url, headers=headers, timeout=60)
        response.raise_for_status()
        return response.json()
    
    def scan_recordings(self, force_rescan=False):
        """Scan all users for recordings without transcripts"""
        # Check if we have recent scan
        if not force_rescan and self.progress['pending_recordings']:
            logger.info(f"Using cached scan - {len(self.progress['pending_recordings'])} pending recordings")
            return
        
        logger.info("Scanning all users for recordings...")
        
        # Get existing transcript filenames
        existing_transcripts = set()
        for vtt in TRANSCRIPTS_DIR.glob('*.vtt'):
            # Extract date/time prefix from filename
            existing_transcripts.add(vtt.stem[:15])  # YYYYMMDD_HHMMSS
        
        # Get all users
        users = self.graph_request("https://graph.microsoft.com/v1.0/users?$select=id,displayName,mail")
        
        all_recordings = []
        for user in users.get('value', []):
            user_id = user['id']
            user_name = user.get('displayName', 'Unknown')
            user_recordings = 0
            
            try:
                # Use the /Recordings folder path (same as working script)
                url = f"https://graph.microsoft.com/v1.0/users/{user_id}/drive/root:/Recordings:/children?$top=500"
                headers = {'Authorization': f'Bearer {self.get_access_token()}'}
                response = requests.get(url, headers=headers, timeout=30)
                
                if response.status_code == 200:
                    items = response.json().get('value', [])
                    
                    for item in items:
                        name = item.get('name', '')
                        if not name.lower().endswith('.mp4'):
                            continue
                            
                        created = item.get('createdDateTime', '')
                        
                        # Create expected transcript prefix
                        if created:
                            date_str = created[:10].replace('-', '')
                            time_str = created[11:19].replace(':', '')
                            prefix = f"{date_str}_{time_str}"
                            
                            # Skip if transcript already exists
                            if prefix in existing_transcripts:
                                continue
                            
                            # Skip if already completed
                            rec_id = f"{user_id}_{item['id']}"
                            if rec_id in self.progress['completed']:
                                continue
                                
                            all_recordings.append({
                                'id': rec_id,
                                'user_id': user_id,
                                'user_name': user_name,
                                'file_id': item['id'],
                                'file_name': name,
                                'created': created,
                                'size': item.get('size', 0)
                            })
                            user_recordings += 1
                            
                    if user_recordings > 0:
                        logger.info(f"  {user_name}: {user_recordings} recordings found")
                        
                elif response.status_code == 404:
                    # No Recordings folder - skip
                    pass
                else:
                    logger.warning(f"  {user_name}: Error {response.status_code}")
                            
            except Exception as e:
                logger.warning(f"  Error scanning {user_name}: {e}")
                continue
        
        # Sort by date (newest first)
        all_recordings.sort(key=lambda x: x.get('created', ''), reverse=True)
        
        self.progress['pending_recordings'] = all_recordings
        self.progress['last_scan'] = datetime.now().isoformat()
        self.save_progress()
        
        logger.info(f"Found {len(all_recordings)} recordings without transcripts")
    
    def download_recording(self, recording):
        """Download a single recording"""
        user_id = recording['user_id']
        file_id = recording['file_id']
        file_name = recording['file_name']
        
        # Get download URL
        url = f"https://graph.microsoft.com/v1.0/users/{user_id}/drive/items/{file_id}"
        headers = {'Authorization': f'Bearer {self.get_access_token()}'}
        
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()
        
        download_url = response.json().get('@microsoft.graph.downloadUrl')
        if not download_url:
            raise Exception("No download URL available")
        
        # Create safe filename
        safe_name = re.sub(r'[<>:"/\\|?*]', '', file_name)[:50]
        created = recording.get('created', '')
        date_prefix = created[:10].replace('-', '') if created else 'unknown'
        time_prefix = created[11:19].replace(':', '') if created else ''
        
        local_path = RECORDINGS_DIR / f"{date_prefix}_{time_prefix}_{safe_name}"
        
        # Download with progress
        logger.info(f"  Downloading {recording['size']/1024/1024:.1f} MB...")
        with requests.get(download_url, stream=True, timeout=600) as r:
            r.raise_for_status()
            with open(local_path, 'wb') as f:
                for chunk in r.iter_content(chunk_size=8192*8):
                    f.write(chunk)
        
        logger.info(f"  Downloaded: {local_path.name}")
        return local_path
    
    def extract_audio(self, video_path):
        """Extract audio from video using FFmpeg"""
        audio_path = video_path.with_suffix('.wav')
        
        cmd = [
            'ffmpeg', '-y', '-i', str(video_path),
            '-vn', '-acodec', 'pcm_s16le', '-ar', '16000', '-ac', '1',
            str(audio_path)
        ]
        
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
        
        if result.returncode != 0:
            raise Exception(f"FFmpeg error: {result.stderr}")
        
        logger.info(f"  Extracted audio: {audio_path.name}")
        return audio_path
    
    def transcribe_audio(self, audio_path):
        """Transcribe audio using Azure Speech SDK"""
        import azure.cognitiveservices.speech as speechsdk
        
        speech_config = speechsdk.SpeechConfig(subscription=SPEECH_KEY, region=SPEECH_REGION)
        speech_config.speech_recognition_language = "en-US"
        
        audio_config = speechsdk.AudioConfig(filename=str(audio_path))
        speech_recognizer = speechsdk.SpeechRecognizer(
            speech_config=speech_config,
            audio_config=audio_config
        )
        
        transcript_parts = []
        done = False
        error_msg = None
        
        def recognized_cb(evt):
            if evt.result.reason == speechsdk.ResultReason.RecognizedSpeech:
                transcript_parts.append(evt.result.text)
                # Show progress
                if len(transcript_parts) % 10 == 0:
                    print(f"  ... {len(transcript_parts)} segments transcribed", end='\r')
        
        def stop_cb(evt):
            nonlocal done
            done = True
        
        def canceled_cb(evt):
            nonlocal done, error_msg
            done = True
            if evt.result.cancellation_details.reason == speechsdk.CancellationReason.Error:
                error_msg = evt.result.cancellation_details.error_details
        
        speech_recognizer.recognized.connect(recognized_cb)
        speech_recognizer.session_stopped.connect(stop_cb)
        speech_recognizer.canceled.connect(canceled_cb)
        
        logger.info("  Transcribing... (may take several minutes)")
        speech_recognizer.start_continuous_recognition()
        
        # Wait for completion (max 30 minutes for long recordings)
        start_time = time.time()
        while not done and (time.time() - start_time) < 1800:
            time.sleep(1)
        
        speech_recognizer.stop_continuous_recognition()
        
        if error_msg:
            raise Exception(f"Transcription error: {error_msg}")
        
        full_transcript = ' '.join(transcript_parts)
        logger.info(f"  Transcription complete! ({len(full_transcript)} chars, {len(transcript_parts)} segments)")
        return full_transcript
    
    def save_transcript(self, transcript_text, recording):
        """Save transcript as VTT file"""
        created = recording.get('created', '')
        user_name = recording.get('user_name', 'Unknown')
        file_name = recording.get('file_name', 'recording')
        
        date_str = created[:10].replace('-', '') if created else 'unknown'
        time_str = created[11:19].replace(':', '') if created else ''
        safe_name = re.sub(r'[<>:"/\\|?*]', '', file_name.replace('.mp4', ''))[:40]
        
        vtt_filename = f"{date_str}_{time_str}_{safe_name}.vtt"
        vtt_path = TRANSCRIPTS_DIR / vtt_filename
        
        vtt_content = f"""WEBVTT

NOTE Transcribed from OneDrive recording using Azure Speech-to-Text
NOTE Source: {user_name} - {file_name}
NOTE Date: {created}

00:00:00.000 --> 99:59:59.999
{transcript_text}
"""
        
        with open(vtt_path, 'w', encoding='utf-8') as f:
            f.write(vtt_content)
        
        logger.info(f"  ✓ Saved: {vtt_filename}")
        return vtt_path
    
    def process_one(self):
        """Process ONE recording and save progress"""
        if not self.progress['pending_recordings']:
            logger.info("No pending recordings! Run with --rescan to scan again.")
            return False
        
        # Get next recording
        recording = self.progress['pending_recordings'][0]
        rec_id = recording['id']
        
        logger.info("="*60)
        logger.info(f"Processing: {recording['user_name']} - {recording['file_name']}")
        logger.info(f"Date: {recording['created']}")
        logger.info(f"Size: {recording['size']/1024/1024:.1f} MB")
        logger.info(f"Pending: {len(self.progress['pending_recordings'])} recordings")
        logger.info("="*60)
        
        video_path = None
        audio_path = None
        
        try:
            # Step 1: Download
            video_path = self.download_recording(recording)
            
            # Step 2: Extract audio
            audio_path = self.extract_audio(video_path)
            
            # Step 3: Transcribe
            transcript = self.transcribe_audio(audio_path)
            
            if not transcript or len(transcript) < 50:
                raise Exception("Transcription too short or empty")
            
            # Step 4: Save transcript
            self.save_transcript(transcript, recording)
            
            # Mark as completed
            self.progress['completed'].append(rec_id)
            self.progress['pending_recordings'].pop(0)
            self.progress['stats']['total_completed'] += 1
            self.save_progress()
            
            logger.info(f"✓ SUCCESS! {self.progress['stats']['total_completed']} completed, {len(self.progress['pending_recordings'])} remaining")
            return True
            
        except Exception as e:
            logger.error(f"✗ FAILED: {e}")
            
            # Mark as failed and move to end of queue (retry later)
            self.progress['failed'].append(rec_id)
            self.progress['pending_recordings'].pop(0)
            self.progress['stats']['total_failed'] += 1
            self.save_progress()
            
            return False
            
        finally:
            # Cleanup temp files
            if audio_path and audio_path.exists():
                audio_path.unlink()
            # Keep video file for now
    
    def show_status(self):
        """Show current progress status"""
        print("\n" + "="*60)
        print("TRANSCRIPTION PROGRESS STATUS")
        print("="*60)
        print(f"Last scan: {self.progress.get('last_scan', 'Never')}")
        print(f"Completed: {self.progress['stats'].get('total_completed', 0)}")
        print(f"Failed: {self.progress['stats'].get('total_failed', 0)}")
        print(f"Pending: {len(self.progress.get('pending_recordings', []))}")
        
        if self.progress.get('pending_recordings'):
            print("\nNext 5 recordings to process:")
            for i, rec in enumerate(self.progress['pending_recordings'][:5], 1):
                print(f"  {i}. {rec['user_name']} - {rec['file_name'][:40]}...")
                print(f"     Date: {rec['created'][:10]}, Size: {rec['size']/1024/1024:.1f} MB")
        print("="*60)
    
    def list_pending(self):
        """List all pending recordings"""
        if not self.progress.get('pending_recordings'):
            print("No pending recordings!")
            return
            
        print(f"\nPending recordings ({len(self.progress['pending_recordings'])} total):")
        for i, rec in enumerate(self.progress['pending_recordings'], 1):
            print(f"{i:3}. [{rec['created'][:10]}] {rec['user_name'][:20]:20} - {rec['file_name'][:40]}")


def main():
    parser = argparse.ArgumentParser(description='Transcribe recordings one by one')
    parser.add_argument('--count', type=int, default=1, help='Number of recordings to process')
    parser.add_argument('--rescan', action='store_true', help='Force rescan of all users')
    parser.add_argument('--status', action='store_true', help='Show progress status')
    parser.add_argument('--list', action='store_true', help='List pending recordings')
    args = parser.parse_args()
    
    manager = TranscriptionManager()
    
    if args.status:
        manager.show_status()
        return
    
    if args.list:
        manager.scan_recordings()
        manager.list_pending()
        return
    
    # Scan for recordings
    manager.scan_recordings(force_rescan=args.rescan)
    
    # Process recordings
    for i in range(args.count):
        if not manager.progress['pending_recordings']:
            logger.info("All recordings processed!")
            break
            
        logger.info(f"\n{'='*60}")
        logger.info(f"BATCH PROGRESS: {i+1}/{args.count}")
        logger.info(f"{'='*60}")
        
        manager.process_one()
        
        # Small delay between recordings
        if i < args.count - 1:
            time.sleep(3)
    
    manager.show_status()


if __name__ == '__main__':
    main()
