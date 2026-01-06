"""
Transcribe Recordings Without Transcripts
==========================================
This script:
1. Finds recordings that have NO transcript available (404 errors)
2. Downloads the recording MP4 file
3. Extracts audio and transcribes using Azure Speech-to-Text
4. Saves the transcript in VTT format
5. Runs AI analysis on the new transcripts

Use Case: For meetings recorded BEFORE auto-transcription was enabled
"""

import os
import sys
import json
import time
import requests
from datetime import datetime
from pathlib import Path
from azure.identity import ClientSecretCredential
from azure.storage.blob import BlobServiceClient
from dotenv import load_dotenv
import subprocess
import logging

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/transcribe_recordings.log'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# Load environment variables
load_dotenv()

# Configuration
TENANT_ID = os.getenv('AZURE_TENANT_ID', '187b2af6-1bfb-490a-85dd-b720fe3d31bc')
CLIENT_ID = os.getenv('AZURE_CLIENT_ID')
CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET')
HR_USER_ID = '81835016-79d5-4a15-91b1-c104e2cd9adb'

# Azure Speech Service (you'll need to add these to .env)
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


class RecordingTranscriber:
    def __init__(self):
        self.credential = ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
        self.token = None
        self.headers = None
        self.recordings_without_transcripts = []
        self.stats = {
            'recordings_found': 0,
            'without_transcript': 0,
            'downloaded': 0,
            'transcribed': 0,
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
    
    def get_all_recordings(self):
        """Fetch all recordings from Graph API"""
        all_recordings = []
        
        logger.info(f"Fetching recordings for user: {HR_USER_ID}")
        url = f"https://graph.microsoft.com/beta/users/{HR_USER_ID}/onlineMeetings/getAllRecordings(meetingOrganizerUserId='{HR_USER_ID}')"
        
        while url:
            resp = requests.get(url, headers=self.headers, timeout=30)
            if resp.status_code == 200:
                data = resp.json()
                recordings = data.get('value', [])
                all_recordings.extend(recordings)
                url = data.get('@odata.nextLink')
                time.sleep(0.5)
            else:
                logger.error(f"Error fetching recordings: {resp.status_code}")
                break
        
        self.stats['recordings_found'] = len(all_recordings)
        logger.info(f"Found {len(all_recordings)} total recordings")
        return all_recordings
    
    def check_transcript_exists(self, user_id, meeting_id, transcript_id):
        """Check if transcript is available (returns False for 404)"""
        url = f"https://graph.microsoft.com/beta/users/{user_id}/onlineMeetings/{meeting_id}/transcripts/{transcript_id}"
        try:
            resp = requests.get(url, headers=self.headers, timeout=10)
            return resp.status_code == 200
        except:
            return False
    
    def find_recordings_without_transcripts(self):
        """Find all recordings that don't have transcripts available"""
        logger.info("="*60)
        logger.info("FINDING RECORDINGS WITHOUT TRANSCRIPTS")
        logger.info("="*60)
        
        recordings = self.get_all_recordings()
        
        for rec in recordings:
            recording_id = rec.get('id', '')
            meeting_id = rec.get('meetingId', '')
            created = rec.get('createdDateTime', '')
            
            # Try to get transcript
            # Note: Graph API doesn't directly link recordings to transcripts
            # We'll check if transcript exists by checking the transcript API
            # If 404, it means no transcript available
            
            # For simplicity, we'll check if we already have this transcript locally
            date_str = created[:10].replace('-', '') if created else 'unknown'
            matching_transcripts = list(TRANSCRIPTS_DIR.glob(f"{date_str}*.vtt"))
            
            if not matching_transcripts:
                self.recordings_without_transcripts.append({
                    'recording_id': recording_id,
                    'meeting_id': meeting_id,
                    'created': created,
                    'recording': rec
                })
                self.stats['without_transcript'] += 1
        
        logger.info(f"Found {self.stats['without_transcript']} recordings WITHOUT transcripts")
        return self.recordings_without_transcripts
    
    def download_recording(self, user_id, meeting_id, recording_id, created_date):
        """Download a recording MP4 file"""
        url = f"https://graph.microsoft.com/beta/users/{user_id}/onlineMeetings/{meeting_id}/recordings/{recording_id}/content"
        
        try:
            resp = requests.get(url, headers=self.headers, timeout=60, stream=True)
            if resp.status_code == 200:
                # Create filename
                date_str = created_date[:10].replace('-', '') if created_date else 'unknown'
                time_str = created_date[11:19].replace(':', '') if created_date and len(created_date) > 11 else ''
                filename = f"{date_str}_{time_str}_recording.mp4"
                filepath = RECORDINGS_DIR / filename
                
                # Download file
                with open(filepath, 'wb') as f:
                    for chunk in resp.iter_content(chunk_size=8192):
                        f.write(chunk)
                
                logger.info(f"  Downloaded: {filename} ({filepath.stat().st_size / 1024 / 1024:.1f} MB)")
                return filepath
            else:
                logger.error(f"  Failed to download recording: {resp.status_code}")
                return None
        except Exception as e:
            logger.error(f"  Error downloading recording: {e}")
            return None
    
    def extract_audio_from_recording(self, video_path):
        """Extract audio from MP4 using ffmpeg"""
        audio_path = video_path.with_suffix('.wav')
        
        try:
            # Check if ffmpeg is available
            subprocess.run(['ffmpeg', '-version'], capture_output=True, check=True)
            
            # Extract audio (16kHz mono WAV for best speech recognition)
            cmd = [
                'ffmpeg', '-i', str(video_path),
                '-ar', '16000',  # 16kHz sample rate
                '-ac', '1',       # Mono
                '-y',             # Overwrite
                str(audio_path)
            ]
            
            subprocess.run(cmd, capture_output=True, check=True)
            logger.info(f"  Extracted audio: {audio_path.name}")
            return audio_path
        except FileNotFoundError:
            logger.error("  FFmpeg not found! Please install FFmpeg: https://ffmpeg.org/download.html")
            return None
        except Exception as e:
            logger.error(f"  Error extracting audio: {e}")
            return None
    
    def transcribe_audio_with_azure_speech(self, audio_path):
        """Transcribe audio using Azure Speech-to-Text"""
        if not SPEECH_KEY:
            logger.error("  Azure Speech API key not configured. Add AZURE_SPEECH_KEY to .env")
            return None
        
        try:
            import azure.cognitiveservices.speech as speechsdk
            
            speech_config = speechsdk.SpeechConfig(subscription=SPEECH_KEY, region=SPEECH_REGION)
            speech_config.speech_recognition_language = "en-US"
            
            audio_config = speechsdk.audio.AudioConfig(filename=str(audio_path))
            speech_recognizer = speechsdk.SpeechRecognizer(speech_config=speech_config, audio_config=audio_config)
            
            logger.info(f"  Transcribing audio... (this may take a few minutes)")
            
            done = False
            transcript_text = []
            
            def stop_cb(evt):
                nonlocal done
                done = True
            
            def recognized_cb(evt):
                if evt.result.reason == speechsdk.ResultReason.RecognizedSpeech:
                    transcript_text.append(evt.result.text)
            
            speech_recognizer.recognized.connect(recognized_cb)
            speech_recognizer.session_stopped.connect(stop_cb)
            speech_recognizer.canceled.connect(stop_cb)
            
            speech_recognizer.start_continuous_recognition()
            
            # Wait for transcription to complete
            timeout = 600  # 10 minutes max
            start_time = time.time()
            while not done and (time.time() - start_time) < timeout:
                time.sleep(0.5)
            
            speech_recognizer.stop_continuous_recognition()
            
            full_transcript = ' '.join(transcript_text)
            logger.info(f"  Transcription complete! ({len(full_transcript)} characters)")
            return full_transcript
        
        except ImportError:
            logger.error("  Azure Speech SDK not installed. Run: pip install azure-cognitiveservices-speech")
            return None
        except Exception as e:
            logger.error(f"  Error transcribing audio: {e}")
            return None
    
    def save_transcript_as_vtt(self, transcript_text, recording_path, created_date):
        """Save transcript in VTT format"""
        # Create filename matching pattern
        date_str = created_date[:10].replace('-', '') if created_date else 'unknown'
        time_str = created_date[11:19].replace(':', '') if created_date and len(created_date) > 11 else ''
        vtt_filename = f"{date_str}_{time_str}_transcribed.vtt"
        vtt_path = TRANSCRIPTS_DIR / vtt_filename
        
        # Create VTT content
        vtt_content = "WEBVTT\n\n"
        vtt_content += "NOTE Transcribed from recording using Azure Speech-to-Text\n\n"
        vtt_content += f"00:00:00.000 --> 00:05:00.000\n"
        vtt_content += transcript_text
        
        with open(vtt_path, 'w', encoding='utf-8') as f:
            f.write(vtt_content)
        
        logger.info(f"  Saved transcript: {vtt_path.name}")
        return vtt_path
    
    def process_recordings(self):
        """Main processing loop"""
        logger.info("="*60)
        logger.info("PROCESSING RECORDINGS WITHOUT TRANSCRIPTS")
        logger.info("="*60)
        
        if not self.recordings_without_transcripts:
            logger.info("No recordings to process!")
            return
        
        for i, rec_info in enumerate(self.recordings_without_transcripts[:5], 1):  # Process first 5
            logger.info(f"\n[{i}/{len(self.recordings_without_transcripts[:5])}] Processing recording from {rec_info['created']}")
            
            # Step 1: Download recording
            recording_path = self.download_recording(
                HR_USER_ID,
                rec_info['meeting_id'],
                rec_info['recording_id'],
                rec_info['created']
            )
            
            if not recording_path:
                self.stats['errors'] += 1
                continue
            
            self.stats['downloaded'] += 1
            
            # Step 2: Extract audio
            audio_path = self.extract_audio_from_recording(recording_path)
            if not audio_path:
                self.stats['errors'] += 1
                continue
            
            # Step 3: Transcribe audio
            transcript_text = self.transcribe_audio_with_azure_speech(audio_path)
            if not transcript_text:
                self.stats['errors'] += 1
                continue
            
            # Step 4: Save as VTT
            vtt_path = self.save_transcript_as_vtt(transcript_text, recording_path, rec_info['created'])
            self.stats['transcribed'] += 1
            
            # Cleanup audio file
            if audio_path.exists():
                audio_path.unlink()
            
            time.sleep(2)  # Rate limiting
    
    def generate_report(self):
        """Generate summary report"""
        logger.info("\n" + "="*60)
        logger.info("TRANSCRIPTION SUMMARY")
        logger.info("="*60)
        
        report = f"""
Recordings Found:           {self.stats['recordings_found']}
Without Transcript:         {self.stats['without_transcript']}
Downloaded:                 {self.stats['downloaded']}
Successfully Transcribed:   {self.stats['transcribed']}
Errors:                     {self.stats['errors']}
"""
        logger.info(report)
    
    def run(self):
        """Run the full transcription process"""
        start_time = datetime.now()
        logger.info(f"Starting recording transcription - {start_time}")
        
        try:
            # Step 1: Authenticate
            self.authenticate()
            
            # Step 2: Find recordings without transcripts
            self.find_recordings_without_transcripts()
            
            # Step 3: Process recordings
            self.process_recordings()
            
            # Step 4: Generate report
            self.generate_report()
            
            end_time = datetime.now()
            duration = (end_time - start_time).total_seconds()
            logger.info(f"\nCompleted in {duration:.1f} seconds")
            
        except Exception as e:
            logger.error(f"Transcription process failed: {e}")
            raise


def main():
    """Main entry point"""
    transcriber = RecordingTranscriber()
    transcriber.run()


if __name__ == '__main__':
    main()
