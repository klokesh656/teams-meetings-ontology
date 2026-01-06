"""
Download Louise check-in recordings from OneDrive and transcribe them.
"""

import os
import sys
import json
import asyncio
import aiohttp
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv
from azure.identity import ClientSecretCredential
from msgraph import GraphServiceClient
import azure.cognitiveservices.speech as speechsdk
import subprocess

# Load environment variables
load_dotenv()

# Configuration
TENANT_ID = os.getenv("AZURE_TENANT_ID")
CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
SPEECH_KEY = os.getenv("AZURE_SPEECH_KEY")
SPEECH_REGION = os.getenv("AZURE_SPEECH_REGION", "eastus")

RECORDINGS_DIR = Path("recordings")
TRANSCRIPTS_DIR = Path("transcripts")
PROGRESS_FILE = Path("louise_download_progress.json")

# Load the report
REPORT_FILE = Path("output/louise_checkins_report_20260101_004344.json")


def load_progress():
    """Load download/transcription progress."""
    if PROGRESS_FILE.exists():
        with open(PROGRESS_FILE, 'r') as f:
            return json.load(f)
    return {"downloaded": [], "transcribed": [], "failed": []}


def save_progress(progress):
    """Save progress."""
    with open(PROGRESS_FILE, 'w') as f:
        json.dump(progress, f, indent=2)


def get_graph_client():
    """Create Microsoft Graph client."""
    credential = ClientSecretCredential(
        tenant_id=TENANT_ID,
        client_id=CLIENT_ID,
        client_secret=CLIENT_SECRET
    )
    scopes = ["https://graph.microsoft.com/.default"]
    return GraphServiceClient(credentials=credential, scopes=scopes)


async def download_file_direct(download_url: str, local_path: Path, session: aiohttp.ClientSession):
    """Download a file using direct download URL (no auth needed for tempauth URLs)."""
    try:
        async with session.get(download_url) as response:
            if response.status == 200:
                local_path.parent.mkdir(parents=True, exist_ok=True)
                total_size = int(response.headers.get('content-length', 0))
                downloaded = 0
                with open(local_path, 'wb') as f:
                    while True:
                        chunk = await response.content.read(65536)  # 64KB chunks
                        if not chunk:
                            break
                        f.write(chunk)
                        downloaded += len(chunk)
                        if total_size > 0:
                            pct = (downloaded / total_size) * 100
                            print(f"\r    Progress: {pct:.1f}% ({downloaded / 1024 / 1024:.1f} MB)", end="")
                print()  # New line after progress
                return True
            else:
                print(f"    ❌ Download failed: HTTP {response.status}")
                return False
    except Exception as e:
        print(f"    ❌ Download error: {e}")
        return False


async def download_file_from_onedrive(client, download_url: str, local_path: Path, session: aiohttp.ClientSession):
    """Download a file from OneDrive using the download URL."""
    try:
        # Get access token
        credential = ClientSecretCredential(
            tenant_id=TENANT_ID,
            client_id=CLIENT_ID,
            client_secret=CLIENT_SECRET
        )
        token = credential.get_token("https://graph.microsoft.com/.default")
        
        headers = {
            "Authorization": f"Bearer {token.token}"
        }
        
        async with session.get(download_url, headers=headers) as response:
            if response.status == 200:
                local_path.parent.mkdir(parents=True, exist_ok=True)
                with open(local_path, 'wb') as f:
                    while True:
                        chunk = await response.content.read(8192)
                        if not chunk:
                            break
                        f.write(chunk)
                return True
            else:
                print(f"    ❌ Download failed: HTTP {response.status}")
                return False
    except Exception as e:
        print(f"    ❌ Download error: {e}")
        return False


async def get_drive_item_download_url(client, user_id: str, item_id: str):
    """Get the download URL for a drive item."""
    try:
        item = await client.users.by_user_id(user_id).drive.items.by_drive_item_id(item_id).get()
        if item and hasattr(item, 'additional_data') and '@microsoft.graph.downloadUrl' in item.additional_data:
            return item.additional_data['@microsoft.graph.downloadUrl']
        return None
    except Exception as e:
        print(f"    ❌ Error getting download URL: {e}")
        return None


def extract_audio(video_path: Path, audio_path: Path) -> bool:
    """Extract audio from video using FFmpeg."""
    try:
        cmd = [
            "ffmpeg", "-i", str(video_path),
            "-vn", "-acodec", "pcm_s16le",
            "-ar", "16000", "-ac", "1",
            "-y", str(audio_path)
        ]
        result = subprocess.run(cmd, capture_output=True, text=True)
        return result.returncode == 0 and audio_path.exists()
    except Exception as e:
        print(f"    ❌ Audio extraction error: {e}")
        return False


def transcribe_audio(audio_path: Path, output_path: Path) -> bool:
    """Transcribe audio using Azure Speech SDK."""
    try:
        speech_config = speechsdk.SpeechConfig(
            subscription=SPEECH_KEY,
            region=SPEECH_REGION
        )
        speech_config.speech_recognition_language = "en-US"
        speech_config.request_word_level_timestamps()
        
        audio_config = speechsdk.AudioConfig(filename=str(audio_path))
        recognizer = speechsdk.SpeechRecognizer(
            speech_config=speech_config,
            audio_config=audio_config
        )
        
        all_results = []
        done = False
        
        def on_recognized(evt):
            if evt.result.reason == speechsdk.ResultReason.RecognizedSpeech:
                all_results.append({
                    "text": evt.result.text,
                    "offset": evt.result.offset,
                    "duration": evt.result.duration
                })
        
        def on_canceled(evt):
            nonlocal done
            done = True
        
        def on_stopped(evt):
            nonlocal done
            done = True
        
        recognizer.recognized.connect(on_recognized)
        recognizer.canceled.connect(on_canceled)
        recognizer.session_stopped.connect(on_stopped)
        
        recognizer.start_continuous_recognition()
        
        while not done:
            import time
            time.sleep(0.5)
        
        recognizer.stop_continuous_recognition()
        
        # Generate VTT
        if all_results:
            vtt_content = generate_vtt(all_results)
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(vtt_content)
            return True
        return False
        
    except Exception as e:
        print(f"    ❌ Transcription error: {e}")
        return False


def generate_vtt(results):
    """Generate VTT content from transcription results."""
    vtt_lines = ["WEBVTT", ""]
    
    for i, result in enumerate(results, 1):
        start_time = result["offset"] / 10_000_000  # Convert from 100ns to seconds
        duration = result["duration"] / 10_000_000
        end_time = start_time + duration
        
        start_str = format_vtt_time(start_time)
        end_str = format_vtt_time(end_time)
        
        vtt_lines.append(str(i))
        vtt_lines.append(f"{start_str} --> {end_str}")
        vtt_lines.append(result["text"])
        vtt_lines.append("")
    
    return "\n".join(vtt_lines)


def format_vtt_time(seconds):
    """Format seconds as VTT timestamp."""
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    secs = seconds % 60
    return f"{hours:02d}:{minutes:02d}:{secs:06.3f}"


async def main():
    print("=" * 70)
    print("LOUISE CHECK-IN RECORDINGS - DOWNLOAD & TRANSCRIBE")
    print("=" * 70)
    
    # Load the report
    if not REPORT_FILE.exists():
        print(f"❌ Report file not found: {REPORT_FILE}")
        return
    
    with open(REPORT_FILE, 'r') as f:
        report = json.load(f)
    
    # Get meetings that can be transcribed from all_meetings
    all_meetings = report.get("all_meetings", [])
    to_transcribe = [m for m in all_meetings if m.get("can_transcribe", False)]
    print(f"\nFound {len(to_transcribe)} recordings that can be transcribed")
    
    # Separate local vs OneDrive recordings
    local_recordings = []
    onedrive_recordings = []
    
    for rec in to_transcribe:
        if rec.get("has_local_recording", False):
            local_recordings.append(rec)
        elif rec.get("has_onedrive_recording", False) and rec.get("onedrive_info"):
            onedrive_recordings.append(rec)
    
    print(f"  Local recordings: {len(local_recordings)}")
    print(f"  OneDrive recordings: {len(onedrive_recordings)}")
    
    # Load progress
    progress = load_progress()
    
    # Process OneDrive recordings first - download them
    print("\n" + "=" * 70)
    print("PHASE 1: DOWNLOADING FROM ONEDRIVE")
    print("=" * 70)
    
    downloaded_count = 0
    async with aiohttp.ClientSession() as session:
        for i, rec in enumerate(onedrive_recordings, 1):
            va_name = rec.get("va_name", "Unknown")
            date = rec.get("date", "unknown")
            meeting_key = f"{date}_{va_name}"
            
            if meeting_key in progress["downloaded"]:
                print(f"\n[{i}/{len(onedrive_recordings)}] Already downloaded: Louise x {va_name}")
                continue
            
            onedrive_info = rec.get("onedrive_info", {})
            size_mb = onedrive_info.get("size_mb", 0)
            download_url = onedrive_info.get("download_url", "")
            
            print(f"\n[{i}/{len(onedrive_recordings)}] {date} - Louise x {va_name}")
            print(f"    Size: {size_mb:.1f} MB")
            
            # Create local filename
            date_str = date.replace("-", "")
            va_name_safe = va_name.replace(" ", "_")
            filename = f"{date_str}_Integration_Team_Check-in_Louise_x_{va_name_safe}.mp4"
            local_path = RECORDINGS_DIR / filename
            
            if local_path.exists():
                print(f"    ✓ Already exists locally")
                progress["downloaded"].append(meeting_key)
                save_progress(progress)
                continue
            
            if not download_url:
                print("    ⚠️ No download URL available, skipping")
                progress["failed"].append({"key": meeting_key, "reason": "No download URL"})
                save_progress(progress)
                continue
            
            # Download using the direct download URL from the report
            print(f"    Downloading...")
            success = await download_file_direct(download_url, local_path, session)
            
            if success:
                print(f"    ✅ Downloaded: {filename}")
                progress["downloaded"].append(meeting_key)
                downloaded_count += 1
            else:
                progress["failed"].append({"key": meeting_key, "reason": "Download failed"})
            
            save_progress(progress)
    
    print(f"\n✅ Downloaded {downloaded_count} new recordings")
    
    # Now transcribe all local recordings (including newly downloaded)
    print("\n" + "=" * 70)
    print("PHASE 2: TRANSCRIBING LOCAL RECORDINGS")
    print("=" * 70)
    
    # Find all Louise recordings in local folder without transcripts
    all_local_louise = []
    for mp4_file in RECORDINGS_DIR.glob("*.mp4"):
        if "louise" in mp4_file.name.lower():
            # Check if transcript exists
            vtt_name = mp4_file.stem + ".vtt"
            vtt_path = TRANSCRIPTS_DIR / vtt_name
            if not vtt_path.exists():
                all_local_louise.append(mp4_file)
    
    print(f"\nFound {len(all_local_louise)} local Louise recordings without transcripts")
    
    transcribed_count = 0
    for i, mp4_file in enumerate(all_local_louise, 1):
        meeting_key = mp4_file.stem
        
        if meeting_key in progress["transcribed"]:
            print(f"\n[{i}/{len(all_local_louise)}] Already transcribed: {mp4_file.name[:50]}...")
            continue
        
        print(f"\n[{i}/{len(all_local_louise)}] Transcribing: {mp4_file.name[:60]}...")
        
        # Extract audio
        audio_path = mp4_file.with_suffix(".wav")
        print("    Extracting audio...")
        
        if not extract_audio(mp4_file, audio_path):
            print("    ❌ Audio extraction failed")
            progress["failed"].append({"key": meeting_key, "reason": "Audio extraction failed"})
            save_progress(progress)
            continue
        
        # Transcribe
        vtt_path = TRANSCRIPTS_DIR / (mp4_file.stem + ".vtt")
        print("    Transcribing with Azure Speech...")
        
        if transcribe_audio(audio_path, vtt_path):
            print(f"    ✅ Transcript saved: {vtt_path.name}")
            progress["transcribed"].append(meeting_key)
            transcribed_count += 1
        else:
            print("    ❌ Transcription failed")
            progress["failed"].append({"key": meeting_key, "reason": "Transcription failed"})
        
        # Clean up audio file
        if audio_path.exists():
            audio_path.unlink()
        
        save_progress(progress)
    
    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    print(f"  Downloaded: {downloaded_count} recordings")
    print(f"  Transcribed: {transcribed_count} recordings")
    print(f"  Failed: {len(progress.get('failed', []))} recordings")
    print(f"\nProgress saved to: {PROGRESS_FILE}")


if __name__ == "__main__":
    asyncio.run(main())
