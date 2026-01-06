"""
Transcribe remaining HR check-in recordings using Azure Speech Service.
Processes all local check-in recordings that don't have VTT transcripts yet.
Converts MP4 to WAV using FFmpeg before transcription.
"""
import os
import sys
import json
import time
import re
import subprocess
import tempfile
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv
import azure.cognitiveservices.speech as speechsdk

load_dotenv()

# Azure Speech Configuration
AZURE_SPEECH_KEY = os.getenv('AZURE_SPEECH_KEY')
AZURE_SPEECH_REGION = os.getenv('AZURE_SPEECH_REGION', 'eastus')

# Paths
RECORDINGS_DIR = Path('recordings')
TRANSCRIPTS_DIR = Path('transcripts')
TRANSCRIPTS_DIR.mkdir(exist_ok=True)

PROGRESS_FILE = Path('transcription_progress_hr.json')


def format_time_vtt(seconds):
    """Format seconds to VTT timestamp format: HH:MM:SS.mmm"""
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    secs = seconds % 60
    return f"{hours:02d}:{minutes:02d}:{secs:06.3f}"


def convert_to_wav(input_path, output_path):
    """Convert audio/video file to WAV format using FFmpeg"""
    cmd = [
        'ffmpeg', '-y', '-i', str(input_path),
        '-vn',  # No video
        '-acodec', 'pcm_s16le',  # PCM 16-bit
        '-ar', '16000',  # 16kHz sample rate
        '-ac', '1',  # Mono
        str(output_path)
    ]
    
    result = subprocess.run(
        cmd, 
        capture_output=True, 
        text=True,
        creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
    )
    
    return result.returncode == 0


def is_checkin_meeting(name):
    """Check if the name indicates a check-in meeting"""
    name_lower = name.lower()
    return 'check-in' in name_lower or 'checkin' in name_lower


def transcribe_audio_file(audio_path, output_vtt_path):
    """Transcribe an audio/video file using Azure Speech Service with continuous recognition."""
    
    speech_config = speechsdk.SpeechConfig(
        subscription=AZURE_SPEECH_KEY, 
        region=AZURE_SPEECH_REGION
    )
    speech_config.speech_recognition_language = "en-US"
    speech_config.request_word_level_timestamps()
    
    audio_config = speechsdk.audio.AudioConfig(filename=str(audio_path))
    speech_recognizer = speechsdk.SpeechRecognizer(
        speech_config=speech_config, 
        audio_config=audio_config
    )
    
    all_results = []
    done = False
    
    def handle_result(evt):
        if evt.result.reason == speechsdk.ResultReason.RecognizedSpeech:
            result = evt.result
            offset_seconds = result.offset / 10_000_000
            duration_seconds = result.duration / 10_000_000
            
            all_results.append({
                'text': result.text,
                'offset': offset_seconds,
                'duration': duration_seconds
            })
    
    def handle_canceled(evt):
        nonlocal done
        if evt.reason == speechsdk.CancellationReason.EndOfStream:
            pass
        elif evt.reason == speechsdk.CancellationReason.Error:
            print(f"    ❌ Error: {evt.error_details}")
        done = True
    
    def handle_session_stopped(evt):
        nonlocal done
        done = True
    
    speech_recognizer.recognized.connect(handle_result)
    speech_recognizer.canceled.connect(handle_canceled)
    speech_recognizer.session_stopped.connect(handle_session_stopped)
    
    speech_recognizer.start_continuous_recognition()
    
    while not done:
        time.sleep(0.5)
    
    speech_recognizer.stop_continuous_recognition()
    
    if not all_results:
        return False, "No speech detected"
    
    # Write VTT file
    with open(output_vtt_path, 'w', encoding='utf-8') as f:
        f.write("WEBVTT\n\n")
        
        for i, result in enumerate(all_results, 1):
            start_time = format_time_vtt(result['offset'])
            end_time = format_time_vtt(result['offset'] + result['duration'])
            
            f.write(f"{i}\n")
            f.write(f"{start_time} --> {end_time}\n")
            f.write(f"{result['text']}\n\n")
    
    return True, f"Transcribed {len(all_results)} segments"


def load_progress():
    """Load transcription progress"""
    if PROGRESS_FILE.exists():
        with open(PROGRESS_FILE, 'r') as f:
            return json.load(f)
    return {'completed': [], 'failed': [], 'skipped': []}


def save_progress(progress):
    """Save transcription progress"""
    with open(PROGRESS_FILE, 'w') as f:
        json.dump(progress, f, indent=2)


def find_recordings_to_transcribe():
    """Find all check-in recordings that don't have transcripts yet"""
    recordings = []
    
    # Get existing transcript basenames
    existing_transcripts = set()
    for vtt in TRANSCRIPTS_DIR.glob('*.vtt'):
        existing_transcripts.add(vtt.stem)
    
    # Search all recordings
    for item in RECORDINGS_DIR.rglob('*'):
        if item.is_file() and item.suffix.lower() in ['.mp4', '.m4a', '.webm']:
            # Check if it's a check-in meeting
            if is_checkin_meeting(item.name) or is_checkin_meeting(item.parent.name):
                # Check if transcript already exists
                if item.stem not in existing_transcripts:
                    recordings.append(item)
    
    return sorted(recordings, key=lambda x: x.name)


def main():
    print("="*70)
    print("HR CHECK-IN TRANSCRIPTION")
    print(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*70)
    
    if not AZURE_SPEECH_KEY:
        print("❌ AZURE_SPEECH_KEY not found in environment")
        return
    
    print(f"✅ Azure Speech Key configured")
    print(f"   Region: {AZURE_SPEECH_REGION}")
    
    # Find recordings to transcribe
    recordings = find_recordings_to_transcribe()
    print(f"\n📊 Found {len(recordings)} recordings to transcribe\n")
    
    if not recordings:
        print("✅ All recordings have been transcribed!")
        return
    
    # Load progress
    progress = load_progress()
    
    # Calculate total size
    total_size = sum(r.stat().st_size for r in recordings) / (1024*1024*1024)
    print(f"   Total size: {total_size:.2f} GB")
    
    # Transcribe each recording
    success_count = 0
    fail_count = 0
    
    for i, recording in enumerate(recordings, 1):
        filename = recording.name
        size_mb = recording.stat().st_size / (1024*1024)
        
        # Skip if already processed
        if filename in progress['completed'] or filename in progress['failed']:
            print(f"\n[{i}/{len(recordings)}] ⏭️ Skipping (already processed): {filename}")
            continue
        
        print(f"\n[{i}/{len(recordings)}] 🎙️ Transcribing: {filename}")
        print(f"    Size: {size_mb:.1f} MB")
        
        output_path = TRANSCRIPTS_DIR / f"{recording.stem}.vtt"
        
        try:
            start_time = time.time()
            
            # Convert to WAV if needed (MP4, M4A files need conversion)
            if recording.suffix.lower() in ['.mp4', '.m4a', '.webm', '.mkv']:
                print(f"    🔄 Converting to WAV...")
                wav_path = Path(tempfile.gettempdir()) / f"{recording.stem}_temp.wav"
                if not convert_to_wav(recording, wav_path):
                    print(f"    ❌ Failed to convert to WAV")
                    progress['failed'].append(filename)
                    fail_count += 1
                    save_progress(progress)
                    continue
                audio_path = wav_path
            else:
                audio_path = recording
            
            success, message = transcribe_audio_file(audio_path, output_path)
            elapsed = time.time() - start_time
            
            # Clean up temp WAV file
            if audio_path != recording and audio_path.exists():
                audio_path.unlink()
            
            if success:
                print(f"    ✅ {message}")
                print(f"    ⏱️ Time: {elapsed:.1f}s")
                print(f"    💾 Saved: {output_path.name}")
                progress['completed'].append(filename)
                success_count += 1
            else:
                print(f"    ⚠️ {message}")
                progress['failed'].append(filename)
                fail_count += 1
                
        except Exception as e:
            print(f"    ❌ Error: {e}")
            progress['failed'].append(filename)
            fail_count += 1
        
        # Save progress after each file
        save_progress(progress)
    
    # Final summary
    print("\n" + "="*70)
    print("TRANSCRIPTION COMPLETE")
    print("="*70)
    print(f"  ✅ Successful: {success_count}")
    print(f"  ❌ Failed: {fail_count}")
    print(f"  📁 Total transcripts: {len(list(TRANSCRIPTS_DIR.glob('*.vtt')))}")


if __name__ == '__main__':
    main()
