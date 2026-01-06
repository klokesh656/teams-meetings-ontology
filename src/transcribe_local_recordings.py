"""
Transcribe Local Recordings Without VTT Files
=============================================
Transcribes MP4 recordings in the recordings/ folder that don't have
corresponding VTT transcript files.

Usage:
    python src/transcribe_local_recordings.py              # Process all missing
    python src/transcribe_local_recordings.py --list       # List what needs transcription
    python src/transcribe_local_recordings.py --count 3    # Process only 3 recordings
"""

import os
import sys
import json
import time
import argparse
import subprocess
import logging
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/transcribe_local.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# Load environment variables
load_dotenv()

SPEECH_KEY = os.getenv('AZURE_SPEECH_KEY')
SPEECH_REGION = os.getenv('AZURE_SPEECH_REGION', 'eastus')

# Paths
BASE_DIR = Path(__file__).parent.parent
RECORDINGS_DIR = BASE_DIR / 'recordings'
TRANSCRIPTS_DIR = BASE_DIR / 'transcripts'
PROGRESS_FILE = BASE_DIR / 'local_transcription_progress.json'

TRANSCRIPTS_DIR.mkdir(exist_ok=True)


def find_recordings_without_transcripts():
    """Find local recordings that don't have VTT transcripts"""
    
    # Get all existing VTT files - extract date_time prefix
    existing_transcripts = set()
    for vtt in TRANSCRIPTS_DIR.glob('*.vtt'):
        # Extract YYYYMMDD_HHMMSS prefix
        parts = vtt.stem.split('_')
        if len(parts) >= 2 and len(parts[0]) == 8 and len(parts[1]) == 6:
            prefix = f"{parts[0]}_{parts[1]}"
            existing_transcripts.add(prefix)
    
    logger.info(f"Found {len(existing_transcripts)} existing transcripts")
    
    # Find recordings without transcripts
    missing = []
    
    # Check for MP4 files directly in recordings folder
    for item in sorted(RECORDINGS_DIR.iterdir()):
        # Handle both direct MP4 files and folders containing MP4s
        mp4_file = None
        name_for_parsing = None
        
        if item.is_file() and item.suffix.lower() == '.mp4':
            mp4_file = item
            name_for_parsing = item.stem
        elif item.is_dir():
            # Check for MP4 files in the folder
            mp4_files = list(item.glob('*.mp4'))
            if mp4_files:
                mp4_file = mp4_files[0]
                name_for_parsing = item.name
        
        if not mp4_file:
            continue
        
        # Extract date_time from name (YYYYMMDD_HHMMSS_...)
        parts = name_for_parsing.split('_')
        if len(parts) >= 2 and len(parts[0]) == 8 and len(parts[1]) == 6:
            prefix = f"{parts[0]}_{parts[1]}"
            
            # Check if transcript exists
            if prefix not in existing_transcripts:
                missing.append({
                    'folder': item if item.is_dir() else item.parent,
                    'mp4_file': mp4_file,
                    'prefix': prefix,
                    'size_mb': mp4_file.stat().st_size / 1024 / 1024,
                    'meeting_name': '_'.join(parts[2:]) if len(parts) > 2 else 'Unknown',
                    'full_name': name_for_parsing
                })
    
    return missing


def extract_audio(video_path):
    """Extract audio from MP4 using FFmpeg"""
    audio_path = video_path.with_suffix('.wav')
    
    logger.info(f"  Extracting audio...")
    
    cmd = [
        'ffmpeg', '-y', '-i', str(video_path),
        '-vn',                    # No video
        '-acodec', 'pcm_s16le',   # PCM format
        '-ar', '16000',           # 16kHz sample rate (optimal for speech)
        '-ac', '1',               # Mono
        str(audio_path)
    ]
    
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
    
    if result.returncode != 0:
        raise Exception(f"FFmpeg error: {result.stderr[:500]}")
    
    logger.info(f"  Audio extracted: {audio_path.stat().st_size / 1024 / 1024:.1f} MB")
    return audio_path


def transcribe_audio(audio_path, recording_info):
    """Transcribe audio using Azure Speech-to-Text with timestamps"""
    import azure.cognitiveservices.speech as speechsdk
    
    speech_config = speechsdk.SpeechConfig(subscription=SPEECH_KEY, region=SPEECH_REGION)
    speech_config.speech_recognition_language = "en-US"
    speech_config.request_word_level_timestamps()
    
    audio_config = speechsdk.AudioConfig(filename=str(audio_path))
    speech_recognizer = speechsdk.SpeechRecognizer(
        speech_config=speech_config,
        audio_config=audio_config
    )
    
    segments = []  # List of (start_time, end_time, text)
    done = False
    error_msg = None
    
    def recognized_cb(evt):
        if evt.result.reason == speechsdk.ResultReason.RecognizedSpeech:
            # Get timing info
            offset_ticks = evt.result.offset  # 100-nanosecond units
            duration_ticks = evt.result.duration
            
            start_ms = offset_ticks // 10000
            end_ms = (offset_ticks + duration_ticks) // 10000
            
            segments.append({
                'start_ms': start_ms,
                'end_ms': end_ms,
                'text': evt.result.text
            })
            
            # Show progress
            if len(segments) % 10 == 0:
                elapsed_min = start_ms / 60000
                print(f"  ... {len(segments)} segments, {elapsed_min:.1f} min processed", end='\r')
    
    def stop_cb(evt):
        nonlocal done
        done = True
    
    def canceled_cb(evt):
        nonlocal done, error_msg
        done = True
        if evt.cancellation_details.reason == speechsdk.CancellationReason.Error:
            error_msg = evt.cancellation_details.error_details
    
    speech_recognizer.recognized.connect(recognized_cb)
    speech_recognizer.session_stopped.connect(stop_cb)
    speech_recognizer.canceled.connect(canceled_cb)
    
    logger.info("  Transcribing... (this may take several minutes)")
    speech_recognizer.start_continuous_recognition()
    
    # Wait for completion (max 45 minutes for long recordings)
    start_time = time.time()
    while not done and (time.time() - start_time) < 2700:
        time.sleep(1)
    
    speech_recognizer.stop_continuous_recognition()
    print()  # New line after progress
    
    if error_msg:
        raise Exception(f"Transcription error: {error_msg}")
    
    logger.info(f"  Transcription complete! {len(segments)} segments")
    return segments


def format_timestamp(ms):
    """Convert milliseconds to VTT timestamp (HH:MM:SS.mmm)"""
    hours = ms // 3600000
    ms = ms % 3600000
    minutes = ms // 60000
    ms = ms % 60000
    seconds = ms // 1000
    millis = ms % 1000
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}.{millis:03d}"


def save_vtt_transcript(segments, recording_info):
    """Save transcript in VTT format with proper timestamps"""
    prefix = recording_info['prefix']
    meeting_name = recording_info.get('full_name', recording_info['meeting_name'])[:50]
    
    # Clean meeting name for filename
    safe_name = ''.join(c if c.isalnum() or c in ' -_' else '' for c in meeting_name)
    safe_name = safe_name.strip()[:60]
    
    vtt_filename = f"{safe_name}.vtt"
    vtt_path = TRANSCRIPTS_DIR / vtt_filename
    
    # Build VTT content
    lines = ["WEBVTT", ""]
    lines.append(f"NOTE Transcribed from local recording using Azure Speech-to-Text")
    lines.append(f"NOTE Source: {recording_info['folder'].name}")
    lines.append(f"NOTE Transcribed: {datetime.now().isoformat()}")
    lines.append("")
    
    for i, seg in enumerate(segments, 1):
        start = format_timestamp(seg['start_ms'])
        end = format_timestamp(seg['end_ms'])
        
        lines.append(str(i))
        lines.append(f"{start} --> {end}")
        lines.append(seg['text'])
        lines.append("")
    
    with open(vtt_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    
    logger.info(f"  ✓ Saved: {vtt_filename}")
    return vtt_path


def load_progress():
    """Load progress from file"""
    if PROGRESS_FILE.exists():
        with open(PROGRESS_FILE, 'r') as f:
            return json.load(f)
    return {'completed': [], 'failed': [], 'stats': {'total_completed': 0, 'total_failed': 0}}


def save_progress(progress):
    """Save progress to file"""
    with open(PROGRESS_FILE, 'w') as f:
        json.dump(progress, f, indent=2, default=str)


def process_recording(recording_info, progress):
    """Process a single recording"""
    full_name = recording_info.get('full_name', recording_info['mp4_file'].stem)
    mp4_path = recording_info['mp4_file']
    
    logger.info("="*70)
    logger.info(f"Processing: {full_name}")
    logger.info(f"MP4 File: {mp4_path.name}")
    logger.info(f"Size: {recording_info['size_mb']:.1f} MB")
    logger.info("="*70)
    
    audio_path = None
    
    try:
        # Step 1: Extract audio
        audio_path = extract_audio(mp4_path)
        
        # Step 2: Transcribe
        segments = transcribe_audio(audio_path, recording_info)
        
        if not segments:
            raise Exception("No speech detected in recording")
        
        # Step 3: Save VTT
        vtt_path = save_vtt_transcript(segments, recording_info)
        
        # Update progress
        progress['completed'].append(full_name)
        progress['stats']['total_completed'] += 1
        save_progress(progress)
        
        logger.info(f"✓ SUCCESS!")
        return True
        
    except Exception as e:
        logger.error(f"✗ FAILED: {e}")
        progress['failed'].append({'folder': full_name, 'error': str(e)})
        progress['stats']['total_failed'] += 1
        save_progress(progress)
        return False
        
    finally:
        # Cleanup audio file to save space
        if audio_path and audio_path.exists():
            try:
                audio_path.unlink()
                logger.info("  Cleaned up temporary audio file")
            except:
                pass


def main():
    parser = argparse.ArgumentParser(description='Transcribe local recordings without VTT files')
    parser.add_argument('--list', action='store_true', help='List recordings needing transcription')
    parser.add_argument('--count', type=int, default=0, help='Process only N recordings (0 = all)')
    parser.add_argument('--filter', type=str, help='Filter recordings by name pattern (e.g., "Integration Team")')
    parser.add_argument('--prefixes', type=str, help='Comma-separated date prefixes to process (e.g., "20251203_190046,20251211_134434")')
    args = parser.parse_args()
    
    # Check dependencies
    if not SPEECH_KEY:
        logger.error("AZURE_SPEECH_KEY not found in .env file!")
        logger.error("Please add your Azure Speech Service key to .env")
        sys.exit(1)
    
    # Check FFmpeg
    try:
        subprocess.run(['ffmpeg', '-version'], capture_output=True, check=True)
    except FileNotFoundError:
        logger.error("FFmpeg not found! Please install FFmpeg:")
        logger.error("  Windows: Download from https://ffmpeg.org/download.html")
        logger.error("  Or use: winget install ffmpeg")
        sys.exit(1)
    
    # Find recordings needing transcription
    missing = find_recordings_without_transcripts()
    
    if not missing:
        logger.info("All recordings already have transcripts!")
        return
    
    # Apply filters
    if args.filter:
        missing = [r for r in missing if args.filter.lower() in r.get('full_name', '').lower()]
        logger.info(f"Filtered to {len(missing)} recordings matching '{args.filter}'")
    
    if args.prefixes:
        prefix_list = [p.strip() for p in args.prefixes.split(',')]
        missing = [r for r in missing if r['prefix'] in prefix_list]
        logger.info(f"Filtered to {len(missing)} recordings matching specified prefixes")
    
    if not missing:
        logger.info("No recordings match the filter criteria!")
        return
    
    # Calculate totals
    total_size = sum(r['size_mb'] for r in missing)
    # Rough estimate: 1 hour of audio = 60MB MP4 = $0.98 (at $0.98/hour for real-time STT)
    estimated_hours = total_size / 60
    estimated_cost = estimated_hours * 0.98
    
    if args.list:
        print("\n" + "="*70)
        print(f"RECORDINGS NEEDING TRANSCRIPTION: {len(missing)}")
        print(f"Total Size: {total_size:.1f} MB")
        print(f"Estimated Audio: ~{estimated_hours:.1f} hours")
        print(f"Estimated Cost: ~${estimated_cost:.2f}")
        print("="*70)
        
        for i, rec in enumerate(missing, 1):
            print(f"\n{i}. {rec['folder'].name}")
            print(f"   Size: {rec['size_mb']:.1f} MB | MP4: {rec['mp4_file'].name}")
        
        print("\n" + "="*70)
        print("Run without --list to start transcription")
        return
    
    # Load progress
    progress = load_progress()
    
    # Skip already completed
    missing = [r for r in missing if r.get('full_name', r['mp4_file'].stem) not in progress['completed']]
    
    if not missing:
        logger.info("All recordings already processed!")
        return
    
    # Limit count if specified
    to_process = missing[:args.count] if args.count > 0 else missing
    
    logger.info(f"\nWill process {len(to_process)} recordings")
    logger.info(f"Total size: {sum(r['size_mb'] for r in to_process):.1f} MB")
    logger.info(f"Press Ctrl+C to stop at any time\n")
    
    time.sleep(3)  # Give user time to cancel
    
    # Process recordings
    success_count = 0
    fail_count = 0
    
    for i, rec in enumerate(to_process, 1):
        logger.info(f"\n[{i}/{len(to_process)}] Processing...")
        
        if process_recording(rec, progress):
            success_count += 1
        else:
            fail_count += 1
        
        # Brief pause between recordings
        if i < len(to_process):
            time.sleep(2)
    
    # Summary
    print("\n" + "="*70)
    print("TRANSCRIPTION COMPLETE")
    print("="*70)
    print(f"Successful: {success_count}")
    print(f"Failed: {fail_count}")
    print(f"Total completed: {progress['stats']['total_completed']}")
    print("="*70)
    
    if success_count > 0:
        print("\nRun 'python src/update_excel_with_transcripts.py' to add new transcripts to Excel")


if __name__ == '__main__':
    main()
