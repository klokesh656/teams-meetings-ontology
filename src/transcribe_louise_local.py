"""
Transcribe local Louise recordings that don't have transcripts yet.
Simple script focused only on transcription.
"""

import os
import sys
import json
import time
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv
import azure.cognitiveservices.speech as speechsdk
import subprocess

# Load environment variables
load_dotenv()

SPEECH_KEY = os.getenv("AZURE_SPEECH_KEY")
SPEECH_REGION = os.getenv("AZURE_SPEECH_REGION", "eastus")

RECORDINGS_DIR = Path("recordings")
TRANSCRIPTS_DIR = Path("transcripts")
PROGRESS_FILE = Path("transcription_progress_louise.json")


def load_progress():
    """Load transcription progress."""
    if PROGRESS_FILE.exists():
        with open(PROGRESS_FILE, 'r') as f:
            return json.load(f)
    return {"transcribed": [], "failed": []}


def save_progress(progress):
    """Save progress."""
    with open(PROGRESS_FILE, 'w') as f:
        json.dump(progress, f, indent=2)


def extract_audio(video_path: Path, audio_path: Path) -> bool:
    """Extract audio from video using FFmpeg."""
    try:
        cmd = [
            "ffmpeg", "-i", str(video_path),
            "-vn", "-acodec", "pcm_s16le",
            "-ar", "16000", "-ac", "1",
            "-y", str(audio_path)
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
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
        last_activity = time.time()
        
        def on_recognized(evt):
            nonlocal last_activity
            if evt.result.reason == speechsdk.ResultReason.RecognizedSpeech:
                all_results.append({
                    "text": evt.result.text,
                    "offset": evt.result.offset,
                    "duration": evt.result.duration
                })
                last_activity = time.time()
                print(f"\r    Recognized: {len(all_results)} segments", end="")
        
        def on_canceled(evt):
            nonlocal done
            if evt.cancellation_details.reason == speechsdk.CancellationReason.EndOfStream:
                pass  # Normal end
            else:
                print(f"\n    Canceled: {evt.cancellation_details.reason}")
            done = True
        
        def on_stopped(evt):
            nonlocal done
            done = True
        
        recognizer.recognized.connect(on_recognized)
        recognizer.canceled.connect(on_canceled)
        recognizer.session_stopped.connect(on_stopped)
        
        recognizer.start_continuous_recognition()
        
        # Wait with timeout
        timeout = 3600  # 1 hour max per file
        inactivity_timeout = 300  # 5 minutes of no new results
        start_time = time.time()
        
        while not done:
            time.sleep(1)
            elapsed = time.time() - start_time
            inactive = time.time() - last_activity
            
            if elapsed > timeout:
                print(f"\n    ⚠️ Timeout after {timeout}s")
                break
            if inactive > inactivity_timeout and len(all_results) > 0:
                print(f"\n    ⚠️ No new results for {inactivity_timeout}s, finishing")
                break
        
        recognizer.stop_continuous_recognition()
        print()  # New line after progress
        
        # Generate VTT
        if all_results:
            vtt_content = generate_vtt(all_results)
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(vtt_content)
            return True
        return False
        
    except Exception as e:
        print(f"\n    ❌ Transcription error: {e}")
        return False


def generate_vtt(results):
    """Generate VTT content from transcription results."""
    vtt_lines = ["WEBVTT", ""]
    
    for i, result in enumerate(results, 1):
        start_time = result["offset"] / 10_000_000
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


def safe_delete_file(file_path: Path, max_retries: int = 10):
    """Safely delete a file with retry logic."""
    for attempt in range(max_retries):
        try:
            if file_path.exists():
                file_path.unlink()
            return True
        except PermissionError:
            print(f"    ⚠️ File locked, waiting... (attempt {attempt + 1})")
            time.sleep(2)
    print(f"    ⚠️ Could not delete {file_path.name}, leaving it")
    return False


def main():
    print("=" * 70)
    print("TRANSCRIBE LOCAL LOUISE RECORDINGS")
    print("=" * 70)
    
    # Load progress
    progress = load_progress()
    
    # Find all Louise MP4s without transcripts
    to_transcribe = []
    for mp4_file in RECORDINGS_DIR.glob("*.mp4"):
        if "louise" not in mp4_file.name.lower():
            continue
        
        # Check if transcript exists
        vtt_path = TRANSCRIPTS_DIR / (mp4_file.stem + ".vtt")
        if vtt_path.exists():
            continue
        
        # Check if already processed
        if mp4_file.stem in progress.get("transcribed", []):
            continue
        
        # Check file size
        if mp4_file.stat().st_size < 100000:  # < 100KB
            continue
        
        to_transcribe.append(mp4_file)
    
    print(f"\nFound {len(to_transcribe)} recordings to transcribe")
    print(f"Previously transcribed: {len(progress.get('transcribed', []))}")
    print(f"Previously failed: {len(progress.get('failed', []))}")
    
    if not to_transcribe:
        print("\n✅ All recordings already transcribed!")
        return
    
    # Estimate time/cost
    total_size_mb = sum(f.stat().st_size / 1024 / 1024 for f in to_transcribe)
    est_hours = total_size_mb / 60  # Rough estimate: 60MB ≈ 1 hour
    est_cost = est_hours * 1.0  # $1/hour
    
    print(f"\nTotal size: {total_size_mb:.1f} MB")
    print(f"Estimated audio: ~{est_hours:.1f} hours")
    print(f"Estimated cost: ~${est_cost:.2f}")
    
    # Transcribe
    transcribed_count = 0
    for i, mp4_file in enumerate(to_transcribe, 1):
        print(f"\n[{i}/{len(to_transcribe)}] {mp4_file.name[:65]}...")
        print(f"    Size: {mp4_file.stat().st_size / 1024 / 1024:.1f} MB")
        
        # Extract audio
        audio_path = mp4_file.with_suffix(".wav")
        print("    Extracting audio...")
        
        if not extract_audio(mp4_file, audio_path):
            print("    ❌ Audio extraction failed")
            progress["failed"].append({"file": mp4_file.name, "reason": "Audio extraction failed"})
            save_progress(progress)
            continue
        
        # Transcribe
        vtt_path = TRANSCRIPTS_DIR / (mp4_file.stem + ".vtt")
        print("    Transcribing with Azure Speech...")
        
        if transcribe_audio(audio_path, vtt_path):
            print(f"    ✅ Transcript saved: {vtt_path.name}")
            progress["transcribed"].append(mp4_file.stem)
            transcribed_count += 1
        else:
            print("    ❌ Transcription failed (no speech detected?)")
            progress["failed"].append({"file": mp4_file.name, "reason": "Transcription failed"})
        
        # Clean up audio file
        safe_delete_file(audio_path)
        
        save_progress(progress)
    
    # Summary
    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    print(f"  Transcribed this run: {transcribed_count}")
    print(f"  Total transcribed: {len(progress.get('transcribed', []))}")
    print(f"  Total failed: {len(progress.get('failed', []))}")
    print(f"\nProgress saved to: {PROGRESS_FILE}")


if __name__ == "__main__":
    main()
