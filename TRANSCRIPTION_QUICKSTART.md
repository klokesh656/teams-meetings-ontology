# Recording Transcription - Quick Start Guide

## What This Does
Transcribes Teams meeting recordings that **don't have existing transcripts** using Azure Speech-to-Text.

## 🎯 Problem Solved
You have 49 recordings but many don't have transcripts available (404 errors). This script will:
- Find recordings without transcripts
- Download the MP4 files
- Extract audio
- Transcribe using AI (Azure Speech)
- Create VTT transcript files
- Make them available for analysis

## 📋 Prerequisites

### 1. Install FFmpeg (Required for audio extraction)

**Option A: Using Chocolatey (Recommended)**
```powershell
# Install Chocolatey if you don't have it
Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))

# Install FFmpeg
choco install ffmpeg -y

# Restart PowerShell and verify
ffmpeg -version
```

**Option B: Manual Installation**
1. Download: https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip
2. Extract to `C:\ffmpeg`
3. Add to PATH:
   - Press `Win + X` → System
   - Click "Advanced system settings"
   - Click "Environment Variables"
   - Under "System Variables", find "Path", click "Edit"
   - Click "New" and add: `C:\ffmpeg\bin`
   - Click OK, OK, OK
4. **Restart PowerShell** and verify:
   ```powershell
   ffmpeg -version
   ```

### 2. Create Azure Speech Service (Free Tier Available!)

**Quick Setup via Azure Portal:**
1. Go to: https://portal.azure.com/#create/Microsoft.CognitiveServicesSpeechServices
2. Fill in:
   - **Subscription**: Your Azure subscription
   - **Resource Group**: `rg-transcripts` (or create new)
   - **Region**: `East US`
   - **Name**: `speech-transcripts-service`
   - **Pricing tier**: **Free F0** (500 minutes/month FREE!)
3. Click "Review + create" → "Create"
4. After creation:
   - Go to resource
   - Click "Keys and Endpoint"
   - Copy **Key 1**
   - Note the **Region** (e.g., `eastus`)

### 3. Update .env File

Open your `.env` file and update these lines:
```env
AZURE_SPEECH_KEY="paste_your_key_here"
AZURE_SPEECH_REGION="eastus"
```

### 4. Install Python Dependencies

```powershell
pip install azure-cognitiveservices-speech
```

## 🚀 Running the Script

```powershell
python src/transcribe_recordings.py
```

## 📊 What Happens

```
Step 1: Authenticate with Microsoft Graph API ✓
Step 2: Fetch all recordings (49 found) ✓
Step 3: Check which recordings have NO transcripts (~40 without) ✓
Step 4: For each recording without transcript:
  - Download MP4 file
  - Extract audio (WAV format)
  - Transcribe using Azure Speech AI
  - Save as VTT file in transcripts/ folder
  - Clean up temporary files
Step 5: Generate summary report
```

## 📁 Output Structure

```
transcripts/
├── 20250818_170052_transcribed.vtt  ← NEW from recording
├── 20250908_175043_transcribed.vtt  ← NEW from recording
└── ...

recordings/
├── 20250818_170052_recording.mp4    ← Downloaded MP4
└── ...

logs/
└── transcribe_recordings.log        ← Detailed log
```

## 💰 Cost Estimation

### Free Tier (F0) - Recommended to Start
- **500 minutes FREE** per month
- ~16 meetings of 30 minutes each
- Perfect for testing!

### Standard Tier (S0) - If you need more
- **$1.00 per audio hour**
- Example: 40 recordings × 30 min = 20 hours = **$20**

**Pro Tip**: Start with Free tier, upgrade only if needed!

## 🔧 Troubleshooting

### "FFmpeg not found"
→ Install FFmpeg (see Prerequisites #1)
→ **Restart PowerShell** after installation
→ Verify: `ffmpeg -version`

### "Azure Speech API key not configured"
→ Check `.env` file has `AZURE_SPEECH_KEY` set
→ Verify key is correct (copy from Azure Portal)

### "Error downloading recording: 404"
→ Recording may have expired or been deleted
→ Script will skip and continue with next recording

### Slow transcription
→ Normal! Speech-to-Text takes ~1-2 minutes per 30-min recording
→ Script processes in batches (5 at a time by default)

## 🎯 After Transcription

Once you have new transcripts, you can:

1. **Analyze with AI**:
   ```powershell
   python src/analyze_transcripts.py
   ```

2. **Update Excel Report**:
   ```powershell
   python src/upload_and_create_master.py
   ```

3. **Set up Daily Automation** (future):
   - Daily sync will check for new recordings
   - Auto-transcribe if no transcript exists
   - Run AI analysis
   - Update Excel

## 📝 Notes

- **Processing Limit**: Script processes **first 5 recordings** by default (for testing)
- **Change Limit**: Edit line in `transcribe_recordings.py`:
  ```python
  for i, rec_info in enumerate(self.recordings_without_transcripts[:5], 1):  # Change :5 to :10 for 10
  ```
- **Recordings Folder**: Downloaded MP4s are saved to `recordings/` folder
- **Cleanup**: Audio files are auto-deleted after transcription
- **Logs**: Check `logs/transcribe_recordings.log` for details

## 🆘 Need Help?

Detailed setup guide: See `SPEECH_SERVICE_SETUP.md`

## Quick Command Reference

```powershell
# Install dependencies
pip install -r requirements.txt

# Check FFmpeg
ffmpeg -version

# Run transcription
python src/transcribe_recordings.py

# Check results
dir transcripts\*transcribed.vtt

# View log
type logs\transcribe_recordings.log
```

---

**Ready to transcribe? Follow Prerequisites 1-4, then run the script! 🎤**
