# Complete Solution: Recording Transcription & Daily Sync

## 🎯 Problem Statement
Your client has:
- **59 existing transcripts** (already analyzed)
- **49 recordings** with many missing transcripts (404 errors)
- Need to **transcribe recordings** that don't have transcripts
- Need **daily automation** to check for new recordings and analyze them

## 📦 Solution Overview

### Created Files

#### 1. **`src/transcribe_recordings.py`** - NEW! 🆕
Transcribes recordings that don't have existing transcripts using Azure Speech-to-Text.

**What it does:**
- Finds recordings without matching VTT files
- Downloads MP4 recordings from Graph API
- Extracts audio using FFmpeg
- Transcribes with Azure Speech-to-Text
- Saves as VTT files in `transcripts/` folder
- Processes first 5 recordings (configurable)

**Requirements:**
- Azure Speech Service (Free tier: 500 min/month)
- FFmpeg installed
- Python package: `azure-cognitiveservices-speech`

#### 2. **`src/daily_sync.py`** - FIXED! ✅
Daily automation script with improved error handling.

**What it does:**
- Downloads new transcripts from Graph API
- Handles 404 errors gracefully (skips expired transcripts)
- Uploads to Azure Blob Storage
- Runs AI analysis on new transcripts
- Updates Excel report

**Fixed issues:**
- Added timeout handling (30 sec per request)
- Graceful 404 error handling
- ASCII-safe logging (no emoji encoding errors)
- Continues processing after errors

#### 3. **`TRANSCRIPTION_QUICKSTART.md`** - Setup Guide
Step-by-step guide to set up and run transcription.

**Includes:**
- FFmpeg installation (3 methods)
- Azure Speech Service setup
- Cost estimation (Free vs Paid tier)
- Troubleshooting tips

#### 4. **`SPEECH_SERVICE_SETUP.md`** - Detailed Documentation
Complete technical documentation for Azure Speech Service.

**Includes:**
- Azure Portal setup steps
- Azure CLI commands
- Advanced configuration options
- Multi-speaker recognition

## 🚀 Getting Started

### Step 1: Install FFmpeg (Required for audio extraction)

**Quick Install with Chocolatey:**
```powershell
# Install Chocolatey (if not installed)
Set-ExecutionPolicy Bypass -Scope Process -Force
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))

# Install FFmpeg
choco install ffmpeg -y

# Restart PowerShell and verify
ffmpeg -version
```

**Alternative: Manual Install**
1. Download: https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip
2. Extract to `C:\ffmpeg`
3. Add `C:\ffmpeg\bin` to PATH
4. Restart PowerShell

### Step 2: Create Azure Speech Service

**Option A: Azure Portal (5 minutes)**
1. Go to: https://portal.azure.com/#create/Microsoft.CognitiveServicesSpeechServices
2. Fill in:
   - Resource Group: `rg-transcripts`
   - Region: `East US`
   - Name: `speech-transcripts-service`
   - Pricing: **Free F0** (500 min/month FREE!)
3. After creation → Keys and Endpoint → Copy Key 1

**Option B: Azure CLI**
```bash
az cognitiveservices account create \
  --name speech-transcripts-service \
  --resource-group rg-transcripts \
  --kind SpeechServices \
  --sku F0 \
  --location eastus
```

### Step 3: Update .env File

Add these lines to your `.env`:
```env
AZURE_SPEECH_KEY="your_key_from_step_2"
AZURE_SPEECH_REGION="eastus"
```

### Step 4: Install Python Dependencies

```powershell
pip install azure-cognitiveservices-speech
```

### Step 5: Run Transcription Script

```powershell
python src/transcribe_recordings.py
```

**Expected Output:**
```
Authenticating with Microsoft Graph API...
Authentication successful!
============================================================
FINDING RECORDINGS WITHOUT TRANSCRIPTS
============================================================
Found 49 total recordings
Found 35 recordings WITHOUT transcripts
============================================================
PROCESSING RECORDINGS WITHOUT TRANSCRIPTS
============================================================
[1/5] Processing recording from 2025-08-18T17:00:52Z
  Downloaded: 20250818_170052_recording.mp4 (45.2 MB)
  Extracted audio: 20250818_170052_recording.wav
  Transcribing audio... (this may take a few minutes)
  Transcription complete! (12456 characters)
  Saved transcript: 20250818_170052_transcribed.vtt
[2/5] Processing recording from...
```

## 📊 Workflow Overview

```
┌─────────────────────────────────────────────────────────────┐
│                    COMPLETE WORKFLOW                        │
└─────────────────────────────────────────────────────────────┘

1. TRANSCRIBE RECORDINGS (NEW!)
   ├─ Find recordings without transcripts
   ├─ Download MP4 files
   ├─ Extract audio with FFmpeg
   ├─ Transcribe with Azure Speech AI
   └─ Save as VTT files
   
2. DAILY SYNC (FIXED!)
   ├─ Check for new transcripts
   ├─ Download (skip 404 errors)
   ├─ Upload to Azure Blob Storage
   └─ Update Excel report
   
3. AI ANALYSIS (EXISTING)
   ├─ Analyze transcripts with Azure OpenAI
   ├─ Extract sentiment, churn risk, opportunities
   └─ Update Excel with scores

4. AUTOMATION (FUTURE)
   ├─ Windows Task Scheduler
   └─ Runs daily at 8 AM
```

## 💰 Cost Breakdown

### Azure Speech Service
- **Free Tier (F0)**: 500 minutes/month = **$0**
  - Good for: ~16 meetings @ 30 min each
  - Recommended: Start here!
  
- **Standard (S0)**: $1.00/hour = **Pay as you go**
  - Example: 40 recordings × 30 min = 20 hours = **$20**

### Azure OpenAI (Existing)
- Already configured, no additional cost for this feature

### Azure Blob Storage (Existing)
- Already configured, no additional cost

**Total New Cost**: $0 (with Free tier) or ~$20-50/month (Standard tier)

## 🛠️ Troubleshooting

### Error: "FFmpeg not found"
**Solution:**
- Install FFmpeg (see Step 1)
- **Restart PowerShell** after installation
- Verify: `ffmpeg -version`

### Error: "Azure Speech API key not configured"
**Solution:**
- Check `.env` has `AZURE_SPEECH_KEY="your_key"`
- Verify key from Azure Portal → Speech Service → Keys and Endpoint

### Error: "Failed to download recording: 404"
**Solution:**
- Recording expired or deleted (normal behavior)
- Script will skip and continue
- Check `logs/transcribe_recordings.log` for details

### Slow transcription
**Expected:**
- ~1-2 minutes per 30-minute recording
- This is normal for speech-to-text processing
- Script processes 5 recordings at a time

### Unicode encoding errors (FIXED!)
**Solution:**
- Updated `daily_sync.py` with ASCII-safe logging
- No more emoji encoding errors on Windows

## 📁 File Structure After Setup

```
Upwork/Issac - Copilot agent/
├── src/
│   ├── transcribe_recordings.py       ← NEW: Transcribe recordings
│   ├── daily_sync.py                  ← FIXED: Error handling
│   ├── analyze_transcripts.py         ← EXISTING: AI analysis
│   └── ...
│
├── transcripts/                       ← VTT files
│   ├── 20250818_170052_Meeting.vtt   ← Existing
│   ├── 20250818_170052_transcribed.vtt ← NEW from recording!
│   └── ...
│
├── recordings/                        ← NEW: Downloaded MP4s
│   ├── 20250818_170052_recording.mp4
│   └── ...
│
├── logs/
│   ├── daily_sync.log
│   └── transcribe_recordings.log      ← NEW: Transcription log
│
├── output/
│   └── meeting_transcripts_latest_analyzed.xlsx
│
├── TRANSCRIPTION_QUICKSTART.md        ← Quick start guide
├── SPEECH_SERVICE_SETUP.md            ← Detailed setup
├── .env                               ← Updated with Speech key
└── requirements.txt                   ← Updated with Speech SDK
```

## 🎯 Next Steps

### Immediate Tasks:
1. **Install FFmpeg** (Step 1 above)
2. **Create Azure Speech Service** (Step 2 above)
3. **Update .env** with Speech key (Step 3 above)
4. **Run transcription script** (Step 5 above)

### After Transcription:
1. **Run AI Analysis** on new transcripts:
   ```powershell
   python src/analyze_transcripts.py
   ```

2. **Update Excel** with new data:
   ```powershell
   python src/upload_and_create_master.py
   ```

3. **Test Daily Sync** with fixed error handling:
   ```powershell
   python src/daily_sync.py
   ```

### Future Automation:
1. Set up **Windows Task Scheduler** (use `setup_scheduler.ps1`)
2. Schedule daily runs at 8 AM
3. Automatically transcribe + analyze + update Excel

## 📝 Key Changes Summary

### What Was Fixed:
1. ✅ **404 Error Handling**: Script now skips expired transcripts
2. ✅ **Timeout Handling**: 30-second timeout per request
3. ✅ **Logging Improvements**: ASCII-safe, no emoji errors
4. ✅ **Error Recovery**: Continues processing after failures

### What Was Added:
1. 🆕 **Recording Transcription**: Azure Speech-to-Text integration
2. 🆕 **Audio Extraction**: FFmpeg integration
3. 🆕 **Setup Guides**: Complete documentation
4. 🆕 **Cost Estimation**: Free tier recommendation

### What Remains:
- Run transcription on ~35 recordings without transcripts
- Test daily sync with fixed error handling
- Set up Windows Task Scheduler (optional)

## 🆘 Support Resources

- **FFmpeg**: https://ffmpeg.org/download.html
- **Azure Speech**: https://portal.azure.com/#create/Microsoft.CognitiveServicesSpeechServices
- **Documentation**: See `SPEECH_SERVICE_SETUP.md`
- **Quick Start**: See `TRANSCRIPTION_QUICKSTART.md`

---

## 🚀 Ready to Start?

Follow these 5 commands:
```powershell
# 1. Install FFmpeg
choco install ffmpeg -y

# 2. Restart PowerShell, then install dependencies
pip install azure-cognitiveservices-speech

# 3. Configure .env (add AZURE_SPEECH_KEY)
notepad .env

# 4. Run transcription
python src/transcribe_recordings.py

# 5. Check results
dir transcripts\*transcribed.vtt
```

**All set! Your client's recordings will be transcribed and ready for AI analysis! 🎉**
