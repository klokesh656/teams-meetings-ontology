# Azure Speech Service Setup Guide

## Overview
This guide will help you set up Azure Speech Service to transcribe recordings that don't have existing transcripts.

## Prerequisites
- Azure subscription with access to create Speech Services
- FFmpeg installed on your system (for audio extraction)

## Step 1: Create Azure Speech Service

### Option A: Azure Portal (Recommended)
1. Go to [Azure Portal](https://portal.azure.com)
2. Click **"Create a resource"**
3. Search for **"Speech Services"**
4. Click **"Create"**
5. Fill in the details:
   - **Subscription**: Select your subscription
   - **Resource Group**: Use existing or create new (e.g., `rg-transcripts`)
   - **Region**: Choose **East US** (or your preferred region)
   - **Name**: Enter a unique name (e.g., `speech-transcripts-service`)
   - **Pricing tier**: Select **Free F0** (500 minutes/month free) or **Standard S0**
6. Click **"Review + create"** then **"Create"**

### Option B: Azure CLI
```bash
# Login to Azure
az login

# Create resource group (if needed)
az group create --name rg-transcripts --location eastus

# Create Speech Service
az cognitiveservices account create \
    --name speech-transcripts-service \
    --resource-group rg-transcripts \
    --kind SpeechServices \
    --sku F0 \
    --location eastus
```

## Step 2: Get API Key and Region

### From Azure Portal:
1. Navigate to your Speech Service resource
2. Click **"Keys and Endpoint"** in the left menu
3. Copy **Key 1** or **Key 2**
4. Note the **Region** (e.g., `eastus`)

### From Azure CLI:
```bash
# Get keys
az cognitiveservices account keys list \
    --name speech-transcripts-service \
    --resource-group rg-transcripts

# Get endpoint
az cognitiveservices account show \
    --name speech-transcripts-service \
    --resource-group rg-transcripts \
    --query "properties.endpoint"
```

## Step 3: Update .env File

Add these lines to your `.env` file:
```env
AZURE_SPEECH_KEY="your_key_here"
AZURE_SPEECH_REGION="eastus"
```

## Step 4: Install FFmpeg

### Windows:
1. Download FFmpeg from: https://ffmpeg.org/download.html
2. Extract to `C:\ffmpeg`
3. Add to PATH:
   - Open **System Properties** > **Environment Variables**
   - Edit **Path** variable
   - Add `C:\ffmpeg\bin`
4. Verify installation:
   ```powershell
   ffmpeg -version
   ```

### Using Chocolatey (Windows):
```powershell
choco install ffmpeg
```

### Using Scoop (Windows):
```powershell
scoop install ffmpeg
```

### macOS:
```bash
brew install ffmpeg
```

### Linux:
```bash
sudo apt-get install ffmpeg  # Ubuntu/Debian
sudo yum install ffmpeg      # CentOS/RHEL
```

## Step 5: Install Python Dependencies

```bash
pip install azure-cognitiveservices-speech
```

Or install all requirements:
```bash
pip install -r requirements.txt
```

## Step 6: Run the Transcription Script

```bash
python src/transcribe_recordings.py
```

## What the Script Does

1. **Finds Recordings Without Transcripts**: Identifies all recordings that don't have matching VTT files
2. **Downloads Recordings**: Downloads MP4 recordings from Microsoft Graph API
3. **Extracts Audio**: Converts MP4 to WAV format using FFmpeg
4. **Transcribes**: Uses Azure Speech-to-Text to transcribe the audio
5. **Saves as VTT**: Creates VTT transcript files in the `transcripts/` folder
6. **Cleanup**: Removes temporary audio files

## Cost Estimation

### Free Tier (F0)
- **500 minutes** of audio transcription per month
- Suitable for testing and small-scale use

### Standard Tier (S0)
- **Pay-as-you-go**: $1.00 per audio hour
- Example: 50 hours of recordings = $50

### Example Calculation
- Average meeting: 30 minutes
- 100 meetings = 50 hours
- Cost: $50 with Standard tier
- **Tip**: Start with Free tier to test!

## Troubleshooting

### Error: "FFmpeg not found"
- Ensure FFmpeg is installed and in PATH
- Restart your terminal/IDE after installation
- Test with: `ffmpeg -version`

### Error: "Azure Speech API key not configured"
- Check your `.env` file has `AZURE_SPEECH_KEY` set
- Verify the key is correct (copy from Azure Portal)

### Error: "Transcription timeout"
- Increase timeout in script (default: 10 minutes)
- Check your network connection
- Verify Speech Service region matches your configuration

### Poor Transcription Quality
- Use 16kHz audio sample rate (already configured)
- Ensure audio quality is good
- Consider using multiple speakers feature for better accuracy

## Advanced Configuration

### Multi-Speaker Recognition
Add to `transcribe_audio_with_azure_speech()`:
```python
speech_config.set_property(
    speechsdk.PropertyId.SpeechServiceConnection_TranscriptionBackend,
    "diarizationConversational"
)
```

### Custom Language
Change in script:
```python
speech_config.speech_recognition_language = "en-US"  # or "es-ES", "fr-FR", etc.
```

### Batch Transcription (for large files)
For recordings > 10 minutes, consider using Azure Batch Transcription API for better performance.

## Next Steps

After transcription:
1. Run `python src/analyze_transcripts.py` to analyze the new transcripts with AI
2. Run `python src/upload_and_create_master.py` to update the Excel report
3. Set up daily sync to automate the process

## Support

- Azure Speech Documentation: https://docs.microsoft.com/azure/cognitive-services/speech-service/
- FFmpeg Documentation: https://ffmpeg.org/documentation.html
