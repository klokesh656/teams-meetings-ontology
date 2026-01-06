# Microsoft Teams Transcript Extractor & AI Analysis

A comprehensive Python solution for extracting, transcribing, and analyzing Microsoft Teams meeting transcripts using Microsoft Graph API, Azure Speech Service, and Azure OpenAI.

## 🎯 Key Features

### Core Functionality
- **Extract Transcripts** from Microsoft Graph API (VTT format)
- **Transcribe Recordings** that don't have transcripts (Azure Speech-to-Text)
- **AI Analysis** with Azure OpenAI (sentiment, churn risk, opportunities)
- **Daily Automation** to sync new meetings and update reports
- **Azure Blob Storage** integration for scalable storage
- **Excel Reports** with comprehensive metadata and AI scores

### New in v2.0 🆕
- ✅ **Recording Transcription**: Transcribe MP4 recordings using Azure Speech Service
- ✅ **Improved Error Handling**: Graceful 404 handling, timeout protection
- ✅ **Daily Sync Script**: Automated workflow for continuous monitoring
- ✅ **Comprehensive Documentation**: Quick start guides and troubleshooting

## 🚀 Quick Start

### 1. Extract Existing Transcripts
```powershell
python src/download_transcripts.py
```

### 2. Transcribe Recordings Without Transcripts (NEW!)
```powershell
python src/transcribe_recordings.py
```

### 3. Analyze with AI
```powershell
python src/analyze_transcripts.py
```

### 4. Daily Automation
```powershell
python src/daily_sync.py
```

## 📋 Prerequisites

### Required Azure Services
1. **Azure AD Application** with permissions:
   - `OnlineMeetings.Read.All`
   - `OnlineMeetingTranscript.Read.All`
   - `OnlineMeetingRecording.Read.All`
   - `User.Read.All`

2. **Azure Blob Storage** (for transcript storage)

3. **Azure OpenAI** (for AI analysis)
   - Deployment: GPT-4 or similar
   - Analysis prompts configured

4. **Azure Speech Service** (NEW! for recording transcription)
   - Free tier: 500 minutes/month
   - Standard tier: $1.00/hour

### System Requirements
- **Python 3.8+**
- **FFmpeg** (for audio extraction from recordings)
- **Windows Long Path Support** (Windows only)
     ```powershell
     New-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\FileSystem" -Name "LongPathsEnabled" -Value 1 -PropertyType DWORD -Force
     ```

## Installation

1. **Clone or download this repository**

2. **Install dependencies**:
   ```powershell
   pip install msgraph-sdk azure-identity msal python-dotenv requests pandas openpyxl azure-storage-blob
   ```

3. **Configure environment variables**:
   - Copy `.env.example` to `.env`
   - Fill in your Azure AD application details:
     ```
     AZURE_CLIENT_ID=your_client_id_here
     AZURE_CLIENT_SECRET=your_client_secret_here
     AZURE_TENANT_ID=your_tenant_id_here
     TEAMS_USER_ID=user@yourdomain.com
     OUTPUT_DIR=transcripts
     
     # Optional: For Azure Blob Storage
     AZURE_STORAGE_CONNECTION_STRING=your_connection_string_here
     ```

## CLI Commands

### List Users
```powershell
python src/transcript_extractor.py --list-users
```
Lists users in your organization to identify who has meetings.

### Scan Recordings
```powershell
python src/transcript_extractor.py --scan-recordings [--max N] [--download]
```
Scans users' OneDrive Recordings folders for meeting files.
- `--max N`: Maximum users to scan (default: 30)
- `--download`: Download transcript files locally

### Export Metadata to Excel
```powershell
python src/transcript_extractor.py --export-metadata [--include-recordings] [--max N] [--output FILE]
```
Creates an Excel file with meeting metadata for Power Automate integration.
- `--include-recordings`: Include MP4 recording files (not just transcripts)
- `--max N`: Maximum users to scan
- `--output FILE`: Output filename (default: meeting_metadata.xlsx)
- `--parse`: Parse VTT content for participants and duration

### Upload to Blob Storage
```powershell
python src/transcript_extractor.py --upload-blobs [--container NAME] [--max N]
```
Uploads transcripts to Azure Blob Storage and updates Excel with blob URLs.
- Requires `AZURE_STORAGE_CONNECTION_STRING` in `.env`
- `--container NAME`: Blob container name (default: transcripts)

### Search OneDrive
```powershell
python src/transcript_extractor.py --search-drives [emails] [--download]
```
Search specific users' OneDrive for transcript files.

## Excel Output Schema

The `--export-metadata` command creates an Excel file with the following columns:

| Column | Description |
|--------|-------------|
| meeting_id | Unique identifier |
| meeting_date | Date of the meeting |
| meeting_time | Time of the meeting |
| meeting_subject | Extracted from filename |
| organizer_email | User who organized the meeting |
| organizer_name | Display name of organizer |
| duration_seconds | Duration (from VTT parsing) |
| participant_count | Number of participants |
| participants | Comma-separated list |
| file_name | Original filename |
| sharepoint_url | Link to file in SharePoint |
| transcript_blob_url | Azure Blob Storage URL |
| **AI Analysis Columns** | |
| sentiment_score | (0-100, filled by Power Automate) |
| churn_risk_score | (0-100, filled by Power Automate) |
| upsell_potential_score | (0-100, filled by Power Automate) |
| execution_reliability_score | (0-100, filled by Power Automate) |
| operational_complexity_score | (0-100, filled by Power Automate) |
| events_detected | Event types from AI analysis |
| key_topics | Main topics discussed |
| action_items | Follow-up actions |
| ai_summary | AI-generated summary |

## Power Automate Integration

See `ARCHITECTURE.md` for the complete pipeline design including:
- Excel trigger configuration
- Azure OpenAI integration
- AI prompt template for transcript analysis
- Event type detection system

## Output Structure

Downloaded transcripts are organized as follows:

```
transcripts/
├── user_at_domain_com/
│   ├── Meeting Subject-20231201_140000-Meeting Recording.vtt
│   └── Meeting Subject-20231201_140000-Meeting Recording.mp4
└── 2024-12-06_abc12345/
    ├── transcript_xyz78901.vtt
    └── metadata_xyz78901.json
```

### No Transcripts Found
- Verify that transcription was enabled for the meetings
- Check that the user ID (email) is correct
- Ensure the meetings have ended (transcripts may not be immediately available)

### Long Path Errors (Windows)
- Enable Windows Long Path support as described in Prerequisites
- Restart your terminal/IDE after enabling
- Try installing packages one at a time if issues persist

## API Limitations

- Microsoft Graph API has rate limits (throttling)
- Transcript availability depends on meeting settings
- Transcripts are in VTT (WebVTT) format

## Security Notes

- Never commit your `.env` file to version control
- Keep your client secret secure
- Use Azure Key Vault for production deployments
- Regularly rotate client secrets

## License

MIT License - feel free to use and modify as needed.

## Support

For issues with:
- Microsoft Graph API: Check [Microsoft Graph documentation](https://learn.microsoft.com/en-us/graph/api/resources/calltranscript)
- Azure AD setup: See [Azure AD app registration guide](https://learn.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
