# Meeting Transcript Analysis Architecture

## Overview

This architecture extracts Teams meeting transcripts, stores them in Azure Blob Storage, creates a metadata Excel file, and uses Power Automate to analyze transcripts with AI.

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                         MEETING TRANSCRIPT PIPELINE                          │
└─────────────────────────────────────────────────────────────────────────────┘

┌──────────────┐    ┌──────────────┐    ┌──────────────┐    ┌──────────────┐
│   Microsoft  │    │   Python     │    │    Azure     │    │    Excel     │
│   Graph API  │───▶│   Script     │───▶│    Blob      │───▶│   Metadata   │
│  (OneDrive)  │    │  (Extract)   │    │   Storage    │    │    File      │
└──────────────┘    └──────────────┘    └──────────────┘    └──────────────┘
                                               │                    │
                                               │                    │
                                               ▼                    ▼
                                        ┌──────────────┐    ┌──────────────┐
                                        │    Power     │◀───│   Trigger    │
                                        │   Automate   │    │  (Schedule/  │
                                        │    Flow      │    │   Manual)    │
                                        └──────────────┘    └──────────────┘
                                               │
                                               ▼
                                        ┌──────────────┐
                                        │   Azure      │
                                        │   OpenAI /   │
                                        │   GPT-4      │
                                        └──────────────┘
                                               │
                                               ▼
                                        ┌──────────────┐
                                        │   Results    │
                                        │  (Scores +   │
                                        │   Events)    │
                                        └──────────────┘
```

## Components

### 1. Python Script (transcript_extractor.py)
- **Input**: Microsoft Graph API (OneDrive/SharePoint)
- **Actions**:
  - Scan users' Recordings folders
  - Download transcript files (.vtt, .docx)
  - Upload to Azure Blob Storage
  - Generate Excel metadata file
- **Output**: 
  - Transcripts in Blob Storage
  - Excel file with meeting metadata

### 2. Azure Blob Storage
- **Container**: `transcripts`
- **Structure**:
  ```
  transcripts/
  ├── 2025-12/
  │   ├── meeting_20251201_user1.vtt
  │   ├── meeting_20251201_user1.txt  (converted)
  │   └── ...
  └── metadata/
      └── meetings_metadata.xlsx
  ```

### 3. Excel Metadata File
| Column | Description |
|--------|-------------|
| meeting_id | Unique identifier |
| meeting_date | Date of meeting |
| meeting_time | Start time |
| organizer_name | Meeting host |
| organizer_email | Host email |
| participants | List of attendees |
| duration_minutes | Meeting length |
| subject | Meeting title |
| transcript_blob_url | Link to transcript in Blob |
| recording_blob_url | Link to recording (optional) |
| processed | Boolean - has been analyzed |
| sentiment_score | -1 to +1 (filled by AI) |
| churn_risk_score | 0 to 100 (filled by AI) |
| upsell_potential | 0 to 100 (filled by AI) |
| execution_reliability | 0 to 100 (filled by AI) |
| operational_complexity | 0 to 100 (filled by AI) |
| events_detected | JSON array of events |
| analysis_summary | AI-generated summary |
| last_analyzed | Timestamp |

### 4. Power Automate Flow

#### Flow Design:
```
Trigger: Recurrence (Daily) or Manual
    │
    ▼
Get Excel rows where processed = FALSE
    │
    ▼
For Each unprocessed row:
    │
    ├─▶ Get transcript from Blob Storage URL
    │
    ├─▶ Call Azure OpenAI / GPT-4 with prompt
    │   (Include transcript + scoring instructions)
    │
    ├─▶ Parse AI response (JSON)
    │
    └─▶ Update Excel row with scores + events
```

## Analysis Prompt Template

```
You are a meeting transcript analyzer. Analyze the following meeting transcript and provide:

## UNIVERSAL SCORES (always calculate these):

1. **Sentiment Score** (-1.0 to +1.0)
   - Negative = frustrated, angry, disappointed
   - Neutral = 0
   - Positive = happy, satisfied, enthusiastic
   
2. **Churn Risk Score** (0-100)
   - 0 = No risk, very satisfied
   - 100 = High risk, likely to leave
   - Consider: complaints, delays mentioned, tone shifts, ultimatums
   
3. **Upsell/Opportunity Potential** (0-100)
   - 0 = No opportunity
   - 100 = Strong buying signals
   - Consider: expansion hints, new needs, interest in additional services
   
4. **Execution Reliability** (0-100)
   - 0 = Many complaints about delivery
   - 100 = All expectations met/exceeded
   - Consider: complaints vs praises about your team's work
   
5. **Operational Complexity** (0-100)
   - 0 = Simple, few items
   - 100 = Very complex, many tasks/deadlines/issues
   - Consider: upcoming tasks, deadlines, chaos signals

## EVENTS DETECTED (extract all that apply):

Identify specific events mentioned. For each event, provide:
- event_type: Category (see examples below)
- description: Brief description
- speaker: Who mentioned it (if known)
- severity: low/medium/high

Example event types (not exhaustive - add new ones as needed):
- Complaint
- Delay
- Decision
- Scope Change
- VA Performance Issue
- Payment Issue
- Positive Feedback
- New Idea
- Process Confusion
- Feature Request
- Deadline
- Escalation
- Resolution

## OUTPUT FORMAT (JSON):

{
  "sentiment_score": 0.0,
  "churn_risk_score": 0,
  "upsell_potential": 0,
  "execution_reliability": 0,
  "operational_complexity": 0,
  "summary": "2-3 sentence summary of the meeting",
  "key_topics": ["topic1", "topic2"],
  "action_items": ["action1", "action2"],
  "events": [
    {
      "event_type": "Complaint",
      "description": "Client mentioned delays in report delivery",
      "speaker": "John Smith",
      "severity": "medium"
    }
  ]
}

## TRANSCRIPT:

{transcript_content}
```

## Implementation Steps

### Phase 1: Metadata Extraction (Python)
1. ✅ Scan recordings folders (done)
2. Download transcript files
3. Parse VTT/DOCX for metadata
4. Upload to Azure Blob Storage
5. Generate Excel file

### Phase 2: Power Automate Setup
1. Create Blob Storage connection
2. Create Excel Online connection
3. Create Azure OpenAI connection
4. Build the flow logic
5. Test with sample transcripts

### Phase 3: Dashboard (Optional)
1. Power BI dashboard connected to Excel
2. Visualize scores over time
3. Alert on high churn risk
4. Track sentiment trends

## Azure Resources Needed

| Resource | Purpose | Estimated Cost |
|----------|---------|----------------|
| Azure Blob Storage | Store transcripts | ~$0.02/GB/month |
| Azure OpenAI (GPT-4) | Analyze transcripts | ~$0.03/1K tokens |
| Power Automate | Orchestration | Included with M365 |
| Excel Online | Metadata storage | Included with M365 |

## Security Considerations

1. **Blob Storage**: Use SAS tokens with expiry for transcript URLs
2. **Excel**: Store in SharePoint with appropriate permissions
3. **API Keys**: Store in Azure Key Vault or Power Automate secure inputs
4. **Data Retention**: Define policy for transcript storage duration

## Next Steps

1. Run `--export-metadata` to generate initial Excel file
2. Set up Azure Blob Storage account
3. Run `--upload-blobs` to upload transcripts
4. Create Power Automate flow
5. Test with a few transcripts
6. Scale to full dataset
