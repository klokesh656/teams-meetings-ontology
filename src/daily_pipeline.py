#!/usr/bin/env python3
"""
Daily Automated Pipeline for VA Check-in Analysis
=================================================
This script runs daily to:
1. Download new meeting transcripts from Microsoft Graph API
2. Analyze new transcripts with the Churn Risk Checklist
3. Update analysis history incrementally
4. Regenerate CSV files for Power BI/Power Automate
5. Upload results to Azure Blob Storage
6. (Optional) Index data in Azure AI Search

Designed to run as:
- Azure Function (Timer Trigger)
- Azure Container Instance (Scheduled)
- Local scheduled task (for testing)
"""

import os
import json
import re
import csv
import logging
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import quote
import requests

from dotenv import load_dotenv
from azure.identity import ClientSecretCredential
from azure.storage.blob import BlobServiceClient

# Load environment variables
load_dotenv()

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# ============================================================================
# CONFIGURATION
# ============================================================================

# Azure AD
TENANT_ID = os.getenv('AZURE_TENANT_ID')
CLIENT_ID = os.getenv('AZURE_CLIENT_ID')
CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET')
HR_USER_ID = '81835016-79d5-4a15-91b1-c104e2cd9adb'

# Azure Storage
STORAGE_CONNECTION_STRING = os.getenv('AZURE_STORAGE_CONNECTION_STRING')
STORAGE_ACCOUNT = "aidevelopement"
TRANSCRIPT_CONTAINER = "transcripts"
OUTPUT_CONTAINER = "pipeline-outputs"

# Azure OpenAI
AZURE_OPENAI_ENDPOINT = os.getenv('AZURE_OPENAI_ENDPOINT')
AZURE_OPENAI_KEY = os.getenv('AZURE_OPENAI_KEY')
AZURE_OPENAI_DEPLOYMENT = os.getenv('AZURE_OPENAI_DEPLOYMENT', 'gpt-4.1')

# Paths (for local development)
BASE_DIR = Path(__file__).parent.parent
TRANSCRIPTS_DIR = BASE_DIR / "transcripts"
OUTPUT_DIR = BASE_DIR / "output"
LOGS_DIR = BASE_DIR / "logs"

# Pipeline state file
PIPELINE_STATE_FILE = OUTPUT_DIR / "pipeline_state.json"

# ============================================================================
# CHURN RISK CHECKLIST (27 Signals)
# ============================================================================

CHURN_RISK_CHECKLIST = """
## VA-Side Signals (VA001-VA012)
- VA001: Resignation hints or job search mentions
- VA002: Attendance issues (late, absent, unreliable)
- VA003: Health concerns affecting work
- VA004: Overwhelming workload or burnout signs
- VA005: Skill gaps or training needs
- VA006: Communication breakdowns
- VA007: Disengagement or low motivation
- VA008: Personal/family issues affecting work
- VA009: Compensation or benefits concerns
- VA010: Career growth frustration
- VA011: Work-life balance issues
- VA012: Tool/system access problems

## Client-Side Signals (CL001-CL010)
- CL001: Negative feedback or complaints
- CL002: Scope creep or unclear expectations
- CL003: Non-responsive or unavailable client
- CL004: Payment or contract issues
- CL005: Micromanagement or trust issues
- CL006: Business downturn or budget cuts
- CL007: Organizational changes at client
- CL008: Multiple VA requests or comparison
- CL009: Reduced hours or task volume
- CL010: Communication style mismatch

## Relationship/HR Signals (RH001-RH005)
- RH001: Conflict between VA and client
- RH002: Integration team escalations
- RH003: Contract/engagement at risk
- RH004: Missing regular check-ins
- RH005: No visibility (missing SOD/EOD reports)
"""


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def get_graph_headers():
    """Get Microsoft Graph API authentication headers."""
    credential = ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
    token = credential.get_token('https://graph.microsoft.com/.default')
    return {
        'Authorization': f'Bearer {token.token}',
        'Content-Type': 'application/json'
    }


def get_blob_service():
    """Get Azure Blob Storage service client."""
    return BlobServiceClient.from_connection_string(STORAGE_CONNECTION_STRING)


def load_pipeline_state():
    """Load pipeline state from file."""
    if PIPELINE_STATE_FILE.exists():
        with open(PIPELINE_STATE_FILE, 'r') as f:
            return json.load(f)
    return {
        'last_run': None,
        'last_transcript_date': None,
        'total_transcripts': 0,
        'total_analyzed': 0
    }


def save_pipeline_state(state):
    """Save pipeline state to file."""
    OUTPUT_DIR.mkdir(exist_ok=True)
    with open(PIPELINE_STATE_FILE, 'w') as f:
        json.dump(state, f, indent=2, default=str)


def generate_blob_url(filename):
    """Generate Azure Blob URL for a transcript file."""
    encoded_name = quote(filename, safe='')
    return f"https://{STORAGE_ACCOUNT}.blob.core.windows.net/{TRANSCRIPT_CONTAINER}/{encoded_name}"


def generate_meeting_id(source_file, va_name='', meeting_date=''):
    """Generate a consistent Meeting ID based on source file.
    
    This creates a unique, stable ID that remains the same across pipeline runs,
    enabling Power BI relationships and Power Automate status tracking.
    """
    import hashlib
    # Use source file as primary key - it's unique per meeting
    if source_file:
        # Create short hash from filename for uniqueness
        file_hash = hashlib.md5(source_file.encode()).hexdigest()[:6].upper()
        return f"MTG-{file_hash}"
    # Fallback to VA+date if no source file
    elif va_name and meeting_date:
        combined = f"{va_name}_{meeting_date}"
        file_hash = hashlib.md5(combined.encode()).hexdigest()[:6].upper()
        return f"MTG-{file_hash}"
    return "MTG-UNKNOWN"


def load_review_status():
    """Load review status from tracking file."""
    status_file = OUTPUT_DIR / "meeting_review_status.json"
    if status_file.exists():
        with open(status_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def save_review_status(status):
    """Save review status to tracking file."""
    status_file = OUTPUT_DIR / "meeting_review_status.json"
    with open(status_file, 'w', encoding='utf-8') as f:
        json.dump(status, f, indent=2)


# ============================================================================
# STEP 1: DOWNLOAD NEW TRANSCRIPTS
# ============================================================================

def download_new_transcripts(days_back=7):
    """Download transcripts from the last N days that we don't have."""
    logger.info(f"📥 Checking for new transcripts (last {days_back} days)...")
    
    headers = get_graph_headers()
    TRANSCRIPTS_DIR.mkdir(exist_ok=True)
    
    # Get existing files
    existing_files = {f.stem.lower() for f in TRANSCRIPTS_DIR.glob("*.vtt")}
    existing_dates = set()
    for f in TRANSCRIPTS_DIR.glob("*.vtt"):
        match = re.search(r'(\d{8}_\d{4})', f.stem)
        if match:
            existing_dates.add(match.group(1))
    
    # Fetch transcript list from Graph API
    url = f"https://graph.microsoft.com/beta/users/{HR_USER_ID}/onlineMeetings/getAllTranscripts(meetingOrganizerUserId='{HR_USER_ID}')"
    
    all_transcripts = []
    while url:
        resp = requests.get(url, headers=headers, timeout=60)
        if resp.status_code != 200:
            logger.error(f"Failed to fetch transcripts: {resp.status_code}")
            break
        data = resp.json()
        all_transcripts.extend(data.get('value', []))
        url = data.get('@odata.nextLink')
    
    logger.info(f"   Found {len(all_transcripts)} total transcripts in API")
    
    # Filter to recent ones
    cutoff_date = (datetime.utcnow() - timedelta(days=days_back)).isoformat() + 'Z'
    recent = [t for t in all_transcripts if t.get('createdDateTime', '') >= cutoff_date]
    
    logger.info(f"   {len(recent)} transcripts in last {days_back} days")
    
    # Download new transcripts
    downloaded = []
    for t in recent:
        created = t.get('createdDateTime', '')[:19].replace('T', ' ')
        transcript_id = t.get('id', '')
        meeting_id = t.get('meetingId', '')
        
        date_str = created[:10].replace('-', '')
        time_str = created[11:16].replace(':', '')
        date_time_key = f"{date_str}_{time_str}"
        
        if date_time_key in existing_dates:
            continue
        
        # Get meeting subject
        subject = get_meeting_subject(headers, meeting_id)
        
        # Download transcript content
        content = download_transcript_content(headers, meeting_id, transcript_id)
        
        if content:
            safe_subject = re.sub(r'[<>:"/\\|?*]', '', subject)[:60]
            filename = f"{date_str}_{time_str}_{safe_subject}.vtt"
            filepath = TRANSCRIPTS_DIR / filename
            
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(content)
            
            logger.info(f"   ✅ Downloaded: {filename}")
            downloaded.append({
                'filename': filename,
                'date': created[:10],
                'subject': subject
            })
            
            # Also upload to Azure Blob
            upload_to_blob(filepath, filename)
    
    logger.info(f"   Downloaded {len(downloaded)} new transcripts")
    return downloaded


def get_meeting_subject(headers, meeting_id):
    """Get meeting subject from Graph API."""
    url = f"https://graph.microsoft.com/beta/users/{HR_USER_ID}/onlineMeetings/{meeting_id}"
    try:
        resp = requests.get(url, headers=headers, timeout=30)
        if resp.status_code == 200:
            return resp.json().get('subject', 'Unknown Meeting')
    except:
        pass
    return 'Unknown Meeting'


def download_transcript_content(headers, meeting_id, transcript_id):
    """Download transcript content."""
    url = f"https://graph.microsoft.com/beta/users/{HR_USER_ID}/onlineMeetings/{meeting_id}/transcripts/{transcript_id}/content"
    try:
        resp = requests.get(url, headers=headers, params={'$format': 'text/vtt'}, timeout=60)
        if resp.status_code == 200:
            return resp.text
    except Exception as e:
        logger.warning(f"   Failed to download: {e}")
    return None


def upload_to_blob(filepath, blob_name):
    """Upload file to Azure Blob Storage."""
    try:
        blob_service = get_blob_service()
        container_client = blob_service.get_container_client(TRANSCRIPT_CONTAINER)
        
        # Create container if not exists
        try:
            container_client.create_container()
        except:
            pass
        
        blob_client = container_client.get_blob_client(blob_name)
        with open(filepath, 'rb') as f:
            blob_client.upload_blob(f, overwrite=True)
        
        logger.debug(f"   Uploaded to blob: {blob_name}")
    except Exception as e:
        logger.warning(f"   Blob upload failed: {e}")


# ============================================================================
# STEP 2: ANALYZE NEW TRANSCRIPTS
# ============================================================================

def analyze_new_transcripts(new_files=None):
    """Analyze new check-in transcripts."""
    logger.info("🔍 Analyzing new check-in transcripts...")
    
    # Load existing analysis history
    history_file = OUTPUT_DIR / "checkin_analysis_history.json"
    if history_file.exists():
        with open(history_file, 'r', encoding='utf-8') as f:
            history = json.load(f)
    else:
        history = {'analyses': [], 'last_updated': None}
    
    # Get already analyzed files
    analyzed_keys = set()
    for a in history.get('analyses', []):
        key = f"{a.get('va_name', '')}_{a.get('meeting_date', '')}"
        analyzed_keys.add(key.lower())
    
    # Find check-in files to analyze
    if new_files:
        files_to_analyze = [TRANSCRIPTS_DIR / f['filename'] for f in new_files 
                          if 'check-in' in f.get('subject', '').lower() or 'check in' in f.get('subject', '').lower()]
    else:
        # Analyze all unanalyzed check-in files
        files_to_analyze = []
        for f in TRANSCRIPTS_DIR.glob("*.vtt"):
            if 'check-in' not in f.name.lower() and 'check in' not in f.name.lower():
                continue
            
            # Extract VA name and date
            va_name, date_str = extract_va_and_date(f.name)
            if not va_name:
                continue
            
            key = f"{va_name}_{date_str}".lower()
            if key not in analyzed_keys:
                files_to_analyze.append(f)
    
    logger.info(f"   {len(files_to_analyze)} new check-ins to analyze")
    
    # Analyze each file
    new_analyses = []
    for filepath in files_to_analyze[:30]:  # Limit per run
        try:
            analysis = analyze_single_transcript(filepath)
            if analysis:
                new_analyses.append(analysis)
                logger.info(f"   ✅ Analyzed: {filepath.name[:50]}... -> {analysis['overall_risk_level'].upper()}")
        except Exception as e:
            logger.error(f"   ❌ Failed to analyze {filepath.name}: {e}")
    
    # Update history
    history['analyses'].extend(new_analyses)
    history['last_updated'] = datetime.now().isoformat()
    
    with open(history_file, 'w', encoding='utf-8') as f:
        json.dump(history, f, indent=2, default=str)
    
    logger.info(f"   Analyzed {len(new_analyses)} new transcripts")
    return new_analyses


def extract_va_and_date(filename):
    """Extract VA name and date from filename."""
    va_name = None
    date_str = "Unknown"
    
    # Extract date
    if filename[:8].isdigit():
        date_str = f"{filename[:4]}-{filename[4:6]}-{filename[6:8]}"
    
    # Extract VA name
    if " x " in filename.lower():
        parts = filename.split(" x ")
        if len(parts) >= 2:
            va_name = parts[-1].replace(".vtt", "").split("-")[0].strip()
            va_name = va_name.replace("_", " ").strip()
    
    return va_name, date_str


def analyze_single_transcript(filepath):
    """Analyze a single transcript using Azure OpenAI."""
    # Read transcript
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
    
    if len(content) < 100:
        return None
    
    # Extract VA name and date
    va_name, date_str = extract_va_and_date(filepath.name)
    if not va_name:
        return None
    
    # Prepare prompt
    prompt = f"""Analyze this VA check-in meeting transcript for churn risk signals.

CHURN RISK CHECKLIST:
{CHURN_RISK_CHECKLIST}

TRANSCRIPT:
{content[:15000]}

Respond in JSON format:
{{
    "va_name": "{va_name}",
    "client_name": "extracted client name or Unknown",
    "meeting_date": "{date_str}",
    "overall_risk_level": "low|medium|high|critical",
    "va_status": "green|yellow|red",
    "client_health": "healthy|at_risk|critical",
    "detected_signals": [
        {{"signal_id": "VA001", "evidence": "quote from transcript", "confidence": "high|medium|low"}}
    ],
    "executive_summary": "2-3 sentence summary",
    "key_findings": ["finding 1", "finding 2", "finding 3"],
    "ai_suggestions": [
        {{"issue": "issue description", "suggestion": "recommended action", "urgency": "immediate|within_48h|this_week|monitor", "category": "category"}}
    ]
}}"""

    # Call Azure OpenAI
    try:
        headers = {
            'Content-Type': 'application/json',
            'api-key': AZURE_OPENAI_KEY
        }
        
        payload = {
            'messages': [
                {'role': 'system', 'content': 'You are an expert HR analyst specializing in VA retention and churn risk analysis.'},
                {'role': 'user', 'content': prompt}
            ],
            'temperature': 0.3,
            'max_tokens': 2000
        }
        
        url = f"{AZURE_OPENAI_ENDPOINT}/openai/deployments/{AZURE_OPENAI_DEPLOYMENT}/chat/completions?api-version=2024-02-15-preview"
        
        resp = requests.post(url, headers=headers, json=payload, timeout=120)
        
        if resp.status_code == 200:
            result = resp.json()
            content = result['choices'][0]['message']['content']
            
            # Parse JSON from response
            json_match = re.search(r'\{[\s\S]*\}', content)
            if json_match:
                analysis = json.loads(json_match.group())
                analysis['source_file'] = filepath.name
                analysis['analyzed_at'] = datetime.now().isoformat()
                return analysis
        else:
            logger.error(f"OpenAI API error: {resp.status_code}")
    
    except Exception as e:
        logger.error(f"Analysis failed: {e}")
    
    return None


# ============================================================================
# STEP 3: GENERATE CSV FILES
# ============================================================================

def generate_csv_files():
    """Generate all CSV files for Power BI and Power Automate."""
    logger.info("📊 Generating CSV files...")
    
    # Load data
    history_file = OUTPUT_DIR / "checkin_analysis_history.json"
    if not history_file.exists():
        logger.warning("No analysis history found")
        return
    
    with open(history_file, 'r', encoding='utf-8') as f:
        history = json.load(f)
    
    analyses = history.get('analyses', [])
    logger.info(f"   Processing {len(analyses)} analyses")
    
    # Generate each CSV
    generate_va_risk_summary(analyses)
    generate_critical_alerts(analyses)
    generate_all_meetings_detail(analyses)
    generate_va_client_mapping()
    generate_kpi_summary(analyses)
    generate_pending_suggestions(analyses)
    
    logger.info("   ✅ All CSV files generated")


def generate_va_risk_summary(analyses):
    """Generate VA risk summary CSV."""
    review_status = load_review_status()
    
    va_stats = {}
    for a in analyses:
        va = a.get('va_name', 'Unknown')
        meeting_id = generate_meeting_id(a.get('source_file', ''), va, a.get('meeting_date', ''))
        
        if va not in va_stats:
            va_stats[va] = {
                'va_name': va,
                'total_checkins': 0,
                'critical': 0, 'high': 0, 'medium': 0, 'low': 0,
                'total_signals': 0,
                'latest_date': '',
                'latest_risk': '',
                'latest_source': '',
                'latest_meeting_id': '',
                'clients': set(),
                'reviewed': 0,
                'pending': 0
            }
        
        stats = va_stats[va]
        stats['total_checkins'] += 1
        stats['total_signals'] += len(a.get('detected_signals', []))
        
        # Track review status
        status = review_status.get(meeting_id, {}).get('status', 'Pending')
        if status == 'Reviewed':
            stats['reviewed'] += 1
        else:
            stats['pending'] += 1
        
        risk = a.get('overall_risk_level', 'medium').lower()
        stats[risk] = stats.get(risk, 0) + 1
        
        date = a.get('meeting_date', '')
        if date >= stats['latest_date']:
            stats['latest_date'] = date
            stats['latest_risk'] = risk
            stats['latest_source'] = a.get('source_file', '')
            stats['latest_meeting_id'] = meeting_id
        
        client = a.get('client_name', '')
        if client and client != 'Unknown':
            stats['clients'].add(client[:30])
    
    output_file = OUTPUT_DIR / "va_risk_summary.csv"
    with open(output_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow([
            'Latest Meeting ID', 'VA Name', 'Total Check-ins', 'Reviewed', 'Pending',
            'Critical', 'High', 'Medium', 'Low', 'Total Signals', 'Risk Score', 
            'Current Risk', 'Latest Date', 'Clients', 'Attention', 'Transcript Link'
        ])
        
        for va, stats in sorted(va_stats.items(), key=lambda x: x[1]['critical']*100 + x[1]['high']*10, reverse=True):
            risk_score = stats['critical']*100 + stats['high']*25 + stats['medium']*5 + stats['low']
            attention = 'CRITICAL' if stats['critical'] > 0 else 'HIGH' if stats['high'] > 0 else 'MONITOR' if stats['medium'] > 2 else 'OK'
            
            writer.writerow([
                stats['latest_meeting_id'],
                stats['va_name'],
                stats['total_checkins'],
                stats['reviewed'],
                stats['pending'],
                stats['critical'],
                stats['high'],
                stats['medium'],
                stats['low'],
                stats['total_signals'],
                risk_score,
                stats['latest_risk'].upper(),
                stats['latest_date'],
                '; '.join(list(stats['clients'])[:3]),
                attention,
                generate_blob_url(stats['latest_source']) if stats['latest_source'] else ''
            ])
    
    logger.info(f"   ✅ va_risk_summary.csv ({len(va_stats)} VAs)")


def generate_critical_alerts(analyses):
    """Generate critical alerts CSV."""
    output_file = OUTPUT_DIR / "critical_alerts.csv"
    alerts = []
    review_status = load_review_status()
    
    for a in analyses:
        risk = a.get('overall_risk_level', '').lower()
        if risk in ['critical', 'high']:
            alerts.append(a)
    
    with open(output_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow([
            'Meeting ID', 'Priority', 'VA Name', 'Client', 'Date', 'Risk Level',
            'Review Status', 'Signals', 'Summary', 'Key Findings', 
            'Client Input', 'Top Suggestion', 'Source File', 'Blob Link'
        ])
        
        for a in sorted(alerts, key=lambda x: (0 if x.get('overall_risk_level','')=='critical' else 1, x.get('meeting_date','')), reverse=True):
            meeting_id = generate_meeting_id(a.get('source_file', ''), a.get('va_name', ''), a.get('meeting_date', ''))
            priority = 'P1-CRITICAL' if a.get('overall_risk_level','').lower() == 'critical' else 'P2-HIGH'
            status_info = review_status.get(meeting_id, {})
            
            # Get top suggestion
            suggestions = a.get('ai_suggestions', [])
            top_suggestion = suggestions[0].get('suggestion', '')[:150] if suggestions else ''
            
            writer.writerow([
                meeting_id,
                priority,
                a.get('va_name', ''),
                a.get('client_name', 'Unknown'),
                a.get('meeting_date', ''),
                a.get('overall_risk_level', '').upper(),
                status_info.get('status', 'Pending'),
                len(a.get('detected_signals', [])),
                a.get('executive_summary', '')[:300],
                '; '.join(a.get('key_findings', [])[:2])[:200],
                status_info.get('client_input', ''),
                top_suggestion,
                a.get('source_file', ''),
                generate_blob_url(a.get('source_file', ''))
            ])
    
    logger.info(f"   ✅ critical_alerts.csv ({len(alerts)} alerts)")


def generate_all_meetings_detail(analyses):
    """Generate all meetings detail CSV with consistent Meeting IDs."""
    output_file = OUTPUT_DIR / "all_meetings_detail.csv"
    review_status = load_review_status()
    
    with open(output_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow([
            'Meeting ID', 'VA Name', 'Client', 'Date', 'Risk Level', 'Risk Score',
            'Review Status', 'Signals', 'VA Status', 'Client Health', 'Summary',
            'Top Suggestions', 'Client Input', 'Key Findings', 'Source File', 'Blob Link'
        ])
        
        for a in sorted(analyses, key=lambda x: x.get('meeting_date',''), reverse=True):
            meeting_id = generate_meeting_id(a.get('source_file', ''), a.get('va_name', ''), a.get('meeting_date', ''))
            risk_weights = {'critical': 100, 'high': 75, 'medium': 25, 'low': 5}
            risk_score = risk_weights.get(a.get('overall_risk_level','medium').lower(), 10)
            status_info = review_status.get(meeting_id, {})
            
            # Get top 2 suggestions
            suggestions = a.get('ai_suggestions', [])
            top_suggestions = '; '.join([s.get('suggestion', '')[:100] for s in suggestions[:2]])
            
            writer.writerow([
                meeting_id,
                a.get('va_name', ''),
                a.get('client_name', 'Unknown'),
                a.get('meeting_date', ''),
                a.get('overall_risk_level', '').upper(),
                risk_score,
                status_info.get('status', 'Pending'),
                len(a.get('detected_signals', [])),
                a.get('va_status', ''),
                a.get('client_health', ''),
                a.get('executive_summary', '')[:400],
                top_suggestions[:300],
                status_info.get('client_input', ''),
                '; '.join(a.get('key_findings', [])[:3])[:300],
                a.get('source_file', ''),
                generate_blob_url(a.get('source_file', ''))
            ])
    
    logger.info(f"   ✅ all_meetings_detail.csv ({len(analyses)} meetings)")


def generate_va_client_mapping():
    """Generate VA-Client mapping CSV for team to fill."""
    output_file = OUTPUT_DIR / "va_client_mapping.csv"
    
    # Load history
    history_file = OUTPUT_DIR / "checkin_analysis_history.json"
    history = {}
    if history_file.exists():
        with open(history_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
            for a in data.get('analyses', []):
                key = f"{a.get('va_name','')}_{a.get('meeting_date','')}"
                history[key.lower()] = a
    
    # Get check-in files
    meetings = []
    for f in TRANSCRIPTS_DIR.glob("*.vtt"):
        if 'check-in' not in f.name.lower() and 'check in' not in f.name.lower():
            continue
        
        va_name, date_str = extract_va_and_date(f.name)
        if not va_name:
            continue
        
        key = f"{va_name}_{date_str}".lower()
        hist = history.get(key, {})
        
        meetings.append({
            'filename': f.name,
            'va_name': va_name,
            'date': date_str,
            'client': hist.get('client_name', ''),
            'blob_url': generate_blob_url(f.name)
        })
    
    meetings.sort(key=lambda x: x['date'], reverse=True)
    
    review_status = load_review_status()
    
    with open(output_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow([
            'Meeting ID', 'VA Name', 'Client Name (FILL IN)', 'Client (Auto)', 'Date',
            'Review Status', 'Subject', 'Source File', 'Blob Link', 'Notes'
        ])
        
        for m in meetings:
            meeting_id = generate_meeting_id(m['filename'], m['va_name'], m['date'])
            status_info = review_status.get(meeting_id, {})
            subject = m['filename'].replace('.vtt', '').replace('_', ' ')[:60]
            writer.writerow([
                meeting_id,
                m['va_name'],
                '',  # For team to fill
                m['client'] if m['client'] != 'Unknown' else '',
                m['date'],
                status_info.get('status', 'Pending'),
                subject,
                m['filename'],
                m['blob_url'],
                status_info.get('notes', '')
            ])
    
    logger.info(f"   ✅ va_client_mapping.csv ({len(meetings)} meetings)")


def generate_kpi_summary(analyses):
    """Generate KPI summary CSV."""
    output_file = OUTPUT_DIR / "kpi_dashboard_summary.csv"
    
    total_vas = len(set(a.get('va_name') for a in analyses))
    total = len(analyses)
    critical = sum(1 for a in analyses if a.get('overall_risk_level','').lower() == 'critical')
    high = sum(1 for a in analyses if a.get('overall_risk_level','').lower() == 'high')
    medium = sum(1 for a in analyses if a.get('overall_risk_level','').lower() == 'medium')
    low = sum(1 for a in analyses if a.get('overall_risk_level','').lower() == 'low')
    
    risk_rate = (critical + high) / total * 100 if total > 0 else 0
    
    with open(output_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['KPI', 'Value', 'Description', 'Updated'])
        now = datetime.now().strftime('%Y-%m-%d %H:%M')
        
        writer.writerow(['Total VAs', total_vas, 'Unique VAs monitored', now])
        writer.writerow(['Total Check-ins', total, 'Analyzed meetings', now])
        writer.writerow(['Critical Risk', critical, 'Immediate action needed', now])
        writer.writerow(['High Risk', high, 'Follow-up within 48h', now])
        writer.writerow(['Medium Risk', medium, 'Monitor closely', now])
        writer.writerow(['Low Risk', low, 'Healthy status', now])
        writer.writerow(['At-Risk Rate %', f'{risk_rate:.1f}', 'Critical+High %', now])
    
    logger.info(f"   ✅ kpi_dashboard_summary.csv")


def generate_pending_suggestions(analyses):
    """Generate pending suggestions CSV with Meeting IDs."""
    output_file = OUTPUT_DIR / "pending_suggestions_review.csv"
    review_status = load_review_status()
    
    with open(output_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow([
            'Suggestion ID', 'Meeting ID', 'VA Name', 'Client', 'Date', 'Risk', 
            'Review Status', 'Issue', 'Suggestion', 'Urgency', 'Category', 
            'Source File', 'Blob Link'
        ])
        
        count = 0
        for a in analyses:
            meeting_id = generate_meeting_id(a.get('source_file', ''), a.get('va_name', ''), a.get('meeting_date', ''))
            status_info = review_status.get(meeting_id, {})
            
            for sugg in a.get('ai_suggestions', []):
                count += 1
                writer.writerow([
                    f'SUGG-{count:04d}',
                    meeting_id,
                    a.get('va_name', ''),
                    a.get('client_name', 'Unknown'),
                    a.get('meeting_date', ''),
                    a.get('overall_risk_level', '').upper(),
                    status_info.get('status', 'Pending'),
                    sugg.get('issue', '')[:200],
                    sugg.get('suggestion', '')[:400],
                    sugg.get('urgency', 'monitor'),
                    sugg.get('category', ''),
                    a.get('source_file', ''),
                    generate_blob_url(a.get('source_file', ''))
                ])
    
    logger.info(f"   ✅ pending_suggestions_review.csv ({count} suggestions)")


# ============================================================================
# STEP 4: UPLOAD OUTPUTS TO AZURE
# ============================================================================

def upload_outputs_to_azure():
    """Upload all output files to Azure Blob Storage."""
    logger.info("☁️ Uploading outputs to Azure Blob Storage...")
    
    try:
        blob_service = get_blob_service()
        container_client = blob_service.get_container_client(OUTPUT_CONTAINER)
        
        try:
            container_client.create_container()
        except:
            pass
        
        # Upload CSV files
        csv_files = list(OUTPUT_DIR.glob("*.csv"))
        for f in csv_files:
            blob_name = f"csv/{datetime.now().strftime('%Y%m%d')}/{f.name}"
            blob_client = container_client.get_blob_client(blob_name)
            with open(f, 'rb') as data:
                blob_client.upload_blob(data, overwrite=True)
            logger.info(f"   ✅ Uploaded: {blob_name}")
        
        # Upload analysis history
        history_file = OUTPUT_DIR / "checkin_analysis_history.json"
        if history_file.exists():
            blob_name = f"data/checkin_analysis_history.json"
            blob_client = container_client.get_blob_client(blob_name)
            with open(history_file, 'rb') as data:
                blob_client.upload_blob(data, overwrite=True)
            logger.info(f"   ✅ Uploaded: {blob_name}")
        
        # Also upload "latest" copies for Power BI direct access
        for f in csv_files:
            blob_name = f"latest/{f.name}"
            blob_client = container_client.get_blob_client(blob_name)
            with open(f, 'rb') as data:
                blob_client.upload_blob(data, overwrite=True)
        
        logger.info("   ✅ All outputs uploaded")
        
    except Exception as e:
        logger.error(f"   ❌ Upload failed: {e}")


# ============================================================================
# STEP 5: GENERATE DAILY REPORT
# ============================================================================

def generate_daily_report(downloaded, analyzed):
    """Generate daily pipeline report."""
    LOGS_DIR.mkdir(exist_ok=True)
    
    report_file = LOGS_DIR / f"daily_report_{datetime.now().strftime('%Y%m%d')}.txt"
    
    with open(report_file, 'w', encoding='utf-8') as f:
        f.write("=" * 60 + "\n")
        f.write("DAILY VA CHECK-IN PIPELINE REPORT\n")
        f.write("=" * 60 + "\n")
        f.write(f"Run Time: {datetime.now()}\n\n")
        
        f.write("[DOWNLOAD] TRANSCRIPTS DOWNLOADED\n")
        f.write(f"   New transcripts: {len(downloaded)}\n")
        for d in downloaded[:10]:
            f.write(f"   - {d['filename'][:50]}...\n")
        
        f.write("\n[ANALYZE] ANALYSES COMPLETED\n")
        f.write(f"   New analyses: {len(analyzed)}\n")
        
        critical = [a for a in analyzed if a.get('overall_risk_level','').lower() == 'critical']
        high = [a for a in analyzed if a.get('overall_risk_level','').lower() == 'high']
        
        if critical:
            f.write("\n[CRITICAL] CRITICAL ALERTS:\n")
            for a in critical:
                f.write(f"   - {a.get('va_name')}: {a.get('executive_summary','')[:100]}...\n")
        
        if high:
            f.write("\n[HIGH] HIGH RISK:\n")
            for a in high:
                f.write(f"   - {a.get('va_name')}: {a.get('executive_summary','')[:100]}...\n")
        
        f.write("\n" + "=" * 60 + "\n")
    
    logger.info(f"Daily report saved: {report_file}")
    return report_file


# ============================================================================
# MAIN PIPELINE
# ============================================================================

def run_daily_pipeline(days_back=7):
    """Run the complete daily pipeline."""
    start_time = datetime.now()
    
    logger.info("=" * 60)
    logger.info("🚀 STARTING DAILY VA CHECK-IN PIPELINE")
    logger.info("=" * 60)
    logger.info(f"   Start Time: {start_time}")
    
    # Load state
    state = load_pipeline_state()
    logger.info(f"   Last Run: {state.get('last_run', 'Never')}")
    
    try:
        # Step 1: Download new transcripts
        downloaded = download_new_transcripts(days_back)
        
        # Step 2: Analyze new check-in transcripts
        analyzed = analyze_new_transcripts(downloaded if downloaded else None)
        
        # Step 3: Generate CSV files
        generate_csv_files()
        
        # Step 4: Upload to Azure
        upload_outputs_to_azure()
        
        # Step 5: Generate report
        generate_daily_report(downloaded, analyzed)
        
        # Update state
        state['last_run'] = datetime.now().isoformat()
        state['total_transcripts'] = state.get('total_transcripts', 0) + len(downloaded)
        state['total_analyzed'] = state.get('total_analyzed', 0) + len(analyzed)
        save_pipeline_state(state)
        
        end_time = datetime.now()
        duration = (end_time - start_time).total_seconds()
        
        logger.info("=" * 60)
        logger.info("✅ PIPELINE COMPLETED SUCCESSFULLY")
        logger.info("=" * 60)
        logger.info(f"   Duration: {duration:.1f} seconds")
        logger.info(f"   Downloaded: {len(downloaded)} transcripts")
        logger.info(f"   Analyzed: {len(analyzed)} check-ins")
        
        return {
            'status': 'success',
            'downloaded': len(downloaded),
            'analyzed': len(analyzed),
            'duration': duration
        }
        
    except Exception as e:
        logger.error(f"❌ Pipeline failed: {e}")
        return {
            'status': 'error',
            'error': str(e)
        }


# ============================================================================
# ENTRY POINT
# ============================================================================

if __name__ == '__main__':
    import argparse
    
    parser = argparse.ArgumentParser(description='Daily VA Check-in Pipeline')
    parser.add_argument('--days', type=int, default=7, help='Days to look back for transcripts')
    parser.add_argument('--skip-download', action='store_true', help='Skip transcript download')
    parser.add_argument('--skip-analyze', action='store_true', help='Skip analysis')
    parser.add_argument('--csv-only', action='store_true', help='Only regenerate CSVs')
    
    args = parser.parse_args()
    
    if args.csv_only:
        generate_csv_files()
    else:
        run_daily_pipeline(args.days)
