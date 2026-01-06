"""
Daily Sync Script - Automated Transcript Download & AI Analysis
================================================================
This script:
1. Connects to Microsoft Graph API
2. Downloads new transcripts (not already in local folder)
3. Uploads new transcripts to Azure Blob Storage
4. Runs AI analysis on new transcripts
5. Updates the master Excel file
6. Uploads updated Excel to Azure Blob Storage

Can be run manually or scheduled via Windows Task Scheduler / cron
"""

import os
import sys
import json
import time
import requests
from datetime import datetime, timedelta
from pathlib import Path
from azure.identity import ClientSecretCredential
from azure.storage.blob import BlobServiceClient, generate_blob_sas, BlobSasPermissions
from dotenv import load_dotenv
import pandas as pd
from openai import AzureOpenAI
import re
import logging

# Setup logging (use ASCII-safe format for Windows console)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/daily_sync.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# Configure console handler to handle encoding errors gracefully
for handler in logger.handlers:
    if isinstance(handler, logging.StreamHandler) and handler.stream == sys.stdout:
        handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

# Load environment variables
load_dotenv()

# Configuration
TENANT_ID = os.getenv('AZURE_TENANT_ID', '187b2af6-1bfb-490a-85dd-b720fe3d31bc')
CLIENT_ID = os.getenv('AZURE_CLIENT_ID')
CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET')

# User IDs for transcript access (add more as needed)
USER_IDS = [
    '81835016-79d5-4a15-91b1-c104e2cd9adb',  # HR account
]

# Azure Blob Storage
BLOB_CONNECTION_STRING = os.getenv('AZURE_STORAGE_CONNECTION_STRING')
BLOB_ACCOUNT_NAME = 'aidevelopement'
BLOB_ACCOUNT_KEY = os.getenv('AZURE_STORAGE_KEY')
TRANSCRIPTS_CONTAINER = 'transcripts'
REPORTS_CONTAINER = 'reports'

# Azure OpenAI
AZURE_OPENAI_ENDPOINT = os.getenv('AZURE_OPENAI_ENDPOINT', 'https://foundary-1-lokesh.cognitiveservices.azure.com/')
AZURE_OPENAI_KEY = os.getenv('AZURE_OPENAI_KEY')
AZURE_OPENAI_DEPLOYMENT = os.getenv('AZURE_OPENAI_DEPLOYMENT', 'gpt-4.1')

# Paths
TRANSCRIPTS_DIR = Path('transcripts')
OUTPUT_DIR = Path('output')
LOGS_DIR = Path('logs')

# Ensure directories exist
TRANSCRIPTS_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
LOGS_DIR.mkdir(exist_ok=True)


class DailySync:
    def __init__(self):
        self.credential = ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
        self.token = None
        self.headers = None
        self.blob_service = None
        self.openai_client = None
        self.new_transcripts = []
        self.stats = {
            'transcripts_found': 0,
            'new_downloaded': 0,
            'uploaded_to_blob': 0,
            'analyzed': 0,
            'errors': 0
        }
    
    def authenticate(self):
        """Authenticate with Microsoft Graph API"""
        logger.info("Authenticating with Microsoft Graph API...")
        self.token = self.credential.get_token('https://graph.microsoft.com/.default').token
        self.headers = {
            'Authorization': f'Bearer {self.token}',
            'Content-Type': 'application/json'
        }
        logger.info("✅ Graph API authentication successful")
    
    def connect_blob_storage(self):
        """Connect to Azure Blob Storage"""
        logger.info("Connecting to Azure Blob Storage...")
        if BLOB_CONNECTION_STRING:
            self.blob_service = BlobServiceClient.from_connection_string(BLOB_CONNECTION_STRING)
        else:
            self.blob_service = BlobServiceClient(
                account_url=f"https://{BLOB_ACCOUNT_NAME}.blob.core.windows.net",
                credential=BLOB_ACCOUNT_KEY
            )
        logger.info("✅ Blob Storage connected")
    
    def connect_openai(self):
        """Connect to Azure OpenAI"""
        logger.info("Connecting to Azure OpenAI...")
        self.openai_client = AzureOpenAI(
            azure_endpoint=AZURE_OPENAI_ENDPOINT,
            api_key=AZURE_OPENAI_KEY,
            api_version="2024-12-01-preview"
        )
        logger.info("✅ Azure OpenAI connected")
    
    def get_existing_transcripts(self):
        """Get list of already downloaded transcript IDs"""
        existing = set()
        for f in TRANSCRIPTS_DIR.glob('*.vtt'):
            existing.add(f.stem)
        return existing
    
    def get_all_transcripts(self):
        """Fetch all transcripts from Graph API"""
        all_transcripts = []
        
        for user_id in USER_IDS:
            logger.info(f"Fetching transcripts for user: {user_id}")
            url = f"https://graph.microsoft.com/beta/users/{user_id}/onlineMeetings/getAllTranscripts(meetingOrganizerUserId='{user_id}')"
            
            while url:
                resp = requests.get(url, headers=self.headers)
                if resp.status_code == 200:
                    data = resp.json()
                    transcripts = data.get('value', [])
                    all_transcripts.extend(transcripts)
                    url = data.get('@odata.nextLink')
                    time.sleep(0.5)  # Rate limiting
                else:
                    logger.error(f"Error fetching transcripts: {resp.status_code} - {resp.text[:200]}")
                    break
        
        self.stats['transcripts_found'] = len(all_transcripts)
        logger.info(f"Found {len(all_transcripts)} total transcripts")
        return all_transcripts
    
    def get_meeting_details(self, user_id, meeting_id):
        """Get meeting details including subject"""
        url = f"https://graph.microsoft.com/beta/users/{user_id}/onlineMeetings/{meeting_id}"
        resp = requests.get(url, headers=self.headers)
        if resp.status_code == 200:
            return resp.json()
        return {}
    
    def download_transcript(self, user_id, meeting_id, transcript_id, meeting_subject, created_date, max_retries=3):
        """Download a single transcript with retry logic"""
        url = f"https://graph.microsoft.com/beta/users/{user_id}/onlineMeetings/{meeting_id}/transcripts/{transcript_id}/content?$format=text/vtt"
        
        for attempt in range(max_retries):
            try:
                resp = requests.get(url, headers=self.headers, timeout=60)
                if resp.status_code == 200:
                    # Create filename
                    safe_subject = re.sub(r'[<>:"/\\|?*]', '', meeting_subject or 'Unknown Meeting')[:50]
                    date_str = created_date[:10].replace('-', '') if created_date else 'unknown'
                    time_str = created_date[11:19].replace(':', '') if created_date and len(created_date) > 11 else ''
                    filename = f"{date_str}_{time_str}_{safe_subject}.vtt"
                    filepath = TRANSCRIPTS_DIR / filename
                    
                    with open(filepath, 'w', encoding='utf-8') as f:
                        f.write(resp.text)
                    
                    logger.info(f"  [OK] Downloaded: {filename}")
                    return filepath
                elif resp.status_code == 404:
                    logger.warning(f"  [SKIP] Transcript not available (404) - may be expired or deleted")
                    return None
                else:
                    logger.error(f"  [ERROR] Failed to download transcript: {resp.status_code}")
                    if attempt < max_retries - 1:
                        logger.info(f"  [RETRY] Attempt {attempt + 2}/{max_retries}...")
                        import time
                        time.sleep(2)
                        continue
                    return None
            except requests.exceptions.Timeout:
                logger.error(f"  [TIMEOUT] Request timed out (attempt {attempt + 1}/{max_retries})")
                if attempt < max_retries - 1:
                    import time
                    time.sleep(3)
                    continue
                return None
            except Exception as e:
                logger.error(f"  [ERROR] Exception downloading transcript: {e}")
                if attempt < max_retries - 1:
                    import time
                    time.sleep(2)
                    continue
                return None
        return None
    
    def download_new_transcripts(self):
        """Download only new transcripts"""
        logger.info("="*60)
        logger.info("DOWNLOADING NEW TRANSCRIPTS")
        logger.info("="*60)
        
        existing = self.get_existing_transcripts()
        all_transcripts = self.get_all_transcripts()
        
        for t in all_transcripts:
            transcript_id = t.get('id', '')
            meeting_id = t.get('meetingId', '')
            created = t.get('createdDateTime', '')
            
            # Create a simple hash for checking
            simple_id = f"{created[:10]}_{meeting_id[:20]}" if created else meeting_id[:30]
            
            # Check if we already have a transcript from this meeting/date
            already_have = False
            for existing_file in TRANSCRIPTS_DIR.glob('*.vtt'):
                if created[:10].replace('-', '') in existing_file.name:
                    already_have = True
                    break
            
            if not already_have:
                # Get meeting details
                user_id = USER_IDS[0]  # Primary user
                meeting_details = self.get_meeting_details(user_id, meeting_id)
                subject = meeting_details.get('subject', 'Unknown Meeting')
                
                filepath = self.download_transcript(user_id, meeting_id, transcript_id, subject, created)
                if filepath:
                    self.new_transcripts.append({
                        'filepath': filepath,
                        'subject': subject,
                        'created': created,
                        'meeting_id': meeting_id
                    })
                    self.stats['new_downloaded'] += 1
                
                time.sleep(1)  # Rate limiting
        
        logger.info(f"Downloaded {self.stats['new_downloaded']} new transcripts")
    
    def upload_to_blob(self, filepath, container_name):
        """Upload a file to Azure Blob Storage"""
        try:
            container_client = self.blob_service.get_container_client(container_name)
            blob_name = filepath.name
            blob_client = container_client.get_blob_client(blob_name)
            
            with open(filepath, 'rb') as f:
                blob_client.upload_blob(f, overwrite=True)
            
            # Generate SAS URL
            sas_token = generate_blob_sas(
                account_name=BLOB_ACCOUNT_NAME,
                container_name=container_name,
                blob_name=blob_name,
                account_key=BLOB_ACCOUNT_KEY,
                permission=BlobSasPermissions(read=True),
                expiry=datetime.utcnow() + timedelta(days=365)
            )
            blob_url = f"https://{BLOB_ACCOUNT_NAME}.blob.core.windows.net/{container_name}/{blob_name}?{sas_token}"
            
            return blob_url
        except Exception as e:
            logger.error(f"Error uploading {filepath}: {e}")
            return None
    
    def upload_new_transcripts(self):
        """Upload new transcripts to Azure Blob Storage"""
        logger.info("="*60)
        logger.info("UPLOADING NEW TRANSCRIPTS TO BLOB STORAGE")
        logger.info("="*60)
        
        for t in self.new_transcripts:
            filepath = t['filepath']
            blob_url = self.upload_to_blob(filepath, TRANSCRIPTS_CONTAINER)
            if blob_url:
                t['blob_url'] = blob_url
                self.stats['uploaded_to_blob'] += 1
                logger.info(f"  ✅ Uploaded: {filepath.name}")
        
        logger.info(f"Uploaded {self.stats['uploaded_to_blob']} transcripts to blob storage")
    
    def analyze_transcript(self, content):
        """Analyze transcript with Azure OpenAI"""
        system_prompt = """You are an expert meeting analyst. Analyze the following meeting transcript and provide a structured evaluation.

Return a JSON object with these exact fields:
{
    "sentiment_score": <0-100, overall meeting positivity>,
    "churn_risk": <0-100, likelihood of client/employee leaving>,
    "opportunity_score": <0-100, potential for growth/upsell>,
    "execution_reliability": <0-100, team's ability to deliver>,
    "operational_complexity": <0-100, complexity of discussed operations>,
    "events": [<list of key events: "complaint", "praise", "decision", "escalation", "concern", etc>],
    "summary": "<2-3 sentence summary of the meeting>",
    "key_concerns": [<list of main concerns raised>],
    "key_positives": [<list of positive points>],
    "action_items": [<list of action items identified>]
}

Be objective and base scores on actual transcript content."""

        try:
            response = self.openai_client.chat.completions.create(
                model=AZURE_OPENAI_DEPLOYMENT,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": f"Analyze this meeting transcript:\n\n{content[:15000]}"}
                ],
                temperature=0.3,
                response_format={"type": "json_object"}
            )
            
            result = json.loads(response.choices[0].message.content)
            return result
        except Exception as e:
            logger.error(f"Error analyzing transcript: {e}")
            return None
    
    def analyze_new_transcripts(self):
        """Analyze all new transcripts with AI"""
        logger.info("="*60)
        logger.info("ANALYZING NEW TRANSCRIPTS WITH AI")
        logger.info("="*60)
        
        for t in self.new_transcripts:
            filepath = t['filepath']
            logger.info(f"Analyzing: {filepath.name}")
            
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read()
            
            analysis = self.analyze_transcript(content)
            if analysis:
                t['analysis'] = analysis
                self.stats['analyzed'] += 1
                logger.info(f"  ✅ Sentiment: {analysis.get('sentiment_score')}, Churn: {analysis.get('churn_risk')}")
            else:
                self.stats['errors'] += 1
            
            time.sleep(2)  # Rate limiting for OpenAI
        
        logger.info(f"Analyzed {self.stats['analyzed']} transcripts")
    
    def update_master_excel(self):
        """Update master Excel with new transcripts and analysis"""
        logger.info("="*60)
        logger.info("UPDATING MASTER EXCEL")
        logger.info("="*60)
        
        excel_path = OUTPUT_DIR / 'meeting_transcripts_master.xlsx'
        
        # Load existing or create new
        if excel_path.exists():
            df = pd.read_excel(excel_path)
        else:
            df = pd.DataFrame(columns=[
                'Meeting Subject', 'Date', 'Transcript File', 'Blob URL',
                'Sentiment Score', 'Churn Risk', 'Opportunity Score',
                'Execution Reliability', 'Operational Complexity',
                'Events', 'Summary', 'Key Concerns', 'Key Positives',
                'Action Items', 'Analyzed At'
            ])
        
        # Add new transcripts
        for t in self.new_transcripts:
            analysis = t.get('analysis', {})
            new_row = {
                'Meeting Subject': t.get('subject', 'Unknown'),
                'Date': t.get('created', '')[:10] if t.get('created') else '',
                'Transcript File': t['filepath'].name,
                'Blob URL': t.get('blob_url', ''),
                'Sentiment Score': analysis.get('sentiment_score', ''),
                'Churn Risk': analysis.get('churn_risk', ''),
                'Opportunity Score': analysis.get('opportunity_score', ''),
                'Execution Reliability': analysis.get('execution_reliability', ''),
                'Operational Complexity': analysis.get('operational_complexity', ''),
                'Events': ', '.join(analysis.get('events', [])),
                'Summary': analysis.get('summary', ''),
                'Key Concerns': ', '.join(analysis.get('key_concerns', [])),
                'Key Positives': ', '.join(analysis.get('key_positives', [])),
                'Action Items': ', '.join(analysis.get('action_items', [])),
                'Analyzed At': datetime.now().isoformat()
            }
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        
        # Save Excel
        df.to_excel(excel_path, index=False)
        logger.info(f"✅ Updated master Excel: {excel_path}")
        
        # Upload to blob
        blob_url = self.upload_to_blob(excel_path, REPORTS_CONTAINER)
        if blob_url:
            logger.info(f"✅ Uploaded Excel to blob storage")
        
        return excel_path
    
    def generate_daily_report(self):
        """Generate a daily summary report"""
        logger.info("="*60)
        logger.info("DAILY SYNC SUMMARY")
        logger.info("="*60)
        
        report = f"""
Daily Sync Report - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
{'='*50}

Transcripts Found:     {self.stats['transcripts_found']}
New Downloaded:        {self.stats['new_downloaded']}
Uploaded to Blob:      {self.stats['uploaded_to_blob']}
Analyzed with AI:      {self.stats['analyzed']}
Errors:                {self.stats['errors']}

"""
        
        if self.new_transcripts:
            report += "New Transcripts Processed:\n"
            report += "-"*40 + "\n"
            for t in self.new_transcripts:
                analysis = t.get('analysis', {})
                report += f"  • {t.get('subject', 'Unknown')}\n"
                report += f"    Sentiment: {analysis.get('sentiment_score', 'N/A')}, "
                report += f"Churn Risk: {analysis.get('churn_risk', 'N/A')}\n"
        else:
            report += "No new transcripts found.\n"
        
        # High-risk alerts
        high_risk = [t for t in self.new_transcripts 
                     if t.get('analysis', {}).get('churn_risk', 0) >= 50]
        if high_risk:
            report += "\n⚠️ HIGH CHURN RISK ALERTS:\n"
            report += "-"*40 + "\n"
            for t in high_risk:
                report += f"  🚨 {t.get('subject')}: {t['analysis']['churn_risk']}% churn risk\n"
        
        logger.info(report)
        
        # Save report
        report_path = LOGS_DIR / f"daily_report_{datetime.now().strftime('%Y%m%d')}.txt"
        with open(report_path, 'w') as f:
            f.write(report)
        
        return report
    
    def run(self):
        """Run the full daily sync process"""
        start_time = datetime.now()
        logger.info("="*60)
        logger.info(f"STARTING DAILY SYNC - {start_time}")
        logger.info("="*60)
        
        try:
            # Step 1: Authenticate
            self.authenticate()
            
            # Step 2: Connect to services
            self.connect_blob_storage()
            self.connect_openai()
            
            # Step 3: Download new transcripts
            self.download_new_transcripts()
            
            if self.new_transcripts:
                # Step 4: Upload to blob storage
                self.upload_new_transcripts()
                
                # Step 5: Analyze with AI
                self.analyze_new_transcripts()
                
                # Step 6: Update master Excel
                self.update_master_excel()
            
            # Step 7: Generate report
            self.generate_daily_report()
            
            end_time = datetime.now()
            duration = (end_time - start_time).total_seconds()
            logger.info(f"✅ Daily sync completed in {duration:.1f} seconds")
            
        except Exception as e:
            logger.error(f"❌ Daily sync failed: {e}")
            raise


def main():
    """Main entry point"""
    sync = DailySync()
    sync.run()


if __name__ == '__main__':
    main()
