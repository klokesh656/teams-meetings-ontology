"""
Analyze meeting transcripts using Azure OpenAI.
Reads transcripts from Azure Blob Storage and evaluates them using the 5 universal scores.

Scores:
1. Sentiment Score (Client Mood): -1 to +1 or 0 to 100
2. Churn Risk Score: 0 to 100
3. Opportunity/Upsell Potential: 0 to 100
4. Execution Reliability: 0 to 100
5. Operational Complexity/Workload: 0 to 100

Events Extracted:
- Complaint, Delay, Decision, Scope change, VA performance issue
- Payment issue, Positive feedback, New idea, Process confusion, Feature request
"""
import os
import json
import requests
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv
from azure.storage.blob import BlobServiceClient
from openai import AzureOpenAI

load_dotenv()

# Azure credentials
STORAGE_CONNECTION_STRING = os.getenv('AZURE_STORAGE_CONNECTION_STRING')

# Azure OpenAI Configuration
AZURE_OPENAI_ENDPOINT = os.getenv('AZURE_OPENAI_ENDPOINT', '')
AZURE_OPENAI_KEY = os.getenv('AZURE_OPENAI_KEY', '')
AZURE_OPENAI_DEPLOYMENT = os.getenv('AZURE_OPENAI_DEPLOYMENT', 'gpt-4')
AZURE_OPENAI_API_VERSION = os.getenv('AZURE_OPENAI_API_VERSION', '2024-12-01-preview')

# Containers
TRANSCRIPTS_CONTAINER = 'transcripts'
REPORTS_CONTAINER = 'reports'

# Analysis prompt template
ANALYSIS_PROMPT = """You are an expert meeting analyst. Analyze the following meeting transcript and provide structured evaluation.

## TRANSCRIPT:
{transcript}

## ANALYSIS REQUIRED:

### 1. SENTIMENT SCORE (Client Mood)
- Range: 0 to 100 (0 = very negative, 50 = neutral, 100 = very positive)
- Captures tone, frustration, positivity
- Consider: word choice, complaints vs compliments, overall mood

### 2. CHURN RISK SCORE
- Range: 0 to 100 (0 = no risk, 100 = high risk of leaving)
- Based on: phrases indicating dissatisfaction, complaints, delays, tone shifts
- Look for: "looking elsewhere", "not happy", "disappointed", "considering alternatives"

### 3. OPPORTUNITY/UPSELL POTENTIAL
- Range: 0 to 100 (0 = no opportunity, 100 = strong upsell potential)
- Based on: hints of expansion, new needs, interest in additional support
- Look for: "need more help", "growing", "new projects", "additional resources"

### 4. EXECUTION RELIABILITY
- Range: 0 to 100 (0 = poor execution, 100 = excellent execution)
- How well the team is meeting expectations
- Derived from: complaints vs praises about work quality, deadlines, communication

### 5. OPERATIONAL COMPLEXITY/WORKLOAD
- Range: 0 to 100 (0 = simple/light, 100 = very complex/heavy)
- How much is on the table
- Based on: tasks mentioned, deadlines, chaos signals, multiple priorities

### 6. EVENTS DETECTED
Identify any of these event types mentioned in the meeting:
- Complaint: Client expressing dissatisfaction
- Delay: Project/task delays mentioned
- Decision: Important decisions made
- Scope change: Changes to project scope
- VA performance issue: Issues with virtual assistant performance
- Payment issue: Payment or billing concerns
- Positive feedback: Compliments or satisfaction expressed
- New idea: New suggestions or proposals
- Process confusion: Confusion about processes or procedures
- Feature request: Requests for new features or capabilities

### 7. KEY SUMMARY
Provide a 2-3 sentence summary of the meeting highlighting the most important points.

## OUTPUT FORMAT (JSON):
{{
    "sentiment_score": <0-100>,
    "churn_risk_score": <0-100>,
    "opportunity_score": <0-100>,
    "execution_reliability_score": <0-100>,
    "operational_complexity_score": <0-100>,
    "events": ["event1", "event2", ...],
    "summary": "Brief meeting summary...",
    "key_concerns": ["concern1", "concern2"],
    "key_positives": ["positive1", "positive2"],
    "action_items": ["action1", "action2"]
}}

Respond ONLY with valid JSON, no additional text.
"""


class TranscriptAnalyzer:
    def __init__(self):
        self.blob_service = BlobServiceClient.from_connection_string(STORAGE_CONNECTION_STRING)
        
        # Initialize Azure OpenAI client
        if AZURE_OPENAI_ENDPOINT and AZURE_OPENAI_KEY:
            self.openai_client = AzureOpenAI(
                api_key=AZURE_OPENAI_KEY,
                api_version=AZURE_OPENAI_API_VERSION,
                azure_endpoint=AZURE_OPENAI_ENDPOINT
            )
            self.ai_available = True
            print("✅ Azure OpenAI connected")
        else:
            self.openai_client = None
            self.ai_available = False
            print("⚠️ Azure OpenAI not configured")
    
    def read_transcript_from_blob(self, blob_url: str) -> str:
        """Read transcript content from blob storage URL"""
        # Extract blob name from URL
        # URL format: https://account.blob.core.windows.net/container/blobname?sas
        try:
            # Parse the URL
            from urllib.parse import urlparse, unquote
            parsed = urlparse(blob_url)
            path_parts = parsed.path.split('/', 2)
            if len(path_parts) >= 3:
                container = path_parts[1]
                blob_name = unquote(path_parts[2])
                
                container_client = self.blob_service.get_container_client(container)
                blob_client = container_client.get_blob_client(blob_name)
                
                content = blob_client.download_blob().readall().decode('utf-8')
                return content
        except Exception as e:
            print(f"Error reading blob: {e}")
        return ""
    
    def analyze_transcript(self, transcript_content: str) -> dict:
        """Analyze transcript using Azure OpenAI"""
        if not self.ai_available:
            return self._get_placeholder_analysis()
        
        # Truncate very long transcripts to fit in context
        max_chars = 30000  # Leave room for prompt
        if len(transcript_content) > max_chars:
            transcript_content = transcript_content[:max_chars] + "\n... [TRUNCATED]"
        
        prompt = ANALYSIS_PROMPT.format(transcript=transcript_content)
        
        try:
            response = self.openai_client.chat.completions.create(
                model=AZURE_OPENAI_DEPLOYMENT,
                messages=[
                    {"role": "system", "content": "You are a meeting analyst. Respond only with valid JSON."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.3,
                max_tokens=1500
            )
            
            result_text = response.choices[0].message.content.strip()
            
            # Parse JSON response
            # Handle potential markdown code blocks
            if result_text.startswith('```'):
                result_text = result_text.split('```')[1]
                if result_text.startswith('json'):
                    result_text = result_text[4:]
            
            return json.loads(result_text)
            
        except Exception as e:
            print(f"Error analyzing transcript: {e}")
            return self._get_placeholder_analysis()
    
    def _get_placeholder_analysis(self) -> dict:
        """Return placeholder analysis when AI is not available"""
        return {
            "sentiment_score": None,
            "churn_risk_score": None,
            "opportunity_score": None,
            "execution_reliability_score": None,
            "operational_complexity_score": None,
            "events": [],
            "summary": "Analysis pending - Azure OpenAI not configured",
            "key_concerns": [],
            "key_positives": [],
            "action_items": []
        }
    
    def analyze_from_master_excel(self, excel_path: str, output_path: str = None):
        """Read master Excel, analyze transcripts, and update with results"""
        print("=" * 70)
        print("ANALYZING MEETING TRANSCRIPTS")
        print("=" * 70)
        
        # Read the master Excel
        df = pd.read_excel(excel_path)
        print(f"Loaded {len(df)} meetings from {excel_path}")
        
        # Add analysis columns if not present
        analysis_columns = [
            'Sentiment Score', 'Churn Risk', 'Opportunity Score',
            'Execution Reliability', 'Operational Complexity',
            'Events', 'Summary', 'Key Concerns', 'Key Positives', 'Action Items',
            'Analyzed At'
        ]
        for col in analysis_columns:
            if col not in df.columns:
                df[col] = None
        
        # Analyze each transcript
        analyzed = 0
        for idx, row in df.iterrows():
            blob_url = row.get('Blob Storage Link', '')
            already_analyzed = pd.notna(row.get('Analyzed At'))
            
            if not blob_url or already_analyzed:
                continue
            
            subject = row.get('Subject', 'Unknown')
            print(f"\n[{idx+1}/{len(df)}] Analyzing: {subject[:50]}...")
            
            # Read transcript
            transcript = self.read_transcript_from_blob(blob_url)
            if not transcript:
                print(f"   ⚠️ Could not read transcript")
                continue
            
            # Analyze
            analysis = self.analyze_transcript(transcript)
            
            # Update DataFrame
            df.at[idx, 'Sentiment Score'] = analysis.get('sentiment_score')
            df.at[idx, 'Churn Risk'] = analysis.get('churn_risk_score')
            df.at[idx, 'Opportunity Score'] = analysis.get('opportunity_score')
            df.at[idx, 'Execution Reliability'] = analysis.get('execution_reliability_score')
            df.at[idx, 'Operational Complexity'] = analysis.get('operational_complexity_score')
            df.at[idx, 'Events'] = ', '.join(analysis.get('events', []))
            df.at[idx, 'Summary'] = analysis.get('summary', '')
            df.at[idx, 'Key Concerns'] = ', '.join(analysis.get('key_concerns', []))
            df.at[idx, 'Key Positives'] = ', '.join(analysis.get('key_positives', []))
            df.at[idx, 'Action Items'] = ', '.join(analysis.get('action_items', []))
            df.at[idx, 'Analyzed At'] = datetime.now().isoformat()
            
            print(f"   ✅ Sentiment: {analysis.get('sentiment_score')}, Churn: {analysis.get('churn_risk_score')}, Opportunity: {analysis.get('opportunity_score')}")
            analyzed += 1
        
        # Save updated Excel
        if output_path is None:
            output_path = excel_path.replace('.xlsx', '_analyzed.xlsx')
        
        df.to_excel(output_path, index=False)
        print(f"\n{'='*70}")
        print(f"SUMMARY: Analyzed {analyzed} transcripts")
        print(f"Saved to: {output_path}")
        
        return df


def main():
    """Main function to demonstrate the analyzer"""
    analyzer = TranscriptAnalyzer()
    
    # Check for master Excel in output folder - use latest file
    master_excel = 'output/meeting_transcripts_latest.xlsx'
    
    if os.path.exists(master_excel):
        print(f"Found master Excel: {master_excel}")
        analyzer.analyze_from_master_excel(master_excel)
    else:
        print(f"Master Excel not found at {master_excel}")
        print("Run upload_and_create_master.py first to create the master Excel")
        
        # Demo: analyze a single transcript file
        transcripts_dir = 'transcripts'
        if os.path.exists(transcripts_dir):
            vtt_files = [f for f in os.listdir(transcripts_dir) if f.endswith('.vtt')]
            if vtt_files:
                sample_file = os.path.join(transcripts_dir, vtt_files[0])
                print(f"\nDemo: Analyzing sample transcript: {vtt_files[0]}")
                
                with open(sample_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                analysis = analyzer.analyze_transcript(content)
                print(json.dumps(analysis, indent=2))


if __name__ == '__main__':
    main()
