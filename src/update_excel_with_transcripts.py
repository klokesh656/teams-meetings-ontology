"""
Update Master Excel with newly transcribed recordings.
Scans the transcripts folder for VTT files that are not in the Excel file,
runs AI analysis on them, and adds them to the master Excel.
"""
import os
import json
import re
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv
import pandas as pd
from openai import AzureOpenAI

load_dotenv()

# Directories
TRANSCRIPTS_DIR = Path('transcripts')
OUTPUT_DIR = Path('output')
OUTPUT_DIR.mkdir(exist_ok=True)

# Azure OpenAI Configuration
AZURE_OPENAI_ENDPOINT = os.getenv('AZURE_OPENAI_ENDPOINT', '')
AZURE_OPENAI_KEY = os.getenv('AZURE_OPENAI_KEY', '')
AZURE_OPENAI_DEPLOYMENT = os.getenv('AZURE_OPENAI_DEPLOYMENT', 'gpt-4')
AZURE_OPENAI_API_VERSION = os.getenv('AZURE_OPENAI_API_VERSION', '2024-12-01-preview')

# Analysis prompt
ANALYSIS_PROMPT = """You are an expert meeting analyst. Analyze the following meeting transcript and provide structured evaluation.

## TRANSCRIPT:
{transcript}

## ANALYSIS REQUIRED:

### 1. SENTIMENT SCORE (Client Mood)
- Range: 0 to 100 (0 = very negative, 50 = neutral, 100 = very positive)

### 2. CHURN RISK SCORE
- Range: 0 to 100 (0 = no risk, 100 = high risk of leaving)

### 3. OPPORTUNITY/UPSELL POTENTIAL
- Range: 0 to 100 (0 = no opportunity, 100 = strong upsell potential)

### 4. EXECUTION RELIABILITY
- Range: 0 to 100 (0 = poor execution, 100 = excellent execution)

### 5. OPERATIONAL COMPLEXITY/WORKLOAD
- Range: 0 to 100 (0 = simple/light, 100 = very complex/heavy)

### 6. EVENTS DETECTED
Identify any of these event types:
- Complaint, Delay, Decision, Scope change, VA performance issue
- Payment issue, Positive feedback, New idea, Process confusion, Feature request

### 7. KEY SUMMARY
Provide a 2-3 sentence summary.

## OUTPUT FORMAT (JSON only, no markdown):
{{
    "sentiment_score": <0-100>,
    "churn_risk_score": <0-100>,
    "opportunity_score": <0-100>,
    "execution_reliability_score": <0-100>,
    "operational_complexity_score": <0-100>,
    "events": ["event1", "event2"],
    "summary": "Brief meeting summary...",
    "key_concerns": ["concern1", "concern2"],
    "key_positives": ["positive1", "positive2"],
    "action_items": ["action1", "action2"]
}}
"""


class ExcelUpdater:
    def __init__(self):
        # Use the latest analyzed Excel or create new master
        self.excel_path = self._find_latest_excel()
        self.output_path = OUTPUT_DIR / f'meeting_transcripts_master_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        self.openai_client = None
    
    def _find_latest_excel(self):
        """Find the most recent Excel file to use as base"""
        excel_files = list(OUTPUT_DIR.glob('*.xlsx'))
        if not excel_files:
            return OUTPUT_DIR / 'meeting_transcripts_master.xlsx'
        
        # Prefer analyzed file, then latest by modification time
        analyzed = [f for f in excel_files if 'analyzed' in f.name.lower()]
        if analyzed:
            return max(analyzed, key=lambda x: x.stat().st_mtime)
        
        return max(excel_files, key=lambda x: x.stat().st_mtime)
        
    def connect_openai(self):
        """Initialize Azure OpenAI client"""
        try:
            self.openai_client = AzureOpenAI(
                azure_endpoint=AZURE_OPENAI_ENDPOINT,
                api_key=AZURE_OPENAI_KEY,
                api_version=AZURE_OPENAI_API_VERSION
            )
            print("✅ Connected to Azure OpenAI")
            return True
        except Exception as e:
            print(f"❌ Failed to connect to Azure OpenAI: {e}")
            return False
    
    def load_excel(self):
        """Load existing Excel or create new DataFrame"""
        if self.excel_path.exists():
            df = pd.read_excel(self.excel_path)
            print(f"✅ Loaded existing Excel: {self.excel_path.name}")
            print(f"   Rows: {len(df)}")
        else:
            df = pd.DataFrame(columns=[
                'Meeting Subject', 'Date', 'Transcript File', 'Blob URL',
                'Sentiment Score', 'Churn Risk', 'Opportunity Score',
                'Execution Reliability', 'Operational Complexity',
                'Events', 'Summary', 'Key Concerns', 'Key Positives',
                'Action Items', 'Analyzed At', 'Source'
            ])
            print("📄 Created new Excel DataFrame")
        return df
    
    def get_existing_files(self, df):
        """Get set of transcript files already in Excel"""
        if 'Transcript File' in df.columns:
            return set(df['Transcript File'].dropna().tolist())
        return set()
    
    def find_new_transcripts(self, existing_files):
        """Find VTT files not yet in Excel"""
        new_files = []
        for vtt_file in TRANSCRIPTS_DIR.glob('*.vtt'):
            if vtt_file.name not in existing_files:
                new_files.append(vtt_file)
        return sorted(new_files, key=lambda x: x.stat().st_mtime, reverse=True)
    
    def extract_meeting_info(self, filename):
        """Extract meeting date and subject from filename"""
        # Pattern: 20251204_131321_Interview - Jon Jevi Dela Cruz.vtt
        match = re.match(r'(\d{8})_(\d{6})_(.+)\.vtt', filename)
        if match:
            date_str = match.group(1)
            subject = match.group(3)
            # Format date
            try:
                date = datetime.strptime(date_str, '%Y%m%d').strftime('%Y-%m-%d')
            except:
                date = date_str
            return date, subject
        return '', filename.replace('.vtt', '')
    
    def read_transcript(self, filepath):
        """Read VTT file content"""
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                return f.read()
        except Exception as e:
            print(f"❌ Error reading {filepath}: {e}")
            return None
    
    def analyze_transcript(self, content, filename):
        """Analyze transcript using Azure OpenAI"""
        if not self.openai_client:
            return {}
        
        # Truncate if too long (GPT-4 context limit)
        if len(content) > 100000:
            content = content[:50000] + "\n...[TRUNCATED]...\n" + content[-50000:]
        
        try:
            prompt = ANALYSIS_PROMPT.format(transcript=content)
            
            response = self.openai_client.chat.completions.create(
                model=AZURE_OPENAI_DEPLOYMENT,
                messages=[
                    {"role": "system", "content": "You are an expert meeting analyst. Always respond with valid JSON only."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.3,
                max_tokens=2000
            )
            
            result_text = response.choices[0].message.content.strip()
            
            # Clean up markdown if present
            if result_text.startswith('```'):
                result_text = re.sub(r'^```(?:json)?\n?', '', result_text)
                result_text = re.sub(r'\n?```$', '', result_text)
            
            return json.loads(result_text)
        except json.JSONDecodeError as e:
            print(f"  ⚠️ JSON parse error for {filename}: {e}")
            return {}
        except Exception as e:
            print(f"  ❌ Analysis error for {filename}: {e}")
            return {}
    
    def process_transcript(self, filepath, df):
        """Process a single transcript file"""
        filename = filepath.name
        date, subject = self.extract_meeting_info(filename)
        
        print(f"\n📝 Processing: {filename}")
        
        # Read content
        content = self.read_transcript(filepath)
        if not content:
            return None
        
        print(f"  📄 Read {len(content)} characters")
        
        # Analyze with AI
        print(f"  🤖 Analyzing with Azure OpenAI...")
        analysis = self.analyze_transcript(content, filename)
        
        if analysis:
            print(f"  ✅ Analysis complete - Sentiment: {analysis.get('sentiment_score', 'N/A')}")
        else:
            print(f"  ⚠️ No analysis results")
        
        # Create row
        new_row = {
            'Meeting Subject': subject,
            'Date': date,
            'Transcript File': filename,
            'Blob URL': '',  # Will be filled by upload script
            'Sentiment Score': analysis.get('sentiment_score', ''),
            'Churn Risk': analysis.get('churn_risk_score', ''),
            'Opportunity Score': analysis.get('opportunity_score', ''),
            'Execution Reliability': analysis.get('execution_reliability_score', ''),
            'Operational Complexity': analysis.get('operational_complexity_score', ''),
            'Events': ', '.join(analysis.get('events', [])),
            'Summary': analysis.get('summary', ''),
            'Key Concerns': ', '.join(analysis.get('key_concerns', [])),
            'Key Positives': ', '.join(analysis.get('key_positives', [])),
            'Action Items': ', '.join(analysis.get('action_items', [])),
            'Analyzed At': datetime.now().isoformat(),
            'Source': 'Speech-to-Text'  # Mark as transcribed from recording
        }
        
        return new_row
    
    def run(self, max_count=None, analyze=True):
        """Main update process"""
        print("="*60)
        print("UPDATING MASTER EXCEL WITH NEW TRANSCRIPTS")
        print("="*60)
        print(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # Connect to OpenAI if analyzing
        if analyze:
            if not self.connect_openai():
                print("⚠️ Continuing without AI analysis")
                analyze = False
        
        # Load Excel
        df = self.load_excel()
        
        # Find new transcripts
        existing_files = self.get_existing_files(df)
        print(f"📊 Existing transcripts in Excel: {len(existing_files)}")
        
        new_transcripts = self.find_new_transcripts(existing_files)
        print(f"🆕 New transcripts found: {len(new_transcripts)}")
        
        if not new_transcripts:
            print("\n✅ Excel is up to date - no new transcripts to add")
            return
        
        # Limit count if specified
        if max_count and len(new_transcripts) > max_count:
            new_transcripts = new_transcripts[:max_count]
            print(f"⏳ Processing first {max_count} transcripts")
        
        # Process each transcript
        processed = 0
        errors = 0
        
        for filepath in new_transcripts:
            try:
                if analyze:
                    new_row = self.process_transcript(filepath, df)
                else:
                    # Just add metadata without analysis
                    filename = filepath.name
                    date, subject = self.extract_meeting_info(filename)
                    new_row = {
                        'Meeting Subject': subject,
                        'Date': date,
                        'Transcript File': filename,
                        'Source': 'Speech-to-Text'
                    }
                
                if new_row:
                    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                    processed += 1
            except Exception as e:
                print(f"❌ Error processing {filepath.name}: {e}")
                errors += 1
        
        # Save Excel
        df.to_excel(self.output_path, index=False)
        print(f"\n{'='*60}")
        print("SUMMARY")
        print("="*60)
        print(f"✅ Processed: {processed}")
        print(f"❌ Errors: {errors}")
        print(f"📊 Total rows in Excel: {len(df)}")
        print(f"💾 Saved to: {self.output_path}")
        
        # Also save as latest
        latest_path = OUTPUT_DIR / 'meeting_transcripts_latest_analyzed.xlsx'
        df.to_excel(latest_path, index=False)
        print(f"💾 Also saved to: {latest_path.name}")


def main():
    import argparse
    parser = argparse.ArgumentParser(description='Update Excel with new transcripts')
    parser.add_argument('--max', type=int, help='Maximum number of transcripts to process')
    parser.add_argument('--no-analyze', action='store_true', help='Skip AI analysis')
    parser.add_argument('--status', action='store_true', help='Show status only')
    args = parser.parse_args()
    
    updater = ExcelUpdater()
    
    if args.status:
        df = updater.load_excel()
        existing = updater.get_existing_files(df)
        new_files = updater.find_new_transcripts(existing)
        print(f"\n📊 Status:")
        print(f"   In Excel: {len(existing)}")
        print(f"   New files: {len(new_files)}")
        if new_files:
            print(f"\n   Recent new files:")
            for f in new_files[:10]:
                print(f"   - {f.name}")
    else:
        updater.run(max_count=args.max, analyze=not args.no_analyze)


if __name__ == '__main__':
    main()
