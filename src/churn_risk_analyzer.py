"""
Churn Risk Analysis System
==========================
Analyzes Outlier Discussion transcripts using Azure OpenAI to:
1. Extract churn risk signals and patterns
2. Build a structured dataset with KPIs
3. Create searchable index for Copilot Agent
4. Generate Power BI compatible data

Output formats:
- JSON for Azure AI Search indexing
- CSV for Power BI
- Structured churn risk checklist
"""

import os
import re
import json
import csv
import hashlib
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv
from openai import AzureOpenAI

load_dotenv()

# Configuration
AZURE_OPENAI_ENDPOINT = os.getenv('AZURE_OPENAI_ENDPOINT')
AZURE_OPENAI_KEY = os.getenv('AZURE_OPENAI_KEY')
AZURE_OPENAI_DEPLOYMENT = os.getenv('AZURE_OPENAI_DEPLOYMENT', 'gpt-4.1')

TRANSCRIPTS_DIR = Path('outliers meetings trasncripts')
OUTPUT_DIR = Path('output')

# Initialize Azure OpenAI client
client = AzureOpenAI(
    azure_endpoint=AZURE_OPENAI_ENDPOINT,
    api_key=AZURE_OPENAI_KEY,
    api_version="2024-02-15-preview"
)


# Churn Risk Checklist Schema
CHURN_RISK_CHECKLIST = {
    "va_risk_signals": [
        {"id": "VA001", "signal": "Non-responsive to integration team", "severity": "high", "category": "communication"},
        {"id": "VA002", "signal": "Task scope expanded without acknowledgment", "severity": "high", "category": "workload"},
        {"id": "VA003", "signal": "Mentioned feeling overwhelmed", "severity": "medium", "category": "emotional"},
        {"id": "VA004", "signal": "Casual mention of resignation/leaving", "severity": "critical", "category": "retention"},
        {"id": "VA005", "signal": "Attendance issues (tardiness, absences)", "severity": "medium", "category": "reliability"},
        {"id": "VA006", "signal": "Technical/equipment problems unresolved", "severity": "medium", "category": "technical"},
        {"id": "VA007", "signal": "Not visible to client (lack of SOD/EOD)", "severity": "high", "category": "visibility"},
        {"id": "VA008", "signal": "Over-reliance on AI without sense-checking", "severity": "medium", "category": "performance"},
        {"id": "VA009", "signal": "Requested compensation review", "severity": "medium", "category": "compensation"},
        {"id": "VA010", "signal": "First 30 days - low engagement", "severity": "high", "category": "onboarding"},
        {"id": "VA011", "signal": "MIA (Missing in Action) pattern", "severity": "critical", "category": "reliability"},
        {"id": "VA012", "signal": "Coaching sessions with no progress", "severity": "medium", "category": "development"},
    ],
    "client_risk_signals": [
        {"id": "CL001", "signal": "Slow/no response to OA communications", "severity": "high", "category": "engagement"},
        {"id": "CL002", "signal": "Unilateral task changes without discussion", "severity": "medium", "category": "scope"},
        {"id": "CL003", "signal": "Negative feedback delivered indirectly", "severity": "high", "category": "communication"},
        {"id": "CL004", "signal": "Hired directly from another agency", "severity": "critical", "category": "competition"},
        {"id": "CL005", "signal": "Internal layoffs or restructuring", "severity": "high", "category": "stability"},
        {"id": "CL006", "signal": "Surprise cancellation after positive feedback", "severity": "critical", "category": "trust"},
        {"id": "CL007", "signal": "Information hidden from client (felt blind-sided)", "severity": "high", "category": "transparency"},
        {"id": "CL008", "signal": "Part-time client resistant to full-time", "severity": "low", "category": "growth"},
        {"id": "CL009", "signal": "Multiple VA replacements requested", "severity": "high", "category": "satisfaction"},
        {"id": "CL010", "signal": "No check-in for 30+ days", "severity": "medium", "category": "engagement"},
    ],
    "relationship_health_signals": [
        {"id": "RH001", "signal": "OA seen as vendor not partner", "severity": "high", "category": "perception"},
        {"id": "RH002", "signal": "Feedback loop broken (no response to surveys)", "severity": "medium", "category": "communication"},
        {"id": "RH003", "signal": "VA told client about issue before OA", "severity": "high", "category": "information_flow"},
        {"id": "RH004", "signal": "Client competing for VA time (second job)", "severity": "medium", "category": "loyalty"},
        {"id": "RH005", "signal": "Trust deficit - client verifying VA work", "severity": "medium", "category": "trust"},
    ]
}


def parse_transcript_date(filename):
    """Extract date from transcript filename"""
    # Format: "Outlier Discussion Transcripts DD month YYYY.txt"
    match = re.search(r'(\d{1,2})\s+(january|february|march|april|may|june|july|august|september|october|november|december)\s+(\d{4})', 
                      filename.lower())
    if match:
        day, month, year = match.groups()
        month_num = {
            'january': 1, 'february': 2, 'march': 3, 'april': 4,
            'may': 5, 'june': 6, 'july': 7, 'august': 8,
            'september': 9, 'october': 10, 'november': 11, 'december': 12
        }[month]
        return f"{year}-{month_num:02d}-{int(day):02d}"
    return None


def analyze_transcript_with_ai(transcript_text, transcript_date, filename):
    """Use Azure OpenAI to analyze transcript for churn risk signals"""
    
    prompt = f"""Analyze this Outlier Discussion meeting transcript from {transcript_date} and extract structured data.

TRANSCRIPT:
{transcript_text[:15000]}  # Limit to avoid token limits

EXTRACT THE FOLLOWING IN JSON FORMAT:

1. "meeting_summary": Brief 2-3 sentence summary of the meeting
2. "participants": List of participant names mentioned
3. "vas_discussed": Array of VAs discussed with structure:
   - "name": VA name
   - "client": Client company name
   - "status": "green", "yellow", "orange", or "red"
   - "risk_signals": Array of risk signal IDs from checklist (VA001-VA012)
   - "issues": Brief description of any issues
   - "actions_taken": What actions were discussed/taken
   
4. "clients_discussed": Array of clients discussed with structure:
   - "name": Client company name
   - "risk_signals": Array of risk signal IDs (CL001-CL010)
   - "health_status": "healthy", "at_risk", "critical"
   - "issues": Brief description of concerns
   
5. "key_insights": Array of important insights or patterns mentioned
6. "action_items": Array of action items discussed
7. "churn_risks_identified": Array of specific churn risks with:
   - "type": "va" or "client"
   - "name": Name of VA or client
   - "risk_level": "low", "medium", "high", "critical"
   - "signal": Description of the risk signal
   - "recommended_action": What should be done

8. "kpis_mentioned": Any metrics or KPIs discussed (response times, attendance, etc.)
9. "best_practices_shared": Any best practices or successful strategies mentioned

RISK SIGNAL REFERENCE:
VA Signals: VA001=Non-responsive, VA002=Scope creep, VA003=Overwhelmed, VA004=Resignation mention, 
VA005=Attendance issues, VA006=Tech problems, VA007=Not visible, VA008=AI over-reliance, 
VA009=Comp review request, VA010=Poor onboarding, VA011=MIA pattern, VA012=No coaching progress

Client Signals: CL001=Slow response, CL002=Unilateral changes, CL003=Indirect feedback, 
CL004=Direct hiring, CL005=Internal instability, CL006=Surprise cancellation, CL007=Felt blind-sided,
CL008=Resistant to growth, CL009=Multiple replacements, CL010=No recent check-in

Return ONLY valid JSON, no markdown formatting."""

    try:
        response = client.chat.completions.create(
            model=AZURE_OPENAI_DEPLOYMENT,
            messages=[
                {"role": "system", "content": "You are an expert HR analyst specializing in VA (Virtual Assistant) management and churn risk analysis. Extract structured data from meeting transcripts accurately."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3,
            max_tokens=4000
        )
        
        result_text = response.choices[0].message.content.strip()
        
        # Clean up JSON if wrapped in markdown
        if result_text.startswith('```'):
            result_text = re.sub(r'^```json?\n?', '', result_text)
            result_text = re.sub(r'\n?```$', '', result_text)
        
        # Fix common JSON issues
        result_text = re.sub(r',\s*}', '}', result_text)  # Remove trailing commas
        result_text = re.sub(r',\s*]', ']', result_text)  # Remove trailing commas in arrays
        
        return json.loads(result_text)
        
    except Exception as e:
        print(f"   ❌ AI Analysis Error: {e}")
        return None


def create_document_id(text):
    """Create unique document ID for search indexing"""
    return hashlib.md5(text.encode()).hexdigest()[:16]


def process_all_transcripts():
    """Process all outlier discussion transcripts"""
    print("\n" + "="*70)
    print("🔍 CHURN RISK ANALYSIS SYSTEM")
    print("="*70)
    
    # Data structures for output
    all_analyses = []
    va_risk_dataset = []
    client_risk_dataset = []
    meeting_insights = []
    action_items_all = []
    
    # Find all transcript files
    transcript_files = list(TRANSCRIPTS_DIR.glob('*.txt'))
    print(f"\n📁 Found {len(transcript_files)} transcript files\n")
    
    for idx, file_path in enumerate(transcript_files, 1):
        filename = file_path.name
        transcript_date = parse_transcript_date(filename) or "unknown"
        
        print(f"[{idx}/{len(transcript_files)}] Analyzing: {filename}")
        print(f"   📅 Date: {transcript_date}")
        
        # Read transcript
        with open(file_path, 'r', encoding='utf-8') as f:
            transcript_text = f.read()
        
        # Analyze with AI
        analysis = analyze_transcript_with_ai(transcript_text, transcript_date, filename)
        
        if analysis:
            # Add metadata
            analysis['transcript_date'] = transcript_date
            analysis['filename'] = filename
            analysis['document_id'] = create_document_id(filename + transcript_date)
            analysis['processed_at'] = datetime.now().isoformat()
            
            all_analyses.append(analysis)
            
            # Extract VA risk records
            for va in analysis.get('vas_discussed', []):
                va_record = {
                    'document_id': create_document_id(f"{va.get('name', 'unknown')}_{transcript_date}"),
                    'meeting_date': transcript_date,
                    'va_name': va.get('name', ''),
                    'client_name': va.get('client', ''),
                    'status': va.get('status', 'unknown'),
                    'risk_signals': ','.join(va.get('risk_signals', [])),
                    'risk_signal_count': len(va.get('risk_signals', [])),
                    'issues': va.get('issues', ''),
                    'actions_taken': va.get('actions_taken', ''),
                    'source_file': filename
                }
                va_risk_dataset.append(va_record)
            
            # Extract client risk records
            for client in analysis.get('clients_discussed', []):
                client_record = {
                    'document_id': create_document_id(f"{client.get('name', 'unknown')}_{transcript_date}"),
                    'meeting_date': transcript_date,
                    'client_name': client.get('name', ''),
                    'health_status': client.get('health_status', 'unknown'),
                    'risk_signals': ','.join(client.get('risk_signals', [])),
                    'risk_signal_count': len(client.get('risk_signals', [])),
                    'issues': client.get('issues', ''),
                    'source_file': filename
                }
                client_risk_dataset.append(client_record)
            
            # Extract insights
            for insight in analysis.get('key_insights', []):
                meeting_insights.append({
                    'document_id': create_document_id(f"insight_{insight[:30]}_{transcript_date}"),
                    'meeting_date': transcript_date,
                    'insight': insight,
                    'source_file': filename
                })
            
            # Extract action items
            for action in analysis.get('action_items', []):
                action_items_all.append({
                    'document_id': create_document_id(f"action_{action[:30]}_{transcript_date}"),
                    'meeting_date': transcript_date,
                    'action_item': action,
                    'source_file': filename
                })
            
            print(f"   ✅ Extracted: {len(analysis.get('vas_discussed', []))} VAs, {len(analysis.get('clients_discussed', []))} clients")
            print(f"   📊 Churn risks: {len(analysis.get('churn_risks_identified', []))}")
        else:
            print(f"   ⚠️ Could not analyze transcript")
    
    return all_analyses, va_risk_dataset, client_risk_dataset, meeting_insights, action_items_all


def create_copilot_search_index(all_analyses, va_dataset, client_dataset, insights):
    """Create JSON documents suitable for Azure AI Search indexing"""
    
    search_documents = []
    
    # Document Type 1: Meeting Summaries (for general queries)
    for analysis in all_analyses:
        doc = {
            "@search.action": "upload",
            "id": f"meeting_{analysis['document_id']}",
            "document_type": "meeting_summary",
            "title": f"Outlier Discussion - {analysis['transcript_date']}",
            "date": analysis['transcript_date'],
            "content": analysis.get('meeting_summary', ''),
            "participants": analysis.get('participants', []),
            "key_insights": analysis.get('key_insights', []),
            "action_items": analysis.get('action_items', []),
            "vas_count": len(analysis.get('vas_discussed', [])),
            "clients_count": len(analysis.get('clients_discussed', [])),
            "churn_risks_count": len(analysis.get('churn_risks_identified', [])),
            "best_practices": analysis.get('best_practices_shared', []),
            "searchable_text": json.dumps(analysis, ensure_ascii=False)
        }
        search_documents.append(doc)
    
    # Document Type 2: VA Risk Profiles
    for va in va_dataset:
        doc = {
            "@search.action": "upload",
            "id": f"va_{va['document_id']}",
            "document_type": "va_risk_profile",
            "title": f"VA Risk: {va['va_name']} at {va['client_name']}",
            "date": va['meeting_date'],
            "va_name": va['va_name'],
            "client_name": va['client_name'],
            "status": va['status'],
            "risk_signals": va['risk_signals'],
            "risk_level": "high" if va['risk_signal_count'] >= 2 else "medium" if va['risk_signal_count'] == 1 else "low",
            "issues": va['issues'],
            "actions_taken": va['actions_taken'],
            "searchable_text": f"{va['va_name']} {va['client_name']} {va['status']} {va['issues']} {va['risk_signals']}"
        }
        search_documents.append(doc)
    
    # Document Type 3: Client Health Profiles
    for client in client_dataset:
        doc = {
            "@search.action": "upload",
            "id": f"client_{client['document_id']}",
            "document_type": "client_health_profile",
            "title": f"Client Health: {client['client_name']}",
            "date": client['meeting_date'],
            "client_name": client['client_name'],
            "health_status": client['health_status'],
            "risk_signals": client['risk_signals'],
            "issues": client['issues'],
            "searchable_text": f"{client['client_name']} {client['health_status']} {client['issues']} {client['risk_signals']}"
        }
        search_documents.append(doc)
    
    # Document Type 4: Best Practices & Insights
    for insight in insights:
        doc = {
            "@search.action": "upload",
            "id": f"insight_{insight['document_id']}",
            "document_type": "insight",
            "title": f"Insight from {insight['meeting_date']}",
            "date": insight['meeting_date'],
            "content": insight['insight'],
            "searchable_text": insight['insight']
        }
        search_documents.append(doc)
    
    # Document Type 5: Churn Risk Checklist Reference
    for category, signals in CHURN_RISK_CHECKLIST.items():
        for signal in signals:
            doc = {
                "@search.action": "upload",
                "id": f"checklist_{signal['id']}",
                "document_type": "churn_checklist",
                "title": f"Churn Signal: {signal['signal']}",
                "signal_id": signal['id'],
                "signal_description": signal['signal'],
                "severity": signal['severity'],
                "category": signal['category'],
                "signal_type": category.replace('_signals', ''),
                "searchable_text": f"{signal['id']} {signal['signal']} {signal['severity']} {signal['category']}"
            }
            search_documents.append(doc)
    
    return search_documents


def create_power_bi_datasets(va_dataset, client_dataset, all_analyses):
    """Create CSV files suitable for Power BI import"""
    
    OUTPUT_DIR.mkdir(exist_ok=True)
    
    # 1. VA Risk Dataset
    va_csv_path = OUTPUT_DIR / 'churn_risk_va_dataset.csv'
    if va_dataset:
        with open(va_csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=va_dataset[0].keys())
            writer.writeheader()
            writer.writerows(va_dataset)
        print(f"   📊 VA Dataset: {va_csv_path}")
    
    # 2. Client Risk Dataset
    client_csv_path = OUTPUT_DIR / 'churn_risk_client_dataset.csv'
    if client_dataset:
        with open(client_csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=client_dataset[0].keys())
            writer.writeheader()
            writer.writerows(client_dataset)
        print(f"   📊 Client Dataset: {client_csv_path}")
    
    # 3. KPI Summary Dataset
    kpi_records = []
    for analysis in all_analyses:
        kpi_records.append({
            'meeting_date': analysis['transcript_date'],
            'total_vas_discussed': len(analysis.get('vas_discussed', [])),
            'total_clients_discussed': len(analysis.get('clients_discussed', [])),
            'churn_risks_identified': len(analysis.get('churn_risks_identified', [])),
            'action_items_count': len(analysis.get('action_items', [])),
            'insights_count': len(analysis.get('key_insights', [])),
            'red_status_count': sum(1 for va in analysis.get('vas_discussed', []) if va.get('status') == 'red'),
            'yellow_status_count': sum(1 for va in analysis.get('vas_discussed', []) if va.get('status') in ['yellow', 'orange']),
            'green_status_count': sum(1 for va in analysis.get('vas_discussed', []) if va.get('status') == 'green'),
        })
    
    kpi_csv_path = OUTPUT_DIR / 'churn_risk_kpi_summary.csv'
    if kpi_records:
        with open(kpi_csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=kpi_records[0].keys())
            writer.writeheader()
            writer.writerows(kpi_records)
        print(f"   📊 KPI Summary: {kpi_csv_path}")
    
    # 4. Churn Risk Checklist Reference
    checklist_records = []
    for category, signals in CHURN_RISK_CHECKLIST.items():
        for signal in signals:
            checklist_records.append({
                'signal_id': signal['id'],
                'signal_type': category.replace('_signals', ''),
                'signal_description': signal['signal'],
                'severity': signal['severity'],
                'category': signal['category']
            })
    
    checklist_csv_path = OUTPUT_DIR / 'churn_risk_checklist_reference.csv'
    with open(checklist_csv_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=checklist_records[0].keys())
        writer.writeheader()
        writer.writerows(checklist_records)
    print(f"   📊 Checklist Reference: {checklist_csv_path}")
    
    return va_csv_path, client_csv_path, kpi_csv_path, checklist_csv_path


def save_all_outputs(all_analyses, va_dataset, client_dataset, insights, actions, search_documents):
    """Save all output files"""
    
    OUTPUT_DIR.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    
    print("\n" + "="*70)
    print("💾 SAVING OUTPUT FILES")
    print("="*70)
    
    # 1. Full Analysis JSON (for reference)
    analysis_path = OUTPUT_DIR / f'churn_analysis_full_{timestamp}.json'
    with open(analysis_path, 'w', encoding='utf-8') as f:
        json.dump(all_analyses, f, indent=2, ensure_ascii=False)
    print(f"\n   📄 Full Analysis: {analysis_path}")
    
    # 2. Copilot Search Index JSON
    search_index_path = OUTPUT_DIR / 'copilot_churn_risk_index.json'
    with open(search_index_path, 'w', encoding='utf-8') as f:
        json.dump(search_documents, f, indent=2, ensure_ascii=False)
    print(f"   🔍 Search Index: {search_index_path}")
    
    # 3. Power BI Datasets
    print("\n   📊 Power BI Datasets:")
    create_power_bi_datasets(va_dataset, client_dataset, all_analyses)
    
    # 4. Churn Risk Checklist JSON
    checklist_path = OUTPUT_DIR / 'churn_risk_checklist.json'
    with open(checklist_path, 'w', encoding='utf-8') as f:
        json.dump(CHURN_RISK_CHECKLIST, f, indent=2, ensure_ascii=False)
    print(f"\n   ✅ Churn Checklist: {checklist_path}")
    
    # 5. Latest link
    latest_path = OUTPUT_DIR / 'churn_analysis_latest.json'
    with open(latest_path, 'w', encoding='utf-8') as f:
        json.dump({
            'generated_at': datetime.now().isoformat(),
            'transcript_count': len(all_analyses),
            'total_va_records': len(va_dataset),
            'total_client_records': len(client_dataset),
            'total_insights': len(insights),
            'total_action_items': len(actions),
            'search_documents_count': len(search_documents),
            'analyses': all_analyses
        }, f, indent=2, ensure_ascii=False)
    print(f"   🔗 Latest Analysis: {latest_path}")
    
    return analysis_path, search_index_path


def print_summary(all_analyses, va_dataset, client_dataset, insights, actions):
    """Print analysis summary"""
    
    print("\n" + "="*70)
    print("📊 ANALYSIS SUMMARY")
    print("="*70)
    
    print(f"\n   📁 Transcripts Analyzed: {len(all_analyses)}")
    print(f"   👥 VA Records Created: {len(va_dataset)}")
    print(f"   🏢 Client Records Created: {len(client_dataset)}")
    print(f"   💡 Insights Extracted: {len(insights)}")
    print(f"   ✅ Action Items Found: {len(actions)}")
    
    # Risk distribution
    if va_dataset:
        status_counts = {}
        for va in va_dataset:
            status = va.get('status', 'unknown')
            status_counts[status] = status_counts.get(status, 0) + 1
        
        print(f"\n   📈 VA Status Distribution:")
        for status, count in sorted(status_counts.items()):
            emoji = {'green': '🟢', 'yellow': '🟡', 'orange': '🟠', 'red': '🔴'}.get(status, '⚪')
            print(f"      {emoji} {status.capitalize()}: {count}")
    
    # Top risk signals
    if va_dataset:
        signal_counts = {}
        for va in va_dataset:
            signals = va.get('risk_signals', '').split(',')
            for signal in signals:
                signal = signal.strip()
                if signal:
                    signal_counts[signal] = signal_counts.get(signal, 0) + 1
        
        if signal_counts:
            print(f"\n   🚨 Top Risk Signals Detected:")
            for signal, count in sorted(signal_counts.items(), key=lambda x: -x[1])[:5]:
                # Get signal description
                desc = next((s['signal'] for cat in CHURN_RISK_CHECKLIST.values() 
                           for s in cat if s['id'] == signal), signal)
                print(f"      {signal}: {desc[:40]}... ({count}x)")


def main():
    """Main execution"""
    print("\n" + "="*70)
    print("🚀 STARTING CHURN RISK ANALYSIS SYSTEM")
    print(f"   Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*70)
    
    # Process all transcripts
    all_analyses, va_dataset, client_dataset, insights, actions = process_all_transcripts()
    
    if not all_analyses:
        print("\n❌ No transcripts were successfully analyzed")
        return
    
    # Create search index documents
    print("\n" + "="*70)
    print("🔍 CREATING COPILOT SEARCH INDEX")
    print("="*70)
    search_documents = create_copilot_search_index(all_analyses, va_dataset, client_dataset, insights)
    print(f"   ✅ Created {len(search_documents)} searchable documents")
    
    # Save all outputs
    save_all_outputs(all_analyses, va_dataset, client_dataset, insights, actions, search_documents)
    
    # Print summary
    print_summary(all_analyses, va_dataset, client_dataset, insights, actions)
    
    print("\n" + "="*70)
    print("✅ CHURN RISK ANALYSIS COMPLETE")
    print("="*70)
    print("\n📋 OUTPUT FILES READY FOR:")
    print("   • Azure AI Search: copilot_churn_risk_index.json")
    print("   • Power BI: churn_risk_*.csv files")
    print("   • Reference: churn_risk_checklist.json")
    print("\n")


if __name__ == '__main__':
    main()
