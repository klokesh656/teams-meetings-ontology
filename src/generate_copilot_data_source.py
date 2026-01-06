#!/usr/bin/env python3
"""
Copilot Agent Data Source Generator

Generates a unified JSON data source optimized for Azure AI Search
to power the Copilot Agent with comprehensive meeting and churn risk data.
"""

import json
import os
import hashlib
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Any, Optional

# Paths
OUTPUT_DIR = Path('output')
TRANSCRIPTS_DIR = Path('transcripts')
OUTLIERS_DIR = Path('outliers meetings trasncripts')

def generate_id(content: str) -> str:
    """Generate a unique ID from content hash."""
    return hashlib.md5(content.encode()).hexdigest()[:16]

def parse_date(date_str: str) -> Optional[str]:
    """Parse date string to ISO format."""
    if not date_str:
        return None
    try:
        # Try various formats
        for fmt in ['%Y-%m-%d', '%Y-%m-%dT%H:%M:%S', '%Y-%m-%d %H:%M:%S', '%Y%m%d']:
            try:
                dt = datetime.strptime(date_str.split('.')[0].split('Z')[0], fmt)
                return dt.isoformat() + 'Z'
            except ValueError:
                continue
        return None
    except Exception:
        return None

def extract_risk_level_score(risk_level: str) -> int:
    """Convert risk level to numeric score."""
    risk_map = {
        'critical': 100,
        'high': 80,
        'medium': 50,
        'low': 20,
        'none': 0,
        'green': 10,
        'yellow': 40,
        'orange': 60,
        'red': 80,
        'at_risk': 70
    }
    return risk_map.get(risk_level.lower() if risk_level else '', 0)

def load_json_file(filepath: Path) -> Optional[Dict]:
    """Load JSON file safely."""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"Error loading {filepath}: {e}")
        return None

def process_meetings_knowledge_base() -> List[Dict]:
    """Process copilot knowledge base meetings."""
    documents = []
    kb_file = OUTPUT_DIR / 'copilot_knowledge_base_latest.json'
    
    if not kb_file.exists():
        print(f"Warning: {kb_file} not found")
        return documents
    
    data = load_json_file(kb_file)
    if not data or 'meetings' not in data:
        return documents
    
    for meeting in data['meetings']:
        doc = {
            'id': f"meeting_{meeting.get('id', generate_id(meeting.get('subject', '')))}",
            'document_type': 'meeting_transcript',
            'entity_type': 'meeting',
            'title': meeting.get('subject', 'Unknown Meeting'),
            'content': meeting.get('searchable_content', ''),
            'summary': meeting.get('summary', ''),
            'organizer': meeting.get('organizer', ''),
            'date_string': meeting.get('date', ''),
            'date': parse_date(meeting.get('date', '')),
            'sentiment_score': meeting.get('sentiment_score', 0),
            'risk_score': meeting.get('churn_risk', 0),
            'risk_level': 'high' if meeting.get('churn_risk', 0) > 60 else 'medium' if meeting.get('churn_risk', 0) > 30 else 'low',
            'opportunity_score': meeting.get('opportunity_score', 0),
            'events_detected': [e.strip() for e in meeting.get('events', '').split(',') if e.strip()],
            'key_concerns': meeting.get('key_concerns', ''),
            'key_positives': meeting.get('key_positives', ''),
            'action_items': [a.strip() for a in meeting.get('action_items', '').split(',') if a.strip()],
            'transcript_file': meeting.get('transcript_file', ''),
            'blob_url': meeting.get('blob_url', ''),
            'meeting_type': extract_meeting_type(meeting.get('subject', '')),
            'tags': extract_tags(meeting),
            'created_at': meeting.get('analyzed_at', datetime.now().isoformat() + 'Z')
        }
        
        # Extract VA and client names from subject
        va_name, client_name = extract_names_from_subject(meeting.get('subject', ''))
        doc['va_name'] = va_name
        doc['client_name'] = client_name
        
        documents.append(doc)
    
    print(f"Processed {len(documents)} meetings from knowledge base")
    return documents

def process_churn_analysis() -> List[Dict]:
    """Process churn analysis data."""
    documents = []
    churn_file = OUTPUT_DIR / 'churn_analysis_latest.json'
    
    if not churn_file.exists():
        print(f"Warning: {churn_file} not found")
        return documents
    
    data = load_json_file(churn_file)
    if not data or 'analyses' not in data:
        return documents
    
    for analysis in data['analyses']:
        # Create document for the meeting analysis
        doc_id = analysis.get('document_id', generate_id(analysis.get('filename', '')))
        
        meeting_doc = {
            'id': f"churn_{doc_id}",
            'document_type': 'churn_analysis',
            'entity_type': 'meeting',
            'title': f"Outlier Discussion - {analysis.get('transcript_date', 'Unknown')}",
            'content': analysis.get('meeting_summary', ''),
            'summary': analysis.get('meeting_summary', ''),
            'participants': analysis.get('participants', []),
            'date_string': analysis.get('transcript_date', ''),
            'date': parse_date(analysis.get('transcript_date', '')),
            'key_insights': analysis.get('key_insights', []),
            'action_items': analysis.get('action_items', []),
            'best_practices': analysis.get('best_practices_shared', []),
            'kpis_mentioned': analysis.get('kpis_mentioned', []),
            'source_file': analysis.get('filename', ''),
            'meeting_type': 'outlier_discussion',
            'tags': ['outlier', 'churn_analysis', 'strategy_meeting'],
            'created_at': analysis.get('processed_at', datetime.now().isoformat() + 'Z')
        }
        documents.append(meeting_doc)
        
        # Create documents for each VA discussed
        for va in analysis.get('vas_discussed', []):
            va_doc = {
                'id': f"va_{doc_id}_{generate_id(va.get('name', ''))}",
                'document_type': 'va_status',
                'entity_type': 'va',
                'title': f"VA Status: {va.get('name', 'Unknown')}",
                'va_name': va.get('name', ''),
                'client_name': va.get('client', ''),
                'content': f"Status: {va.get('status', '')}. Issues: {va.get('issues', '')}. Actions: {va.get('actions_taken', '')}",
                'summary': va.get('issues', ''),
                'issues': va.get('issues', ''),
                'actions_taken': va.get('actions_taken', ''),
                'health_status': va.get('status', ''),
                'risk_signals': va.get('risk_signals', []),
                'risk_level': va.get('status', ''),
                'risk_score': extract_risk_level_score(va.get('status', '')),
                'date_string': analysis.get('transcript_date', ''),
                'date': parse_date(analysis.get('transcript_date', '')),
                'tags': ['va', 'status_update', va.get('status', '').lower()],
                'created_at': analysis.get('processed_at', datetime.now().isoformat() + 'Z')
            }
            documents.append(va_doc)
        
        # Create documents for each client discussed
        for client in analysis.get('clients_discussed', []):
            client_doc = {
                'id': f"client_{doc_id}_{generate_id(client.get('name', ''))}",
                'document_type': 'client_status',
                'entity_type': 'client',
                'title': f"Client Status: {client.get('name', 'Unknown')}",
                'client_name': client.get('name', ''),
                'content': f"Health: {client.get('health_status', '')}. Issues: {client.get('issues', '')}",
                'summary': client.get('issues', ''),
                'issues': client.get('issues', ''),
                'health_status': client.get('health_status', ''),
                'risk_signals': client.get('risk_signals', []),
                'risk_level': client.get('health_status', ''),
                'risk_score': extract_risk_level_score(client.get('health_status', '')),
                'date_string': analysis.get('transcript_date', ''),
                'date': parse_date(analysis.get('transcript_date', '')),
                'tags': ['client', 'status_update', client.get('health_status', '').lower().replace(' ', '_')],
                'created_at': analysis.get('processed_at', datetime.now().isoformat() + 'Z')
            }
            documents.append(client_doc)
        
        # Create documents for churn risks identified
        for risk in analysis.get('churn_risks_identified', []):
            risk_doc = {
                'id': f"risk_{doc_id}_{generate_id(risk.get('name', '') + risk.get('signal', ''))}",
                'document_type': 'churn_risk',
                'entity_type': risk.get('type', 'unknown'),
                'title': f"Churn Risk: {risk.get('name', 'Unknown')}",
                'va_name': risk.get('name', '') if risk.get('type') == 'va' else '',
                'client_name': risk.get('name', '') if risk.get('type') == 'client' else '',
                'content': f"Risk Signal: {risk.get('signal', '')}. Recommended: {risk.get('recommended_action', '')}",
                'summary': risk.get('signal', ''),
                'issues': risk.get('signal', ''),
                'risk_level': risk.get('risk_level', ''),
                'risk_score': extract_risk_level_score(risk.get('risk_level', '')),
                'recommended_action': risk.get('recommended_action', ''),
                'date_string': analysis.get('transcript_date', ''),
                'date': parse_date(analysis.get('transcript_date', '')),
                'tags': ['churn_risk', risk.get('risk_level', '').lower(), risk.get('type', '')],
                'created_at': analysis.get('processed_at', datetime.now().isoformat() + 'Z')
            }
            documents.append(risk_doc)
    
    print(f"Processed {len(documents)} churn analysis documents")
    return documents

def process_batch_analysis() -> List[Dict]:
    """Process batch analysis results."""
    documents = []
    
    # Find the latest batch analysis file
    batch_files = list(OUTPUT_DIR.glob('batch_analysis_results_*.json'))
    if not batch_files:
        print("Warning: No batch analysis files found")
        return documents
    
    latest_file = max(batch_files, key=lambda p: p.stat().st_mtime)
    data = load_json_file(latest_file)
    
    if not data or 'processed' not in data:
        return documents
    
    for item in data['processed']:
        va_name = item.get('va_name', '')
        item_date = item.get('date', '')
        doc = {
            'id': f"batch_{generate_id(va_name + '_' + item_date)}",
            'document_type': 'va_checkin_analysis',
            'entity_type': 'va',
            'title': f"Check-in Analysis: {item.get('va_name', 'Unknown')} - {item.get('date', '')}",
            'va_name': item.get('va_name', ''),
            'client_name': item.get('client_name', ''),
            'date_string': item.get('date', ''),
            'date': parse_date(item.get('date', '')),
            'risk_level': item.get('risk_level', ''),
            'risk_score': extract_risk_level_score(item.get('risk_level', '')),
            'content': f"VA: {item.get('va_name', '')} with client {item.get('client_name', '')}. Risk level: {item.get('risk_level', '')}. Signals: {item.get('signals_count', 0)}. Suggestions: {item.get('suggestions_count', 0)}",
            'summary': f"Risk: {item.get('risk_level', '')} | Signals: {item.get('signals_count', 0)} | Suggestions: {item.get('suggestions_count', 0)}",
            'tags': ['checkin', 'analysis', item.get('risk_level', '').lower()],
            'created_at': datetime.now().isoformat() + 'Z'
        }
        documents.append(doc)
    
    print(f"Processed {len(documents)} batch analysis documents")
    return documents

def process_suggestions() -> List[Dict]:
    """Process pending suggestions."""
    documents = []
    suggestions_file = OUTPUT_DIR / 'pending_suggestions.json'
    
    if not suggestions_file.exists():
        print(f"Warning: {suggestions_file} not found")
        return documents
    
    data = load_json_file(suggestions_file)
    if not data or 'suggestions' not in data:
        return documents
    
    for suggestion in data['suggestions']:
        doc = {
            'id': f"suggestion_{suggestion.get('suggestion_id', generate_id(suggestion.get('suggestion', '')))}",
            'document_type': 'suggestion',
            'entity_type': 'action_item',
            'title': f"Suggestion: {suggestion.get('va_name', 'Unknown')} - {suggestion.get('category', '')}",
            'va_name': suggestion.get('va_name', ''),
            'client_name': suggestion.get('client_name', ''),
            'content': suggestion.get('suggestion', ''),
            'summary': suggestion.get('suggestion', ''),
            'issues': suggestion.get('issue', ''),
            'rationale': suggestion.get('rationale', ''),
            'suggestion_category': suggestion.get('category', ''),
            'urgency': suggestion.get('urgency', ''),
            'status': suggestion.get('status', ''),
            'date_string': suggestion.get('meeting_date', ''),
            'date': parse_date(suggestion.get('meeting_date', '')),
            'reviewed_by': suggestion.get('reviewed_by', ''),
            'stakeholder_notes': suggestion.get('stakeholder_notes', ''),
            'recommended_action': suggestion.get('final_solution', suggestion.get('suggestion', '')),
            'tags': ['suggestion', suggestion.get('category', '').lower(), suggestion.get('urgency', '').lower(), suggestion.get('status', '').lower()],
            'created_at': suggestion.get('created_at', datetime.now().isoformat() + 'Z')
        }
        documents.append(doc)
    
    print(f"Processed {len(documents)} suggestions")
    return documents

def process_approved_solutions() -> List[Dict]:
    """Process approved solutions."""
    documents = []
    solutions_file = OUTPUT_DIR / 'approved_solutions.json'
    
    if not solutions_file.exists():
        print(f"Warning: {solutions_file} not found")
        return documents
    
    data = load_json_file(solutions_file)
    if not data or 'solutions' not in data:
        return documents
    
    for solution in data.get('solutions', []):
        doc = {
            'id': f"solution_{generate_id(solution.get('suggestion', ''))}",
            'document_type': 'approved_solution',
            'entity_type': 'best_practice',
            'title': f"Approved Solution: {solution.get('va_name', 'Unknown')}",
            'va_name': solution.get('va_name', ''),
            'client_name': solution.get('client_name', ''),
            'content': solution.get('final_solution', solution.get('suggestion', '')),
            'summary': solution.get('suggestion', ''),
            'issues': solution.get('issue', ''),
            'rationale': solution.get('rationale', ''),
            'suggestion_category': solution.get('category', ''),
            'status': 'approved',
            'reviewed_by': solution.get('reviewed_by', ''),
            'stakeholder_notes': solution.get('stakeholder_notes', ''),
            'date_string': solution.get('meeting_date', ''),
            'date': parse_date(solution.get('meeting_date', '')),
            'tags': ['approved', 'solution', 'best_practice', solution.get('category', '').lower()],
            'created_at': solution.get('approved_at', datetime.now().isoformat() + 'Z')
        }
        documents.append(doc)
    
    print(f"Processed {len(documents)} approved solutions")
    return documents

def extract_meeting_type(subject: str) -> str:
    """Extract meeting type from subject."""
    subject_lower = subject.lower()
    if 'check-in' in subject_lower or 'checkin' in subject_lower:
        return 'check_in'
    elif 'orientation' in subject_lower or 'onboarding' in subject_lower:
        return 'onboarding'
    elif 'interview' in subject_lower:
        return 'interview'
    elif 'readiness' in subject_lower:
        return 'readiness_check'
    elif 'gtm' in subject_lower:
        return 'gtm_orientation'
    elif 'catch up' in subject_lower or 'catch-up' in subject_lower:
        return 'catch_up'
    elif 'performance' in subject_lower:
        return 'performance_review'
    elif 'outlier' in subject_lower:
        return 'outlier_discussion'
    else:
        return 'general'

def extract_names_from_subject(subject: str) -> tuple:
    """Extract VA and client names from meeting subject."""
    va_name = ''
    client_name = ''
    
    # Common patterns: "... x VA_NAME", "... x VA_NAME (CLIENT)"
    if ' x ' in subject.lower():
        parts = subject.split(' x ')
        if len(parts) > 1:
            name_part = parts[-1].strip()
            # Check for client in parentheses
            if '(' in name_part:
                va_name = name_part.split('(')[0].strip()
                client_name = name_part.split('(')[-1].replace(')', '').strip()
            else:
                va_name = name_part.split('-')[0].strip()
    
    return va_name, client_name

def extract_tags(meeting: Dict) -> List[str]:
    """Extract relevant tags from meeting data."""
    tags = []
    
    # Add meeting type tag
    meeting_type = extract_meeting_type(meeting.get('subject', ''))
    tags.append(meeting_type)
    
    # Add risk-related tags
    churn_risk = meeting.get('churn_risk', 0)
    if churn_risk > 60:
        tags.append('high_risk')
    elif churn_risk > 30:
        tags.append('medium_risk')
    else:
        tags.append('low_risk')
    
    # Add sentiment tags
    sentiment = meeting.get('sentiment_score', 50)
    if sentiment > 75:
        tags.append('positive_sentiment')
    elif sentiment < 40:
        tags.append('negative_sentiment')
    
    # Add event tags
    events = meeting.get('events', '')
    if 'complaint' in events.lower():
        tags.append('complaint')
    if 'positive' in events.lower():
        tags.append('positive_feedback')
    if 'delay' in events.lower():
        tags.append('delay_issue')
    if 'confusion' in events.lower():
        tags.append('process_confusion')
    
    return tags

def generate_data_source():
    """Generate the complete data source for Azure AI Search."""
    all_documents = []
    
    print("=" * 60)
    print("Copilot Agent Data Source Generator")
    print("=" * 60)
    
    # Process all data sources
    all_documents.extend(process_meetings_knowledge_base())
    all_documents.extend(process_churn_analysis())
    all_documents.extend(process_batch_analysis())
    all_documents.extend(process_suggestions())
    all_documents.extend(process_approved_solutions())
    
    # Create the output data structure
    output = {
        'metadata': {
            'generated_at': datetime.now().isoformat(),
            'total_documents': len(all_documents),
            'document_types': list(set(d['document_type'] for d in all_documents)),
            'entity_types': list(set(d['entity_type'] for d in all_documents)),
            'version': '2.0',
            'description': 'Comprehensive data source for Copilot Agent Azure AI Search index'
        },
        'statistics': {
            'meetings': len([d for d in all_documents if d['document_type'] == 'meeting_transcript']),
            'churn_analyses': len([d for d in all_documents if d['document_type'] == 'churn_analysis']),
            'va_status_records': len([d for d in all_documents if d['document_type'] == 'va_status']),
            'client_status_records': len([d for d in all_documents if d['document_type'] == 'client_status']),
            'churn_risks': len([d for d in all_documents if d['document_type'] == 'churn_risk']),
            'suggestions': len([d for d in all_documents if d['document_type'] == 'suggestion']),
            'approved_solutions': len([d for d in all_documents if d['document_type'] == 'approved_solution']),
            'va_checkin_analyses': len([d for d in all_documents if d['document_type'] == 'va_checkin_analysis'])
        },
        'documents': all_documents
    }
    
    # Save the data source
    output_file = OUTPUT_DIR / 'copilot_agent_data_source.json'
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(output, f, indent=2, ensure_ascii=False, default=str)
    
    print("\n" + "=" * 60)
    print(f"✅ Generated {len(all_documents)} documents")
    print(f"📁 Saved to: {output_file}")
    print("\nDocument Statistics:")
    for doc_type, count in output['statistics'].items():
        print(f"  - {doc_type}: {count}")
    print("=" * 60)
    
    return output

if __name__ == '__main__':
    generate_data_source()
