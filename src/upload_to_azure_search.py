#!/usr/bin/env python3
"""
Upload knowledge base to Azure AI Search for Copilot Agent integration.
Creates index and uploads meeting transcript data.
"""

import os
import json
import hashlib
from datetime import datetime
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Azure AI Search configuration
SEARCH_SERVICE_NAME = os.getenv("AZURE_SEARCH_SERVICE_NAME", "")
SEARCH_API_KEY = os.getenv("AZURE_SEARCH_API_KEY", "")
SEARCH_INDEX_NAME = os.getenv("AZURE_SEARCH_INDEX_NAME", "copilot-knowledge-base")

# Paths
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
OUTPUT_DIR = os.path.join(PROJECT_DIR, "output")


def find_latest_knowledge_base():
    """Find the latest knowledge base JSON file."""
    kb_files = [f for f in os.listdir(OUTPUT_DIR) if f.startswith("copilot_knowledge_base_") and f.endswith(".json")]
    if not kb_files:
        return None
    kb_files.sort(reverse=True)
    return os.path.join(OUTPUT_DIR, kb_files[0])


def load_knowledge_base(file_path):
    """Load the knowledge base JSON."""
    with open(file_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def generate_document_id(meeting):
    """Generate a unique document ID."""
    unique_str = f"{meeting.get('meeting_id', '')}{meeting.get('subject', '')}{meeting.get('date', '')}"
    return hashlib.md5(unique_str.encode()).hexdigest()


def parse_date(date_str):
    """Parse date string to ISO format for Azure Search."""
    if not date_str:
        return None
    
    # Try different date formats
    formats = [
        "%Y-%m-%d",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%dT%H:%M:%SZ",
        "%Y-%m-%d %H:%M:%S",
        "%m/%d/%Y",
        "%d/%m/%Y"
    ]
    
    for fmt in formats:
        try:
            dt = datetime.strptime(date_str, fmt)
            return dt.strftime("%Y-%m-%dT%H:%M:%SZ")
        except ValueError:
            continue
    
    # If it's already ISO format with timezone, return as-is
    if 'T' in str(date_str) and 'Z' in str(date_str):
        return date_str
    
    return None


def transform_meeting_to_document(meeting):
    """Transform a meeting record to an Azure Search document."""
    doc_id = generate_document_id(meeting)
    
    # Parse events detected - handle both string and list formats
    events = meeting.get("events_detected", [])
    if isinstance(events, str):
        events = [e.strip() for e in events.split(",") if e.strip()]
    elif not isinstance(events, list):
        events = []
    # Ensure all events are strings
    events = [str(e) for e in events if e]
    
    # Parse date
    date_value = parse_date(meeting.get("date", ""))
    
    # Parse sentiment score - ensure it's an integer
    sentiment = meeting.get("sentiment_score", 50)
    if sentiment is None:
        sentiment = 50
    elif isinstance(sentiment, str):
        try:
            sentiment = int(sentiment)
        except ValueError:
            sentiment = 50  # Default neutral
    else:
        try:
            sentiment = int(sentiment)
        except (TypeError, ValueError):
            sentiment = 50
    
    # Ensure all string fields are actually strings
    def safe_str(val):
        if val is None:
            return ""
        return str(val)
    
    return {
        "@search.action": "upload",
        "id": doc_id,
        "meeting_id": safe_str(meeting.get("meeting_id", "")),
        "subject": safe_str(meeting.get("subject", "")),
        "organizer": safe_str(meeting.get("organizer", "")),
        "date": date_value,
        "sentiment_score": sentiment,
        "churn_risk": safe_str(meeting.get("churn_risk", "unknown")),
        "events_detected": events,
        "summary": safe_str(meeting.get("summary", "")),
        "risks_detected": safe_str(meeting.get("risks_detected", "")),
        "opportunities_detected": safe_str(meeting.get("opportunities_detected", "")),
        "action_items": safe_str(meeting.get("action_items", "")),
        "searchable_content": safe_str(meeting.get("searchable_content", "")),
        "transcript_file": safe_str(meeting.get("transcript_file", ""))
    }


def create_index(search_client):
    """Create the Azure Search index."""
    import requests
    
    # Define index schema inline to avoid JSON file issues
    schema = {
        "name": SEARCH_INDEX_NAME,
        "fields": [
            {"name": "id", "type": "Edm.String", "key": True, "searchable": False, "filterable": True, "retrievable": True},
            {"name": "meeting_id", "type": "Edm.String", "searchable": True, "filterable": True, "retrievable": True},
            {"name": "subject", "type": "Edm.String", "searchable": True, "filterable": True, "sortable": True, "retrievable": True, "analyzer": "standard.lucene"},
            {"name": "organizer", "type": "Edm.String", "searchable": True, "filterable": True, "sortable": True, "facetable": True, "retrievable": True},
            {"name": "date", "type": "Edm.DateTimeOffset", "filterable": True, "sortable": True, "facetable": True, "retrievable": True},
            {"name": "sentiment_score", "type": "Edm.Int32", "filterable": True, "sortable": True, "facetable": True, "retrievable": True},
            {"name": "churn_risk", "type": "Edm.String", "filterable": True, "sortable": True, "facetable": True, "retrievable": True},
            {"name": "events_detected", "type": "Collection(Edm.String)", "searchable": True, "filterable": True, "facetable": True, "retrievable": True},
            {"name": "summary", "type": "Edm.String", "searchable": True, "retrievable": True, "analyzer": "standard.lucene"},
            {"name": "risks_detected", "type": "Edm.String", "searchable": True, "retrievable": True},
            {"name": "opportunities_detected", "type": "Edm.String", "searchable": True, "retrievable": True},
            {"name": "action_items", "type": "Edm.String", "searchable": True, "retrievable": True},
            {"name": "searchable_content", "type": "Edm.String", "searchable": True, "retrievable": True, "analyzer": "standard.lucene"},
            {"name": "transcript_file", "type": "Edm.String", "filterable": True, "retrievable": True}
        ],
        "suggesters": [
            {"name": "sg", "searchMode": "analyzingInfixMatching", "sourceFields": ["subject", "organizer"]}
        ]
    }
    
    url = f"https://{SEARCH_SERVICE_NAME}.search.windows.net/indexes/{SEARCH_INDEX_NAME}?api-version=2024-07-01"
    headers = {
        "Content-Type": "application/json",
        "api-key": SEARCH_API_KEY
    }
    
    # Try to delete existing index first
    delete_url = f"https://{SEARCH_SERVICE_NAME}.search.windows.net/indexes/{SEARCH_INDEX_NAME}?api-version=2024-07-01"
    print(f"  Deleting existing index (if any)...")
    requests.delete(delete_url, headers=headers)
    
    # Create new index
    print(f"  Creating index: {SEARCH_INDEX_NAME}")
    response = requests.put(url, headers=headers, json=schema)
    
    if response.status_code in [200, 201]:
        print(f"✅ Index '{SEARCH_INDEX_NAME}' created successfully")
        return True
    else:
        print(f"❌ Failed to create index: {response.status_code}")
        print(response.text)
        return False


def upload_documents(documents):
    """Upload documents to Azure Search."""
    import requests
    
    url = f"https://{SEARCH_SERVICE_NAME}.search.windows.net/indexes/{SEARCH_INDEX_NAME}/docs/index?api-version=2024-07-01"
    headers = {
        "Content-Type": "application/json",
        "api-key": SEARCH_API_KEY
    }
    
    # Upload in batches of 100
    batch_size = 100
    total_uploaded = 0
    
    for i in range(0, len(documents), batch_size):
        batch = documents[i:i + batch_size]
        payload = {"value": batch}
        
        response = requests.post(url, headers=headers, json=payload)
        
        if response.status_code in [200, 201]:
            result = response.json()
            succeeded = sum(1 for r in result.get("value", []) if r.get("status"))
            total_uploaded += succeeded
            print(f"  📤 Batch {i//batch_size + 1}: Uploaded {succeeded}/{len(batch)} documents")
        else:
            print(f"  ❌ Batch {i//batch_size + 1} failed: {response.status_code}")
            print(f"     {response.text[:500]}")
    
    return total_uploaded


def test_search(query="integration check-in"):
    """Test the search index with a sample query."""
    import requests
    
    url = f"https://{SEARCH_SERVICE_NAME}.search.windows.net/indexes/{SEARCH_INDEX_NAME}/docs/search?api-version=2024-07-01"
    headers = {
        "Content-Type": "application/json",
        "api-key": SEARCH_API_KEY
    }
    
    payload = {
        "search": query,
        "top": 5,
        "select": "subject,organizer,date,sentiment_score,summary"
    }
    
    response = requests.post(url, headers=headers, json=payload)
    
    if response.status_code == 200:
        result = response.json()
        print(f"\n🔍 Test search for '{query}':")
        for doc in result.get("value", []):
            print(f"  - {doc.get('subject', 'N/A')}")
            print(f"    Organizer: {doc.get('organizer', 'N/A')}, Sentiment: {doc.get('sentiment_score', 'N/A')}")
        return True
    else:
        print(f"❌ Search test failed: {response.status_code}")
        return False


def main():
    """Main function to upload knowledge base to Azure Search."""
    print("=" * 60)
    print("UPLOADING KNOWLEDGE BASE TO AZURE AI SEARCH")
    print("=" * 60)
    print(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Check configuration
    if not SEARCH_SERVICE_NAME or not SEARCH_API_KEY:
        print("\n❌ Azure Search configuration missing!")
        print("Please set the following environment variables in .env:")
        print("  AZURE_SEARCH_SERVICE_NAME=your-search-service-name")
        print("  AZURE_SEARCH_API_KEY=your-admin-api-key")
        print("  AZURE_SEARCH_INDEX_NAME=copilot-knowledge-base (optional)")
        return
    
    print(f"\n📋 Configuration:")
    print(f"  Service: {SEARCH_SERVICE_NAME}")
    print(f"  Index: {SEARCH_INDEX_NAME}")
    
    # Find and load knowledge base
    kb_path = find_latest_knowledge_base()
    if not kb_path:
        print("❌ No knowledge base file found in output/")
        return
    
    print(f"\n📂 Loading: {os.path.basename(kb_path)}")
    kb_data = load_knowledge_base(kb_path)
    
    meetings = kb_data.get("meetings", [])
    print(f"  Found {len(meetings)} meetings")
    
    # Transform meetings to documents
    print("\n🔄 Transforming documents...")
    documents = []
    for meeting in meetings:
        try:
            doc = transform_meeting_to_document(meeting)
            documents.append(doc)
        except Exception as e:
            print(f"  ⚠️ Error transforming meeting: {e}")
    
    print(f"  Prepared {len(documents)} documents")
    
    # Create index
    print("\n📊 Creating search index...")
    if not create_index(None):
        print("⚠️ Index creation may have failed, attempting upload anyway...")
    
    # Upload documents
    print("\n📤 Uploading documents...")
    uploaded = upload_documents(documents)
    
    print(f"\n✅ Upload complete: {uploaded}/{len(documents)} documents indexed")
    
    # Test search
    test_search()
    
    # Save upload report
    report = {
        "timestamp": datetime.now().isoformat(),
        "search_service": SEARCH_SERVICE_NAME,
        "index_name": SEARCH_INDEX_NAME,
        "source_file": os.path.basename(kb_path),
        "total_meetings": len(meetings),
        "documents_uploaded": uploaded
    }
    
    report_path = os.path.join(OUTPUT_DIR, "azure_search_upload_report.json")
    with open(report_path, 'w', encoding='utf-8') as f:
        json.dump(report, f, indent=2)
    
    print(f"\n📄 Report saved: {report_path}")
    print("\n" + "=" * 60)
    print("NEXT STEPS FOR COPILOT AGENT:")
    print("=" * 60)
    print("1. Go to Azure Portal > Azure AI Search")
    print(f"2. Select service: {SEARCH_SERVICE_NAME}")
    print(f"3. Navigate to Indexes > {SEARCH_INDEX_NAME}")
    print("4. In Copilot Studio, add this as a knowledge source")
    print("   - Use 'Azure AI Search' connector")
    print(f"   - Service URL: https://{SEARCH_SERVICE_NAME}.search.windows.net")
    print(f"   - Index: {SEARCH_INDEX_NAME}")


if __name__ == "__main__":
    main()
