"""
Upload Excel to Azure Blob Storage and Create Search Index for Copilot Agent
=============================================================================
This script:
1. Uploads the latest Excel file to Azure Blob Storage in a 'reports' container
2. Creates a JSON index from the Excel data for Azure AI Search / Copilot integration
3. Generates embeddings for semantic search (optional)
"""

import os
import json
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv
from azure.storage.blob import BlobServiceClient
from azure.core.exceptions import ResourceExistsError
from openai import AzureOpenAI

load_dotenv()

# Configuration
AZURE_STORAGE_CONNECTION_STRING = os.getenv("AZURE_STORAGE_CONNECTION_STRING")
AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")
AZURE_OPENAI_KEY = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_OPENAI_DEPLOYMENT = os.getenv("AZURE_OPENAI_DEPLOYMENT", "gpt-4.1")

# Container names
REPORTS_CONTAINER = "reports"
INDEX_CONTAINER = "copilot-index"

def get_latest_excel():
    """Find the latest analyzed Excel file"""
    output_dir = "output"
    
    # Check for the latest analyzed file first
    latest_file = os.path.join(output_dir, "meeting_transcripts_latest_analyzed.xlsx")
    if os.path.exists(latest_file):
        return latest_file
    
    # Otherwise find the newest Excel file
    excel_files = [f for f in os.listdir(output_dir) if f.endswith('.xlsx')]
    if not excel_files:
        raise FileNotFoundError("No Excel files found in output directory")
    
    # Sort by modification time
    excel_files.sort(key=lambda x: os.path.getmtime(os.path.join(output_dir, x)), reverse=True)
    return os.path.join(output_dir, excel_files[0])

def upload_to_blob(file_path: str, container_name: str, blob_name: str = None):
    """Upload a file to Azure Blob Storage"""
    if not AZURE_STORAGE_CONNECTION_STRING:
        raise ValueError("AZURE_STORAGE_CONNECTION_STRING not set in environment")
    
    blob_service_client = BlobServiceClient.from_connection_string(AZURE_STORAGE_CONNECTION_STRING)
    
    # Create container if it doesn't exist
    try:
        container_client = blob_service_client.create_container(container_name)
        print(f"✅ Created container: {container_name}")
    except ResourceExistsError:
        container_client = blob_service_client.get_container_client(container_name)
    
    # Use filename if blob_name not specified
    if blob_name is None:
        blob_name = os.path.basename(file_path)
    
    # Upload the file
    blob_client = container_client.get_blob_client(blob_name)
    
    with open(file_path, "rb") as data:
        blob_client.upload_blob(data, overwrite=True)
    
    blob_url = blob_client.url
    print(f"✅ Uploaded: {blob_name}")
    print(f"   URL: {blob_url}")
    
    return blob_url

def create_search_index(excel_path: str):
    """
    Create a search index from Excel data for Copilot Agent integration.
    Returns a list of documents ready for Azure AI Search or similar.
    """
    print("\n📊 Creating search index from Excel data...")
    
    df = pd.read_excel(excel_path)
    print(f"   Loaded {len(df)} rows from Excel")
    print(f"   Columns: {df.columns.tolist()}")
    
    documents = []
    
    for idx, row in df.iterrows():
        # Get values with proper column names (case-insensitive lookup)
        def get_val(col_name, default=''):
            for col in df.columns:
                if col.lower().replace(' ', '_') == col_name.lower().replace(' ', '_'):
                    val = row.get(col)
                    return val if pd.notna(val) else default
                if col.lower() == col_name.lower():
                    val = row.get(col)
                    return val if pd.notna(val) else default
            return default
        
        # Get date and time
        date_val = get_val('Date', '')
        time_val = get_val('Time', '')
        start_datetime = f"{date_val} {time_val}".strip() if date_val else ''
        
        # Create a searchable document for each meeting
        doc = {
            "id": f"meeting_{idx}",
            "meeting_id": str(get_val('Meeting ID', f'meeting_{idx}')),
            "subject": str(get_val('Subject', '')),
            "organizer": str(get_val('Organizer', '')),
            "start_datetime": start_datetime,
            "date": str(date_val),
            "time": str(time_val),
            "duration_min": int(get_val('Duration (min)', 0)) if get_val('Duration (min)', 0) else 0,
            "participants": str(get_val('Participants', '')),
            "team": str(get_val('Team', '')),
            "transcript_file": str(get_val('Transcript File', '')),
            "blob_url": str(get_val('Blob Storage Link', '')),
            
            # AI Analysis scores (updated column names)
            "sentiment_score": int(get_val('Sentiment Score', 0)) if get_val('Sentiment Score', 0) else 0,
            "churn_risk": int(get_val('Churn Risk', 0)) if get_val('Churn Risk', 0) else 0,
            "opportunity_score": int(get_val('Opportunity Score', 0)) if get_val('Opportunity Score', 0) else 0,
            "execution_reliability": int(get_val('Execution Reliability', 0)) if get_val('Execution Reliability', 0) else 0,
            "operational_complexity": int(get_val('Operational Complexity', 0)) if get_val('Operational Complexity', 0) else 0,
            
            # AI Analysis text fields
            "events": str(get_val('Events', '')),
            "summary": str(get_val('Summary', '')),
            "key_concerns": str(get_val('Key Concerns', '')),
            "key_positives": str(get_val('Key Positives', '')),
            "action_items": str(get_val('Action Items', '')),
            "analyzed_at": str(get_val('Analyzed At', '')),
            
            # Combined searchable text for semantic search
            "searchable_content": f"{get_val('Subject', '')} {get_val('Organizer', '')} {get_val('Events', '')} {get_val('Summary', '')} {get_val('Key Concerns', '')} {get_val('Key Positives', '')} {get_val('Action Items', '')}"
        }
        
        documents.append(doc)
    
    print(f"   Created {len(documents)} searchable documents")
    return documents

def create_copilot_knowledge_base(documents: list):
    """
    Create a structured knowledge base for Copilot Agent.
    This creates multiple indexes for different query types.
    """
    print("\n🤖 Creating Copilot knowledge base...")
    
    knowledge_base = {
        "metadata": {
            "created_at": datetime.now().isoformat(),
            "total_meetings": len(documents),
            "version": "1.0",
            "description": "Meeting transcripts analysis index for Copilot Agent"
        },
        
        # Full document index
        "meetings": documents,
        
        # Quick lookup indexes
        "by_organizer": {},
        "by_date": {},
        "by_sentiment": {
            "high": [],    # 70+
            "medium": [],  # 50-69
            "low": []      # <50
        },
        "by_churn_risk": {
            "high": [],    # 50+
            "medium": [],  # 25-49
            "low": []      # <25
        },
        
        # Aggregate statistics
        "statistics": {
            "total_meetings": len(documents),
            "avg_sentiment": 0,
            "avg_churn_risk": 0,
            "avg_opportunity_score": 0,
            "avg_execution_reliability": 0,
            "avg_operational_complexity": 0,
            "organizer_counts": {},
            "meetings_by_month": {},
            "meetings_by_date": {}
        }
    }
    
    # Build indexes
    sentiment_sum = 0
    churn_sum = 0
    opportunity_sum = 0
    execution_sum = 0
    complexity_sum = 0
    count_with_scores = 0
    
    for doc in documents:
        organizer = doc.get('organizer', 'Unknown')
        subject = doc.get('subject', '')
        sentiment = doc.get('sentiment_score', 0)
        churn_risk = doc.get('churn_risk', 0)
        
        # By organizer index
        if organizer not in knowledge_base["by_organizer"]:
            knowledge_base["by_organizer"][organizer] = []
        knowledge_base["by_organizer"][organizer].append({
            "id": doc["id"],
            "subject": subject,
            "date": doc.get("date", ""),
            "sentiment_score": sentiment,
            "churn_risk": churn_risk
        })
        
        # By date index
        date_str = doc.get('date', '')
        if date_str and date_str != 'nan':
            if date_str not in knowledge_base["by_date"]:
                knowledge_base["by_date"][date_str] = []
            knowledge_base["by_date"][date_str].append({
                "id": doc["id"],
                "subject": subject,
                "organizer": organizer,
                "sentiment_score": sentiment
            })
            
            # Monthly aggregation
            try:
                month = date_str[:7] if len(date_str) >= 7 else "unknown"
                knowledge_base["statistics"]["meetings_by_month"][month] = \
                    knowledge_base["statistics"]["meetings_by_month"].get(month, 0) + 1
                knowledge_base["statistics"]["meetings_by_date"][date_str] = \
                    knowledge_base["statistics"]["meetings_by_date"].get(date_str, 0) + 1
            except:
                pass
        
        # By sentiment index
        if sentiment >= 70:
            knowledge_base["by_sentiment"]["high"].append({"id": doc["id"], "subject": subject, "score": sentiment})
        elif sentiment >= 50:
            knowledge_base["by_sentiment"]["medium"].append({"id": doc["id"], "subject": subject, "score": sentiment})
        elif sentiment > 0:
            knowledge_base["by_sentiment"]["low"].append({"id": doc["id"], "subject": subject, "score": sentiment})
        
        # By churn risk index
        if churn_risk >= 50:
            knowledge_base["by_churn_risk"]["high"].append({"id": doc["id"], "subject": subject, "risk": churn_risk})
        elif churn_risk >= 25:
            knowledge_base["by_churn_risk"]["medium"].append({"id": doc["id"], "subject": subject, "risk": churn_risk})
        elif churn_risk > 0:
            knowledge_base["by_churn_risk"]["low"].append({"id": doc["id"], "subject": subject, "risk": churn_risk})
        
        # Statistics
        if sentiment > 0:
            sentiment_sum += sentiment
            churn_sum += churn_risk
            opportunity_sum += doc.get('opportunity_score', 0)
            execution_sum += doc.get('execution_reliability', 0)
            complexity_sum += doc.get('operational_complexity', 0)
            count_with_scores += 1
        
        # Organizer counts
        knowledge_base["statistics"]["organizer_counts"][organizer] = \
            knowledge_base["statistics"]["organizer_counts"].get(organizer, 0) + 1
    
    # Calculate averages
    if count_with_scores > 0:
        knowledge_base["statistics"]["avg_sentiment"] = round(sentiment_sum / count_with_scores, 1)
        knowledge_base["statistics"]["avg_churn_risk"] = round(churn_sum / count_with_scores, 1)
        knowledge_base["statistics"]["avg_opportunity_score"] = round(opportunity_sum / count_with_scores, 1)
        knowledge_base["statistics"]["avg_execution_reliability"] = round(execution_sum / count_with_scores, 1)
        knowledge_base["statistics"]["avg_operational_complexity"] = round(complexity_sum / count_with_scores, 1)
    
    print(f"   ✅ Built organizer index: {len(knowledge_base['by_organizer'])} organizers")
    print(f"   ✅ Built date index: {len(knowledge_base['by_date'])} dates")
    print(f"   ✅ Sentiment distribution: High={len(knowledge_base['by_sentiment']['high'])}, Medium={len(knowledge_base['by_sentiment']['medium'])}, Low={len(knowledge_base['by_sentiment']['low'])}")
    print(f"   ✅ Churn risk distribution: High={len(knowledge_base['by_churn_risk']['high'])}, Medium={len(knowledge_base['by_churn_risk']['medium'])}, Low={len(knowledge_base['by_churn_risk']['low'])}")
    
    return knowledge_base

def generate_embeddings(documents: list, batch_size: int = 20):
    """
    Generate embeddings for semantic search using Azure OpenAI.
    This is optional but enables better semantic search in Copilot.
    """
    if not AZURE_OPENAI_ENDPOINT or not AZURE_OPENAI_KEY:
        print("⚠️  Azure OpenAI not configured, skipping embeddings")
        return documents
    
    print("\n🧠 Generating embeddings for semantic search...")
    
    try:
        client = AzureOpenAI(
            azure_endpoint=AZURE_OPENAI_ENDPOINT,
            api_key=AZURE_OPENAI_KEY,
            api_version="2024-12-01-preview"
        )
        
        # Check if embedding model is available
        # Using text-embedding-ada-002 or text-embedding-3-small
        embedding_deployment = os.getenv("AZURE_OPENAI_EMBEDDING_DEPLOYMENT", "text-embedding-ada-002")
        
        for i, doc in enumerate(documents):
            if i % 10 == 0:
                print(f"   Processing {i+1}/{len(documents)}...")
            
            text = doc.get("searchable_content", "")[:8000]  # Limit text length
            
            try:
                response = client.embeddings.create(
                    input=text,
                    model=embedding_deployment
                )
                doc["embedding"] = response.data[0].embedding
            except Exception as e:
                # Skip embeddings if model not available
                if "DeploymentNotFound" in str(e) or "model" in str(e).lower():
                    print(f"   ⚠️  Embedding model not available, skipping embeddings")
                    return documents
                print(f"   ⚠️  Error generating embedding for doc {i}: {e}")
        
        print(f"   ✅ Generated embeddings for {len(documents)} documents")
        
    except Exception as e:
        print(f"   ⚠️  Embedding generation skipped: {e}")
    
    return documents

def save_and_upload_index(knowledge_base: dict, output_dir: str = "output"):
    """Save the index locally and upload to Azure Blob"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Save main knowledge base
    kb_filename = f"copilot_knowledge_base_{timestamp}.json"
    kb_path = os.path.join(output_dir, kb_filename)
    
    with open(kb_path, 'w', encoding='utf-8') as f:
        json.dump(knowledge_base, f, indent=2, ensure_ascii=False, default=str)
    print(f"\n💾 Saved knowledge base: {kb_path}")
    
    # Also save a "latest" version
    latest_path = os.path.join(output_dir, "copilot_knowledge_base_latest.json")
    with open(latest_path, 'w', encoding='utf-8') as f:
        json.dump(knowledge_base, f, indent=2, ensure_ascii=False, default=str)
    
    # Create a simplified index for quick queries
    simple_index = {
        "metadata": knowledge_base["metadata"],
        "statistics": knowledge_base["statistics"],
        "meetings_summary": [
            {
                "id": m["id"],
                "subject": m["subject"],
                "organizer": m["organizer"],
                "date": m["start_datetime"],
                "sentiment": m["sentiment_score"]
            }
            for m in knowledge_base["meetings"]
        ]
    }
    
    simple_path = os.path.join(output_dir, "copilot_index_simple.json")
    with open(simple_path, 'w', encoding='utf-8') as f:
        json.dump(simple_index, f, indent=2, ensure_ascii=False, default=str)
    print(f"💾 Saved simple index: {simple_path}")
    
    # Upload to Azure Blob
    print("\n☁️  Uploading to Azure Blob Storage...")
    
    try:
        # Upload full knowledge base
        upload_to_blob(kb_path, INDEX_CONTAINER, f"knowledge_base/{kb_filename}")
        upload_to_blob(latest_path, INDEX_CONTAINER, "knowledge_base/copilot_knowledge_base_latest.json")
        
        # Upload simple index
        upload_to_blob(simple_path, INDEX_CONTAINER, "knowledge_base/copilot_index_simple.json")
        
    except Exception as e:
        print(f"❌ Upload failed: {e}")
        raise
    
    return kb_path, simple_path

def main():
    print("=" * 60)
    print("UPLOAD EXCEL & CREATE COPILOT INDEX")
    print("=" * 60)
    print(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Step 1: Find latest Excel
    print("\n📂 Finding latest Excel file...")
    excel_path = get_latest_excel()
    print(f"   Found: {excel_path}")
    
    # Step 2: Upload Excel to reports container
    print("\n☁️  Uploading Excel to Azure Blob Storage...")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    try:
        # Upload to reports folder
        excel_url = upload_to_blob(
            excel_path, 
            REPORTS_CONTAINER, 
            f"excel_reports/meeting_analysis_{timestamp}.xlsx"
        )
        
        # Also upload as "latest"
        upload_to_blob(
            excel_path,
            REPORTS_CONTAINER,
            "excel_reports/meeting_analysis_latest.xlsx"
        )
    except Exception as e:
        print(f"❌ Excel upload failed: {e}")
        return
    
    # Step 3: Create search index
    documents = create_search_index(excel_path)
    
    # Step 4: Create knowledge base for Copilot
    knowledge_base = create_copilot_knowledge_base(documents)
    
    # Step 5: Generate embeddings (optional)
    # Uncomment if you have an embedding model deployed
    # documents = generate_embeddings(documents)
    # knowledge_base["meetings"] = documents
    
    # Step 6: Save and upload index
    kb_path, simple_path = save_and_upload_index(knowledge_base)
    
    # Summary
    print("\n" + "=" * 60)
    print("SUMMARY")
    print("=" * 60)
    print(f"✅ Excel uploaded to: reports/excel_reports/")
    print(f"✅ Knowledge base created with {len(documents)} meetings")
    print(f"✅ Index uploaded to: copilot-index/knowledge_base/")
    print(f"\n📊 Statistics:")
    print(f"   Total meetings: {knowledge_base['statistics']['total_meetings']}")
    print(f"   Avg sentiment: {knowledge_base['statistics']['avg_sentiment']}")
    print(f"   Avg churn risk: {knowledge_base['statistics']['avg_churn_risk']}")
    print(f"   Avg opportunity score: {knowledge_base['statistics']['avg_opportunity_score']}")
    print(f"   Avg execution reliability: {knowledge_base['statistics']['avg_execution_reliability']}")
    print(f"   Organizers: {len(knowledge_base['statistics']['organizer_counts'])}")
    print(f"   Dates covered: {len(knowledge_base['by_date'])}")
    
    print("\n🤖 Copilot Integration:")
    print("   The knowledge base JSON files can be used with:")
    print("   - Azure AI Search (create an index from the JSON)")
    print("   - Copilot Studio (import as knowledge source)")
    print("   - Custom Copilot Agent (use as RAG data source)")
    
    print("\n📁 Azure Blob URLs:")
    print(f"   Excel: https://aidevelopement.blob.core.windows.net/reports/excel_reports/meeting_analysis_latest.xlsx")
    print(f"   Index: https://aidevelopement.blob.core.windows.net/copilot-index/knowledge_base/copilot_knowledge_base_latest.json")

if __name__ == "__main__":
    main()
