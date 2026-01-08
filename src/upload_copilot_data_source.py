#!/usr/bin/env python3
"""
Azure AI Search Uploader for Copilot Agent

Uploads the generated data source to Azure AI Search for use with Copilot Agent.
"""

import json
import os
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Any
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Azure Search Configuration
AZURE_SEARCH_ENDPOINT = os.getenv('AZURE_SEARCH_ENDPOINT', '')
AZURE_SEARCH_KEY = os.getenv('AZURE_SEARCH_KEY', '')
AZURE_SEARCH_INDEX = os.getenv('AZURE_SEARCH_INDEX', 'copilot-agent-comprehensive')

# Paths
OUTPUT_DIR = Path('output')
INDEX_SCHEMA_FILE = OUTPUT_DIR / 'copilot_agent_index_schema.json'
DATA_SOURCE_FILE = OUTPUT_DIR / 'copilot_agent_data_source.json'


def check_configuration():
    """Check if Azure Search is configured."""
    if not AZURE_SEARCH_ENDPOINT or not AZURE_SEARCH_KEY:
        print("⚠️  Azure Search is not configured.")
        print("\nTo configure Azure Search, add the following to your .env file:")
        print("  AZURE_SEARCH_ENDPOINT=https://your-search-service.search.windows.net")
        print("  AZURE_SEARCH_KEY=your-admin-key")
        print("  AZURE_SEARCH_INDEX=copilot-agent-comprehensive")
        return False
    return True


def create_index_if_not_exists():
    """Create the Azure Search index if it doesn't exist."""
    try:
        from azure.core.credentials import AzureKeyCredential
        from azure.search.documents.indexes import SearchIndexClient
        from azure.search.documents.indexes.models import (
            SearchIndex, SearchField, SearchFieldDataType,
            SimpleField, SearchableField,
            SemanticConfiguration, SemanticField, SemanticPrioritizedFields,
            SemanticSearch
        )
        
        # Create client
        credential = AzureKeyCredential(AZURE_SEARCH_KEY)
        index_client = SearchIndexClient(endpoint=AZURE_SEARCH_ENDPOINT, credential=credential)
        
        # Define fields explicitly for proper type handling
        fields = [
            # Key field
            SimpleField(name="id", type=SearchFieldDataType.String, key=True, filterable=True),
            
            # Simple string fields (filterable/facetable)
            SimpleField(name="document_type", type=SearchFieldDataType.String, filterable=True, facetable=True),
            SimpleField(name="entity_type", type=SearchFieldDataType.String, filterable=True, facetable=True),
            SimpleField(name="risk_level", type=SearchFieldDataType.String, filterable=True, facetable=True, sortable=True),
            SimpleField(name="urgency", type=SearchFieldDataType.String, filterable=True, facetable=True, sortable=True),
            SimpleField(name="status", type=SearchFieldDataType.String, filterable=True, facetable=True),
            SimpleField(name="health_status", type=SearchFieldDataType.String, filterable=True, facetable=True),
            SimpleField(name="suggestion_category", type=SearchFieldDataType.String, filterable=True, facetable=True),
            SimpleField(name="meeting_type", type=SearchFieldDataType.String, filterable=True, facetable=True),
            SimpleField(name="date_string", type=SearchFieldDataType.String, filterable=True, sortable=True),
            SimpleField(name="transcript_file", type=SearchFieldDataType.String, filterable=True),
            SimpleField(name="source_file", type=SearchFieldDataType.String, filterable=True),
            SimpleField(name="blob_url", type=SearchFieldDataType.String),
            SimpleField(name="reviewed_by", type=SearchFieldDataType.String, filterable=True, facetable=True),
            
            # Searchable string fields
            SearchableField(name="title", type=SearchFieldDataType.String, filterable=True, sortable=True, analyzer_name="en.microsoft"),
            SearchableField(name="content", type=SearchFieldDataType.String, analyzer_name="en.microsoft"),
            SearchableField(name="summary", type=SearchFieldDataType.String, analyzer_name="en.microsoft"),
            SearchableField(name="va_name", type=SearchFieldDataType.String, filterable=True, facetable=True, sortable=True),
            SearchableField(name="client_name", type=SearchFieldDataType.String, filterable=True, facetable=True, sortable=True),
            SearchableField(name="organizer", type=SearchFieldDataType.String, filterable=True, facetable=True),
            SearchableField(name="issues", type=SearchFieldDataType.String, analyzer_name="en.microsoft"),
            SearchableField(name="key_concerns", type=SearchFieldDataType.String, analyzer_name="en.microsoft"),
            SearchableField(name="key_positives", type=SearchFieldDataType.String, analyzer_name="en.microsoft"),
            SearchableField(name="actions_taken", type=SearchFieldDataType.String, analyzer_name="en.microsoft"),
            SearchableField(name="recommended_action", type=SearchFieldDataType.String, analyzer_name="en.microsoft"),
            SearchableField(name="rationale", type=SearchFieldDataType.String, analyzer_name="en.microsoft"),
            SearchableField(name="stakeholder_notes", type=SearchFieldDataType.String),
            
            # Collection (array) fields - use SearchField directly for proper Collection type
            SearchField(name="participants", type=SearchFieldDataType.Collection(SearchFieldDataType.String), searchable=True, filterable=True, facetable=True),
            SearchField(name="risk_signals", type=SearchFieldDataType.Collection(SearchFieldDataType.String), searchable=True, filterable=True, facetable=True),
            SearchField(name="events_detected", type=SearchFieldDataType.Collection(SearchFieldDataType.String), searchable=True, filterable=True, facetable=True),
            SearchField(name="action_items", type=SearchFieldDataType.Collection(SearchFieldDataType.String), searchable=True),
            SearchField(name="suggestions", type=SearchFieldDataType.Collection(SearchFieldDataType.String), searchable=True),
            SearchField(name="key_insights", type=SearchFieldDataType.Collection(SearchFieldDataType.String), searchable=True),
            SearchField(name="best_practices", type=SearchFieldDataType.Collection(SearchFieldDataType.String), searchable=True),
            SearchField(name="kpis_mentioned", type=SearchFieldDataType.Collection(SearchFieldDataType.String), searchable=True, filterable=True, facetable=True),
            SearchField(name="tags", type=SearchFieldDataType.Collection(SearchFieldDataType.String), searchable=True, filterable=True, facetable=True),
            
            # Numeric fields
            SimpleField(name="risk_score", type=SearchFieldDataType.Int32, filterable=True, sortable=True, facetable=True),
            SimpleField(name="sentiment_score", type=SearchFieldDataType.Int32, filterable=True, sortable=True, facetable=True),
            SimpleField(name="opportunity_score", type=SearchFieldDataType.Int32, filterable=True, sortable=True, facetable=True),
            
            # Date fields
            SimpleField(name="date", type=SearchFieldDataType.DateTimeOffset, filterable=True, sortable=True, facetable=True),
            SimpleField(name="created_at", type=SearchFieldDataType.DateTimeOffset, filterable=True, sortable=True),
            SimpleField(name="updated_at", type=SearchFieldDataType.DateTimeOffset, filterable=True, sortable=True),
        ]
        
        # Create semantic configuration
        semantic_config = SemanticConfiguration(
            name="semantic-config",
            prioritized_fields=SemanticPrioritizedFields(
                title_field=SemanticField(field_name="title"),
                content_fields=[
                    SemanticField(field_name="summary"),
                    SemanticField(field_name="content"),
                    SemanticField(field_name="key_concerns"),
                    SemanticField(field_name="issues")
                ],
                keywords_fields=[
                    SemanticField(field_name="va_name"),
                    SemanticField(field_name="client_name")
                ]
            )
        )
        
        semantic_search = SemanticSearch(configurations=[semantic_config])
        
        # Create the index
        index = SearchIndex(
            name=AZURE_SEARCH_INDEX,
            fields=fields,
            semantic_search=semantic_search
        )
        
        result = index_client.create_or_update_index(index)
        print(f"✅ Index '{result.name}' created/updated successfully")
        return True
        
    except ImportError:
        print("⚠️  Azure Search SDK not installed. Install with:")
        print("    pip install azure-search-documents")
        return False
    except Exception as e:
        print(f"❌ Error creating index: {e}")
        import traceback
        traceback.print_exc()
        return False


def upload_documents():
    """Upload documents to Azure Search."""
    try:
        from azure.core.credentials import AzureKeyCredential
        from azure.search.documents import SearchClient
        
        # Load data source
        with open(DATA_SOURCE_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        documents = data.get('documents', [])
        if not documents:
            print("❌ No documents to upload")
            return False
        
        # Create search client
        credential = AzureKeyCredential(AZURE_SEARCH_KEY)
        search_client = SearchClient(
            endpoint=AZURE_SEARCH_ENDPOINT,
            index_name=AZURE_SEARCH_INDEX,
            credential=credential
        )
        
        # Upload in batches
        batch_size = 100
        total_uploaded = 0
        
        # Fields that should be collections (arrays)
        collection_fields = {
            'participants', 'risk_signals', 'events_detected', 'action_items',
            'suggestions', 'key_insights', 'best_practices', 'kpis_mentioned', 'tags'
        }
        
        # Fields that should be strings (not arrays)
        string_fields = {
            'id', 'document_type', 'entity_type', 'title', 'content', 'summary',
            'va_name', 'client_name', 'organizer', 'date_string', 'risk_level',
            'issues', 'key_concerns', 'key_positives', 'actions_taken',
            'suggestion_category', 'urgency', 'status', 'health_status',
            'recommended_action', 'rationale', 'meeting_type', 'transcript_file',
            'source_file', 'blob_url', 'reviewed_by', 'stakeholder_notes'
        }
        
        # DateTime fields that need proper ISO 8601 format with Z suffix
        datetime_fields = {'date', 'created_at', 'updated_at'}
        
        def fix_datetime(value):
            """Convert datetime to proper ISO 8601 format with Z suffix for Azure Search."""
            if not value:
                return None
            if isinstance(value, str):
                # Remove any existing timezone info and add Z
                if value.endswith('Z'):
                    return value
                # Handle formats like '2026-01-05T23:22:11.838381'
                if 'T' in value:
                    # Remove microseconds beyond 7 digits (Azure limit) and add Z
                    parts = value.split('.')
                    if len(parts) == 2:
                        # Take first 7 digits of microseconds max
                        micros = parts[1][:7] if len(parts[1]) > 7 else parts[1]
                        return f"{parts[0]}.{micros}Z"
                    return f"{value}Z"
                return value
            return str(value)
        
        for i in range(0, len(documents), batch_size):
            batch = documents[i:i + batch_size]
            
            # Clean documents for upload
            cleaned_batch = []
            for doc in batch:
                cleaned_doc = {}
                for key, value in doc.items():
                    if value is None:
                        continue
                    if isinstance(value, list) and len(value) == 0:
                        continue
                    if isinstance(value, str) and value == '':
                        continue
                    
                    # Convert fields to proper types
                    if key in collection_fields:
                        # Ensure it's a list
                        if isinstance(value, str):
                            cleaned_doc[key] = [value] if value else []
                        elif isinstance(value, list):
                            cleaned_doc[key] = [str(v) for v in value if v]
                        else:
                            cleaned_doc[key] = [str(value)]
                    elif key in datetime_fields:
                        # Fix datetime format for Azure Search
                        cleaned_doc[key] = fix_datetime(value)
                    elif key in string_fields:
                        # Ensure it's a string
                        if isinstance(value, list):
                            cleaned_doc[key] = ', '.join(str(v) for v in value if v)
                        else:
                            cleaned_doc[key] = str(value)
                    else:
                        # Keep as is (dates, numbers)
                        cleaned_doc[key] = value
                        
                cleaned_batch.append(cleaned_doc)
            
            try:
                result = search_client.upload_documents(documents=cleaned_batch)
                succeeded = sum(1 for r in result if r.succeeded)
                total_uploaded += succeeded
                print(f"  Batch {i//batch_size + 1}: {succeeded}/{len(batch)} documents uploaded")
                
                # Show errors for failed documents
                for r in result:
                    if not r.succeeded:
                        print(f"    Doc {r.key}: {r.error_message}")
                        break  # Only show first error
            except Exception as e:
                print(f"  Batch {i//batch_size + 1}: Error - {e}")
                # Debug: print first document structure
                if cleaned_batch:
                    print(f"    Sample doc keys: {list(cleaned_batch[0].keys())}")
                    for k, v in cleaned_batch[0].items():
                        print(f"      {k}: {type(v).__name__} = {str(v)[:50]}...")
        
        print(f"\n✅ Total documents uploaded: {total_uploaded}/{len(documents)}")
        return True
        
    except ImportError:
        print("⚠️  Azure Search SDK not installed. Install with:")
        print("    pip install azure-search-documents")
        return False
    except Exception as e:
        print(f"❌ Error uploading documents: {e}")
        return False


def generate_upload_report():
    """Generate a report of the upload operation."""
    report = {
        'timestamp': datetime.now().isoformat(),
        'index_name': AZURE_SEARCH_INDEX,
        'endpoint': AZURE_SEARCH_ENDPOINT,
        'schema_file': str(INDEX_SCHEMA_FILE),
        'data_source_file': str(DATA_SOURCE_FILE)
    }
    
    # Load data source stats
    if DATA_SOURCE_FILE.exists():
        with open(DATA_SOURCE_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        report['statistics'] = data.get('statistics', {})
        report['total_documents'] = data.get('metadata', {}).get('total_documents', 0)
    
    report_file = OUTPUT_DIR / 'copilot_agent_upload_report.json'
    with open(report_file, 'w', encoding='utf-8') as f:
        json.dump(report, f, indent=2)
    
    print(f"📝 Upload report saved to: {report_file}")
    return report


def main():
    """Main function to orchestrate the upload process."""
    print("=" * 60)
    print("Azure AI Search Uploader for Copilot Agent")
    print("=" * 60)
    
    # Check configuration
    if not check_configuration():
        print("\n💡 Generating upload-ready files only...")
        generate_upload_report()
        print("\nOnce Azure Search is configured, run this script again to upload.")
        return
    
    # Create index
    print("\n📊 Creating/Updating Index...")
    if not create_index_if_not_exists():
        return
    
    # Upload documents
    print("\n📤 Uploading Documents...")
    if not upload_documents():
        return
    
    # Generate report
    print("\n📝 Generating Report...")
    generate_upload_report()
    
    print("\n" + "=" * 60)
    print("✅ Upload complete!")
    print("=" * 60)


if __name__ == '__main__':
    main()
