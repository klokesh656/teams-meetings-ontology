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
            SimpleField, SearchableField, ComplexField,
            VectorSearch, HnswAlgorithmConfiguration, VectorSearchProfile,
            SemanticConfiguration, SemanticField, SemanticPrioritizedFields,
            SemanticSearch, ScoringProfile, TextWeights,
            FreshnessScoringFunction, FreshnessScoringParameters,
            MagnitudeScoringFunction, MagnitudeScoringParameters,
            ScoringFunctionInterpolation
        )
        
        # Load schema
        with open(INDEX_SCHEMA_FILE, 'r', encoding='utf-8') as f:
            schema = json.load(f)
        
        # Create client
        credential = AzureKeyCredential(AZURE_SEARCH_KEY)
        index_client = SearchIndexClient(endpoint=AZURE_SEARCH_ENDPOINT, credential=credential)
        
        # Build fields from schema
        fields = []
        for field_def in schema['fields']:
            field_name = field_def['name']
            field_type = field_def['type']
            
            # Skip vector fields for now (requires special handling)
            if field_type == 'Collection(Edm.Single)':
                continue
            
            if field_type == 'Collection(Edm.String)':
                fields.append(SearchableField(
                    name=field_name,
                    type=SearchFieldDataType.Collection(SearchFieldDataType.String),
                    filterable=field_def.get('filterable', False),
                    facetable=field_def.get('facetable', False)
                ))
            elif field_type == 'Edm.String':
                if field_def.get('key', False):
                    fields.append(SimpleField(
                        name=field_name,
                        type=SearchFieldDataType.String,
                        key=True,
                        filterable=True
                    ))
                elif field_def.get('searchable', True):
                    fields.append(SearchableField(
                        name=field_name,
                        type=SearchFieldDataType.String,
                        filterable=field_def.get('filterable', False),
                        sortable=field_def.get('sortable', False),
                        facetable=field_def.get('facetable', False),
                        analyzer_name=field_def.get('analyzer', 'en.microsoft')
                    ))
                else:
                    fields.append(SimpleField(
                        name=field_name,
                        type=SearchFieldDataType.String,
                        filterable=field_def.get('filterable', False),
                        sortable=field_def.get('sortable', False),
                        facetable=field_def.get('facetable', False)
                    ))
            elif field_type == 'Edm.DateTimeOffset':
                fields.append(SimpleField(
                    name=field_name,
                    type=SearchFieldDataType.DateTimeOffset,
                    filterable=field_def.get('filterable', True),
                    sortable=field_def.get('sortable', True),
                    facetable=field_def.get('facetable', False)
                ))
            elif field_type == 'Edm.Int32':
                fields.append(SimpleField(
                    name=field_name,
                    type=SearchFieldDataType.Int32,
                    filterable=field_def.get('filterable', True),
                    sortable=field_def.get('sortable', True),
                    facetable=field_def.get('facetable', False)
                ))
        
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
        
        for i in range(0, len(documents), batch_size):
            batch = documents[i:i + batch_size]
            
            # Clean documents for upload (remove None values and empty lists)
            cleaned_batch = []
            for doc in batch:
                cleaned_doc = {}
                for key, value in doc.items():
                    if value is not None:
                        if isinstance(value, list) and len(value) == 0:
                            continue
                        if isinstance(value, str) and value == '':
                            continue
                        cleaned_doc[key] = value
                cleaned_batch.append(cleaned_doc)
            
            try:
                result = search_client.upload_documents(documents=cleaned_batch)
                succeeded = sum(1 for r in result if r.succeeded)
                total_uploaded += succeeded
                print(f"  Batch {i//batch_size + 1}: {succeeded}/{len(batch)} documents uploaded")
            except Exception as e:
                print(f"  Batch {i//batch_size + 1}: Error - {e}")
        
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
