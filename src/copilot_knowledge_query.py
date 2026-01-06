"""
Copilot Agent Knowledge Base Query Interface
=============================================
This module provides functions to query the meeting transcripts knowledge base.
Can be used as a standalone service or integrated into a Copilot Agent.
"""

import json
import os
from datetime import datetime, timedelta
from typing import List, Dict, Optional, Any
from azure.storage.blob import BlobServiceClient
from dotenv import load_dotenv

load_dotenv()

# Configuration
AZURE_STORAGE_CONNECTION_STRING = os.getenv("AZURE_STORAGE_CONNECTION_STRING")
INDEX_CONTAINER = "copilot-index"
KNOWLEDGE_BASE_BLOB = "knowledge_base/copilot_knowledge_base_latest.json"

# Local cache
_knowledge_base_cache = None
_cache_timestamp = None
CACHE_TTL_SECONDS = 300  # 5 minutes


def load_knowledge_base(force_refresh: bool = False) -> Dict:
    """
    Load knowledge base from Azure Blob or local cache.
    Uses caching to avoid repeated downloads.
    """
    global _knowledge_base_cache, _cache_timestamp
    
    # Check cache
    if not force_refresh and _knowledge_base_cache is not None:
        if _cache_timestamp and (datetime.now() - _cache_timestamp).seconds < CACHE_TTL_SECONDS:
            return _knowledge_base_cache
    
    # Try to load from Azure Blob
    try:
        if AZURE_STORAGE_CONNECTION_STRING:
            blob_service_client = BlobServiceClient.from_connection_string(AZURE_STORAGE_CONNECTION_STRING)
            container_client = blob_service_client.get_container_client(INDEX_CONTAINER)
            blob_client = container_client.get_blob_client(KNOWLEDGE_BASE_BLOB)
            
            content = blob_client.download_blob().readall()
            _knowledge_base_cache = json.loads(content)
            _cache_timestamp = datetime.now()
            return _knowledge_base_cache
    except Exception as e:
        print(f"Warning: Could not load from Azure Blob: {e}")
    
    # Fallback to local file
    local_path = "output/copilot_knowledge_base_latest.json"
    if os.path.exists(local_path):
        with open(local_path, 'r', encoding='utf-8') as f:
            _knowledge_base_cache = json.load(f)
            _cache_timestamp = datetime.now()
            return _knowledge_base_cache
    
    raise FileNotFoundError("Knowledge base not found in Azure Blob or locally")


# ============================================================
# QUERY FUNCTIONS FOR COPILOT AGENT
# ============================================================

def get_meeting_statistics() -> Dict:
    """
    Get overall meeting statistics.
    
    Returns summary statistics about all meetings including averages
    for sentiment, churn risk, and other metrics.
    """
    kb = load_knowledge_base()
    return {
        "total_meetings": kb["statistics"]["total_meetings"],
        "average_scores": {
            "sentiment": kb["statistics"]["avg_sentiment"],
            "churn_risk": kb["statistics"]["avg_churn_risk"],
            "opportunity_score": kb["statistics"]["avg_opportunity_score"],
            "execution_reliability": kb["statistics"]["avg_execution_reliability"],
            "operational_complexity": kb["statistics"]["avg_operational_complexity"]
        },
        "meetings_by_month": kb["statistics"]["meetings_by_month"],
        "organizer_counts": kb["statistics"]["organizer_counts"],
        "sentiment_distribution": {
            "high": len(kb["by_sentiment"]["high"]),
            "medium": len(kb["by_sentiment"]["medium"]),
            "low": len(kb["by_sentiment"]["low"])
        },
        "churn_risk_distribution": {
            "high": len(kb["by_churn_risk"]["high"]),
            "medium": len(kb["by_churn_risk"]["medium"]),
            "low": len(kb["by_churn_risk"]["low"])
        }
    }


def get_meetings_by_date(date: str) -> List[Dict]:
    """
    Get all meetings for a specific date.
    
    Args:
        date: Date in YYYY-MM-DD format
        
    Returns:
        List of meetings on that date with their details
    """
    kb = load_knowledge_base()
    
    if date in kb["by_date"]:
        meeting_refs = kb["by_date"][date]
        meetings = []
        for ref in meeting_refs:
            meeting_id = ref["id"]
            # Find full meeting details
            for m in kb["meetings"]:
                if m["id"] == meeting_id:
                    meetings.append({
                        "subject": m["subject"],
                        "organizer": m["organizer"],
                        "time": m.get("time", ""),
                        "sentiment_score": m["sentiment_score"],
                        "churn_risk": m["churn_risk"],
                        "summary": m["summary"],
                        "action_items": m["action_items"]
                    })
                    break
        return meetings
    return []


def get_meetings_by_organizer(organizer: str) -> List[Dict]:
    """
    Get all meetings organized by a specific person.
    
    Args:
        organizer: Organizer name or email (partial match supported)
        
    Returns:
        List of meetings organized by that person
    """
    kb = load_knowledge_base()
    meetings = []
    
    organizer_lower = organizer.lower()
    for org_key, refs in kb["by_organizer"].items():
        if organizer_lower in org_key.lower():
            for ref in refs:
                meeting_id = ref["id"]
                for m in kb["meetings"]:
                    if m["id"] == meeting_id:
                        meetings.append({
                            "subject": m["subject"],
                            "date": m.get("date", ""),
                            "sentiment_score": m["sentiment_score"],
                            "churn_risk": m["churn_risk"],
                            "summary": m["summary"]
                        })
                        break
    return meetings


def get_high_churn_risk_meetings(threshold: int = 50) -> List[Dict]:
    """
    Get meetings with high churn risk.
    
    Args:
        threshold: Minimum churn risk score (default 50)
        
    Returns:
        List of high-risk meetings with details
    """
    kb = load_knowledge_base()
    high_risk = []
    
    for m in kb["meetings"]:
        if m["churn_risk"] >= threshold:
            high_risk.append({
                "subject": m["subject"],
                "date": m.get("date", ""),
                "organizer": m["organizer"],
                "churn_risk": m["churn_risk"],
                "key_concerns": m["key_concerns"],
                "summary": m["summary"]
            })
    
    # Sort by churn risk descending
    high_risk.sort(key=lambda x: x["churn_risk"], reverse=True)
    return high_risk


def get_low_sentiment_meetings(threshold: int = 50) -> List[Dict]:
    """
    Get meetings with low sentiment scores.
    
    Args:
        threshold: Maximum sentiment score (default 50)
        
    Returns:
        List of low-sentiment meetings with details
    """
    kb = load_knowledge_base()
    low_sentiment = []
    
    for m in kb["meetings"]:
        if 0 < m["sentiment_score"] < threshold:
            low_sentiment.append({
                "subject": m["subject"],
                "date": m.get("date", ""),
                "organizer": m["organizer"],
                "sentiment_score": m["sentiment_score"],
                "key_concerns": m["key_concerns"],
                "summary": m["summary"]
            })
    
    # Sort by sentiment ascending
    low_sentiment.sort(key=lambda x: x["sentiment_score"])
    return low_sentiment


def search_meetings(query: str, limit: int = 10) -> List[Dict]:
    """
    Search meetings by keyword in subject, summary, or action items.
    
    Args:
        query: Search query (case-insensitive)
        limit: Maximum number of results
        
    Returns:
        List of matching meetings
    """
    kb = load_knowledge_base()
    query_lower = query.lower()
    results = []
    
    for m in kb["meetings"]:
        searchable = m.get("searchable_content", "").lower()
        if query_lower in searchable:
            results.append({
                "subject": m["subject"],
                "date": m.get("date", ""),
                "organizer": m["organizer"],
                "sentiment_score": m["sentiment_score"],
                "summary": m["summary"],
                "action_items": m["action_items"],
                "events": m["events"]
            })
            if len(results) >= limit:
                break
    
    return results


def get_meeting_details(meeting_id: str) -> Optional[Dict]:
    """
    Get full details of a specific meeting.
    
    Args:
        meeting_id: Meeting ID (e.g., "meeting_0")
        
    Returns:
        Full meeting details or None if not found
    """
    kb = load_knowledge_base()
    
    for m in kb["meetings"]:
        if m["id"] == meeting_id:
            return m
    return None


def get_action_items(date_from: str = None, date_to: str = None) -> List[Dict]:
    """
    Get action items from meetings, optionally filtered by date range.
    
    Args:
        date_from: Start date (YYYY-MM-DD)
        date_to: End date (YYYY-MM-DD)
        
    Returns:
        List of action items with meeting context
    """
    kb = load_knowledge_base()
    action_items = []
    
    for m in kb["meetings"]:
        meeting_date = m.get("date", "")
        
        # Apply date filter if specified
        if date_from and meeting_date < date_from:
            continue
        if date_to and meeting_date > date_to:
            continue
        
        if m["action_items"]:
            action_items.append({
                "meeting_subject": m["subject"],
                "date": meeting_date,
                "organizer": m["organizer"],
                "action_items": m["action_items"]
            })
    
    return action_items


def get_key_concerns(limit: int = 20) -> List[Dict]:
    """
    Get key concerns from meetings with low sentiment or high churn risk.
    
    Args:
        limit: Maximum number of results
        
    Returns:
        List of key concerns with meeting context
    """
    kb = load_knowledge_base()
    concerns = []
    
    for m in kb["meetings"]:
        if m["key_concerns"] and (m["sentiment_score"] < 70 or m["churn_risk"] > 30):
            concerns.append({
                "meeting_subject": m["subject"],
                "date": m.get("date", ""),
                "sentiment_score": m["sentiment_score"],
                "churn_risk": m["churn_risk"],
                "key_concerns": m["key_concerns"]
            })
    
    # Sort by churn risk descending
    concerns.sort(key=lambda x: x["churn_risk"], reverse=True)
    return concerns[:limit]


def get_monthly_summary(month: str) -> Dict:
    """
    Get summary for a specific month.
    
    Args:
        month: Month in YYYY-MM format
        
    Returns:
        Monthly summary with meeting count, averages, and highlights
    """
    kb = load_knowledge_base()
    
    monthly_meetings = []
    for m in kb["meetings"]:
        meeting_date = m.get("date", "")
        if meeting_date.startswith(month):
            monthly_meetings.append(m)
    
    if not monthly_meetings:
        return {"error": f"No meetings found for {month}"}
    
    # Calculate averages
    total = len(monthly_meetings)
    avg_sentiment = sum(m["sentiment_score"] for m in monthly_meetings) / total
    avg_churn = sum(m["churn_risk"] for m in monthly_meetings) / total
    
    # Find highlights
    best_meeting = max(monthly_meetings, key=lambda x: x["sentiment_score"])
    worst_meeting = min(monthly_meetings, key=lambda x: x["sentiment_score"])
    highest_risk = max(monthly_meetings, key=lambda x: x["churn_risk"])
    
    return {
        "month": month,
        "total_meetings": total,
        "average_sentiment": round(avg_sentiment, 1),
        "average_churn_risk": round(avg_churn, 1),
        "best_meeting": {
            "subject": best_meeting["subject"],
            "sentiment": best_meeting["sentiment_score"]
        },
        "needs_attention": {
            "subject": worst_meeting["subject"],
            "sentiment": worst_meeting["sentiment_score"],
            "concerns": worst_meeting["key_concerns"]
        },
        "highest_risk": {
            "subject": highest_risk["subject"],
            "churn_risk": highest_risk["churn_risk"]
        }
    }


# ============================================================
# COPILOT AGENT TOOL DEFINITIONS
# ============================================================

COPILOT_TOOLS = [
    {
        "name": "get_meeting_statistics",
        "description": "Get overall meeting statistics including averages for sentiment, churn risk, and meeting counts by month and organizer.",
        "parameters": {}
    },
    {
        "name": "get_meetings_by_date",
        "description": "Get all meetings for a specific date.",
        "parameters": {
            "date": {
                "type": "string",
                "description": "Date in YYYY-MM-DD format"
            }
        }
    },
    {
        "name": "get_meetings_by_organizer",
        "description": "Get all meetings organized by a specific person.",
        "parameters": {
            "organizer": {
                "type": "string",
                "description": "Organizer name or email (partial match supported)"
            }
        }
    },
    {
        "name": "get_high_churn_risk_meetings",
        "description": "Get meetings with high churn risk that need attention.",
        "parameters": {
            "threshold": {
                "type": "integer",
                "description": "Minimum churn risk score (default 50)"
            }
        }
    },
    {
        "name": "get_low_sentiment_meetings",
        "description": "Get meetings with low sentiment scores that may indicate issues.",
        "parameters": {
            "threshold": {
                "type": "integer",
                "description": "Maximum sentiment score (default 50)"
            }
        }
    },
    {
        "name": "search_meetings",
        "description": "Search meetings by keyword in subject, summary, or action items.",
        "parameters": {
            "query": {
                "type": "string",
                "description": "Search query"
            },
            "limit": {
                "type": "integer",
                "description": "Maximum number of results (default 10)"
            }
        }
    },
    {
        "name": "get_action_items",
        "description": "Get action items from meetings, optionally filtered by date range.",
        "parameters": {
            "date_from": {
                "type": "string",
                "description": "Start date (YYYY-MM-DD)"
            },
            "date_to": {
                "type": "string",
                "description": "End date (YYYY-MM-DD)"
            }
        }
    },
    {
        "name": "get_key_concerns",
        "description": "Get key concerns from meetings with low sentiment or high churn risk.",
        "parameters": {
            "limit": {
                "type": "integer",
                "description": "Maximum number of results (default 20)"
            }
        }
    },
    {
        "name": "get_monthly_summary",
        "description": "Get summary for a specific month including meeting count, averages, and highlights.",
        "parameters": {
            "month": {
                "type": "string",
                "description": "Month in YYYY-MM format"
            }
        }
    }
]


def execute_tool(tool_name: str, parameters: Dict = None) -> Any:
    """
    Execute a Copilot tool by name with given parameters.
    
    Args:
        tool_name: Name of the tool to execute
        parameters: Dictionary of parameters
        
    Returns:
        Tool execution result
    """
    parameters = parameters or {}
    
    tool_map = {
        "get_meeting_statistics": get_meeting_statistics,
        "get_meetings_by_date": lambda: get_meetings_by_date(parameters.get("date", "")),
        "get_meetings_by_organizer": lambda: get_meetings_by_organizer(parameters.get("organizer", "")),
        "get_high_churn_risk_meetings": lambda: get_high_churn_risk_meetings(parameters.get("threshold", 50)),
        "get_low_sentiment_meetings": lambda: get_low_sentiment_meetings(parameters.get("threshold", 50)),
        "search_meetings": lambda: search_meetings(parameters.get("query", ""), parameters.get("limit", 10)),
        "get_action_items": lambda: get_action_items(parameters.get("date_from"), parameters.get("date_to")),
        "get_key_concerns": lambda: get_key_concerns(parameters.get("limit", 20)),
        "get_monthly_summary": lambda: get_monthly_summary(parameters.get("month", ""))
    }
    
    if tool_name in tool_map:
        return tool_map[tool_name]() if callable(tool_map[tool_name]) else tool_map[tool_name]
    else:
        return {"error": f"Unknown tool: {tool_name}"}


# ============================================================
# CLI INTERFACE
# ============================================================

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="Query Meeting Knowledge Base")
    parser.add_argument("--stats", action="store_true", help="Show overall statistics")
    parser.add_argument("--date", type=str, help="Get meetings for date (YYYY-MM-DD)")
    parser.add_argument("--organizer", type=str, help="Get meetings by organizer")
    parser.add_argument("--high-risk", action="store_true", help="Show high churn risk meetings")
    parser.add_argument("--low-sentiment", action="store_true", help="Show low sentiment meetings")
    parser.add_argument("--search", type=str, help="Search meetings by keyword")
    parser.add_argument("--concerns", action="store_true", help="Show key concerns")
    parser.add_argument("--month", type=str, help="Get monthly summary (YYYY-MM)")
    parser.add_argument("--actions", action="store_true", help="Show all action items")
    parser.add_argument("--tools", action="store_true", help="List available Copilot tools")
    
    args = parser.parse_args()
    
    if args.tools:
        print("\n🤖 Available Copilot Tools:")
        print("=" * 60)
        for tool in COPILOT_TOOLS:
            print(f"\n📌 {tool['name']}")
            print(f"   {tool['description']}")
            if tool['parameters']:
                print("   Parameters:")
                for param, info in tool['parameters'].items():
                    print(f"     - {param}: {info['description']}")
    
    elif args.stats:
        stats = get_meeting_statistics()
        print("\n📊 Meeting Statistics")
        print("=" * 60)
        print(json.dumps(stats, indent=2))
    
    elif args.date:
        meetings = get_meetings_by_date(args.date)
        print(f"\n📅 Meetings on {args.date}")
        print("=" * 60)
        print(json.dumps(meetings, indent=2))
    
    elif args.organizer:
        meetings = get_meetings_by_organizer(args.organizer)
        print(f"\n👤 Meetings by {args.organizer}")
        print("=" * 60)
        print(json.dumps(meetings, indent=2))
    
    elif args.high_risk:
        meetings = get_high_churn_risk_meetings()
        print("\n⚠️ High Churn Risk Meetings")
        print("=" * 60)
        print(json.dumps(meetings, indent=2))
    
    elif args.low_sentiment:
        meetings = get_low_sentiment_meetings()
        print("\n😟 Low Sentiment Meetings")
        print("=" * 60)
        print(json.dumps(meetings, indent=2))
    
    elif args.search:
        results = search_meetings(args.search)
        print(f"\n🔍 Search Results for '{args.search}'")
        print("=" * 60)
        print(json.dumps(results, indent=2))
    
    elif args.concerns:
        concerns = get_key_concerns()
        print("\n⚠️ Key Concerns")
        print("=" * 60)
        print(json.dumps(concerns, indent=2))
    
    elif args.month:
        summary = get_monthly_summary(args.month)
        print(f"\n📈 Monthly Summary for {args.month}")
        print("=" * 60)
        print(json.dumps(summary, indent=2))
    
    elif args.actions:
        actions = get_action_items()
        print("\n✅ Action Items")
        print("=" * 60)
        print(json.dumps(actions, indent=2))
    
    else:
        # Default: show statistics
        print("\n🤖 Meeting Knowledge Base Query Interface")
        print("=" * 60)
        print("Use --help to see available options")
        print("\nQuick Stats:")
        stats = get_meeting_statistics()
        print(f"  Total Meetings: {stats['total_meetings']}")
        print(f"  Avg Sentiment: {stats['average_scores']['sentiment']}")
        print(f"  Avg Churn Risk: {stats['average_scores']['churn_risk']}")
        print(f"  High Risk Meetings: {stats['churn_risk_distribution']['high']}")
