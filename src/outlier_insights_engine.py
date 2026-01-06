"""
Outlier Insights Engine
========================
Automated system to replace weekly Outlier Discussion meetings.

This engine:
1. Analyzes check-in meeting transcripts for outlier-level insights
2. Detects VA/client churn risks automatically using the checklist
3. Generates AI suggestions based on patterns from past Outlier meetings
4. Supports a verification workflow where stakeholders approve suggestions
5. Learns from approved suggestions to improve future recommendations

Workflow:
---------
Check-in Meeting → AI Analysis → Outlier Detection → Suggestions → 
Stakeholder Review → Approval/Rejection → Knowledge Base Update

"""

import os
import re
import json
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

OUTPUT_DIR = Path('output')
KNOWLEDGE_BASE_FILE = OUTPUT_DIR / 'outlier_knowledge_base.json'
PENDING_SUGGESTIONS_FILE = OUTPUT_DIR / 'pending_suggestions.json'
APPROVED_SOLUTIONS_FILE = OUTPUT_DIR / 'approved_solutions.json'
ANALYSIS_HISTORY_FILE = OUTPUT_DIR / 'checkin_analysis_history.json'

# Initialize Azure OpenAI client
client = AzureOpenAI(
    azure_endpoint=AZURE_OPENAI_ENDPOINT,
    api_key=AZURE_OPENAI_KEY,
    api_version="2024-02-15-preview"
)


# =============================================================================
# CHURN RISK DETECTION FRAMEWORK (from Outlier meetings analysis)
# =============================================================================

CHURN_RISK_SIGNALS = {
    "va_signals": {
        "VA001": {"signal": "Non-responsive to integration team", "severity": "high", "category": "communication"},
        "VA002": {"signal": "Task scope expanded without acknowledgment", "severity": "high", "category": "workload"},
        "VA003": {"signal": "Mentioned feeling overwhelmed", "severity": "medium", "category": "emotional"},
        "VA004": {"signal": "Casual mention of resignation/leaving", "severity": "critical", "category": "retention"},
        "VA005": {"signal": "Attendance issues (tardiness, absences)", "severity": "medium", "category": "reliability"},
        "VA006": {"signal": "Technical/equipment problems unresolved", "severity": "medium", "category": "technical"},
        "VA007": {"signal": "Not visible to client (lack of SOD/EOD)", "severity": "high", "category": "visibility"},
        "VA008": {"signal": "Over-reliance on AI without sense-checking", "severity": "medium", "category": "performance"},
        "VA009": {"signal": "Requested compensation review", "severity": "medium", "category": "compensation"},
        "VA010": {"signal": "First 30 days - low engagement", "severity": "high", "category": "onboarding"},
        "VA011": {"signal": "MIA (Missing in Action) pattern", "severity": "critical", "category": "reliability"},
        "VA012": {"signal": "Coaching sessions with no progress", "severity": "medium", "category": "development"},
    },
    "client_signals": {
        "CL001": {"signal": "Slow/no response to OA communications", "severity": "high", "category": "engagement"},
        "CL002": {"signal": "Unilateral task changes without discussion", "severity": "medium", "category": "scope"},
        "CL003": {"signal": "Negative feedback delivered indirectly", "severity": "high", "category": "communication"},
        "CL004": {"signal": "Hired directly from another agency", "severity": "critical", "category": "competition"},
        "CL005": {"signal": "Internal layoffs or restructuring", "severity": "high", "category": "stability"},
        "CL006": {"signal": "Surprise cancellation after positive feedback", "severity": "critical", "category": "trust"},
        "CL007": {"signal": "Information hidden from client (felt blind-sided)", "severity": "high", "category": "transparency"},
        "CL008": {"signal": "Part-time client resistant to full-time", "severity": "low", "category": "growth"},
        "CL009": {"signal": "Multiple VA replacements requested", "severity": "high", "category": "satisfaction"},
        "CL010": {"signal": "No check-in for 30+ days", "severity": "medium", "category": "engagement"},
    },
    "relationship_signals": {
        "RH001": {"signal": "OA seen as vendor not partner", "severity": "high", "category": "perception"},
        "RH002": {"signal": "Feedback loop broken (no response to surveys)", "severity": "medium", "category": "communication"},
        "RH003": {"signal": "VA told client about issue before OA", "severity": "high", "category": "information_flow"},
        "RH004": {"signal": "Client competing for VA time (second job)", "severity": "medium", "category": "loyalty"},
        "RH005": {"signal": "Trust deficit - client verifying VA work", "severity": "medium", "category": "trust"},
    }
}

# =============================================================================
# SOLUTION PATTERNS (learned from Outlier Discussion meetings)
# =============================================================================

SOLUTION_PATTERNS = {
    "communication": [
        {"pattern": "Non-responsive VA", "solution": "Switch to WhatsApp with disappearing messages for urgent contact. Direct call if no response within 24h. Mark as orange until resolved.", "source": "Oct 31, 2025 Outlier Meeting"},
        {"pattern": "Silence from VA", "solution": "Random personalized check-ins with non-work questions to break ice. Silence is worse than complaints - treat as high risk.", "source": "Nov 10, 2025 Outlier Meeting"},
        {"pattern": "Client not responding", "solution": "Escalate to relationship manager. Schedule direct call. Document all communication attempts for contract review.", "source": "Nov 22, 2025 Outlier Meeting"},
    ],
    "workload": [
        {"pattern": "VA overwhelmed", "solution": "Conduct welfare check-in. Review task scope with client. Consider role alignment or additional support. Escalate feedback anonymously if needed.", "source": "Nov 10, 2025 Outlier Meeting"},
        {"pattern": "Scope creep", "solution": "Document original scope vs current tasks. Request alignment call with client. Discuss compensation adjustment if permanent change.", "source": "Nov 10, 2025 Outlier Meeting"},
    ],
    "reliability": [
        {"pattern": "Attendance issues", "solution": "Request medical certificate if health-related. Document pattern. Discuss root cause privately. Inform client if extended absence.", "source": "Nov 22, 2025 Outlier Meeting"},
        {"pattern": "MIA pattern", "solution": "Immediate alternative contact (WhatsApp, phone, emergency contact). Escalate within 24h. Prepare replacement candidate if 48h+ MIA.", "source": "Oct 31, 2025 Outlier Meeting"},
        {"pattern": "Technical/equipment issues", "solution": "Set 1-week deadline for resolution. Offer equipment support fund if needed. Prepare replacement if unresolved - laptop issues caused terminations.", "source": "Oct 24, 2025 Outlier Meeting"},
    ],
    "retention": [
        {"pattern": "Resignation mentioned", "solution": "Immediate private conversation. Understand root cause. If confirmed, delay informing client until replacement plan ready. Request 2-week notice minimum for handover.", "source": "Nov 22, 2025 Outlier Meeting"},
        {"pattern": "Job security concerns", "solution": "Provide proactive reassurance. Address 'yes culture' by explaining that honest concerns are valued. Clarify task changes are not performance-related.", "source": "Nov 22, 2025 Outlier Meeting"},
    ],
    "performance": [
        {"pattern": "AI over-reliance", "solution": "Training session on sense-checking AI outputs. Quality review of recent work. Mirror successful VA practices (e.g., Anne's approach).", "source": "Oct 24, 2025 Outlier Meeting"},
        {"pattern": "Comprehension issues", "solution": "Daily SOD/EOD check-ins. Clear written expectations. Pair with mentor. Regular progress reviews. Assess fit if no improvement after 2 weeks.", "source": "Oct 24, 2025 Outlier Meeting"},
        {"pattern": "Coaching no progress", "solution": "Change coaching approach. Consider if role is right fit. Document performance gap clearly. Prepare alternative plan.", "source": "Nov 27, 2025 Outlier Meeting"},
    ],
    "onboarding": [
        {"pattern": "First 30 days issues", "solution": "Week 1 Protocol: Immediate escalation of any concerns. Daily touchpoints. First 30 days behavior predicts long-term survival - intensive monitoring.", "source": "Nov 27, 2025 Outlier Meeting"},
        {"pattern": "Low engagement early", "solution": "Welfare check with personalized approach. Identify blockers. Consider reassignment if cultural mismatch.", "source": "Oct 31, 2025 Outlier Meeting"},
    ],
    "client_relationship": [
        {"pattern": "Client felt blind-sided", "solution": "Immediate transparency call. Share all relevant information. Rebuild trust through over-communication for next 30 days.", "source": "Nov 27, 2025 Outlier Meeting"},
        {"pattern": "Surprise cancellation", "solution": "Post-mortem analysis. Review last 60 days of communications for missed signals. Update detection patterns.", "source": "Oct 24, 2025 Outlier Meeting"},
        {"pattern": "Multiple replacements", "solution": "Root cause analysis. Review role requirements. Consider if client expectations are realistic. Discuss pattern with client directly.", "source": "Oct 24, 2025 Outlier Meeting"},
    ],
    "visibility": [
        {"pattern": "VA not visible", "solution": "Implement structured SOD/EOD reporting. Coach on daily visibility. Client may not know VA is working without updates.", "source": "Oct 24, 2025 Outlier Meeting"},
    ],
    "compensation": [
        {"pattern": "Comp review request", "solution": "Benchmark against market. Review if scope has expanded. Discuss growth path. Address transparently to prevent silent departure.", "source": "Nov 10, 2025 Outlier Meeting"},
    ],
    "contingency": [
        {"pattern": "Typhoon/disaster impact", "solution": "Contingency planning: Starlink for backup internet, alternative communication channels, offline work arrangements. Single point of failure risk.", "source": "Nov 15, 2025 Outlier Meeting"},
    ]
}

# =============================================================================
# KNOWLEDGE BASE MANAGEMENT
# =============================================================================

def load_knowledge_base():
    """Load the knowledge base with solution patterns and approved solutions"""
    base = {
        "churn_signals": CHURN_RISK_SIGNALS,
        "solution_patterns": SOLUTION_PATTERNS,
        "approved_solutions": [],
        "learning_examples": []
    }
    
    # Load approved solutions if exists
    if APPROVED_SOLUTIONS_FILE.exists():
        with open(APPROVED_SOLUTIONS_FILE, 'r', encoding='utf-8') as f:
            approved = json.load(f)
            base["approved_solutions"] = approved.get("solutions", [])
            base["learning_examples"] = approved.get("learning_examples", [])
    
    return base


def save_approved_solution(issue_context, suggestion, approval_status, stakeholder_notes=""):
    """Save an approved/rejected solution for future learning"""
    OUTPUT_DIR.mkdir(exist_ok=True)
    
    # Load existing
    data = {"solutions": [], "learning_examples": []}
    if APPROVED_SOLUTIONS_FILE.exists():
        with open(APPROVED_SOLUTIONS_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
    
    # Add new
    entry = {
        "id": hashlib.md5(f"{issue_context}{datetime.now().isoformat()}".encode()).hexdigest()[:12],
        "timestamp": datetime.now().isoformat(),
        "issue_context": issue_context,
        "ai_suggestion": suggestion,
        "approval_status": approval_status,  # "approved", "rejected", "modified"
        "stakeholder_notes": stakeholder_notes,
        "final_solution": suggestion if approval_status == "approved" else stakeholder_notes
    }
    
    data["solutions"].append(entry)
    
    # Add as learning example if approved or modified
    if approval_status in ["approved", "modified"]:
        data["learning_examples"].append({
            "issue": issue_context,
            "solution": entry["final_solution"],
            "approved_date": datetime.now().isoformat()
        })
    
    with open(APPROVED_SOLUTIONS_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    
    return entry


def get_relevant_solutions(category, signals_detected):
    """Get relevant solution patterns based on detected signals"""
    kb = load_knowledge_base()
    relevant = []
    
    # Get patterns from the category
    if category in kb["solution_patterns"]:
        relevant.extend(kb["solution_patterns"][category])
    
    # Add approved solutions that match the signals
    for solution in kb["approved_solutions"]:
        if solution["approval_status"] in ["approved", "modified"]:
            for signal in signals_detected:
                if signal.lower() in solution["issue_context"].lower():
                    relevant.append({
                        "pattern": solution["issue_context"],
                        "solution": solution["final_solution"],
                        "source": f"Approved Solution ({solution['timestamp'][:10]})"
                    })
    
    return relevant


# =============================================================================
# CHECK-IN TRANSCRIPT ANALYSIS
# =============================================================================

def analyze_checkin_for_outliers(transcript_text, va_name, client_name, meeting_date):
    """
    Analyze a check-in meeting transcript for outlier-level insights.
    Returns detected signals, risk assessment, and AI suggestions.
    """
    
    # Load knowledge base for context
    kb = load_knowledge_base()
    
    # Build the prompt with all our learned patterns
    solution_examples = []
    for category, patterns in kb["solution_patterns"].items():
        for p in patterns[:2]:  # Top 2 per category
            solution_examples.append(f"- Issue: {p['pattern']} → Solution: {p['solution']}")
    
    # Add recent approved solutions
    for sol in kb["approved_solutions"][-10:]:  # Last 10 approved
        if sol["approval_status"] in ["approved", "modified"]:
            solution_examples.append(f"- Issue: {sol['issue_context'][:50]}... → Solution: {sol['final_solution'][:100]}...")
    
    prompt = f"""Analyze this VA check-in meeting transcript and provide outlier-level insights.

VA NAME: {va_name}
CLIENT: {client_name}
DATE: {meeting_date}

TRANSCRIPT:
{transcript_text[:12000]}

ANALYZE FOR THESE CHURN RISK SIGNALS:

VA SIGNALS:
- VA001: Non-responsive to integration team (high)
- VA002: Task scope expanded without acknowledgment (high)
- VA003: Mentioned feeling overwhelmed (medium)
- VA004: Casual mention of resignation/leaving (critical)
- VA005: Attendance issues - tardiness, absences (medium)
- VA006: Technical/equipment problems unresolved (medium)
- VA007: Not visible to client - lack of SOD/EOD (high)
- VA008: Over-reliance on AI without sense-checking (medium)
- VA009: Requested compensation review (medium)
- VA010: First 30 days - low engagement (high)
- VA011: MIA - Missing in Action pattern (critical)
- VA012: Coaching sessions with no progress (medium)

CLIENT SIGNALS:
- CL001: Slow/no response to OA communications (high)
- CL002: Unilateral task changes without discussion (medium)
- CL003: Negative feedback delivered indirectly (high)
- CL004: Hired directly from another agency (critical)
- CL005: Internal layoffs or restructuring (high)
- CL006: Surprise cancellation after positive feedback (critical)
- CL007: Information hidden - felt blind-sided (high)
- CL008: Part-time resistant to full-time (low)
- CL009: Multiple VA replacements requested (high)
- CL010: No check-in for 30+ days (medium)

RELATIONSHIP SIGNALS:
- RH001: OA seen as vendor not partner (high)
- RH002: Feedback loop broken (medium)
- RH003: VA told client about issue before OA (high)
- RH004: Client competing for VA time - second job (medium)
- RH005: Trust deficit - client verifying work (medium)

SOLUTION PATTERNS FROM PAST OUTLIER MEETINGS:
{chr(10).join(solution_examples[:15])}

PROVIDE JSON OUTPUT:
{{
    "va_status": "green|yellow|orange|red",
    "client_health": "healthy|at_risk|critical",
    "overall_risk_level": "low|medium|high|critical",
    
    "detected_signals": [
        {{"signal_id": "VA001", "evidence": "Quote or description from transcript", "confidence": "high|medium|low"}}
    ],
    
    "key_findings": [
        "Finding 1: Brief summary of important issue or insight"
    ],
    
    "positive_indicators": [
        "Positive aspect noted in the meeting"
    ],
    
    "ai_suggestions": [
        {{
            "issue": "Brief description of the issue",
            "category": "communication|workload|reliability|retention|performance|onboarding|client_relationship|visibility|compensation|contingency",
            "urgency": "immediate|within_48h|this_week|monitor",
            "suggestion": "Detailed actionable suggestion based on past Outlier meeting solutions",
            "rationale": "Why this suggestion - reference to similar past situation"
        }}
    ],
    
    "escalation_needed": true|false,
    "escalation_reason": "Why escalation to executives is needed",
    
    "executive_summary": "2-3 sentence summary for Isaac/Crissy - what they would discuss in an Outlier meeting"
}}

Focus on actionable insights that would typically be discussed in an Outlier meeting between executives and HR."""

    try:
        response = client.chat.completions.create(
            model=AZURE_OPENAI_DEPLOYMENT,
            messages=[
                {"role": "system", "content": """You are an expert HR analyst specializing in VA (Virtual Assistant) churn risk detection.
Your role is to analyze check-in meetings and provide the same level of insights that executives would discuss in their weekly Outlier meetings.
Be thorough, identify subtle signals, and provide actionable suggestions based on proven solutions from past situations.
If there are no issues, still note positive indicators and preventive recommendations."""},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3,
            max_tokens=3000
        )
        
        result_text = response.choices[0].message.content.strip()
        
        # Clean JSON
        if result_text.startswith('```'):
            result_text = re.sub(r'^```json?\n?', '', result_text)
            result_text = re.sub(r'\n?```$', '', result_text)
        result_text = re.sub(r',\s*}', '}', result_text)
        result_text = re.sub(r',\s*]', ']', result_text)
        
        analysis = json.loads(result_text)
        
        # Add metadata
        analysis["va_name"] = va_name
        analysis["client_name"] = client_name
        analysis["meeting_date"] = meeting_date
        analysis["analyzed_at"] = datetime.now().isoformat()
        analysis["analysis_id"] = hashlib.md5(f"{va_name}{client_name}{meeting_date}".encode()).hexdigest()[:12]
        
        return analysis
        
    except Exception as e:
        print(f"   ❌ Analysis Error: {e}")
        return None


def save_analysis_result(analysis):
    """Save analysis to history for tracking and learning"""
    OUTPUT_DIR.mkdir(exist_ok=True)
    
    # Load history
    history = {"analyses": []}
    if ANALYSIS_HISTORY_FILE.exists():
        with open(ANALYSIS_HISTORY_FILE, 'r', encoding='utf-8') as f:
            history = json.load(f)
    
    # Add new analysis
    history["analyses"].append(analysis)
    history["last_updated"] = datetime.now().isoformat()
    
    with open(ANALYSIS_HISTORY_FILE, 'w', encoding='utf-8') as f:
        json.dump(history, f, indent=2, ensure_ascii=False)
    
    return analysis["analysis_id"]


def save_pending_suggestions(analysis):
    """Save suggestions for stakeholder review"""
    OUTPUT_DIR.mkdir(exist_ok=True)
    
    # Load pending
    pending = {"suggestions": []}
    if PENDING_SUGGESTIONS_FILE.exists():
        with open(PENDING_SUGGESTIONS_FILE, 'r', encoding='utf-8') as f:
            pending = json.load(f)
    
    # Add suggestions from this analysis
    for idx, suggestion in enumerate(analysis.get("ai_suggestions", [])):
        pending["suggestions"].append({
            "suggestion_id": f"{analysis['analysis_id']}_{idx}",
            "analysis_id": analysis["analysis_id"],
            "va_name": analysis["va_name"],
            "client_name": analysis["client_name"],
            "meeting_date": analysis["meeting_date"],
            "issue": suggestion["issue"],
            "category": suggestion["category"],
            "urgency": suggestion["urgency"],
            "suggestion": suggestion["suggestion"],
            "rationale": suggestion["rationale"],
            "status": "pending",  # pending, approved, rejected, modified
            "created_at": datetime.now().isoformat(),
            "reviewed_by": None,
            "reviewed_at": None,
            "stakeholder_notes": ""
        })
    
    pending["last_updated"] = datetime.now().isoformat()
    
    with open(PENDING_SUGGESTIONS_FILE, 'w', encoding='utf-8') as f:
        json.dump(pending, f, indent=2, ensure_ascii=False)


def generate_outlier_report(analysis):
    """Generate a report that replaces the Outlier meeting discussion"""
    
    report = f"""
================================================================================
📊 OUTLIER INSIGHTS REPORT - {analysis['meeting_date']}
================================================================================
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}
VA: {analysis['va_name']} | Client: {analysis['client_name']}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📌 EXECUTIVE SUMMARY (What Isaac/Crissy would discuss)
{analysis.get('executive_summary', 'No critical issues identified.')}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🎯 STATUS ASSESSMENT
   VA Status: {get_status_emoji(analysis.get('va_status', 'green'))} {analysis.get('va_status', 'green').upper()}
   Client Health: {analysis.get('client_health', 'healthy').upper()}
   Overall Risk: {analysis.get('overall_risk_level', 'low').upper()}
   Escalation Needed: {'⚠️ YES' if analysis.get('escalation_needed') else '✅ NO'}
"""
    
    if analysis.get('escalation_needed'):
        report += f"   Escalation Reason: {analysis.get('escalation_reason', 'N/A')}\n"
    
    # Detected Signals
    if analysis.get('detected_signals'):
        report += f"""
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🚨 DETECTED RISK SIGNALS ({len(analysis['detected_signals'])} found)
"""
        for signal in analysis['detected_signals']:
            sig_info = get_signal_info(signal['signal_id'])
            severity_emoji = {"critical": "🔴", "high": "🟠", "medium": "🟡", "low": "🟢"}.get(sig_info.get('severity', 'medium'), "⚪")
            report += f"""
   {severity_emoji} {signal['signal_id']}: {sig_info.get('signal', 'Unknown signal')}
      Evidence: "{signal['evidence'][:100]}..."
      Confidence: {signal['confidence']}
"""
    
    # Key Findings
    if analysis.get('key_findings'):
        report += f"""
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📋 KEY FINDINGS
"""
        for finding in analysis['key_findings']:
            report += f"   • {finding}\n"
    
    # Positive Indicators
    if analysis.get('positive_indicators'):
        report += f"""
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

✅ POSITIVE INDICATORS
"""
        for positive in analysis['positive_indicators']:
            report += f"   • {positive}\n"
    
    # AI Suggestions
    if analysis.get('ai_suggestions'):
        report += f"""
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

💡 AI-GENERATED SUGGESTIONS (Pending Stakeholder Review)
"""
        for idx, suggestion in enumerate(analysis['ai_suggestions'], 1):
            urgency_emoji = {"immediate": "🔴", "within_48h": "🟠", "this_week": "🟡", "monitor": "🟢"}.get(suggestion['urgency'], "⚪")
            report += f"""
   [{idx}] {urgency_emoji} {suggestion['urgency'].upper()} - {suggestion['category'].upper()}
       Issue: {suggestion['issue']}
       
       Suggestion: {suggestion['suggestion']}
       
       Rationale: {suggestion['rationale']}
       
       Status: ⏳ PENDING STAKEHOLDER APPROVAL
"""
    
    report += """
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📝 NEXT STEPS
   1. Review suggestions above
   2. Approve/Reject/Modify each suggestion using the approval tool
   3. Approved solutions will be added to the knowledge base for future AI learning

================================================================================
"""
    
    return report


def get_status_emoji(status):
    """Get emoji for status"""
    return {"green": "🟢", "yellow": "🟡", "orange": "🟠", "red": "🔴"}.get(status, "⚪")


def get_signal_info(signal_id):
    """Get signal information from the framework"""
    for category in CHURN_RISK_SIGNALS.values():
        if signal_id in category:
            return category[signal_id]
    return {"signal": "Unknown", "severity": "medium", "category": "unknown"}


# =============================================================================
# SUGGESTION APPROVAL WORKFLOW
# =============================================================================

def list_pending_suggestions():
    """List all pending suggestions for review"""
    if not PENDING_SUGGESTIONS_FILE.exists():
        return []
    
    with open(PENDING_SUGGESTIONS_FILE, 'r', encoding='utf-8') as f:
        pending = json.load(f)
    
    return [s for s in pending.get("suggestions", []) if s["status"] == "pending"]


def approve_suggestion(suggestion_id, stakeholder_name, notes="", modified_solution=None):
    """Approve a suggestion (optionally with modifications)"""
    if not PENDING_SUGGESTIONS_FILE.exists():
        return False
    
    with open(PENDING_SUGGESTIONS_FILE, 'r', encoding='utf-8') as f:
        pending = json.load(f)
    
    for suggestion in pending["suggestions"]:
        if suggestion["suggestion_id"] == suggestion_id:
            suggestion["status"] = "modified" if modified_solution else "approved"
            suggestion["reviewed_by"] = stakeholder_name
            suggestion["reviewed_at"] = datetime.now().isoformat()
            suggestion["stakeholder_notes"] = notes
            if modified_solution:
                suggestion["final_solution"] = modified_solution
            else:
                suggestion["final_solution"] = suggestion["suggestion"]
            
            # Save to approved solutions for learning
            save_approved_solution(
                issue_context=f"{suggestion['va_name']} at {suggestion['client_name']}: {suggestion['issue']}",
                suggestion=suggestion["suggestion"],
                approval_status=suggestion["status"],
                stakeholder_notes=suggestion.get("final_solution", suggestion["suggestion"])
            )
            break
    
    with open(PENDING_SUGGESTIONS_FILE, 'w', encoding='utf-8') as f:
        json.dump(pending, f, indent=2, ensure_ascii=False)
    
    return True


def reject_suggestion(suggestion_id, stakeholder_name, reason):
    """Reject a suggestion with reason"""
    if not PENDING_SUGGESTIONS_FILE.exists():
        return False
    
    with open(PENDING_SUGGESTIONS_FILE, 'r', encoding='utf-8') as f:
        pending = json.load(f)
    
    for suggestion in pending["suggestions"]:
        if suggestion["suggestion_id"] == suggestion_id:
            suggestion["status"] = "rejected"
            suggestion["reviewed_by"] = stakeholder_name
            suggestion["reviewed_at"] = datetime.now().isoformat()
            suggestion["stakeholder_notes"] = reason
            
            # Save rejection for learning (to avoid similar suggestions)
            save_approved_solution(
                issue_context=f"{suggestion['va_name']} at {suggestion['client_name']}: {suggestion['issue']}",
                suggestion=suggestion["suggestion"],
                approval_status="rejected",
                stakeholder_notes=reason
            )
            break
    
    with open(PENDING_SUGGESTIONS_FILE, 'w', encoding='utf-8') as f:
        json.dump(pending, f, indent=2, ensure_ascii=False)
    
    return True


# =============================================================================
# MAIN ANALYSIS FUNCTION
# =============================================================================

def analyze_checkin_meeting(transcript_path_or_text, va_name, client_name, meeting_date=None):
    """
    Main entry point: Analyze a check-in meeting for outlier-level insights.
    
    Args:
        transcript_path_or_text: Path to transcript file or transcript text
        va_name: Name of the VA
        client_name: Name of the client company
        meeting_date: Date of the meeting (optional, will try to parse from filename)
    
    Returns:
        Analysis result dictionary and printed report
    """
    print("\n" + "="*70)
    print("🔍 OUTLIER INSIGHTS ENGINE - Check-in Analysis")
    print("="*70)
    
    # Load transcript
    if os.path.exists(transcript_path_or_text):
        print(f"   📄 Loading transcript: {transcript_path_or_text}")
        with open(transcript_path_or_text, 'r', encoding='utf-8') as f:
            transcript_text = f.read()
        
        # Try to parse date from filename if not provided
        if not meeting_date:
            filename = os.path.basename(transcript_path_or_text)
            # Try common date patterns
            date_match = re.search(r'(\d{4}-\d{2}-\d{2})|(\d{2}-\d{2}-\d{4})', filename)
            if date_match:
                meeting_date = date_match.group()
    else:
        transcript_text = transcript_path_or_text
    
    meeting_date = meeting_date or datetime.now().strftime('%Y-%m-%d')
    
    print(f"   👤 VA: {va_name}")
    print(f"   🏢 Client: {client_name}")
    print(f"   📅 Date: {meeting_date}")
    print(f"\n   🤖 Analyzing with AI...")
    
    # Perform analysis
    analysis = analyze_checkin_for_outliers(transcript_text, va_name, client_name, meeting_date)
    
    if analysis:
        # Save to history
        save_analysis_result(analysis)
        
        # Save pending suggestions for approval
        if analysis.get("ai_suggestions"):
            save_pending_suggestions(analysis)
        
        # Generate and print report
        report = generate_outlier_report(analysis)
        print(report)
        
        # Save report to file
        report_path = OUTPUT_DIR / f"outlier_report_{va_name.replace(' ', '_')}_{meeting_date}.txt"
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(report)
        print(f"   📁 Report saved: {report_path}")
        
        return analysis
    else:
        print("   ❌ Analysis failed")
        return None


# =============================================================================
# CLI INTERFACE
# =============================================================================

def print_help():
    """Print help information"""
    print("""
╔══════════════════════════════════════════════════════════════════════════════╗
║                     OUTLIER INSIGHTS ENGINE - Help                          ║
╚══════════════════════════════════════════════════════════════════════════════╝

This system replaces weekly Outlier meetings by automatically analyzing 
check-in meetings and generating AI suggestions based on learned patterns.

COMMANDS:
---------
1. Analyze a check-in transcript:
   python src/outlier_insights_engine.py analyze <transcript_file> <va_name> <client_name>

2. List pending suggestions for review:
   python src/outlier_insights_engine.py pending

3. Approve a suggestion:
   python src/outlier_insights_engine.py approve <suggestion_id> <stakeholder_name> [notes]

4. Reject a suggestion:
   python src/outlier_insights_engine.py reject <suggestion_id> <stakeholder_name> <reason>

5. View analysis history:
   python src/outlier_insights_engine.py history

6. View knowledge base stats:
   python src/outlier_insights_engine.py stats

WORKFLOW:
---------
1. Check-in meeting happens → Transcript generated
2. Run: analyze <transcript> <va_name> <client>
3. Review generated report with AI suggestions
4. Stakeholders approve/reject/modify suggestions
5. Approved suggestions feed back into AI learning
6. Next analysis uses approved patterns for better suggestions

""")


def print_pending():
    """Print pending suggestions"""
    pending = list_pending_suggestions()
    
    if not pending:
        print("\n✅ No pending suggestions for review.\n")
        return
    
    print(f"\n📋 PENDING SUGGESTIONS FOR REVIEW ({len(pending)} total)\n")
    print("-" * 80)
    
    for s in pending:
        print(f"""
ID: {s['suggestion_id']}
VA: {s['va_name']} | Client: {s['client_name']} | Date: {s['meeting_date']}
Category: {s['category'].upper()} | Urgency: {s['urgency'].upper()}

Issue: {s['issue']}

Suggestion: {s['suggestion']}

Rationale: {s['rationale']}

Commands:
  approve {s['suggestion_id']} <your_name> [optional notes]
  reject {s['suggestion_id']} <your_name> <reason>
""")
        print("-" * 80)


def print_history():
    """Print analysis history"""
    if not ANALYSIS_HISTORY_FILE.exists():
        print("\n📭 No analysis history yet.\n")
        return
    
    with open(ANALYSIS_HISTORY_FILE, 'r', encoding='utf-8') as f:
        history = json.load(f)
    
    analyses = history.get("analyses", [])
    print(f"\n📊 ANALYSIS HISTORY ({len(analyses)} total)\n")
    
    for a in analyses[-10:]:  # Last 10
        signals_count = len(a.get('detected_signals', []))
        suggestions_count = len(a.get('ai_suggestions', []))
        print(f"  {a['meeting_date']} | {a['va_name'][:15]:15} | {a['client_name'][:15]:15} | "
              f"Status: {get_status_emoji(a.get('va_status', 'green'))} | "
              f"Signals: {signals_count} | Suggestions: {suggestions_count}")
    print()


def print_stats():
    """Print knowledge base statistics"""
    kb = load_knowledge_base()
    
    print("\n📈 KNOWLEDGE BASE STATISTICS\n")
    print(f"   Built-in Solution Patterns: {sum(len(v) for v in kb['solution_patterns'].values())}")
    print(f"   Approved Solutions: {len(kb['approved_solutions'])}")
    print(f"   Learning Examples: {len(kb['learning_examples'])}")
    
    # Count by category
    print("\n   Solution Patterns by Category:")
    for cat, patterns in kb['solution_patterns'].items():
        print(f"      {cat}: {len(patterns)}")
    
    print()


def main():
    """Main entry point"""
    import sys
    
    if len(sys.argv) < 2:
        print_help()
        return
    
    command = sys.argv[1].lower()
    
    if command == "help":
        print_help()
    
    elif command == "analyze":
        if len(sys.argv) < 5:
            print("Usage: python src/outlier_insights_engine.py analyze <transcript_file> <va_name> <client_name> [date]")
            return
        transcript = sys.argv[2]
        va_name = sys.argv[3]
        client_name = sys.argv[4]
        date = sys.argv[5] if len(sys.argv) > 5 else None
        analyze_checkin_meeting(transcript, va_name, client_name, date)
    
    elif command == "pending":
        print_pending()
    
    elif command == "approve":
        if len(sys.argv) < 4:
            print("Usage: python src/outlier_insights_engine.py approve <suggestion_id> <stakeholder_name> [notes]")
            return
        suggestion_id = sys.argv[2]
        stakeholder = sys.argv[3]
        notes = " ".join(sys.argv[4:]) if len(sys.argv) > 4 else ""
        if approve_suggestion(suggestion_id, stakeholder, notes):
            print(f"✅ Suggestion {suggestion_id} approved by {stakeholder}")
        else:
            print(f"❌ Could not find suggestion {suggestion_id}")
    
    elif command == "reject":
        if len(sys.argv) < 5:
            print("Usage: python src/outlier_insights_engine.py reject <suggestion_id> <stakeholder_name> <reason>")
            return
        suggestion_id = sys.argv[2]
        stakeholder = sys.argv[3]
        reason = " ".join(sys.argv[4:])
        if reject_suggestion(suggestion_id, stakeholder, reason):
            print(f"❌ Suggestion {suggestion_id} rejected by {stakeholder}")
        else:
            print(f"❌ Could not find suggestion {suggestion_id}")
    
    elif command == "history":
        print_history()
    
    elif command == "stats":
        print_stats()
    
    else:
        print(f"Unknown command: {command}")
        print_help()


if __name__ == '__main__':
    main()
