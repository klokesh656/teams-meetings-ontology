# Outlier Insights Automation System

## Overview

This system **eliminates the need for weekly Outlier Discussion meetings** between executives and HR by automatically analyzing check-in meeting transcripts and generating the same level of insights that would be discussed in those meetings.

### How It Works

```
Check-in Meeting → Transcript → AI Analysis → Outlier Detection → 
AI Suggestions → Stakeholder Review → Approve/Reject → Knowledge Base Learning
```

## Key Features

### 1. 🔍 Automatic Churn Risk Detection
- Analyzes transcripts for 27 defined risk signals (VA, Client, Relationship)
- Flags issues with severity levels (Critical, High, Medium, Low)
- Generates executive summaries suitable for Isaac/Crissy review

### 2. 💡 AI-Generated Suggestions
- Based on patterns learned from past Outlier Discussion meetings
- Provides actionable recommendations with urgency levels
- References similar past situations and successful interventions

### 3. ✅ Stakeholder Approval Workflow
- Suggestions are pending until reviewed
- Stakeholders can approve, reject, or modify suggestions
- Approved solutions feed back into AI learning

### 4. 📚 Continuous Learning
- Knowledge base grows with each approved solution
- AI improves suggestions based on verified outcomes
- Rejected suggestions help avoid future false positives

---

## Quick Start

### Analyze a Single Check-in
```bash
python src/outlier_insights_engine.py analyze <transcript_file> <va_name> <client_name> [date]

# Example:
python src/outlier_insights_engine.py analyze "transcripts/20251212_Checkin_Jep.vtt" "Jep" "Cesar" "2025-12-12"
```

### Review Pending Suggestions
```bash
python src/outlier_insights_engine.py pending
```

### Approve a Suggestion
```bash
python src/outlier_insights_engine.py approve <suggestion_id> <your_name> [optional notes]

# Example:
python src/outlier_insights_engine.py approve abc123_0 "Isaac" "Good suggestion - also schedule follow-up"
```

### Reject a Suggestion
```bash
python src/outlier_insights_engine.py reject <suggestion_id> <your_name> <reason>

# Example:
python src/outlier_insights_engine.py reject abc123_1 "Crissy" "Already handled by HR team"
```

### Run Batch Analysis
```bash
# Process all new transcripts from last 3 days
python src/auto_outlier_analyzer.py --days 3

# Interactive mode (prompts for missing info)
python src/auto_outlier_analyzer.py --interactive
```

### Or use the batch script:
```bash
run_outlier_analysis.bat
```

---

## Risk Signal Framework

### VA Signals (VA001-VA012)

| ID | Signal | Severity |
|----|--------|----------|
| VA001 | Non-responsive to integration team | High |
| VA002 | Task scope expanded without acknowledgment | High |
| VA003 | Mentioned feeling overwhelmed | Medium |
| VA004 | Casual mention of resignation/leaving | **Critical** |
| VA005 | Attendance issues (tardiness, absences) | Medium |
| VA006 | Technical/equipment problems unresolved | Medium |
| VA007 | Not visible to client (lack of SOD/EOD) | High |
| VA008 | Over-reliance on AI without sense-checking | Medium |
| VA009 | Requested compensation review | Medium |
| VA010 | First 30 days - low engagement | High |
| VA011 | MIA (Missing in Action) pattern | **Critical** |
| VA012 | Coaching sessions with no progress | Medium |

### Client Signals (CL001-CL010)

| ID | Signal | Severity |
|----|--------|----------|
| CL001 | Slow/no response to OA communications | High |
| CL002 | Unilateral task changes without discussion | Medium |
| CL003 | Negative feedback delivered indirectly | High |
| CL004 | Hired directly from another agency | **Critical** |
| CL005 | Internal layoffs or restructuring | High |
| CL006 | Surprise cancellation after positive feedback | **Critical** |
| CL007 | Information hidden - felt blind-sided | High |
| CL008 | Part-time resistant to full-time | Low |
| CL009 | Multiple VA replacements requested | High |
| CL010 | No check-in for 30+ days | Medium |

### Relationship Signals (RH001-RH005)

| ID | Signal | Severity |
|----|--------|----------|
| RH001 | OA seen as vendor not partner | High |
| RH002 | Feedback loop broken (no survey response) | Medium |
| RH003 | VA told client about issue before OA | High |
| RH004 | Client competing for VA time (second job) | Medium |
| RH005 | Trust deficit - client verifying work | Medium |

---

## Solution Patterns (From Outlier Meetings)

The AI suggestions are based on patterns learned from analyzing 7 Outlier Discussion meetings (Oct-Dec 2025). Examples:

### Communication Issues
- **Non-responsive VA**: Switch to WhatsApp, direct call within 24h, mark as orange
- **Client not responding**: Escalate to relationship manager, schedule direct call

### Retention Issues
- **Resignation mentioned**: Immediate private conversation, delay client notification until replacement ready
- **Job security concerns**: Proactive reassurance, address "yes culture", clarify changes not performance-related

### Performance Issues
- **AI over-reliance**: Training session, quality review, mirror successful VA practices
- **Comprehension issues**: Daily SOD/EOD check-ins, written expectations, pair with mentor

### Client Relationship Issues
- **Surprise cancellation**: Post-mortem analysis, review last 60 days for missed signals
- **Multiple replacements**: Root cause analysis, review role requirements

---

## Output Files

| File | Purpose |
|------|---------|
| `output/outlier_report_<VA>_<date>.txt` | Individual analysis reports |
| `output/daily_outlier_summary_<date>.txt` | Daily executive summary |
| `output/pending_suggestions.json` | Suggestions awaiting review |
| `output/approved_solutions.json` | Learning database of verified solutions |
| `output/checkin_analysis_history.json` | All past analyses |

---

## Daily Workflow

### For HR/Integration Team
1. **Morning**: Run `run_outlier_analysis.bat` after transcript sync
2. **Review**: Check daily summary for critical issues
3. **Action**: Handle immediate escalations

### For Stakeholders (Isaac/Crissy)
1. **Review**: `python src/outlier_insights_engine.py pending`
2. **Approve/Reject**: Each suggestion with your name
3. **Notes**: Add any modifications or context

### For the System
1. **Learning**: Approved solutions feed back into AI
2. **Improvement**: Future suggestions become more accurate
3. **Tracking**: All decisions logged for audit

---

## Integration with Existing Systems

### Power BI
Import these CSV files for dashboard:
- `output/churn_risk_va_dataset.csv`
- `output/churn_risk_client_dataset.csv`
- `output/churn_risk_kpi_summary.csv`

### Azure AI Search / Copilot
Use `output/copilot_churn_risk_index.json` for knowledge base queries:
- "What are the risk signals for non-responsive VAs?"
- "Show me all VAs at Atlas with issues"
- "What solutions worked for surprise cancellations?"

---

## Replacing Outlier Meetings

### Before (Manual Process)
1. HR collects issues from check-ins during the week
2. Weekly meeting with Isaac/Crissy to discuss outliers
3. Verbal decisions made, sometimes not documented
4. Actions tracked informally

### After (Automated Process)
1. ✅ Each check-in automatically analyzed in real-time
2. ✅ AI detects same issues that would be discussed
3. ✅ Suggestions generated based on past successful interventions
4. ✅ Stakeholders approve/reject asynchronously
5. ✅ All decisions documented and used for AI learning
6. ✅ Daily summary replaces weekly meeting

### Benefits
- **Immediate detection** vs waiting for weekly meeting
- **Consistent analysis** using defined risk framework
- **Documented decisions** for compliance and learning
- **Scalable** - handles any volume of check-ins
- **Learns from feedback** - improves over time

---

## Troubleshooting

### "Could not extract VA/client info"
- Use `--interactive` mode to manually enter info
- Or rename transcript files to include VA and client names

### "Analysis Error"
- Check Azure OpenAI credentials in `.env`
- Ensure transcript file is readable

### Suggestion seems wrong
- Reject with detailed reason
- This helps the AI avoid similar mistakes

---

## Future Enhancements

1. **Real-time integration** with Teams meeting completion
2. **Email alerts** for critical escalations
3. **Dashboard** for visual tracking of risk trends
4. **Slack/Teams notifications** for pending approvals
5. **Monthly learning reports** showing AI improvement

---

*System Version: 1.0*  
*Created: January 2026*  
*Based on analysis of 7 Outlier Discussion meetings (Oct-Dec 2025)*
