# Azure Deployment Guide: VA Check-in Pipeline

This guide explains how to deploy the automated VA Check-in Analysis Pipeline to Azure.

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────────────┐
│                         Azure Cloud                                  │
├─────────────────────────────────────────────────────────────────────┤
│                                                                      │
│  ┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐  │
│  │  Timer Trigger  │───▶│  Azure Function │───▶│  Azure Blob     │  │
│  │  (Daily 6AM)    │    │  (Python)       │    │  Storage        │  │
│  └─────────────────┘    └────────┬────────┘    └────────┬────────┘  │
│                                  │                      │           │
│                                  ▼                      ▼           │
│                         ┌─────────────────┐    ┌─────────────────┐  │
│                         │  Azure OpenAI   │    │  CSV Files for  │  │
│                         │  (Analysis)     │    │  Power BI       │  │
│                         └─────────────────┘    └─────────────────┘  │
│                                                                      │
│  ┌─────────────────┐                                                │
│  │  Microsoft      │◀── Graph API ──▶ Download Transcripts          │
│  │  Graph API      │                                                │
│  └─────────────────┘                                                │
│                                                                      │
└─────────────────────────────────────────────────────────────────────┘

External Connections:
┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐
│  Power BI       │    │  Power Automate │    │  Teams/Outlook  │
│  (Dashboards)   │    │  (Alerts)       │    │  (Notifications)│
└─────────────────┘    └─────────────────┘    └─────────────────┘
```

## Pipeline Flow

1. **Timer Trigger** - Runs daily at 6:00 AM UTC
2. **Download Transcripts** - Fetches new meetings from Microsoft Graph API
3. **Analyze Check-ins** - Uses Azure OpenAI to detect churn risk signals
4. **Generate CSVs** - Creates files for Power BI/Power Automate
5. **Upload to Blob** - Stores results in Azure Blob Storage
6. **Power BI Refresh** - Dashboards auto-refresh from blob storage

---

## Option 1: Azure Function Deployment (Recommended)

### Prerequisites

- Azure CLI installed
- Azure Functions Core Tools (`npm install -g azure-functions-core-tools@4`)
- Python 3.9+ installed

### Step 1: Create Azure Resources

```powershell
# Login to Azure
az login

# Set subscription
az account set --subscription "YOUR_SUBSCRIPTION_ID"

# Create Resource Group
az group create --name rg-va-pipeline --location eastus

# Create Storage Account
az storage account create \
    --name stvapipeline \
    --resource-group rg-va-pipeline \
    --location eastus \
    --sku Standard_LRS

# Create Function App
az functionapp create \
    --name fn-va-checkin-pipeline \
    --resource-group rg-va-pipeline \
    --storage-account stvapipeline \
    --consumption-plan-location eastus \
    --runtime python \
    --runtime-version 3.9 \
    --functions-version 4 \
    --os-type Linux
```

### Step 2: Configure App Settings

```powershell
# Set environment variables
az functionapp config appsettings set \
    --name fn-va-checkin-pipeline \
    --resource-group rg-va-pipeline \
    --settings \
        "AZURE_TENANT_ID=187b2af6-1bfb-490a-85dd-b720fe3d31bc" \
        "AZURE_CLIENT_ID=7b98108a-a799-45c1-aad6-93af90c1134c" \
        "AZURE_CLIENT_SECRET=YOUR_CLIENT_SECRET" \
        "HR_USER_ID=81835016-79d5-4a15-91b1-c104e2cd9adb" \
        "AZURE_STORAGE_CONNECTION_STRING=YOUR_CONNECTION_STRING" \
        "STORAGE_ACCOUNT=aidevelopement" \
        "TRANSCRIPT_CONTAINER=transcripts" \
        "OUTPUT_CONTAINER=pipeline-outputs" \
        "AZURE_OPENAI_ENDPOINT=https://foundary-1-lokesh.cognitiveservices.azure.com/" \
        "AZURE_OPENAI_KEY=YOUR_OPENAI_KEY" \
        "AZURE_OPENAI_DEPLOYMENT=gpt-4.1" \
        "DAYS_TO_LOOK_BACK=7"
```

### Step 3: Deploy Function

```powershell
cd azure_function

# Deploy to Azure
func azure functionapp publish fn-va-checkin-pipeline
```

### Step 4: Verify Deployment

```powershell
# Check function status
az functionapp show --name fn-va-checkin-pipeline --resource-group rg-va-pipeline

# View logs
az webapp log tail --name fn-va-checkin-pipeline --resource-group rg-va-pipeline
```

---

## Option 2: Azure Container Instance (Alternative)

For more control, you can run the pipeline as a scheduled container.

### Step 1: Create Container Registry

```powershell
# Create ACR
az acr create --name acrvapipeline --resource-group rg-va-pipeline --sku Basic

# Login to ACR
az acr login --name acrvapipeline
```

### Step 2: Build and Push Docker Image

Create `Dockerfile`:

```dockerfile
FROM python:3.9-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY src/ ./src/
COPY .env .

CMD ["python", "src/daily_pipeline.py", "--days", "7"]
```

Build and push:

```powershell
docker build -t acrvapipeline.azurecr.io/va-pipeline:latest .
docker push acrvapipeline.azurecr.io/va-pipeline:latest
```

### Step 3: Create Container Instance with Schedule

```powershell
# Create Logic App for scheduling
az logic workflow create \
    --name la-va-pipeline-scheduler \
    --resource-group rg-va-pipeline \
    --definition @logic-app-definition.json
```

---

## Option 3: Local Scheduled Task (Development/Testing)

### Windows Task Scheduler

1. Open Task Scheduler
2. Create Basic Task: "VA Pipeline Daily Sync"
3. Trigger: Daily at 6:00 AM
4. Action: Start a program
   - Program: `python`
   - Arguments: `src/daily_pipeline.py --days 7`
   - Start in: `C:\path\to\project`

### PowerShell Script (Alternative)

```powershell
# run_pipeline.ps1
$env:Path = "C:\Python39;$env:Path"
Set-Location "C:\Users\ASUS\Desktop\Downloads\Downloads\New folder\Upwork\Issac - Copilot agent"
& .\.venv\Scripts\Activate.ps1
python src/daily_pipeline.py --days 7
```

Schedule with:
```powershell
$trigger = New-ScheduledTaskTrigger -Daily -At 6:00AM
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-File C:\path\to\run_pipeline.ps1"
Register-ScheduledTask -TaskName "VA Pipeline Daily" -Trigger $trigger -Action $action
```

---

## Power BI Integration

### Connect Power BI to Azure Blob Storage

1. Open Power BI Desktop
2. Get Data → Azure → Azure Blob Storage
3. Enter storage account: `aidevelopement`
4. Navigate to `pipeline-outputs/latest/`
5. Select CSV files:
   - `va_risk_summary.csv`
   - `all_meetings_detail.csv`
   - `critical_alerts.csv`
   - `kpi_dashboard_summary.csv`

### Set Up Auto-Refresh

1. Publish to Power BI Service
2. Dataset Settings → Scheduled Refresh
3. Configure Gateway (if needed)
4. Set refresh frequency: Daily after 7:00 AM

---

## Power Automate Integration

### Create Flow for Critical Alerts

1. New Flow → Scheduled Cloud Flow
2. Trigger: Daily at 7:30 AM (after pipeline completes)
3. Actions:
   - Get blob content: `pipeline-outputs/latest/critical_alerts.csv`
   - Parse CSV
   - For each row where Priority = "P1":
     - Send email to stakeholders
     - Create Teams message
     - Create Planner task

### Sample Flow Definition

```json
{
  "definition": {
    "triggers": {
      "Recurrence": {
        "type": "Recurrence",
        "recurrence": {
          "frequency": "Day",
          "interval": 1,
          "schedule": {
            "hours": ["7"],
            "minutes": ["30"]
          }
        }
      }
    },
    "actions": {
      "Get_CSV_from_Blob": {
        "type": "ApiConnection",
        "inputs": {
          "host": {
            "connection": {
              "name": "@parameters('$connections')['azureblob']['connectionId']"
            }
          },
          "method": "get",
          "path": "/datasets/default/files/@{encodeURIComponent('pipeline-outputs/latest/critical_alerts.csv')}/content"
        }
      }
    }
  }
}
```

---

## Monitoring & Alerts

### Application Insights

```powershell
# Enable Application Insights
az monitor app-insights component create \
    --app ai-va-pipeline \
    --location eastus \
    --resource-group rg-va-pipeline

# Link to Function App
az functionapp config appsettings set \
    --name fn-va-checkin-pipeline \
    --resource-group rg-va-pipeline \
    --settings "APPINSIGHTS_INSTRUMENTATIONKEY=YOUR_KEY"
```

### Alert Rules

Create alerts for:
- Function failures
- High execution time (>5 minutes)
- Critical risk cases detected

---

## Cost Estimate

| Resource | Estimated Monthly Cost |
|----------|----------------------|
| Azure Function (Consumption) | ~$5 (minimal executions) |
| Azure Blob Storage | ~$2 (small data) |
| Azure OpenAI | ~$10-30 (depends on transcripts) |
| Application Insights | ~$2 |
| **Total** | **~$20-40/month** |

---

## Troubleshooting

### Common Issues

1. **Authentication Errors**
   - Verify Azure AD app credentials
   - Check Microsoft Graph permissions

2. **Missing Transcripts**
   - Ensure HR_USER_ID is correct
   - Verify Graph API permissions: `OnlineMeetings.Read.All`

3. **OpenAI Rate Limits**
   - Pipeline processes max 30 transcripts per run
   - Increase interval if needed

4. **Blob Upload Failures**
   - Check storage connection string
   - Verify container exists and has write access

### Logs Location

- Azure Function: Application Insights / Log Stream
- Local: `logs/daily_report_YYYYMMDD.txt`
- Blob Storage: `pipeline-outputs/state/pipeline_state.json`

---

## Security Recommendations

1. **Use Azure Key Vault** for secrets
2. **Enable Managed Identity** for Azure resources
3. **Restrict network access** to storage accounts
4. **Enable audit logging** for all resources
5. **Rotate API keys** regularly

---

## Next Steps

1. Deploy the Azure Function
2. Configure Power BI dashboards
3. Set up Power Automate flows for alerts
4. Test end-to-end pipeline
5. Monitor first week of automated runs
