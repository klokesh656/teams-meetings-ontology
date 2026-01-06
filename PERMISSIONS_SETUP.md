# Azure AD Application Permissions Setup

## Required Permissions

Your application needs these Microsoft Graph API permissions:

### Application Permissions (Required)
1. **User.Read.All** - To verify users and list users in your organization
2. **OnlineMeetings.Read.All** - To read online meeting details
3. **CallRecords.Read.All** - To read call records (may be needed for some scenarios)

### Current Issue
You're getting `Authorization_RequestDenied` which means the `User.Read.All` permission is missing.

## How to Add Permissions

### Option 1: Azure Portal (Recommended)

1. **Go to Azure Portal**: https://portal.azure.com
2. **Navigate to Azure Active Directory** → **App registrations**
3. **Find your app**: Search for client ID `7b98108a-a799-45c1-aad6-93af90c1134c`
4. **Click on "API permissions"** in the left menu
5. **Click "Add a permission"**
6. **Select "Microsoft Graph"** → **Application permissions**
7. **Add these permissions**:
   - `User.Read.All` (under User)
   - `OnlineMeetings.Read.All` (under OnlineMeetings)
   - `CallRecords.Read.All` (under CallRecords) - Optional but recommended
8. **Click "Add permissions"**
9. **IMPORTANT**: Click **"Grant admin consent for [Your Tenant]"** button
10. Wait for status to show green checkmarks

### Option 2: Using PowerShell

```powershell
# Install Microsoft Graph PowerShell SDK (if not already installed)
Install-Module Microsoft.Graph -Scope CurrentUser

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Application.ReadWrite.All"

# Get your application
$appId = "7b98108a-a799-45c1-aad6-93af90c1134c"
$app = Get-MgApplication -Filter "appId eq '$appId'"

# Get Microsoft Graph service principal
$graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"

# Define required permissions
$permissions = @(
    @{
        ResourceAppId = "00000003-0000-0000-c000-000000000000" # Microsoft Graph
        ResourceAccess = @(
            @{
                Id = "df021288-bdef-4463-88db-98f22de89214" # User.Read.All
                Type = "Role" # Application permission
            },
            @{
                Id = "c1684f21-1984-47fa-9d61-2dc8c296bb70" # OnlineMeetings.Read.All
                Type = "Role"
            }
        )
    }
)

# Update application permissions
Update-MgApplication -ApplicationId $app.Id -RequiredResourceAccess $permissions

# Grant admin consent (requires admin privileges)
Write-Host "Permissions added. Now grant admin consent in the Azure Portal."
```

## Verify Permissions

After adding permissions and granting admin consent:

1. In Azure Portal, go to your app's "API permissions" page
2. Verify you see:
   - ✅ User.Read.All (with green checkmark)
   - ✅ OnlineMeetings.Read.All (with green checkmark)
   - Status should say "Granted for [Your Tenant]"

## Test After Adding Permissions

Once permissions are granted, test the script:

```powershell
# List users to verify User.Read.All permission
python src/transcript_extractor.py --list-users

# Run the full extraction
python src/transcript_extractor.py
```

## Troubleshooting

### "Admin consent required"
- Only Global Administrators can grant consent
- Contact your Azure AD admin if you don't have permission

### Permissions show but still get 403 errors
- Wait 5-10 minutes for permissions to propagate
- Clear any cached tokens
- Try creating a new client secret

### Still getting errors
- Check that the application is not blocked by Conditional Access policies
- Verify the tenant ID is correct
- Ensure the client secret hasn't expired

## Alternative: Delegated Permissions (User Context)

If you want to run this as a specific user instead of using application permissions:

1. Use **delegated permissions** instead:
   - `User.Read`
   - `OnlineMeetings.Read`
2. You'll need to implement interactive authentication (browser-based login)
3. The script would only access meetings for the signed-in user

Let me know if you need help implementing delegated permissions!
