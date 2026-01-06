"""
Check current permissions for the Azure AD application
"""

import os
import asyncio
from dotenv import load_dotenv
from azure.identity import ClientSecretCredential
from msgraph import GraphServiceClient


async def check_permissions():
    """Check what permissions the application currently has."""
    load_dotenv()
    
    client_id = os.getenv("AZURE_CLIENT_ID")
    client_secret = os.getenv("AZURE_CLIENT_SECRET")
    tenant_id = os.getenv("AZURE_TENANT_ID")
    
    if not all([client_id, client_secret, tenant_id]):
        print("❌ Missing credentials in .env file")
        return
    
    print("Checking Microsoft Graph API Permissions...")
    print("=" * 60)
    print(f"Application ID: {client_id}")
    print(f"Tenant ID: {tenant_id}")
    print("=" * 60)
    
    credential = ClientSecretCredential(
        tenant_id=tenant_id,
        client_id=client_id,
        client_secret=client_secret
    )
    graph_client = GraphServiceClient(credentials=credential)
    
    permissions_status = {
        "User.Read.All": False,
        "OnlineMeetings.Read.All": False,
        "OnlineMeetings.ReadWrite.All": False,
        "CallRecords.Read.All": False
    }
    
    # Test User.Read.All
    print("\nTesting: User.Read.All")
    try:
        users = await graph_client.users.get()
        if users and users.value:
            permissions_status["User.Read.All"] = True
            print("  ✅ GRANTED - Can read user information")
        else:
            print("  ⚠ GRANTED but no users returned")
    except Exception as e:
        error_msg = str(e)
        if "403" in error_msg or "Authorization" in error_msg:
            print("  ❌ MISSING - Cannot read users")
        else:
            print(f"  ⚠ Error: {e}")
    
    # Test OnlineMeetings.Read.All
    print("\nTesting: OnlineMeetings.Read.All")
    try:
        # Try to get meetings for a test user
        users = await graph_client.users.get()
        if users and users.value:
            test_user = users.value[0]
            meetings = await graph_client.users.by_user_id(test_user.id).online_meetings.get()
            permissions_status["OnlineMeetings.Read.All"] = True
            print("  ✅ GRANTED - Can read online meetings")
            if meetings and meetings.value:
                print(f"     Found {len(meetings.value)} meeting(s) for {test_user.user_principal_name}")
            else:
                print("     No meetings found (user may have no meetings)")
    except Exception as e:
        error_msg = str(e)
        if "403" in error_msg or "Authorization" in error_msg or "Forbidden" in error_msg:
            print("  ❌ MISSING - Cannot read online meetings")
            print("     This is the permission you need to add!")
        elif "404" in error_msg:
            print("  ⚠ Endpoint not found - May need different permission")
        else:
            print(f"  ⚠ Error: {e}")
    
    # Summary
    print("\n" + "=" * 60)
    print("PERMISSION SUMMARY:")
    print("=" * 60)
    
    all_granted = True
    for perm, status in permissions_status.items():
        icon = "✅" if status else "❌"
        print(f"{icon} {perm}")
        if not status and "OnlineMeetings" in perm:
            all_granted = False
    
    print("\n" + "=" * 60)
    
    if all_granted or permissions_status["OnlineMeetings.Read.All"]:
        print("✅ All required permissions are granted!")
        print("\nYou can now run: python src/transcript_extractor.py")
    else:
        print("❌ Missing required permissions!")
        print("\nTO FIX:")
        print("1. Go to: https://portal.azure.com")
        print("2. Navigate: Azure AD → App registrations → Your app")
        print("3. Click: API permissions → Add a permission")
        print("4. Select: Microsoft Graph → Application permissions")
        print("5. Add: OnlineMeetings.Read.All")
        print("6. Click: Grant admin consent for [Your Organization]")
        print("\nSee PERMISSIONS_SETUP.md for detailed instructions")


if __name__ == "__main__":
    asyncio.run(check_permissions())
