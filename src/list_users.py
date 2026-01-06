#!/usr/bin/env python3
"""List all users from Microsoft 365 tenant."""

import asyncio
import os
from dotenv import load_dotenv
from msgraph import GraphServiceClient
from azure.identity import ClientSecretCredential

load_dotenv()

credential = ClientSecretCredential(
    os.getenv('AZURE_TENANT_ID'),
    os.getenv('AZURE_CLIENT_ID'),
    os.getenv('AZURE_CLIENT_SECRET')
)
client = GraphServiceClient(credential)

async def get_users():
    users = await client.users.get()
    print('=' * 70)
    print('USERS IN MICROSOFT 365 TENANT')
    print('=' * 70)
    for i, u in enumerate(users.value, 1):
        print(f'{i}. {u.display_name}')
        print(f'   Email: {u.mail or u.user_principal_name}')
        print(f'   ID: {u.id}')
        print()
    print(f'Total Users: {len(users.value)}')

if __name__ == '__main__':
    asyncio.run(get_users())
