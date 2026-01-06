"""
Analyze Outlier Discussion Transcripts and Search Calendar for Meeting Details
==============================================================================
This script analyzes the Outlier Discussion transcripts and searches for 
calendar events related to these executive meetings.
"""

import os
import re
import json
import requests
from datetime import datetime
from pathlib import Path
from azure.identity import ClientSecretCredential
from dotenv import load_dotenv

load_dotenv()

# Configuration
TENANT_ID = os.getenv('AZURE_TENANT_ID')
CLIENT_ID = os.getenv('AZURE_CLIENT_ID')
CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET')

HR_USER_ID = '81835016-79d5-4a15-91b1-c104e2cd9adb'

OUTPUT_DIR = Path('output')


class OutlierMeetingAnalyzer:
    def __init__(self):
        self.credential = ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
        self.token = None
        self.headers = None
        
    def authenticate(self):
        print("🔐 Authenticating with Microsoft Graph API...")
        self.token = self.credential.get_token('https://graph.microsoft.com/.default').token
        self.headers = {
            'Authorization': f'Bearer {self.token}',
            'Content-Type': 'application/json'
        }
        print("✅ Authenticated\n")
    
    def search_outlier_meetings(self):
        """Search for Outlier Discussion meetings in calendar"""
        print("="*70)
        print("SEARCHING FOR OUTLIER DISCUSSION MEETINGS IN CALENDAR")
        print("="*70)
        
        # Search calendar for outlier meetings
        filter_query = "start/dateTime ge '2025-10-01T00:00:00Z'"
        url = f"https://graph.microsoft.com/v1.0/users/{HR_USER_ID}/calendar/events"
        params = {
            '$filter': filter_query,
            '$select': 'id,subject,start,end,organizer,attendees,isOnlineMeeting,onlineMeeting,body,location',
            '$top': 500,
            '$orderby': 'start/dateTime desc'
        }
        
        outlier_meetings = []
        all_meetings_sample = []
        
        try:
            resp = requests.get(url, headers=self.headers, params=params, timeout=60)
            if resp.status_code == 200:
                data = resp.json()
                events = data.get('value', [])
                print(f"   Found {len(events)} calendar events since Oct 1, 2025")
                
                for event in events:
                    subject = event.get('subject', '').lower()
                    
                    # Collect sample of meeting subjects
                    if len(all_meetings_sample) < 50:
                        all_meetings_sample.append({
                            'subject': event.get('subject', ''),
                            'date': event.get('start', {}).get('dateTime', '')[:10]
                        })
                    
                    # Search for outlier-related meetings
                    if any(keyword in subject for keyword in ['outlier', 'executive', 'management', 'leadership', 'review', 'discussion', 'sync', 'weekly']):
                        start = event.get('start', {}).get('dateTime', '')[:10]
                        start_time = event.get('start', {}).get('dateTime', '')[11:19]
                        
                        attendees = event.get('attendees', [])
                        attendee_list = []
                        for a in attendees:
                            email = a.get('emailAddress', {})
                            attendee_list.append({
                                'name': email.get('name', ''),
                                'email': email.get('address', '')
                            })
                        
                        organizer = event.get('organizer', {}).get('emailAddress', {})
                        
                        outlier_meetings.append({
                            'subject': event.get('subject', ''),
                            'date': start,
                            'time': start_time,
                            'organizer': {
                                'name': organizer.get('name', ''),
                                'email': organizer.get('address', '')
                            },
                            'attendees': attendee_list,
                            'is_online': event.get('isOnlineMeeting', False),
                            'location': event.get('location', {}).get('displayName', ''),
                            'event_id': event.get('id', '')
                        })
                
                print(f"\n   ✅ Found {len(outlier_meetings)} potential outlier-related meetings")
                
            else:
                print(f"   ❌ Error: {resp.status_code} - {resp.text[:200]}")
                
        except Exception as e:
            print(f"   ❌ Exception: {e}")
        
        # Print sample of all meetings
        print("\n   📅 Sample of recent meetings (for reference):")
        for m in all_meetings_sample[:20]:
            print(f"      {m['date']}: {m['subject'][:60]}")
        
        if outlier_meetings:
            print("\n   🎯 Outlier-related meetings found:")
            for m in outlier_meetings[:15]:
                print(f"      {m['date']} {m['time']}: {m['subject'][:50]}")
                if m['attendees']:
                    print(f"         Attendees: {', '.join([a['name'] for a in m['attendees'][:5]])}")
        
        return outlier_meetings, all_meetings_sample
    
    def run(self):
        self.authenticate()
        outlier_meetings, all_meetings = self.search_outlier_meetings()
        
        # Save results
        results = {
            'search_date': datetime.now().isoformat(),
            'outlier_meetings': outlier_meetings,
            'meeting_sample': all_meetings
        }
        
        OUTPUT_DIR.mkdir(exist_ok=True)
        output_file = OUTPUT_DIR / 'outlier_meeting_search.json'
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        
        print(f"\n💾 Results saved to: {output_file}")
        
        return results


if __name__ == '__main__':
    analyzer = OutlierMeetingAnalyzer()
    analyzer.run()
