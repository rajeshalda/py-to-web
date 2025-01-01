import os
import json
import base64
import msal
import requests
from typing import Dict, List, Optional
from datetime import datetime, timedelta
from config import TENANT_ID, CLIENT_ID, GRAPH_SCOPES
from utils import parse_datetime, clean_text
import re
import urllib.parse

class GraphClient:
    def __init__(self):
        """Initialize the Graph client."""
        self.base_url = "https://graph.microsoft.com/v1.0"
        self.user_id = None
        self.access_token = None
        self.session = requests.Session()
        
    def authenticate(self) -> bool:
        """Authenticate with Microsoft Graph."""
        try:
            app = msal.PublicClientApplication(
                client_id=CLIENT_ID,
                authority=f"https://login.microsoftonline.com/{TENANT_ID}"
            )
            
            # Try to get token silently first
            accounts = app.get_accounts()
            if accounts:
                result = app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])
            else:
                result = app.acquire_token_interactive(scopes=GRAPH_SCOPES)
            
            if "access_token" not in result:
                print("Authentication error:", result.get("error_description", "Unknown error"))
                return False
            
            self.access_token = result["access_token"]
            self.session.headers.update({
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/json"
            })
            
            # Get user info using /me endpoint
            me = self.session.get(f"{self.base_url}/me")
            if me.status_code != 200:
                print("Failed to get user info:", me.status_code, me.reason)
                return False
            
            user_data = me.json()
            self.user_id = user_data.get("id")  # Use id instead of userPrincipalName
            
            if not self.user_id:
                print("Failed to get user ID")
                return False
            
            print(f"Successfully authenticated as: {user_data.get('userPrincipalName')}")
            return True
            
        except Exception as e:
            print("Authentication error:", str(e))
            return False
    
    def _make_request(self, endpoint: str, method: str = "GET", params: Dict = None) -> Optional[Dict]:
        """Make a request to Microsoft Graph API."""
        if not self.access_token:
            print("No access token available. Please authenticate first.")
            return None
            
        headers = {"Authorization": f"Bearer {self.access_token}"}
        url = f"https://graph.microsoft.com/v1.0{endpoint}"
        
        try:
            response = requests.request(method, url, headers=headers, params=params)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            print(f"API request error: {str(e)}")
            return None
    
    def get_user_meetings(self) -> List[Dict]:
        """Get user's meetings from the past 30 days."""
        try:
            # Get user information first
            user_info = self.session.get(f"{self.base_url}/me").json()
            if not user_info:
                print("Failed to retrieve user information")
                return []
            
            # Calculate date range
            end_time = datetime.utcnow()
            start_time = end_time - timedelta(days=30)
            
            print(f"Target User ID: {self.user_id}")
            print(f"Retrieving meetings from {start_time.isoformat()}Z to {end_time.isoformat()}Z...")
            
            # Build filter for date range
            date_filter = f"start/dateTime ge '{start_time.isoformat()}Z' and end/dateTime le '{end_time.isoformat()}Z'"
            
            # Get meetings
            meetings_url = f"{self.base_url}/users/{self.user_id}/events"
            params = {
                "$filter": date_filter,
                "$select": "subject,start,end,onlineMeeting,bodyPreview"
            }
            
            meetings_response = self.session.get(meetings_url, params=params)
            if meetings_response.status_code != 200:
                print(f"Error getting meetings: {meetings_response.status_code} {meetings_response.reason}")
                return []
            
            meetings_data = meetings_response.json()
            processed_meetings = []
            
            for meeting in meetings_data.get("value", []):
                print(f"\nProcessing meeting: {meeting['subject']}")
                if meeting.get("bodyPreview"):
                    print(f"Meeting Description:\n{meeting['bodyPreview']}")
                
                if meeting.get("onlineMeeting"):
                    # Get attendance data
                    start_time = parse_datetime(meeting["start"]["dateTime"])
                    end_time = parse_datetime(meeting["end"]["dateTime"])
                    attendance_data = self.get_meeting_attendance(
                        meeting["onlineMeeting"]["joinUrl"],
                        start_time,
                        end_time
                    )
                    
                    if attendance_data:
                        # Add attendance description to meeting
                        attendance_desc = ["Attendance Report"]
                        for attendee in attendance_data["attendees"]:
                            name = attendee["name"]
                            email = attendee["email"]
                            duration = attendee["duration"]
                            attendance_desc.append(f"User: {name}, Email: {email}, Total Time: {duration} seconds")
                        meeting["description"] = "\n".join(attendance_desc)
                    else:
                        # Use scheduled duration as fallback
                        duration_seconds = int((end_time - start_time).total_seconds())
                        attendance_desc = [
                            "Attendance Report",
                            "Note: Using scheduled duration.",
                            f"User: {user_info['displayName']}, Email: {self.user_id}, Total Time: {duration_seconds} seconds"
                        ]
                        meeting["description"] = "\n".join(attendance_desc)
                else:
                    # Handle non-Teams meetings
                    start_time = parse_datetime(meeting["start"]["dateTime"])
                    end_time = parse_datetime(meeting["end"]["dateTime"])
                    duration_seconds = int((end_time - start_time).total_seconds())
                    
                    attendance_desc = [
                        "Attendance Report",
                        "Note: Not a Teams meeting - using scheduled duration.",
                        f"User: {user_info['displayName']}, Email: {self.user_id}, Total Time: {duration_seconds} seconds"
                    ]
                    meeting["description"] = "\n".join(attendance_desc)
                    
                processed_meetings.append(meeting)
                
            return processed_meetings
        except Exception as e:
            print(f"Error fetching meetings: {str(e)}")
            return []
            
    def get_meeting_attendance(self, meeting_url, start_time, end_time):
        """Get attendance report for a meeting."""
        try:
            # Extract meeting ID and organizer ID from the URL
            print(f"Debug - Meeting URL: {meeting_url}")
            decoded_url = urllib.parse.unquote(meeting_url)
            
            # Extract meeting ID
            meeting_id_match = re.search(r"19:meeting_([^@]+)@thread\.v2", decoded_url)
            if not meeting_id_match:
                raise ValueError("Could not extract meeting ID from URL")
            meeting_id = f"19:meeting_{meeting_id_match.group(1)}@thread.v2"
            print(f"Debug - Extracted Meeting ID: {meeting_id}")
            
            # Extract organizer ID
            organizer_match = re.search(r'"Oid":"([^"]+)"', decoded_url)
            if not organizer_match:
                raise ValueError("Could not extract organizer ID from URL")
            organizer_oid = organizer_match.group(1)
            print(f"Debug - Organizer OID: {organizer_oid}")
            
            # Format the meeting ID as required
            formatted_string = f"1*{organizer_oid}*0**{meeting_id}"
            base64_meeting_id = base64.b64encode(formatted_string.encode('utf-8')).decode('utf-8')
            print(f"Debug - Base64 Meeting ID: {base64_meeting_id}")
            
            try:
                # Get attendance reports
                reports_url = f"{self.base_url}/users/{self.user_id}/onlineMeetings/{base64_meeting_id}/attendanceReports"
                reports_response = self.session.get(reports_url)
                reports_response.raise_for_status()
                reports_data = reports_response.json()
                
                if not reports_data.get('value'):
                    raise ValueError("No attendance reports found")
                
                report_id = reports_data['value'][0]['id']
                print(f"Debug - Report ID: {report_id}")
                
                # Get attendance records
                records_url = f"{self.base_url}/users/{self.user_id}/onlineMeetings/{base64_meeting_id}/attendanceReports/{report_id}/attendanceRecords"
                records_response = self.session.get(records_url)
                records_response.raise_for_status()
                records_data = records_response.json()
                
                if not records_data.get('value'):
                    raise ValueError("No attendance records found")
                
                return {
                    'attendees': [
                        {
                            'name': record['identity'].get('displayName', 'Unknown User'),
                            'email': record.get('emailAddress', 'No Email'),
                            'duration': record['totalAttendanceInSeconds']
                        }
                        for record in records_data['value']
                    ]
                }
            
            except requests.exceptions.HTTPError as e:
                if e.response.status_code == 403:
                    print(f"API request error: {str(e)}")
                    # Fall back to scheduled duration with current user info
                    duration_seconds = int((end_time - start_time).total_seconds())
                    me_response = self.session.get(f"{self.base_url}/me")
                    if me_response.status_code == 200:
                        user_data = me_response.json()
                        return {
                            'attendees': [
                                {
                                    'name': user_data.get('displayName', 'Current User'),
                                    'email': user_data.get('userPrincipalName', self.user_id),
                                    'duration': duration_seconds
                                }
                            ]
                        }
                raise
        
        except Exception as e:
            print(f"Error getting attendance: {str(e)}")
            # Fall back to scheduled duration with current user info
            duration_seconds = int((end_time - start_time).total_seconds())
            me_response = self.session.get(f"{self.base_url}/me")
            if me_response.status_code == 200:
                user_data = me_response.json()
                return {
                    'attendees': [
                        {
                            'name': user_data.get('displayName', 'Current User'),
                            'email': user_data.get('userPrincipalName', self.user_id),
                            'duration': duration_seconds
                        }
                    ]
                }
            return {
                'attendees': [
                    {
                        'name': 'Current User',
                        'email': self.user_id,
                        'duration': duration_seconds
                    }
                ]
            } 