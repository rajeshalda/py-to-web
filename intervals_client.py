import requests
from typing import Optional, List, Dict, Any
from config import INTERVALS_API_BASE_URL
from utils import encode_basic_auth, clean_text, to_intervals_date

class IntervalsClient:
    def __init__(self, api_token: str):
        self.api_token = api_token
        # Ensure token ends with colon
        if not self.api_token.endswith(':'):
            self.api_token += ':'
        self.headers = {
            "Authorization": f"Basic {encode_basic_auth(self.api_token)}",
            "Content-Type": "application/json"
        }
        self.current_user = None
    
    @property
    def user_id(self) -> Optional[str]:
        """Get the current user's ID."""
        return self.current_user.get('personid') if self.current_user else None
    
    def get_current_user(self) -> Optional[Dict[str, Any]]:
        """Get current user information."""
        try:
            print(f"Making request to {INTERVALS_API_BASE_URL}/me")
            response = requests.get(
                f"{INTERVALS_API_BASE_URL}/me",
                headers=self.headers
            )
            print(f"Response status code: {response.status_code}")
            
            if response.status_code != 200:
                print(f"Error response: {response.text}")
                return None
                
            data = response.json()
            print(f"Response data: {data}")
            
            if data.get('me') and isinstance(data['me'], list) and len(data['me']) > 0:
                user = data['me'][0]
                # Ensure personid exists
                if not user.get("personid") and user.get("id"):
                    user["personid"] = user["id"]
                self.current_user = user
                print(f"Successfully authenticated as: {user.get('firstname')} {user.get('lastname')}")
                return user
            else:
                print("No user data found in response")
        except requests.exceptions.RequestException as e:
            print(f"Error fetching user: {str(e)}")
        return None
    
    def get_tasks(self) -> List[Dict[str, Any]]:
        """Get all tasks."""
        try:
            response = requests.get(
                f"{INTERVALS_API_BASE_URL}/task",
                headers=self.headers
            )
            response.raise_for_status()
            data = response.json()
            
            if isinstance(data.get('task'), list):
                return data['task']
            return []
        except requests.exceptions.RequestException as e:
            print(f"Error fetching tasks: {str(e)}")
        return []
    
    def get_project(self, project_id: int) -> Optional[Dict[str, Any]]:
        """Get project information."""
        try:
            response = requests.get(
                f"{INTERVALS_API_BASE_URL}/project/{project_id}",
                headers=self.headers
            )
            response.raise_for_status()
            data = response.json()
            
            if isinstance(data.get('project'), list) and len(data['project']) > 0:
                return data['project'][0]
            return None
        except requests.exceptions.RequestException as e:
            print(f"Error fetching project: {str(e)}")
        return None
    
    def post_time_entry(self, time_entry: Dict[str, Any]) -> bool:
        """Post a time entry."""
        try:
            # Clean description
            if "description" in time_entry:
                time_entry["description"] = clean_text(time_entry["description"], remove_emoji=True)
            
            response = requests.post(
                f"{INTERVALS_API_BASE_URL}/time",
                headers=self.headers,
                json=time_entry
            )
            response.raise_for_status()
            print(f"Successfully posted time entry ({time_entry.get('time', 0)} hours)")
            return True
        except requests.exceptions.RequestException as e:
            print(f"Error posting time entry: {str(e)}")
            return False 