import re
from datetime import datetime, timedelta
import json
from typing import Optional, Dict, Any
import base64

def clean_text(text: str, remove_emoji: bool = False) -> str:
    """Clean and normalize text."""
    if not text:
        return "Untitled"
    
    # Remove emojis if specified
    if remove_emoji:
        text = re.sub(r'[^\x00-\x7F]+', '', text)
    
    # Clean whitespace
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'[\r\n]+', ' ', text)
    text = text.strip()
    
    return text if text else "Untitled"

def format_duration(total_seconds: int) -> str:
    """Format duration in seconds to human readable string."""
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60
    return f"{hours} hours, {minutes} minutes, {seconds} seconds"

def get_meeting_decimal_time(attendance_seconds: int) -> float:
    """Convert seconds to decimal hours."""
    exact_hours = attendance_seconds / 3600
    return round(exact_hours, 1)

def save_intervals_token(token: str) -> bool:
    """Save Intervals API token securely."""
    from config import TOKEN_FILE
    try:
        # In a production environment, use proper encryption
        with open(TOKEN_FILE, 'w') as f:
            f.write(base64.b64encode(token.encode()).decode())
        return True
    except Exception as e:
        print(f"Error saving token: {str(e)}")
        return False

def get_saved_intervals_token() -> Optional[str]:
    """Load saved Intervals API token."""
    from config import TOKEN_FILE
    try:
        with open(TOKEN_FILE, 'r') as f:
            encoded_token = f.read().strip()
            return base64.b64decode(encoded_token).decode()
    except Exception:
        return None

def parse_datetime(date_str: str) -> datetime:
    """Parse datetime string to datetime object."""
    return datetime.fromisoformat(date_str.replace('Z', '+00:00'))

def to_intervals_date(dt: datetime) -> str:
    """Convert datetime to Intervals API date format."""
    return dt.strftime('%Y-%m-%d')

def encode_basic_auth(token: str) -> str:
    """Encode token for basic auth."""
    if not token.endswith(':'):
        token = f"{token}:"
    return base64.b64encode(token.encode()).decode()

def parse_json_response(response_text: str) -> Dict[str, Any]:
    """Safely parse JSON response."""
    try:
        return json.loads(response_text)
    except json.JSONDecodeError:
        return {} 