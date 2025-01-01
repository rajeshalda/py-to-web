# Teams Meeting Time Tracker

This Python application automatically tracks Microsoft Teams meetings and logs time entries to Intervals. It processes meeting attendance data and matches meetings to corresponding tasks in Intervals for efficient time tracking.

## Features

- 📊 Automatic meeting attendance tracking
- ⏱️ Billable hours calculation based on actual attendance
- 🔄 Automatic task matching using AI
- 📝 Detailed attendance reports
- ⚡ Bulk processing of meetings
- 🎯 Integration with Intervals for time tracking

## Prerequisites

- Python 3.8 or higher
- Microsoft 365 account with Teams
- Intervals account with API access
- Required Python packages (see `requirements.txt`)

## Installation

1. Clone the repository:
   ```bash
   git clone <repository-url>
   cd teams-meeting-tracker
   ```

2. Create and activate a virtual environment (recommended):
   ```bash
   python -m venv venv
   
   # Windows
   .\venv\Scripts\activate
   
   # Linux/Mac
   source venv/bin/activate
   ```

3. Install required packages:
   ```bash
   pip install -r requirements.txt
   ```

## Configuration

1. Create a `.env` file in the project root with the following variables:
   ```env
   # Microsoft Graph API credentials
   MS_CLIENT_ID=your_ms_client_id
   MS_CLIENT_SECRET=your_ms_client_secret
   MS_TENANT_ID=your_tenant_id
   
   # Intervals API credentials
   INTERVALS_API_KEY=your_intervals_api_key
   INTERVALS_ACCOUNT_URL=your_intervals_account_url
   ```

2. Set up Microsoft Graph API:
   - Register an application in Azure AD
   - Grant required permissions:
     - Calendars.Read
     - OnlineMeetings.Read.All
     - User.Read
   - Add redirect URI for authentication
   - Copy Client ID and Secret to `.env` file

3. Set up Intervals API:
   - Generate an API key in your Intervals account
   - Copy the API key and account URL to `.env` file

## Usage

1. Run the main script:
   ```bash
   python main.py
   ```

2. The script will:
   - Authenticate with Microsoft Graph and Intervals
   - Fetch recent Teams meetings
   - Process attendance data
   - Match meetings to Intervals tasks
   - Post time entries to Intervals
   - Generate a summary report

## Output Example

```
=== Meeting Processing Report ===

Processing Summary:
Total Meetings: 7
Successfully Matched: 5
Time Entries Posted: 5
Unmatched Meetings: 2
Match Success Rate: 71.4%
Post Success Rate: 100.0%

Total Billable Hours: 2.7 hours

Task Distribution:
Daily Status Report: 2 entries
UMG Task: 1 entries
WD Task: 1 entries
Infra Meeting Task: 1 entries
```

## Features Details

### Meeting Processing
- Fetches meetings from the last 30 days
- Calculates actual attendance duration
- Handles multiple attendees
- Processes online and offline meetings

### Time Entry Posting
- Posts billable time to Intervals
- Rounds duration to nearest 0.1 hours
- Includes meeting subject and attendance details
- Skips entries with zero duration

### Task Matching
- Uses AI to match meetings to tasks
- Supports direct matching based on subject
- Handles unmatched meetings gracefully

## Troubleshooting

1. Authentication Issues:
   - Verify credentials in `.env` file
   - Check Azure AD app permissions
   - Ensure Intervals API key is valid

2. Meeting Data Issues:
   - Check Teams meeting permissions
   - Verify meeting organizer access
   - Ensure attendance tracking is enabled

3. Task Matching Issues:
   - Review task titles in Intervals
   - Check meeting subject formatting
   - Adjust matching confidence threshold

## Contributing

Contributions are welcome! Please feel free to submit pull requests.

## License

This project is licensed under the MIT License - see the LICENSE file for details. 