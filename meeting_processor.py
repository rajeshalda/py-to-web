import pandas as pd
from datetime import datetime, timedelta
from typing import Dict, List, Any, Optional
from utils import (
    get_saved_intervals_token, save_intervals_token,
    format_duration, get_meeting_decimal_time, parse_datetime
)
from graph_client import GraphClient
from intervals_client import IntervalsClient
from task_matcher import TaskMatcher

def to_intervals_date(dt):
    """Convert a datetime object to Intervals date format (YYYY-MM-DD)"""
    return dt.strftime('%Y-%m-%d')

class MeetingProcessor:
    def __init__(self):
        """Initialize the meeting processor."""
        self.graph_client = None
        self.intervals_client = None
        self.task_matcher = None
        self.results = []
        self.tasks = []  # Initialize tasks list
        self.current_user = None  # Add current user info
    
    def initialize(self) -> bool:
        """Initialize the processor and authenticate with services."""
        print("\nInitializing Meeting Processor...")
        
        # Initialize and authenticate Graph client
        self.graph_client = GraphClient()
        if not self.graph_client.authenticate():
            print("Failed to authenticate with Microsoft Graph")
            return False
        
        # Initialize and authenticate Intervals client
        intervals_token = get_saved_intervals_token()
        if not intervals_token:
            print("No Intervals token found")
            return False
            
        self.intervals_client = IntervalsClient(intervals_token)
        self.current_user = self.intervals_client.get_current_user()
        if not self.current_user:
            print("Failed to authenticate with Intervals")
            return False
            
        print(f"Successfully authenticated as: {self.current_user['firstname']} {self.current_user['lastname']}")
        
        # Initialize task matcher
        self.task_matcher = TaskMatcher()
        
        # Get tasks from Intervals
        print("Building task context...")
        self.tasks = self.intervals_client.get_tasks()
        if not self.tasks:
            print("No tasks found in Intervals")
            return False
            
        # Build task context
        self.task_matcher.build_task_context(self.tasks, self.intervals_client)
        
        return True
    
    def show_attendance_report(self, meeting: Dict[str, Any], attendance_data: Optional[Dict]) -> Dict[str, Any]:
        """Display attendance report for a meeting and return attendance statistics."""
        start_time = parse_datetime(meeting['start']['dateTime'])
        end_time = parse_datetime(meeting['end']['dateTime'])
        scheduled_duration = int((end_time - start_time).total_seconds())
        
        # Get current user info
        user_name = f"{self.current_user['firstname']} {self.current_user['lastname']}"
        user_email = f"{self.current_user['username']}@M365x65088219.onmicrosoft.com"
        
        print("\n" + "=" * 45)
        print("Meeting Attendance Report")
        print("=" * 45 + "\n")
        
        # Meeting Details Section
        print(f"Meeting Subject: {meeting['subject']}")
        print(f"Start Time: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"End Time: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"Scheduled Duration: {scheduled_duration // 60} minutes\n")
        
        print("----- Attendance Details -----\n")
        
        max_attendance_seconds = scheduled_duration
        attendee_details = []
        total_attendance_percentage = 0
        attendee_count = 0
        
        if not attendance_data or not attendance_data.get('attendees'):
            # Use current user's info when no attendance data is available
            attendee_details.append({
                'name': user_name,
                'email': user_email,
                'duration': scheduled_duration,
                'percentage': 100.0
            })
        else:
            for attendee in attendance_data['attendees']:
                # Skip entries without duration
                if 'duration' not in attendee:
                    continue
                
                # Calculate attendance percentage
                duration_seconds = attendee['duration']
                attendance_percentage = (duration_seconds / scheduled_duration) * 100
                
                # Use current user info if name/email is missing or shows as Unknown User
                name = attendee.get('name', user_name)
                if name == 'Unknown User':
                    name = user_name
                    
                email = attendee.get('email', user_email)
                if email == self.graph_client.user_id:
                    email = user_email
                
                attendee_details.append({
                    'name': name,
                    'email': email,
                    'duration': duration_seconds,
                    'percentage': attendance_percentage
                })
        
        # Sort attendees by duration (descending)
        attendee_details.sort(key=lambda x: x['duration'], reverse=True)
        
        # Display attendee details
        for attendee in attendee_details:
            attendee_count += 1
            total_attendance_percentage += attendee['percentage']
            
            duration_minutes = attendee['duration'] // 60
            
            print(f"Attendee: {attendee['name']}")
            print(f"Email: {attendee['email']}")
            print(f"Duration: {duration_minutes} minutes")
            print(f"Attendance: {attendee['percentage']:.0f}%\n")
        
        print("=" * 45)
        
        # Calculate actual duration for display
        actual_duration_hours = max([a['duration'] for a in attendee_details]) // 3600 if attendee_details else 0
        actual_duration_minutes = (max([a['duration'] for a in attendee_details]) % 3600) // 60 if attendee_details else 0
        actual_duration_seconds = max([a['duration'] for a in attendee_details]) % 60 if attendee_details else 0
        
        print(f"Actual duration: {actual_duration_hours} hours, {actual_duration_minutes} minutes, {actual_duration_seconds} seconds")
        exact_hours = max([a['duration'] for a in attendee_details]) / 3600 if attendee_details else 0
        print(f"Exact attendance time: {exact_hours:.1f} hours")
        
        return {
            'attendee_count': attendee_count,
            'avg_attendance': total_attendance_percentage / attendee_count if attendee_count > 0 else 0,
            'max_duration': max([a['duration'] for a in attendee_details]) if attendee_details else scheduled_duration
        }
    
    def process_meeting(self, meeting: Dict[str, Any]) -> Dict[str, Any]:
        """Process a single meeting."""
        print(f"\nProcessing meeting: {meeting['subject']}")
        print(f"Meeting Description:\n{meeting.get('bodyPreview', '')}")

        start_time = parse_datetime(meeting['start']['dateTime'])
        end_time = parse_datetime(meeting['end']['dateTime'])
        attendance_data = None
        
        # Get current user info
        user_name = f"{self.current_user['firstname']} {self.current_user['lastname']}"
        user_email = f"{self.current_user['username']}@M365x65088219.onmicrosoft.com"
        
        if meeting.get('onlineMeeting'):
            try:
                # Get attendance data
                attendance_data = self.graph_client.get_meeting_attendance(
                    meeting['onlineMeeting']['joinUrl'],
                    start_time,
                    end_time
                )
            except Exception as e:
                print(f"Failed to get attendance data: {str(e)}")
                # Create default attendance data with current user
                attendance_data = {
                    'attendees': [{
                        'name': user_name,
                        'email': user_email,
                        'duration': int((end_time - start_time).total_seconds()),
                        'role': 'Organizer'
                    }]
                }

        # Get attendance report
        attendance_report = self.show_attendance_report(meeting, attendance_data)
        duration_seconds = attendance_report['max_duration']

        # Calculate billable hours (rounded to nearest 0.1)
        actual_duration = timedelta(seconds=duration_seconds)
        print(f"Actual duration: {actual_duration}")
        billable_hours = round(duration_seconds / 3600, 1)
        print(f"Billable hours calculated: {billable_hours}")

        # Skip posting if billable hours is 0
        if billable_hours == 0:
            print("Skipping time entry posting - zero duration")
            return {
                'meeting': meeting['subject'],
                'time': start_time.strftime('%Y-%m-%d %H:%M'),
                'task_id': 'N/A',
                'task_title': 'Zero Duration',
                'match_status': 'Not Posted',
                'posted': 'No - Zero Duration',
                'billable_duration': billable_hours,
                'duration': duration_seconds,
                'actual_minutes': round(duration_seconds / 60),
                'scheduled_duration': duration_seconds
            }

        # Match meeting to task
        matched_task_id = self.task_matcher.direct_match(meeting['subject'], self.tasks)
        if not matched_task_id:
            matched_task_id = self.task_matcher.ai_match(meeting['subject'], self.tasks)

        if not matched_task_id or matched_task_id == "NO_MATCH":
            print("No task match found for meeting")
            return {
                'meeting': meeting['subject'],
                'time': start_time.strftime('%Y-%m-%d %H:%M'),
                'task_id': 'N/A',
                'task_title': 'No Match',
                'match_status': 'Not Matched',
                'posted': 'No',
                'billable_duration': billable_hours,
                'duration': duration_seconds,
                'actual_minutes': round(duration_seconds / 60),
                'scheduled_duration': duration_seconds
            }

        matched_task = next((task for task in self.tasks if task['id'] == matched_task_id), None)
        if not matched_task:
            print(f"Task with ID {matched_task_id} not found")
            return None

        print(f"Matched to task: {matched_task['title']}")

        # Post time entry to Intervals
        time_entry = {
            'personid': self.intervals_client.user_id,
            'projectid': matched_task['projectid'],
            'moduleid': matched_task.get('moduleid'),
            'taskid': matched_task['id'],
            'worktypeid': 799573,  # Default worktype ID
            'date': to_intervals_date(start_time),
            'time': billable_hours,
            'description': f"{meeting['subject']} (Actual attendance: {round(duration_seconds / 60)} minutes)",
            'billable': True
        }

        post_success = self.intervals_client.post_time_entry(time_entry)
        post_status = "Yes" if post_success else "Failed"
        
        return {
            'meeting': meeting['subject'],
            'time': start_time.strftime('%Y-%m-%d %H:%M'),
            'task_id': matched_task['id'],
            'task_title': matched_task['title'],
            'match_status': 'Matched',
            'posted': post_status,
            'billable_duration': billable_hours,
            'duration': duration_seconds,
            'actual_minutes': round(duration_seconds / 60),
            'scheduled_duration': duration_seconds
        }
    
    def show_statistics(self):
        """Display processing statistics."""
        if not self.results:
            print("No results to display")
            return
        
        print("\n=== Meeting Processing Report ===")
        
        # Calculate statistics
        total_meetings = len(self.results)
        matched_meetings = sum(1 for r in self.results if r['match_status'] == 'Matched')
        posted_entries = sum(1 for r in self.results if r['posted'] == 'Yes')
        total_billable = sum(r['billable_duration'] for r in self.results if r['posted'] == 'Yes')
        
        # Display summary
        print("\nProcessing Summary:")
        print(f"Total Meetings: {total_meetings}")
        print(f"Successfully Matched: {matched_meetings}")
        print(f"Time Entries Posted: {posted_entries}")
        print(f"Unmatched Meetings: {total_meetings - matched_meetings}")
        
        if total_meetings > 0:
            match_rate = round((matched_meetings / total_meetings) * 100, 1)
            print(f"Match Success Rate: {match_rate}%")
        
        if matched_meetings > 0:
            post_rate = round((posted_entries / matched_meetings) * 100, 1)
            print(f"Post Success Rate: {post_rate}%")
        
        print(f"\nTotal Billable Hours: {total_billable} hours")
        
        # Task distribution
        print("\nTask Distribution:")
        task_counts = {}
        for result in self.results:
            if result['posted'] == 'Yes':
                task_title = result['task_title']
                task_counts[task_title] = task_counts.get(task_title, 0) + 1
        
        for task, count in task_counts.items():
            print(f"{task}: {count} entries")
    
    def export_results(self):
        """Export results to a text file."""
        if not self.results:
            print("No results to export")
            return
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        report_path = f"meeting_processing_{timestamp}.txt"
        
        try:
            # Create detailed report
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write("Meeting Processing Report\n")
                f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                
                f.write("Summary Statistics\n")
                f.write("-----------------\n")
                f.write(f"Total Meetings: {len(self.results)}\n")
                f.write(f"Matched: {sum(1 for r in self.results if r['match_status'] == 'Matched')}\n")
                f.write(f"Posted: {sum(1 for r in self.results if r['posted'] == 'Yes')}\n\n")
                
                f.write("Detailed Results\n")
                f.write("---------------\n\n")
                
                for result in self.results:
                    f.write(f"Meeting: {result['meeting']}\n")
                    f.write(f"Time: {result['time']}\n")
                    f.write(f"Status: {result['match_status']}\n")
                    f.write(f"Task: {result['task_title']}\n")
                    f.write(f"Posted: {result['posted']}\n")
                    f.write(f"Duration: {result['billable_duration']} hours\n")
                    f.write(f"Actual Minutes: {result['actual_minutes']}\n")
                    f.write(f"Scheduled Duration: {format_duration(result['scheduled_duration'])}\n")
                    f.write(f"Number of Attendees: {result['attendee_count']}\n")
                    
                    if result['attendees']:
                        f.write("\nAttendees:\n")
                        for attendee in result['attendees']:
                            f.write(f"  - {attendee['name']} ({attendee['email']})\n")
                            f.write(f"    Duration: {format_duration(attendee['duration'])}\n")
                            f.write(f"    Attendance: {attendee['percentage']}%\n")
                    f.write("\n")
            
            print(f"Detailed report exported to: {report_path}")
            
        except Exception as e:
            print(f"Error exporting results: {str(e)}")
    
    def run(self):
        """Run the meeting processor."""
        start_time = datetime.now()
        
        try:
            if not self.initialize():
                return
            
            # Get meetings
            meetings = self.graph_client.get_user_meetings()
            if not meetings:
                print("No meetings found for processing.")
                return
            
            total_meetings = len(meetings)
            
            # Process each meeting
            for i, meeting in enumerate(meetings, 1):
                result = self.process_meeting(meeting)
                self.results.append(result)
            
            # Show results only (export disabled)
            self.show_statistics()
            
            duration = datetime.now() - start_time
            print(f"\nTotal processing time: {duration.seconds // 60} minutes and {duration.seconds % 60} seconds")
            
        except Exception as e:
            print(f"Error in main process: {str(e)}")
            raise 