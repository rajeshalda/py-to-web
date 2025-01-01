#!/usr/bin/env python3

from meeting_processor import MeetingProcessor

def main():
    """Main entry point for the meeting processor."""
    try:
        processor = MeetingProcessor()
        processor.run()
    except KeyboardInterrupt:
        print("\nProcess interrupted by user")
    except Exception as e:
        print(f"Error: {str(e)}")
        raise

if __name__ == "__main__":
    main() 