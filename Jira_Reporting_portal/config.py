# Configuration settings for Jira automation
JIRA_URL = "https://jira-automation-rbm.atlassian.net"
EMAIL = "aravindp@thinknsolutions.com"
API_TOKEN = "ATATT3xFfGF0XPsC4-AXyqPtv5Oj1wEkkmbgsrY9f6lhIj8Aw8G0FjVr85pcQuJpAlaPC14yFmWQhseQZ5lmCe7E5Z3k2f1XgjVh3VThE8DC9zWwO77PWLIMbQNVcFAUmWz52hzv37chD89qqLYvaSbmwqoCgHJ963F_qZ2Fh2eaTSgqWvc1ccQ=E7788F51"
EXPORT_FOLDER = "Files"
ARCHIVE_FOLDER = "Archive"

class Config:
    """Configuration class for Jira automation"""
    
    def __init__(self):
        self.JIRA_URL = JIRA_URL
        self.EMAIL = EMAIL
        self.API_TOKEN = API_TOKEN
        self.OUTPUT_PATH = EXPORT_FOLDER
        self.ARCHIVE_FOLDER = ARCHIVE_FOLDER
        
        # Updated headers with Environment column (Phase removed) and Phase columns added
        self.EXCEL_HEADERS = [
            'Key',
            'Environment',
            'Summary',
            'Priority',
            'Issue Type',
            'Reporter',
            'Assignee',
            'Status',
            'Created Date',
            'Resolution Date',
            'Last Status Change',
            'Assignee History',
            'Duration (HH:MM)',
            'Phase Detected',
            'Phase Injected'
        ]

def get_project_keys():
    """Get EPIC keys from user input"""
    while True:
        # Get user input and remove leading/trailing whitespace
        keys_input = input("Enter project key(s) (comma-separated for multiple): ").strip()
        
        # Check if input is empty
        if not keys_input:
            print("Please enter at least one project key")
            continue
            
        # Process input: split by comma and clean each key
        keys = [key.strip() for key in keys_input.split(',') if key.strip()]
        
        # Validate that we have at least one valid key
        if not keys:
            print("Please enter at least one valid project key")
            continue
            
        return keys 