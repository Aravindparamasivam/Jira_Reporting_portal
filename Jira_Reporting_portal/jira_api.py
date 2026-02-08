import requests
from config import JIRA_URL, EMAIL, API_TOKEN
import urllib.parse

class JiraAPI:
    
    def __init__(self, email=None, api_token=None):
        # Use provided credentials or fall back to config
        self.email = email or EMAIL
        self.api_token = api_token or API_TOKEN
        self.auth = (self.email, self.api_token)
        self.headers = {
            "Accept": "application/json",
            "Content-Type": "application/json"
        }
    
    def verify_credentials(self):
        """Verify if the credentials are valid by making a test API call"""
        try:
            url = f"{JIRA_URL}/rest/api/3/myself"
            response = requests.get(url, headers=self.headers, auth=self.auth, timeout=10)
            if response.status_code == 200:
                return response.json()
            return None
        except Exception as e:
            print(f"Error verifying credentials: {str(e)}")
            return None

    def get_projects(self):
        try:
            url = f"{JIRA_URL}/rest/api/3/project"
            response = requests.get(url, headers=self.headers, auth=self.auth)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            print(f"Error fetching projects: {str(e)}")
            raise

    def fetch_epic_defects(self, epic_key):
        """Fetch defects from EPIC child work items and subtasks"""
        try:
            # First get child work items of the EPIC
            child_work_items = self._get_epic_child_items(epic_key)
            
            all_defects = []
            
            # For each child work item, get its subtasks that are defects
            for child_item in child_work_items:
                subtask_defects = self._get_subtask_defects(child_item['key'])
                all_defects.extend(subtask_defects)
            
            if len(all_defects) == 0:
                print(f"No defects found for EPIC {epic_key}. Please check if the EPIC exists and has child work items with Bug/Defect subtasks.")
                
            return all_defects
        except requests.exceptions.RequestException as e:
            print(f"Error making request to Jira: {str(e)}")
            raise

    def _get_epic_child_items(self, epic_key):
        """Get child work items for an EPIC"""
        try:
            # Try multiple JQL approaches to find child items
            jql_queries = [
                f'parent = {epic_key}',  # Direct parent relationship
                f'"Epic Link" = {epic_key}',  # Epic Link field
                f'epic = {epic_key}',  # Epic field
                f'parentEpic = {epic_key}',  # Parent Epic field
                f'issue in linkedIssues({epic_key})'  # Linked issues
            ]
            
            all_issues = []
            
            for jql in jql_queries:
                try:
                    url = f"{JIRA_URL}/rest/api/3/search/jql"
                    
                    params = {
                        "jql": jql,
                        "fields": "key,summary,status,issuetype,parent,issuelinks",
                        "maxResults": 100
                    }
                    
                    response = requests.get(url, headers=self.headers, auth=self.auth, params=params)
                    
                    if response.status_code == 200:
                        issues = response.json().get("issues", [])
                        if len(issues) > 0:
                            all_issues.extend(issues)
                            # Stop after first successful query to avoid duplicates
                            break
                        
                except Exception as e:
                    print(f"Error with query '{jql}': {str(e)}")
                    continue
            
            return all_issues
            
        except Exception as e:
            print(f"Error fetching child work items for {epic_key}: {str(e)}")
            return []

    def _get_subtask_defects(self, parent_key):
        """Get subtasks that are defects for a parent issue"""
        try:
            # Try multiple JQL approaches for subtasks (only Defect type)
            jql_queries = [
                f'parent = "{parent_key}" AND issuetype = Defect',
                f'parent = {parent_key} AND issuetype = Defect',
                f'parent = "{parent_key}" AND issuetype = "Defect"',
                f'parent = {parent_key} AND issuetype = "Defect"'
            ]
            
            all_issues = []
            
            for jql in jql_queries:
                try:
                    url = f"{JIRA_URL}/rest/api/3/search/jql"
                    
                    params = {
                        "jql": jql,
                        "fields": "*all",  # Fetch all fields including custom fields
                        "expand": "changelog,parent",  # Include parent information
                        "maxResults": 100
                    }
                    
                    response = requests.get(url, headers=self.headers, auth=self.auth, params=params)
                    
                    if response.status_code == 200:
                        issues = response.json().get("issues", [])
                        if len(issues) > 0:
                            all_issues.extend(issues)
                            # Stop after first successful query to avoid duplicates
                            break
                        
                except Exception as e:
                    print(f"Error fetching subtasks for {parent_key}: {str(e)}")
                    continue
            
            # Filter for actual defects based on issue type (only "Defect" type)
            defects = []
            for issue in all_issues:
                issue_type = issue.get('fields', {}).get('issuetype', {}).get('name', '').strip()
                # Only include issues with "Defect" as the exact issue type
                if issue_type.lower() == 'defect':
                    defects.append(issue)
            
            return defects
            
        except Exception as e:
            print(f"Error fetching subtasks for {parent_key}: {str(e)}")
            return []

    def fetch_issue_by_key(self, issue_key):
        """Fetch a specific issue by its key"""
        try:
            url = f"{JIRA_URL}/rest/api/3/issue/{issue_key}"
            params = {
                "fields": "*all",
                "expand": "changelog"
            }
            
            response = requests.get(url, headers=self.headers, auth=self.auth, params=params)
            response.raise_for_status()
            
            return response.json()
            
        except Exception as e:
            print(f"Error fetching issue {issue_key}: {str(e)}")
            return None

    def get_parent_progress(self, epic_key):
        """Get parent EPIC progress information"""
        try:
            issue = self.fetch_issue_by_key(epic_key)
            if issue and 'fields' in issue:
                # Extract progress information if available
                progress = issue['fields'].get('progress', {})
                if progress:
                    return f"Progress: {progress.get('percent', 0)}%"
            return "Progress: N/A"
        except Exception as e:
            print(f"Error getting parent progress for {epic_key}: {str(e)}")
            return "Progress: N/A" 