from jira_api import JiraAPI
from jira_exporter import JiraExporter
from exporter import FileExporter
from file_handler import archive_existing_files, close_pdf_processes
from config import get_project_keys

def process_project(project_key):
    
    print(f"\nProcessing project: {project_key}")
    print("Fetching issues from Jira...")
    api = JiraAPI()
    issues = api.fetch_epic_defects(project_key)
    
    # Fetch the parent EPIC summary, progress, and status
    try:
        parent_issue = api.fetch_issue_by_key(project_key)
        if parent_issue and parent_issue.get('fields'):
            parent_summary = parent_issue.get('fields', {}).get('summary', '')
            parent_status = parent_issue.get('fields', {}).get('status', {}).get('name', '')
        else:
            parent_summary = f'EPIC {project_key}'
            parent_status = 'Unknown'
    except:
        parent_summary = f'EPIC {project_key}'
        parent_status = 'Unknown'
    
    parent_progress = api.get_parent_progress(project_key)

    print("Extracting relevant fields...")
    exporter = JiraExporter(issues)
    data, main_ticket_summary, parent_progress, parent_status = exporter.extract_data(header_value=parent_summary, parent_progress=parent_progress, parent_status=parent_status)

    print("Exporting to Excel...")
    file_export = FileExporter((data, main_ticket_summary, parent_progress, parent_status), exporter)
    excel_file, pdf_file = file_export.to_excel(project_key)

    print(f"Exported for {project_key}:")
    print(f"Excel: {excel_file}")
    
    return excel_file, pdf_file

def launch_all_files(generated_files):
    """Launch all Excel and PDF files together"""
    import subprocess
    import os
    import webbrowser
    import time
    
    if not generated_files:
        return
    
    print("\n🚀 Launching all generated files...")
    
    # Close any existing PDF viewers before launching new ones
    print("Closing any open PDF files...")
    close_pdf_processes()
    
    # Launch all PDFs first
    for excel_path, pdf_path in generated_files:
        if pdf_path and os.path.exists(pdf_path):
            try:
                abs_pdf_path = os.path.abspath(pdf_path)
                webbrowser.open(f"file://{abs_pdf_path}")
                print(f"PDF dashboard launched: {pdf_path}")
            except Exception as e:
                print(f"Could not launch PDF {pdf_path}: {str(e)}")
    
    # Small delay before launching Excel files
    time.sleep(1)
    
    # Launch all Excel files
    excel_processes = []
    for excel_path, pdf_path in generated_files:
        if excel_path and os.path.exists(excel_path):
            try:
                abs_excel_path = os.path.abspath(excel_path)
                process = subprocess.Popen(['start', 'excel', abs_excel_path], shell=True)
                excel_processes.append((process, abs_excel_path))
                print(f"Excel file launched: {excel_path}")
            except Exception as e:
                print(f"Could not launch Excel {excel_path}: {str(e)}")
    
    # Wait for all Excel windows to open, then maximize them
    if excel_processes:
        time.sleep(3)  # Wait for all Excel windows to open
        
        # Maximize all Excel windows
        try:
            # Get all Excel processes and maximize them
            ps_command = '''
$excelProcesses = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -ne "" }
foreach ($excel in $excelProcesses) {
    try {
        $hwnd = $excel.MainWindowHandle
        Add-Type -TypeDefinition "using System; using System.Runtime.InteropServices; public class Win32 { [DllImport(\\"user32.dll\\")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow); public static readonly int SW_MAXIMIZE = 3; }"
        [Win32]::ShowWindow($hwnd, [Win32]::SW_MAXIMIZE)
    } catch {}
}
'''
            subprocess.run(['powershell', '-Command', ps_command], shell=True, capture_output=True, timeout=10)
        except Exception as e:
            pass  # Silently continue if maximization fails

def main():
    
    print("Archiving existing files...")
    archive_existing_files()

    # Get project keys from user
    project_keys = get_project_keys()
    
    # Process each project and collect all generated files
    generated_files = []
    for project_key in project_keys:
        try:
            excel_file, pdf_file = process_project(project_key)
            if excel_file:
                generated_files.append((excel_file, pdf_file))
        except Exception as e:
            print(f"Error processing project {project_key}: {str(e)}")
            continue
    
    # Launch all files together at the end
    if generated_files:
        launch_all_files(generated_files)

if __name__ == "__main__":
    main() 