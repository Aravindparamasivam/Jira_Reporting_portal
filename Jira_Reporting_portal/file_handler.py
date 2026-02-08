# File management module for handling export and archive operations.
# Manages the creation, movement, and cleanup of exported files.

import os
import shutil
from config import EXPORT_FOLDER, ARCHIVE_FOLDER
import time
from datetime import datetime
import psutil

def get_file_creation_time(file_path):
    return datetime.fromtimestamp(os.path.getctime(file_path))

def close_excel_processes():
    """Close any running Excel processes to prevent file access issues"""
    try:
        closed_count = 0
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info['name'] and 'excel' in proc.info['name'].lower():
                    proc.terminate()
                    closed_count += 1
                    print(f"Closed Excel process: {proc.info['name']} (PID: {proc.info['pid']})")
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
        
        if closed_count > 0:
            print(f"Closed {closed_count} Excel process(es)")
            # Give processes time to close
            time.sleep(2)
        
    except Exception as e:
        print(f"Warning: Could not close Excel processes: {str(e)}")

def close_pdf_processes():
    """Close PDF viewer processes and browser tabs with PDF files from the project"""
    try:
        import subprocess
        closed_count = 0
        
        # List of common PDF viewer process names
        pdf_viewers = ['acrobat', 'acrord32', 'acrord64', 'foxit', 'foxitreader', 'sumatra', 'sumatrapdf']
        
        # Close dedicated PDF viewer applications
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info['name']:
                    proc_name = proc.info['name'].lower()
                    if any(viewer in proc_name for viewer in pdf_viewers):
                        proc.terminate()
                        closed_count += 1
                        print(f"Closed PDF viewer: {proc.info['name']} (PID: {proc.info['pid']})")
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
        
        # Close browser windows/tabs that have PDF files from our project open
        # Get list of PDF files in the export folder
        pdf_files = []
        if os.path.exists(EXPORT_FOLDER):
            for file in os.listdir(EXPORT_FOLDER):
                # Match new filename pattern: yyyymmddhhmm Project_key.pdf
                if file.endswith('.pdf') and len(file) > 13 and file[0:12].isdigit() and ' ' in file:
                    pdf_files.append(file)
        
        # Use PowerShell to close browser windows that match our PDF file pattern
        if pdf_files:
            try:
                # Build PowerShell command to close browser windows with PDF files
                # Match windows that contain timestamp pattern (12 digits) and ".pdf" in the title
                ps_command = '''
Add-Type -TypeDefinition "using System; using System.Runtime.InteropServices; public class Win32 { [DllImport(\"user32.dll\")] public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam); public const uint WM_CLOSE = 0x0010; }"
$pdfPattern = ".pdf"
$closedCount = 0
Get-Process | Where-Object { 
    $_.MainWindowTitle -ne "" -and 
    $_.MainWindowTitle -match "^\d{12}" -and 
    $_.MainWindowTitle -like "*$pdfPattern*" -and 
    ($_.ProcessName -eq "chrome" -or $_.ProcessName -eq "msedge" -or $_.ProcessName -eq "firefox" -or $_.ProcessName -eq "opera" -or $_.ProcessName -eq "brave") 
} | ForEach-Object {
    try {
        [Win32]::PostMessage($_.MainWindowHandle, [Win32]::WM_CLOSE, [IntPtr]::Zero, [IntPtr]::Zero)
        $closedCount++
        Write-Host "Closed browser window: $($_.MainWindowTitle)"
    } catch {}
}
if ($closedCount -gt 0) {
    Write-Host "Closed $closedCount browser window(s) with PDF files"
}
'''
                result = subprocess.run(['powershell', '-Command', ps_command], 
                                      shell=True, capture_output=True, text=True, timeout=5)
                if result.stdout and result.stdout.strip():
                    print(result.stdout.strip())
                if result.stderr and result.stderr.strip():
                    # Silently ignore stderr for this operation
                    pass
            except Exception as e:
                # Silently continue if PowerShell command fails
                pass
        
        if closed_count > 0:
            # Give processes time to close
            time.sleep(1)
        
    except Exception as e:
        print(f"Warning: Could not close PDF processes: {str(e)}")

def archive_existing_files():
    # Create necessary directories
    os.makedirs(EXPORT_FOLDER, exist_ok=True)
    os.makedirs(ARCHIVE_FOLDER, exist_ok=True)

    # Close Excel processes first to prevent file access issues
    print("Closing any open Excel processes...")
    close_excel_processes()

    # Archive current export files
    archived_count = 0
    for file in os.listdir(EXPORT_FOLDER):
        file_path = os.path.join(EXPORT_FOLDER, file)
        if os.path.isfile(file_path):
            try:
                # Move file to archive folder
                shutil.move(file_path, os.path.join(ARCHIVE_FOLDER, file))
                archived_count += 1
                print(f"Archived: {file}")
            except PermissionError:
                print(f"Warning: Could not archive {file} as it is currently in use. Skipping...")
                continue
            except Exception as e:
                print(f"Warning: Error archiving {file}: {str(e)}. Skipping...")
                continue
    
    if archived_count > 0:
        print(f"Successfully archived {archived_count} file(s)")
    else:
        print("No files to archive")

    # Manage archive folder size
    archive_files = []
    for file in os.listdir(ARCHIVE_FOLDER):
        file_path = os.path.join(ARCHIVE_FOLDER, file)
        if os.path.isfile(file_path):
            archive_files.append((file_path, get_file_creation_time(file_path)))

    # Sort files by creation time (oldest first)
    archive_files.sort(key=lambda x: x[1])

    # Remove oldest files if we have more than 10
    while len(archive_files) > 10:
        oldest_file = archive_files.pop(0)
        try:
            os.remove(oldest_file[0])
            print(f"Removed oldest archive file: {os.path.basename(oldest_file[0])}")
        except Exception as e:
            print(f"Warning: Could not remove old archive file {oldest_file[0]}: {str(e)}") 