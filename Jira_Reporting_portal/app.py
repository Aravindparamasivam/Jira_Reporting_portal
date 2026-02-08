from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify, send_file
from functools import wraps
from jira_api import JiraAPI
from jira_exporter import JiraExporter
from exporter import FileExporter
from config import JIRA_URL
import os
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'your-secret-key-change-in-production'  # Change this in production

# Configure session to expire when browser closes (not persistent)
app.config['SESSION_PERMANENT'] = False

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'username' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

@app.route('/')
def index():
    if 'username' in session:
        return redirect(url_for('dashboard'))
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    # If user is already logged in, redirect to dashboard
    if 'username' in session:
        return redirect(url_for('dashboard'))
    
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        api_token = request.form.get('password', '').strip()  # Using password field for API token
        
        if not username or not api_token:
            flash('Please enter both username/email and API token', 'error')
            return render_template('login.html')
        
        # Verify credentials against Jira
        try:
            api = JiraAPI(email=username, api_token=api_token)
            user_info = api.verify_credentials()
            
            if user_info:
                # Store credentials in session
                session['username'] = username
                session['api_token'] = api_token
                session['full_name'] = user_info.get('displayName', username)
                session['email'] = user_info.get('emailAddress', username)
                session['account_id'] = user_info.get('accountId', '')
                return redirect(url_for('dashboard'))
            else:
                flash('Invalid Jira credentials. Please check your username/email and API token.', 'error')
        except Exception as e:
            flash(f'Error connecting to Jira: {str(e)}', 'error')
    
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    flash('You have been logged out successfully', 'success')
    return redirect(url_for('login'))

@app.route('/dashboard')
@login_required
def dashboard():
    return render_template('dashboard.html')

@app.route('/search', methods=['POST'])
@login_required
def search_ticket():
    ticket_id = request.form.get('ticket_id', '').strip().upper()
    
    if not ticket_id:
        flash('Please enter a ticket number', 'error')
        # Clear previous data when input is empty
        session.pop('jira_data', None)
        session.pop('current_epic_key', None)
        session.pop('current_epic_name', None)
        session.pop('current_epic_status', None)
        session.pop('current_epic_progress', None)
        return redirect(url_for('dashboard'))
    
    try:
        # Fetch data from Jira using the same logic as main.py
        api = JiraAPI(email=session.get('username'), api_token=session.get('api_token'))
        issues = api.fetch_epic_defects(ticket_id)
        
        if not issues:
            # Clear previous data when no defects found
            session.pop('jira_data', None)
            session.pop('current_epic_key', None)
            session.pop('current_epic_name', None)
            session.pop('current_epic_status', None)
            session.pop('current_epic_progress', None)
            flash(f'No defects found for EPIC {ticket_id}. Please check if the EPIC exists and has child work items with Bug/Defect subtasks.', 'warning')
            return redirect(url_for('dashboard'))
        
        # Fetch the parent EPIC summary, progress, and status
        try:
            parent_issue = api.fetch_issue_by_key(ticket_id)
            if parent_issue and parent_issue.get('fields'):
                parent_summary = parent_issue.get('fields', {}).get('summary', '')
                parent_status = parent_issue.get('fields', {}).get('status', {}).get('name', '')
            else:
                parent_summary = f'EPIC {ticket_id}'
                parent_status = 'Unknown'
        except Exception as e:
            print(f"Error fetching parent issue: {e}")
            # Clear previous data on error
            session.pop('jira_data', None)
            session.pop('current_epic_key', None)
            session.pop('current_epic_name', None)
            session.pop('current_epic_status', None)
            session.pop('current_epic_progress', None)
            parent_summary = f'EPIC {ticket_id}'
            parent_status = 'Unknown'
        
        parent_progress = api.get_parent_progress(ticket_id)
        
        # Extract data using JiraExporter
        exporter = JiraExporter(issues)
        data, main_ticket_summary, parent_progress, parent_status = exporter.extract_data(
            header_value=parent_summary,
            parent_progress=parent_progress,
            parent_status=parent_status
        )
        
        # Store data in session
        session['current_epic_key'] = ticket_id
        session['current_epic_name'] = main_ticket_summary
        session['current_epic_status'] = parent_status
        session['current_epic_progress'] = parent_progress
        session['jira_data'] = data
        # Note: We'll fetch raw issues on-demand for metrics calculation to avoid session size issues
        
        flash(f'Successfully loaded data for {ticket_id}', 'success')
        return redirect(url_for('dashboard'))
        
    except Exception as e:
        # Clear previous data on error
        session.pop('jira_data', None)
        session.pop('current_epic_key', None)
        session.pop('current_epic_name', None)
        session.pop('current_epic_status', None)
        session.pop('current_epic_progress', None)
        flash(f'Error fetching data: {str(e)}', 'error')
        return redirect(url_for('dashboard'))

@app.route('/report/jira')
@login_required
def jira_report():
    if 'jira_data' not in session:
        flash('Please search for a ticket first', 'error')
        return redirect(url_for('dashboard'))
    
    data = session.get('jira_data', [])
    epic_key = session.get('current_epic_key', '')
    epic_name = session.get('current_epic_name', '')
    
    return render_template('reports/jira_report.html', 
                         data=data, 
                         epic_key=epic_key, 
                         epic_name=epic_name)

@app.route('/report/qa')
@login_required
def qa_report():
    if 'jira_data' not in session:
        flash('Please search for a ticket first', 'error')
        return redirect(url_for('dashboard'))
    
    data = session.get('jira_data', [])
    epic_key = session.get('current_epic_key', '')
    epic_name = session.get('current_epic_name', '')
    epic_status = session.get('current_epic_status', 'Unknown')
    epic_progress = session.get('current_epic_progress', 'Progress: N/A')
    
    # Calculate QA metrics
    total_defects = len(data)
    open_defects = len([d for d in data if d.get('Status', '').upper() != 'DONE'])
    closed_defects = total_defects - open_defects
    qa_complete = (closed_defects / total_defects * 100) if total_defects > 0 else 0
    
    return render_template('reports/qa_report.html',
                         data=data,
                         epic_key=epic_key,
                         epic_name=epic_name,
                         epic_status=epic_status,
                         epic_progress=epic_progress,
                         total_defects=total_defects,
                         open_defects=open_defects,
                         closed_defects=closed_defects,
                         qa_complete=qa_complete)

@app.route('/report/metrics')
@login_required
def metrics_board():
    if 'jira_data' not in session:
        flash('Please search for a ticket first', 'error')
        return redirect(url_for('dashboard'))
    
    data = session.get('jira_data', [])
    epic_key = session.get('current_epic_key', '')
    epic_name = session.get('current_epic_name', '')
    
    # Calculate defect leakage metrics
    # Fetch raw issues on-demand for metrics calculation
    metrics = None
    try:
        api = JiraAPI(email=session.get('username'), api_token=session.get('api_token'))
        raw_issues = api.fetch_epic_defects(epic_key)
        exporter = JiraExporter(raw_issues)
        # Try ODC-based calculation first
        metrics = exporter.calculate_defect_metrics(data)
        
        if not metrics or metrics.get('total_defects', 0) == 0:
            # Fallback to date-based calculation
            metrics = exporter.calculate_defect_metrics_from_dates(data)
    except Exception as e:
        print(f"Error calculating metrics: {e}")
        import traceback
        traceback.print_exc()
        metrics = None
    
    return render_template('reports/metrics_board.html',
                         data=data,
                         metrics=metrics,
                         epic_key=epic_key,
                         epic_name=epic_name)

@app.route('/report/visual')
@login_required
def visual_board():
    if 'jira_data' not in session:
        flash('Please search for a ticket first', 'error')
        return redirect(url_for('dashboard'))
    
    data = session.get('jira_data', [])
    epic_key = session.get('current_epic_key', '')
    epic_name = session.get('current_epic_name', '')
    
    # Calculate statistics for charts
    status_counts = {}
    priority_counts = {}
    environment_counts = {}
    assignee_counts = {}
    
    for item in data:
        # Status counts
        status = item.get('Status', 'Unknown')
        status_counts[status] = status_counts.get(status, 0) + 1
        
        # Priority counts
        priority = item.get('Priority', 'Unknown')
        priority_counts[priority] = priority_counts.get(priority, 0) + 1
        
        # Environment counts
        env = item.get('Environment', 'Not Specified')
        environment_counts[env] = environment_counts.get(env, 0) + 1
        
        # Assignee counts
        assignee = item.get('Assignee', 'Unassigned')
        assignee_counts[assignee] = assignee_counts.get(assignee, 0) + 1
    
    # Calculate time metrics (only for tickets with resolution dates)
    durations = []
    for item in data:
        duration_str = item.get('Duration (HH:MM)', 'N/A')
        resolution_date = item.get('Resolution Date', '')
        # Only include durations for tickets that have a resolution date (not '00:00' or 'N/A')
        if duration_str and duration_str != 'N/A' and resolution_date and resolution_date != 'None':
            try:
                parts = duration_str.split(':')
                if len(parts) == 2:
                    hours = float(parts[0])
                    minutes = float(parts[1])
                    total_hours = hours + (minutes / 60)
                    # Only include non-zero durations (tickets with actual resolution dates)
                    if total_hours > 0:
                        durations.append(total_hours)
            except:
                pass
    
    time_metrics = {
        'min': min(durations) if durations else 0,
        'max': max(durations) if durations else 0,
        'avg': sum(durations) / len(durations) if durations else 0
    }
    
    # Overall statistics
    total_defects = len(data)
    open_defects = len([d for d in data if d.get('Status', '').upper() != 'DONE'])
    closed_defects = total_defects - open_defects
    
    # Calculate defect leakage metrics for donut charts
    metrics = None
    try:
        api = JiraAPI(email=session.get('username'), api_token=session.get('api_token'))
        raw_issues = api.fetch_epic_defects(epic_key)
        exporter = JiraExporter(raw_issues)
        # Try ODC-based calculation first
        metrics = exporter.calculate_defect_metrics(data)
        
        if not metrics or metrics.get('total_defects', 0) == 0:
            # Fallback to date-based calculation
            metrics = exporter.calculate_defect_metrics_from_dates(data)
    except Exception as e:
        print(f"Error calculating metrics: {e}")
        import traceback
        traceback.print_exc()
        metrics = None
    
    return render_template('reports/visual_board.html',
                         data=data,
                         epic_key=epic_key,
                         epic_name=epic_name,
                         status_counts=status_counts,
                         priority_counts=priority_counts,
                         environment_counts=environment_counts,
                         assignee_counts=assignee_counts,
                         time_metrics=time_metrics,
                         total_defects=total_defects,
                         open_defects=open_defects,
                         closed_defects=closed_defects,
                         metrics=metrics)

@app.route('/download/excel')
@login_required
def download_excel():
    epic_key = request.args.get('epic_key', '').strip().upper()
    
    if not epic_key:
        flash('Please enter a ticket number', 'error')
        return redirect(url_for('dashboard'))
    
    try:
        # Use the same logic as main.py's process_project function
        api = JiraAPI(email=session.get('username'), api_token=session.get('api_token'))
        issues = api.fetch_epic_defects(epic_key)
        
        if not issues:
            flash(f'No defects found for EPIC {epic_key}. Please check if the EPIC exists and has child work items with Bug/Defect subtasks.', 'warning')
            return redirect(url_for('dashboard'))
        
        # Fetch the parent EPIC summary, progress, and status
        try:
            parent_issue = api.fetch_issue_by_key(epic_key)
            if parent_issue and parent_issue.get('fields'):
                parent_summary = parent_issue.get('fields', {}).get('summary', '')
                parent_status = parent_issue.get('fields', {}).get('status', {}).get('name', '')
            else:
                parent_summary = f'EPIC {epic_key}'
                parent_status = 'Unknown'
        except:
            parent_summary = f'EPIC {epic_key}'
            parent_status = 'Unknown'
        
        parent_progress = api.get_parent_progress(epic_key)
        
        # Extract data using JiraExporter
        exporter = JiraExporter(issues)
        data, main_ticket_summary, parent_progress, parent_status = exporter.extract_data(
            header_value=parent_summary,
            parent_progress=parent_progress,
            parent_status=parent_status
        )
        
        # Generate Excel file using FileExporter (same as main.py)
        file_export = FileExporter((data, main_ticket_summary, parent_progress, parent_status), exporter)
        excel_file_path, pdf_file = file_export.to_excel(epic_key)
        
        if excel_file_path and os.path.exists(excel_file_path):
            # Generate a clean filename for download
            timestamp = datetime.now().strftime("%Y%m%d%H%M")
            download_filename = f'{timestamp}_{epic_key}.xlsx'
            
            return send_file(
                excel_file_path,
                as_attachment=True,
                download_name=download_filename,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            flash('Error generating Excel file', 'error')
            return redirect(url_for('dashboard'))
        
    except Exception as e:
        flash(f'Error generating Excel file: {str(e)}', 'error')
        return redirect(url_for('dashboard'))

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)

