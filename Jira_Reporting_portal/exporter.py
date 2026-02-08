import pandas as pd
from datetime import datetime
from config import EXPORT_FOLDER
import os
import subprocess
import psutil

def close_existing_files():
    """Close any existing Excel and PDF viewer processes to prevent multiple files from opening"""
    try:
        # Close Excel processes
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info['name'] and 'excel' in proc.info['name'].lower():
                    proc.terminate()
                    print("Closed existing Excel process:", proc.info['name'], "(PID:", proc.info['pid'], ")")
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
        
        # Close PDF viewer processes (common PDF viewers)
        pdf_viewers = ['acrobat', 'adobe', 'foxit', 'sumatra', 'chrome', 'firefox', 'edge']
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info['name']:
                    proc_name = proc.info['name'].lower()
                    if any(viewer in proc_name for viewer in pdf_viewers):
                        # Only close if it's likely a PDF viewer (not the main browser)
                        if 'acrobat' in proc_name or 'adobe' in proc_name or 'foxit' in proc_name or 'sumatra' in proc_name:
                            proc.terminate()
                            print("Closed existing PDF viewer:", proc.info['name'], "(PID:", proc.info['pid'], ")")
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
                
        # Give processes a moment to close
        import time
        time.sleep(1)
        
    except Exception as e:
        print("Warning: Could not close existing processes:", str(e))

class FileExporter:
    
    def __init__(self, data, jira_exporter=None):
        # Handle both tuple and list formats
        if isinstance(data, tuple):
            self.data = data[0]  # Extract the data list
            self.main_ticket_description = data[1]  # Extract the main ticket description
            self.parent_progress = data[2] if len(data) > 2 else ""  # Extract parent progress
            self.parent_status = data[3] if len(data) > 3 else ''  # Extract parent status
        else:
            # Assume data is already the list
            self.data = data
            self.main_ticket_description = ""
            self.parent_progress = ""
            self.parent_status = ""
        
        self.jira_exporter = jira_exporter
        os.makedirs(EXPORT_FOLDER, exist_ok=True)
    
    def close_existing_files(self):
        """Close any existing Excel and PDF viewer processes to prevent multiple files from opening"""
        try:
            # Close Excel processes
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    if proc.info['name'] and 'excel' in proc.info['name'].lower():
                        proc.terminate()
                        print("Closed existing Excel process:", proc.info['name'], "(PID:", proc.info['pid'], ")")
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    pass
            
            # Close PDF viewer processes (common PDF viewers)
            pdf_viewers = ['acrobat', 'adobe', 'foxit', 'sumatra', 'chrome', 'firefox', 'edge']
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    if proc.info['name']:
                        proc_name = proc.info['name'].lower()
                        if any(viewer in proc_name for viewer in pdf_viewers):
                            # Only close if it's likely a PDF viewer (not the main browser)
                            if 'acrobat' in proc_name or 'adobe' in proc_name or 'foxit' in proc_name or 'sumatra' in proc_name:
                                proc.terminate()
                                print("Closed existing PDF viewer:", proc.info['name'], "(PID:", proc.info['pid'], ")")
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    pass
                    
            # Give processes a moment to close
            import time
            time.sleep(1)
            
        except Exception as e:
            print("Warning: Could not close existing processes:", str(e))

    def to_excel(self, project_key):
        
        # Check if data is empty
        if not self.data:
            print(f"No data to export for project {project_key} - creating placeholder file")
            
            # Create a minimal Excel file with just headers and a message
            now = datetime.now()
            timestamp = now.strftime("%Y%m%d%H%M")
            file_name = f'{timestamp} {project_key}.xlsx'
            file_path = os.path.join(EXPORT_FOLDER, file_name)
            
            # Create a simple DataFrame with a message
            message_data = [{
                'Message': f'No defects found for EPIC {project_key}',
                'Suggestion': 'Check if the EPIC exists and has child work items with Bug/Defect subtasks'
            }]
            df = pd.DataFrame(message_data)
            df.to_excel(file_path, index=False)
            return file_path
        
        # Generate timestamped filename with new format: yyyymmddhhmm Project_key
        now = datetime.now()
        timestamp = now.strftime("%Y%m%d%H%M")
        file_name = f'{timestamp} {project_key}.xlsx'
        file_path = os.path.join(EXPORT_FOLDER, file_name)
        
        # Create Excel writer with xlsxwriter engine
        writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
        
        # Convert data to DataFrame
        df = pd.DataFrame(self.data)
        
        # Convert any dictionary values to strings
        for col in df.columns:
            df[col] = df[col].apply(lambda x: str(x) if isinstance(x, dict) else x)
        
        # Write the DataFrame to Excel first to create the worksheet
        df.to_excel(writer, index=False, sheet_name='Jira Report', startrow=3)
        
        # Get the xlsxwriter workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['Jira Report']
        
        # Write the main title header
        worksheet.merge_range('A1:O1', f'{project_key} - Jira Report', workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'bg_color': '#366092',
            'font_color': 'white',
            'valign': 'top',
            'align': 'left',
            'border': 1,
            'font_size': 14
        }))
        
        # Define header format
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'align': 'center',
            'bg_color': '#D9E1F2',
            'border': 1
        })
        
        # Write the main ticket description as a header
        worksheet.merge_range('A2:O2', str(self.main_ticket_description) if self.main_ticket_description else 'No description available', workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'bg_color': '#cbf9f4', 
            'valign': 'top',
            'align': 'left',
            'border': 1
        }))
        
        # Define status-specific cell formats (always center-aligned for status column)
        done_format = workbook.add_format({
            'bg_color': '#92D050',  # Light green
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        todo_format = workbook.add_format({
            'bg_color': '#ff545d',  # Red
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        in_progress_format = workbook.add_format({
            'bg_color': '#FFEB9C',  # Light yellow
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        # Define column-specific cell formats
        center_format = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        left_format = workbook.add_format({
            'border': 1,
            'align': 'left',
            'valign': 'vcenter',
            'text_wrap': True
        })
        
        # Apply column-specific alignment to all data cells
        left_align_columns = ['Summary', 'Assignee History', 'Description', 'Defect Description', 'Phase Detected', 'Phase Injected']  # Include variations
        
        for row_num in range(len(df)):
            for col_num in range(len(df.columns)):
                cell_value = df.iloc[row_num, col_num]
                column_name = df.columns[col_num]
                
                # Choose format based on column name
                if column_name in left_align_columns:
                    worksheet.write(row_num + 4, col_num, cell_value, left_format)
                else:
                    worksheet.write(row_num + 4, col_num, cell_value, center_format)
        
        # Get status column index for conditional formatting
        if 'Status' in df.columns:
            status_col = df.columns.get_loc('Status')
            
            # Apply conditional formatting based on status
            for row_num, status in enumerate(df['Status'], start=4):  # start=4 to account for title and description headers
                status = str(status).upper().strip()
                if status and status != 'NAN' and status != 'NONE':
                    # Green for DONE
                    if 'DONE' in status:
                        worksheet.write(row_num, status_col, status, done_format)
                    # Yellow for In-review, RSA IN PROGRESS, DEV IN PROGRESS, QA IN PROGRESS, UAT IN PROGRESS (check first)
                    elif any(keyword in status for keyword in ['IN REVIEW', 'IN PROGRESS', 'RSA IN PROGRESS', 'DEV IN PROGRESS', 'QA IN PROGRESS', 'UAT IN PROGRESS']):
                        worksheet.write(row_num, status_col, status, in_progress_format)
                    # Red for TO-DO, TODO, QA, RSA, DEV, PROD, Ready for QA, UAT (check last)
                    elif any(keyword in status for keyword in ['TO DO', 'TODO', 'QA', 'RSA', 'DEV', 'PROD', 'READY FOR QA', 'UAT']):
                        worksheet.write(row_num, status_col, status, todo_format)
        
        # Auto-adjust column widths based on content
        for idx, col in enumerate(df.columns):
            # Calculate maximum content length
            max_length = max(
                df[col].astype(str).apply(len).max(),  # max length of data
                len(col)  # length of column name
            )
            # Set column width with padding
            worksheet.set_column(idx, idx, max_length + 2)
        
        # Create the Report tab with Serial Number, Key, Summary, Status
        available_columns = ['Key', 'Summary']
        if 'Status' in df.columns:
            available_columns.append('Status')
        else:
            print("Warning: 'Status' column not found, creating report without status")
            # Add a default Status column with 'Unknown' values
            df['Status'] = 'Unknown'
            available_columns.append('Status')
            
        report_df = df[available_columns].copy()
        report_df.insert(0, 'S.No', range(1, len(report_df) + 1))
        # Rename columns for the Report sheet
        report_df = report_df.rename(columns={
            'Key': 'Defect Key',
            'Summary': 'Defect Description'
        })
        
        # Create QA Report sheet with summary data in A1 and defect table starting from A5
        self._create_qa_report_sheet(writer, workbook, report_df, df, project_key)
        
        
        # Create Metrics Board worksheet
        self._create_metrics_board_worksheet(writer, df, project_key)
        
        # Create Visual Board worksheet with processed metrics data
        chart_data = self._get_chart_data_from_metrics(df)
        self._create_visual_board_worksheet(writer, df, project_key, chart_data)
        
        # Save the Excel file first
        writer.close()
        
        # Use the file_path that was already created with the new naming format
        excel_file_path = file_path
        
        # Small delay to ensure Excel file is properly saved
        import time
        time.sleep(1)
        
        # Generate Power BI dashboard directly (simplified output)
        pdf_file = self._generate_powerbi_report_directly(excel_file_path, project_key)
        
        # If PDF generation failed, return None for PDF path
        if pdf_file is False or pdf_file is None or not isinstance(pdf_file, str):
            pdf_file = None
        else:
            # Ensure PDF file path is absolute and exists
            try:
                if not os.path.isabs(pdf_file):
                    # If relative path, make it absolute
                    pdf_file = os.path.abspath(pdf_file)
                if not os.path.exists(pdf_file):
                    pdf_file = None
            except (TypeError, ValueError):
                pdf_file = None
        
        # Return both Excel and PDF file paths
        return file_path, pdf_file
    
    def _create_qa_report_sheet(self, writer, workbook, report_df, df, project_key):
        """Create QA Report sheet with summary data in A1 and defect table starting from A5"""
        
        # Create QA Report worksheet
        qa_ws = workbook.add_worksheet('QA Report')
        
        # Calculate summary data (similar to QA Summary sheet)
        open_count = 0
        closed_count = 0
        
        if 'Status' in df.columns:
            for status in df['Status']:
                status = str(status).lower()
                if status and status != 'nan' and status != 'none':
                    # Check if status is closed
                    if 'done' in status or 'completed' in status or 'closed' in status or 'resolved' in status:
                        closed_count += 1
                    else:
                        # All other statuses are considered open
                        open_count += 1
        
        # Extract project name from main ticket description or use a default
        project_name = str(self.main_ticket_description) if self.main_ticket_description else "EPIC Defect Summary"
        if not project_name or project_name == "None":
            project_name = "EPIC Defect Summary"
        
        # Create summary data in A1 (similar to QA Summary format)
        summary_data = [
            ['Project', 'Ticket Status', 'Open Defects', 'Closed Defects', 'Prod Status', 'QA Complete %', 'Notes'],
            [project_name, self.parent_status, open_count, closed_count, 'NA', '', '']
        ]
        
        # Define formats for summary section
        summary_header_format = workbook.add_format({
            'bold': True,
            'valign': 'top',
            'bg_color': '#D9E1F2',
            'border': 1,
            'align': 'center'
        })
        
        summary_data_format = workbook.add_format({
            'valign': 'top',
            'border': 1,
            'align': 'center'
        })
        
        # Write summary data starting from F1 (column 5, 0-indexed)
        summary_start_col = 5
        for row_idx, row_data in enumerate(summary_data):
            for col_idx, cell_value in enumerate(row_data):
                if row_idx == 0:  # Header row
                    qa_ws.write(row_idx, summary_start_col + col_idx, cell_value, summary_header_format)
                else:  # Data row
                    qa_ws.write(row_idx, summary_start_col + col_idx, cell_value, summary_data_format)
        
        # Auto-adjust column widths for summary section (starting from column F)
        for col_idx, header in enumerate(summary_data[0]):
            # Calculate maximum content length for this column
            max_content_length = len(header)  # Start with header length
            
            # Check data row content length
            if len(summary_data) > 1:
                data_content = str(summary_data[1][col_idx])
                max_content_length = max(max_content_length, len(data_content))
            
            # Set column width with padding (minimum 12 characters)
            column_width = max(max_content_length + 2, 12)
            qa_ws.set_column(summary_start_col + col_idx, summary_start_col + col_idx, column_width)
        
        # Now write the defect data table starting from A1 (row 0, column 0, 0-indexed)
        start_row = 0
        start_col = 0
        
        # Define QA Report column-specific formats
        qa_center_format = workbook.add_format({
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
        })
        
        qa_left_format = workbook.add_format({
            'border': 1,
            'align': 'left',
            'valign': 'vcenter'
        })
        
        # Apply center alignment to headers
        qa_header_format = workbook.add_format({
                'bold': True,
                'border': 1,
                'align': 'center',
                'valign': 'vcenter',
            'bg_color': '#D9E1F2'
        })
        
        # Write headers at A1 (row 0, column 0)
        for col_num, header in enumerate(report_df.columns):
            qa_ws.write(start_row, start_col + col_num, header, qa_header_format)
        
        # Apply column-specific formatting to all data cells
        qa_left_align_columns = ['Defect Description']  # Summary equivalent in QA Report
        
        for row_num in range(len(report_df)):
            for col_num in range(len(report_df.columns)):
                cell_value = report_df.iloc[row_num, col_num]
                column_name = report_df.columns[col_num]
                
                # Choose format based on column name
                if column_name in qa_left_align_columns:
                    qa_ws.write(start_row + 1 + row_num, start_col + col_num, cell_value, qa_left_format)
                else:
                    qa_ws.write(start_row + 1 + row_num, start_col + col_num, cell_value, qa_center_format)
        
        # Define status-specific cell formats for QA Report
        qa_done_format = workbook.add_format({
            'bg_color': '#92D050',  # Light green
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
            })
            
        qa_todo_format = workbook.add_format({
            'bg_color': '#ff545d',  # Red
                'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        qa_in_progress_format = workbook.add_format({
            'bg_color': '#FFEB9C',  # Light yellow
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        # Apply status-based conditional formatting to QA Report
        if 'Status' in report_df.columns:
            qa_status_col = report_df.columns.get_loc('Status')
            
            for row_num, status in enumerate(report_df['Status'], start=1):
                status = str(status).upper().strip()
                if status and status != 'NAN' and status != 'NONE':
                    # Green for DONE
                    if 'DONE' in status:
                        qa_ws.write(start_row + row_num, start_col + qa_status_col, status, qa_done_format)
                    # Yellow for In-review, RSA IN PROGRESS, DEV IN PROGRESS, QA IN PROGRESS, UAT IN PROGRESS (check first)
                    elif any(keyword in status for keyword in ['IN REVIEW', 'IN PROGRESS', 'RSA IN PROGRESS', 'DEV IN PROGRESS', 'QA IN PROGRESS', 'UAT IN PROGRESS']):
                        qa_ws.write(start_row + row_num, start_col + qa_status_col, status, qa_in_progress_format)
                    # Red for TO-DO, TODO, QA, RSA, DEV, PROD, Ready for QA, UAT (check last)
                    elif any(keyword in status for keyword in ['TO DO', 'TODO', 'QA', 'RSA', 'DEV', 'PROD', 'READY FOR QA', 'UAT']):
                        qa_ws.write(start_row + row_num, start_col + qa_status_col, status, qa_todo_format)
        
        # Auto-adjust column widths for defect table (starting from column A)
        for idx, col in enumerate(report_df.columns):
                max_length = max(
                    report_df[col].astype(str).apply(len).max(),
                    len(col)
                )
                qa_ws.set_column(start_col + idx, start_col + idx, max_length + 2)


    def _calculate_defect_metrics(self, defects_data):
        """Calculate defect leakage metrics from ODC data"""
        try:
            # Initialize the defect leakage matrix
            phases = ['Requirements', 'Design', 'Implementation', 'Unit testing', 'QA', 'UAT', 'Production']
            matrix = {}
            
            # Initialize matrix with zeros
            for introduced in phases:
                matrix[introduced] = {}
                for detected in phases:
                    matrix[introduced][detected] = 0
            
            # Count defects based on ODC data
            for defect in defects_data:
                phase_injected = defect.get('ODC Phase Injected', 'Not Specified')
                phase_detected = defect.get('ODC Phase Detected', 'Not Specified')
                
                # Skip if either phase is not specified
                if phase_injected == 'Not Specified' or phase_detected == 'Not Specified':
                    continue
                
                # Map ODC phases to matrix phases
                injected_mapping = {
                    'Requirements': 'Requirements',
                    'Design': 'Design', 
                    'Code': 'Implementation',
                    'Unit Test': 'Unit testing',
                    'Unit test': 'Unit testing',  # Handle lowercase variant
                    'System Test': 'QA',
                    'UAT': 'UAT',
                    'Production': 'Production'
                }
                
                detected_mapping = {
                    'Requirements': 'Requirements',
                    'Design': 'Design',
                    'Code': 'Implementation', 
                    'Unit Test': 'Unit testing',
                    'Unit Test': 'Unit testing',  # Handle "Unit Test Phase" variant
                    'System Test': 'QA',
                    'UAT': 'UAT',
                    'Production': 'Production'
                }
                
                mapped_injected = injected_mapping.get(phase_injected, phase_injected)
                mapped_detected = detected_mapping.get(phase_detected, phase_detected)
                
                # Only count if both phases are valid
                if mapped_injected in phases and mapped_detected in phases:
                    matrix[mapped_injected][mapped_detected] += 1
            
            # Calculate totals
            totals_injected = {}
            totals_detected = {}
            
            for phase in phases:
                totals_injected[phase] = sum(matrix[phase].values())
                totals_detected[phase] = sum(matrix[p][phase] for p in phases)
            
            # Calculate percentages
            total_defects = sum(totals_injected.values())
            review_defects = sum(matrix['Requirements'].values()) + sum(matrix['Design'].values()) + sum(matrix['Implementation'].values())
            testing_defects = sum(matrix['Unit testing'].values()) + sum(matrix['QA'].values()) + sum(matrix['UAT'].values()) + sum(matrix['Production'].values())
            
            percentage_review = (review_defects / total_defects * 100) if total_defects > 0 else 0
            percentage_testing = (testing_defects / total_defects * 100) if total_defects > 0 else 0
            
            return {
                'matrix': matrix,
                'totals_injected': totals_injected,
                'totals_detected': totals_detected,
                'total_defects': total_defects,
                'percentage_review': percentage_review,
                'percentage_testing': percentage_testing,
                'phases': phases
            }
            
        except Exception as e:
            print(f"Error calculating defect metrics: {e}")
            return None 
    
    def _get_chart_data_from_metrics(self, df):
        """Extract chart data using the same logic as Defect Metrics sheet"""
        chart_data = {}
        
        # Status Breakdown
        status_column = None
        for col in df.columns:
            if col.lower() == 'status':
                status_column = col
                break
        
        if status_column and not df[status_column].empty:
            status_data = df[status_column].dropna()
            status_data = status_data[status_data != '']
            status_data = status_data[status_data != ' ']
            status_counts = status_data.value_counts()
            chart_data['status_breakdown'] = status_counts.to_dict()
        else:
            chart_data['status_breakdown'] = {}
        
        # Priority Distribution
        priority_column = None
        for col in df.columns:
            if col.lower() == 'priority':
                priority_column = col
                break
        
        if priority_column and not df[priority_column].empty:
            priority_data = df[priority_column].dropna()
            priority_data = priority_data[priority_data != '']
            priority_data = priority_data[priority_data != ' ']
            priority_counts = priority_data.value_counts()
            chart_data['priority_distribution'] = priority_counts.to_dict()
        else:
            chart_data['priority_distribution'] = {}
        
        # Environment Analysis
        environment_column = None
        for col in df.columns:
            if col.lower() == 'environment':
                environment_column = col
                break
        
        if environment_column and not df[environment_column].empty:
            environment_data = df[environment_column].dropna()
            environment_data = environment_data[environment_data != '']
            environment_data = environment_data[environment_data != ' ']
            environment_counts = environment_data.value_counts()
            chart_data['environment_analysis'] = environment_counts.to_dict()
        else:
            chart_data['environment_analysis'] = {}
        
        # Assignee Workload
        assignee_column = None
        for col in df.columns:
            if col.lower() == 'assignee':
                assignee_column = col
                break
        
        if assignee_column and not df[assignee_column].empty:
            assignee_data = df[assignee_column].dropna()
            assignee_data = assignee_data[assignee_data != '']
            assignee_data = assignee_data[assignee_data != ' ']
            assignee_counts = assignee_data.value_counts()
            chart_data['assignee_workload'] = assignee_counts.to_dict()
        else:
            chart_data['assignee_workload'] = {}
        
        return chart_data
    
    def _calculate_open_closed_counts(self, df):
        """Calculate open and closed defect counts"""
        if 'Status' not in df.columns:
            return 0, 0
        
        status_counts = df['Status'].value_counts()
        closed_statuses = ['DONE', 'COMPLETED', 'CLOSED', 'RESOLVED']
        
        open_count = 0
        closed_count = 0
        
        for status, count in status_counts.items():
            status_str = str(status).upper().strip()
            # Check if status is closed
            if any(closed in status_str for closed in closed_statuses):
                closed_count += count
            else:
                # All other statuses are considered open
                open_count += count
        
        return open_count, closed_count
    
    def _calculate_resolution_times(self, df):
        """Calculate resolution times for defects"""
        resolution_times = []
        
        if 'Created Date' not in df.columns or 'Resolution Date' not in df.columns:
            return resolution_times
        
        for _, row in df.iterrows():
            created = row.get('Created Date', '')
            resolved = row.get('Resolution Date', '')
            
            if created and resolved and created != 'N/A' and resolved != 'N/A':
                try:
                    # Parse dates (assuming format like "2025-09-10 13:23:13")
                    from datetime import datetime
                    created_dt = datetime.strptime(str(created), "%Y-%m-%d %H:%M:%S")
                    resolved_dt = datetime.strptime(str(resolved), "%Y-%m-%d %H:%M:%S")
                    # Calculate time difference in hours
                    time_diff = resolved_dt - created_dt
                    resolution_time = time_diff.total_seconds() / 3600
                    resolution_times.append(resolution_time)
                except:
                    continue
        
        return resolution_times
    
    def _create_metrics_board_worksheet(self, writer, df, project_key):
        """Create Metrics Board worksheet with defect leakage matrix"""
        try:
            workbook = writer.book
            metrics_ws = workbook.add_worksheet('Metrics Board')
            
            # Define phases
            phases = ['Requirements', 'Design', 'Implementation', 'Unit testing', 'QA', 'UAT', 'Production']
            
            # Start writing data from row 3 (after main header in row 1 and column headers in row 2)
            current_row = 3
            
            # Overall Statistics
            metrics_ws.merge_range(f'A{current_row}:C{current_row}', 'Overall Statistics', section_header_format)
            #current_row += 1
            
            total_defects = len(df)
            open_count, closed_count = self._calculate_open_closed_counts(df)
            completion_percentage = (closed_count / total_defects * 100) if total_defects > 0 else 0
            
            metrics_ws.write(current_row, 0, 'Total Defects', category_format)
            metrics_ws.write(current_row, 1, total_defects, value_format)
            metrics_ws.write(current_row, 2, '', percentage_format)
            current_row += 1
            
            metrics_ws.write(current_row, 0, 'Open Defects', category_format)
            metrics_ws.write(current_row, 1, open_count, value_format)
            metrics_ws.write(current_row, 2, '', percentage_format)
            current_row += 1
            
            metrics_ws.write(current_row, 0, 'Closed Defects', category_format)
            metrics_ws.write(current_row, 1, closed_count, value_format)
            metrics_ws.write(current_row, 2, '', percentage_format)
            current_row += 1
            
            metrics_ws.write(current_row, 0, 'Completion %', category_format)
            metrics_ws.write(current_row, 1, f"{completion_percentage:.1f}%", value_format)
            metrics_ws.write(current_row, 2, '', percentage_format)
            current_row += 1  # No spacing
            current_row += 2
            # Status Breakdown
            metrics_ws.merge_range(f'A{current_row}:C{current_row}', 'Status Breakdown', section_header_format)
           # current_row += 1
            
            # Check for Status column (case insensitive)
            status_column = None
            for col in df.columns:
                if col.lower() == 'status':
                    status_column = col
                    break
            
            if status_column and not df[status_column].empty:
                # Remove null/empty values before counting
                status_data = df[status_column].dropna()
                status_data = status_data[status_data != '']
                status_data = status_data[status_data != ' ']
                
                status_counts = status_data.value_counts()
                # Calculate total count for this section to get 100% total
                section_total = status_counts.sum()
                
                for status, count in status_counts.items():
                    percentage = (count / section_total * 100) if section_total > 0 else 0
                    # Convert status to proper case for display
                    display_status = str(status).title() if str(status).upper() == str(status) else str(status)
                    metrics_ws.write(current_row, 0, display_status, category_format)
                    metrics_ws.write(current_row, 1, count, value_format)
                    metrics_ws.write(current_row, 2, f"{percentage:.1f}%", percentage_format)
                    current_row += 1
            else:
                # If no Status column or empty data, show a message
                metrics_ws.write(current_row, 0, 'No Status Data Available', category_format)
                metrics_ws.write(current_row, 1, '', value_format)
                metrics_ws.write(current_row, 2, '', percentage_format)
                current_row += 1
            # No spacing
            current_row += 2
            # Priority Distribution
            metrics_ws.merge_range(f'A{current_row}:C{current_row}', 'Priority Distribution', section_header_format)
             #current_row += 1
            
            # Check for Priority column (case insensitive)
            priority_column = None
            for col in df.columns:
                if col.lower() == 'priority':
                    priority_column = col
                    break
            
            if priority_column and not df[priority_column].empty:
                # Remove null/empty values before counting
                priority_data = df[priority_column].dropna()
                priority_data = priority_data[priority_data != '']
                priority_data = priority_data[priority_data != ' ']
                
                priority_counts = priority_data.value_counts()
                # Calculate total count for this section to get 100% total
                section_total = priority_counts.sum()
                for priority, count in priority_counts.items():
                    percentage = (count / section_total * 100) if section_total > 0 else 0
                    metrics_ws.write(current_row, 0, str(priority), category_format)
                    metrics_ws.write(current_row, 1, count, value_format)
                    metrics_ws.write(current_row, 2, f"{percentage:.1f}%", percentage_format)
                    current_row += 1
            # No spacing
            current_row += 2
            # Environment Analysis
            metrics_ws.merge_range(f'A{current_row}:C{current_row}', 'Environment Analysis', section_header_format)
            #current_row += 1
            
            # Check for Environment column (case insensitive)
            env_column = None
            for col in df.columns:
                if col.lower() == 'environment':
                    env_column = col
                    break
            
            if env_column and not df[env_column].empty:
                # Remove null/empty values before counting
                env_data = df[env_column].dropna()
                env_data = env_data[env_data != '']
                env_data = env_data[env_data != ' ']
                
                env_counts = env_data.value_counts()
                # Calculate total count for this section to get 100% total
                section_total = env_counts.sum()
                for env, count in env_counts.items():
                    percentage = (count / section_total * 100) if section_total > 0 else 0
                    metrics_ws.write(current_row, 0, str(env), category_format)
                    metrics_ws.write(current_row, 1, count, value_format)
                    metrics_ws.write(current_row, 2, f"{percentage:.1f}%", percentage_format)
                    current_row += 1
            # No spacing
            current_row += 2
            # Assignee Workload
            metrics_ws.merge_range(f'A{current_row}:C{current_row}', 'Assignee Workload', section_header_format)
            #current_row += 1
            
            # Check for Assignee column (case insensitive)
            assignee_column = None
            for col in df.columns:
                if col.lower() == 'assignee':
                    assignee_column = col
                    break
            
            if assignee_column and not df[assignee_column].empty:
                # Remove null/empty values before counting
                assignee_data = df[assignee_column].dropna()
                assignee_data = assignee_data[assignee_data != '']
                assignee_data = assignee_data[assignee_data != ' ']
                
                assignee_counts = assignee_data.value_counts()
                # Calculate total count for this section to get 100% total
                section_total = assignee_counts.sum()
                for assignee, count in assignee_counts.items():
                    percentage = (count / section_total * 100) if section_total > 0 else 0
                    metrics_ws.write(current_row, 0, str(assignee), category_format)
                    metrics_ws.write(current_row, 1, count, value_format)
                    metrics_ws.write(current_row, 2, f"{percentage:.1f}%", percentage_format)
                    current_row += 1
            # No spacing
            current_row += 2
            # Time Metrics
            metrics_ws.merge_range(f'A{current_row}:C{current_row}', 'Time Metrics', section_header_format)
            #current_row += 1
            
            # Calculate resolution times
            resolution_times = self._calculate_resolution_times(df)
            if resolution_times:
                avg_resolution = sum(resolution_times) / len(resolution_times)
                min_resolution = min(resolution_times)
                max_resolution = max(resolution_times)
                
                # Average Resolution Time
                metrics_ws.write(current_row, 0, 'Avg Resolution Time', category_format)
                metrics_ws.write(current_row, 1, f"{avg_resolution:.1f} days", value_format)
                metrics_ws.write(current_row, 2, '', percentage_format)
                current_row += 1
                
                # Minimum Resolution Time
                metrics_ws.write(current_row, 0, 'Min Resolution Time', category_format)
                metrics_ws.write(current_row, 1, f"{min_resolution:.1f} days", value_format)
                metrics_ws.write(current_row, 2, '', percentage_format)
                current_row += 1
                
                # Maximum Resolution Time
                metrics_ws.write(current_row, 0, 'Max Resolution Time', category_format)
                metrics_ws.write(current_row, 1, f"{max_resolution:.1f} days", value_format)
                metrics_ws.write(current_row, 2, '', percentage_format)
                current_row += 1
                
                # Resolved Defects Count
                metrics_ws.write(current_row, 0, 'Resolved Defects', category_format)
                metrics_ws.write(current_row, 1, len(resolution_times), value_format)
                metrics_ws.write(current_row, 2, '', percentage_format)
            else:
                # Average Resolution Time
                metrics_ws.write(current_row, 0, 'Avg Resolution Time', category_format)
                metrics_ws.write(current_row, 1, 'N/A', value_format)
                metrics_ws.write(current_row, 2, '', percentage_format)
                current_row += 1

                # Minimum Resolution Time
                metrics_ws.write(current_row, 0, 'Min Resolution Time', category_format)
                metrics_ws.write(current_row, 1, 'N/A', value_format)
                metrics_ws.write(current_row, 2, '', percentage_format)
                current_row += 1

                # Maximum Resolution Time
                metrics_ws.write(current_row, 0, 'Max Resolution Time', category_format)
                metrics_ws.write(current_row, 1, 'N/A', value_format)
                metrics_ws.write(current_row, 2, '', percentage_format)
                current_row += 1

                # Resolved Defects Count
                metrics_ws.write(current_row, 0, 'Resolved Defects', category_format)
                metrics_ws.write(current_row, 1, 0, value_format)
                metrics_ws.write(current_row, 2, '', percentage_format)
            
            # Auto-adjust column widths
            metrics_ws.set_column('A:A', 25)  # Category column
            metrics_ws.set_column('B:B', 15)  # Value column
            metrics_ws.set_column('C:C', 15)  # Percentage column
            
        except Exception as e:
            print(f"Error creating Defect Metrics worksheet: {str(e)}")
            return
    
    def _calculate_open_closed_counts(self, df):
        """Calculate open and closed defect counts"""
        if 'Status' not in df.columns:
            return 0, 0
        
        status_counts = df['Status'].value_counts()
        closed_statuses = ['DONE', 'COMPLETED', 'CLOSED', 'RESOLVED']
        
        open_count = 0
        closed_count = 0
        
        for status, count in status_counts.items():
            status_str = str(status).upper().strip()
            # Check if status is closed
            if any(closed in status_str for closed in closed_statuses):
                closed_count += count
            else:
                # All other statuses are considered open
                open_count += count
        
        return open_count, closed_count
    
    def _calculate_resolution_times(self, df):
        """Calculate resolution times for defects"""
        resolution_times = []
        
        if 'Created Date' not in df.columns or 'Resolution Date' not in df.columns:
            return resolution_times
        
        for _, row in df.iterrows():
            created = row.get('Created Date', '')
            resolved = row.get('Resolution Date', '')
            
            if created and resolved and created != 'N/A' and resolved != 'N/A':
                try:
                    # Parse dates (assuming format like "2025-09-10 13:23:13")
                    from datetime import datetime
                    created_dt = datetime.strptime(str(created), "%Y-%m-%d %H:%M:%S")
                    resolved_dt = datetime.strptime(str(resolved), "%Y-%m-%d %H:%M:%S")
                    # Calculate time difference in hours
                    time_diff = resolved_dt - created_dt
                    resolution_time = time_diff.total_seconds() / 3600
                    resolution_times.append(resolution_time)
                except:
                    continue
        
        return resolution_times
    
    def _create_metrics_board_worksheet(self, writer, df, project_key):
        """Create Metrics Board worksheet with defect leakage matrix"""
        try:
            workbook = writer.book
            metrics_ws = workbook.add_worksheet('Metrics Board')
            
            # Define phases
            phases = ['Requirements', 'Design', 'Implementation', 'Unit testing', 'QA', 'UAT', 'Production']
            
            # Create defect leakage matrix
            matrix = {}
            for introduced in phases:
                matrix[introduced] = {}
                for detected in phases:
                    matrix[introduced][detected] = 0
            
            # Count defects based on Phase Detected and Phase Injected data
            for _, row in df.iterrows():
                phase_detected = str(row.get('Phase Detected', '')).strip()
                phase_injected = str(row.get('Phase Injected', '')).strip()
                
                # Skip if either phase is empty or 'Not Specified'
                if not phase_detected or not phase_injected or phase_detected == 'Not Specified' or phase_injected == 'Not Specified':
                    continue
                
                # Map ODC phases to matrix phases
                detected_mapping = {
                    'Requirements': 'Requirements',
                    'Design': 'Design',
                    'Code': 'Implementation',
                    'Code': 'Implementation',  # Updated mapping
                    'Unit Test': 'Unit testing',
                    'System Test': 'QA',
                    'QA': 'QA',
                    'UAT (User Acceptance Test)': 'UAT',
                    'UAT': 'UAT',
                    'Production': 'Production'
                }
                
                injected_mapping = {
                    'Requirements': 'Requirements',
                    'Design': 'Design',
                    'Code': 'Implementation',
                    'Code ': 'Implementation',
                    'Unit Test': 'Unit testing',
                    'Unit test': 'Unit testing',  # Handle lowercase variant
                    'Unit testing': 'Unit testing',
                    'QA': 'QA',
                    'UAT': 'UAT',
                    'Production': 'Production'
                }
                
                mapped_detected = detected_mapping.get(phase_detected, phase_detected)
                mapped_injected = injected_mapping.get(phase_injected, phase_injected)
                
                # Only count if both phases are valid
                if mapped_injected in phases and mapped_detected in phases:
                    matrix[mapped_injected][mapped_detected] += 1
            
            # Calculate totals
            totals_injected = {}
            totals_detected = {}
            
            for phase in phases:
                totals_injected[phase] = sum(matrix[phase].values())
                totals_detected[phase] = sum(matrix[p][phase] for p in phases)
            
            # Calculate percentages
            total_defects = sum(totals_injected.values())
            review_defects = totals_detected['Requirements'] + totals_detected['Design'] + totals_detected['Implementation']
            testing_defects = totals_detected['Unit testing'] + totals_detected['QA'] + totals_detected['UAT'] + totals_detected['Production']
            
            percentage_review = (review_defects / total_defects * 100) if total_defects > 0 else 0
            percentage_testing = (testing_defects / total_defects * 100) if total_defects > 0 else 0
            
            # Define formats
            title_format = workbook.add_format({
                'bold': True,
                'font_size': 14,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': 'white',
                'border': 1,
                'text_wrap': True
            })
            
            header_format = workbook.add_format({
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#D9D9D9',
                'border': 1
            })
            
            phase_header_format = workbook.add_format({
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#8FAADC',
                'border': 1
            })
            
            review_header_format = workbook.add_format({
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#D5E8D4',
                'border': 1
            })
            
            testing_header_format = workbook.add_format({
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#F8CECC',
                'border': 1
            })
            
            total_header_format = workbook.add_format({
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#FFE699',
                'border': 1
            })
            
            data_format = workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'border': 1
            })
            
            white_data_format = workbook.add_format({
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': 'white',
                'border': 1
            })
            
            gray_format = workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#D9D9D9',
                'border': 1
            })
            
            # Write title in A3:B5 (merged across columns with wrap text)
            metrics_ws.merge_range('A3:B5', 'Defect Leakage Metrics', title_format)
            
            # Write main headers in row 3 (matching sample)
            metrics_ws.merge_range('C3:I3', 'Phase detected', header_format)
            
            # Write sub-category headers in row 4 (matching sample)
            metrics_ws.merge_range('C4:E4', 'Review', review_header_format)
            metrics_ws.merge_range('F4:I4', 'Testing', testing_header_format)
            
            # Write phase headers in row 5 (matching sample) - using direct cell references
            metrics_ws.write('C5', 'Requirements', phase_header_format)
            metrics_ws.write('D5', 'Design', phase_header_format)
            metrics_ws.write('E5', 'Implementation', phase_header_format)
            metrics_ws.write('F5', 'Unit testing', phase_header_format)
            metrics_ws.write('G5', 'QA', phase_header_format)
            metrics_ws.write('H5', 'UAT', phase_header_format)
            metrics_ws.write('I5', 'Production', phase_header_format)
            
            # Write "Total injected" header in J3:J5 (merged range)
            metrics_ws.merge_range('J3:J5', 'Total injected', total_header_format)
            
            
            # Write "Phase injected" header in A6:A12 (merged vertically with vertical text)
            phase_injected_format = workbook.add_format({
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#D9D9D9',
                'border': 1,
                'rotation': 90  # Vertical text
            })
            metrics_ws.merge_range('A6:A12', 'Phase injected', phase_injected_format)
            
            # Write data rows starting from row 6 (matching sample)
            for i, introduced_phase in enumerate(phases):
                row = i + 5
                metrics_ws.write(row, 1, introduced_phase, phase_header_format)  # Column B
                
                for j, detected_phase in enumerate(phases):
                    col = j + 2  # Starting from column C (after removing column C, now starting from C)
                    count = matrix[introduced_phase][detected_phase]
                    
                    # Determine if cell should be greyed out (inactive)
                    # Grey out cells where defects injected in a later phase cannot be detected in an earlier phase
                    should_be_grey = False
                    
                    if j == 0:  # Column C (Requirements)
                        should_be_grey = row >= 6  # C7:C12 (Design onwards cannot be detected in Requirements)
                    elif j == 1:  # Column D (Design)
                        should_be_grey = row >= 7  # D8:D12 (Implementation onwards cannot be detected in Design)
                    elif j == 2:  # Column E (Implementation)
                        should_be_grey = row >= 8  # E9:E12 (Unit testing onwards cannot be detected in Implementation)
                    elif j == 3:  # Column F (Unit testing)
                        should_be_grey = row >= 9  # F10:F12 (QA onwards cannot be detected in Unit testing)
                    elif j == 4:  # Column G (QA)
                        should_be_grey = row >= 10  # G11:G12 (UAT onwards cannot be detected in QA)
                    elif j == 5:  # Column H (UAT)
                        should_be_grey = row >= 11  # H12 (Production cannot be detected in UAT)
                    # Column I (Production) - no grey cells (all phases can be detected in Production)
                    
                    if should_be_grey:
                        metrics_ws.write(row, col, '', gray_format)
                    else:
                        # Use actual defect count from Jira Report data, show '-' for 0
                        display_value = '-' if count == 0 else count
                        metrics_ws.write(row, col, display_value, data_format)
                
                # Write total injected in column J - only sum active (non-grey) cells
                # Based on row number, determine which columns are active
                if row == 6:  # Requirements row
                    formula = '=SUM(C6:I6)'  # All columns active
                elif row == 7:  # Design row
                    formula = '=SUM(D7:I7)'  # D onwards active
                elif row == 8:  # Implementation row
                    formula = '=SUM(E8:I8)'  # E onwards active
                elif row == 9:  # Unit testing row
                    formula = '=SUM(F9:I9)'  # F onwards active
                elif row == 10:  # QA row
                    formula = '=SUM(G10:I10)'  # G onwards active
                elif row == 11:  # UAT row
                    formula = '=SUM(H11:I11)'  # H onwards active
                elif row == 12:  # Production row
                    formula = '=I12'  # Only I active
                else:
                    formula = f'=SUM(C{row}:I{row})'  # Fallback
                
                # Write formula to column J (index 9) at the current row
                # Use write_formula with cell reference to ensure correct positioning
                cell_ref = f'J{row}'
                if row == 12:  # Skip J12 for now, write it later
                    pass
                else:
                    metrics_ws.write_formula(cell_ref, formula, white_data_format)
            
            # Write J12 after all data is written to ensure I12 has a value
            # Use formula to show 0 if I12 is '-' (text), otherwise show I12 value
            j12_formula = '=IF(I12="-", 0, I12)'
            metrics_ws.write_formula('J12', j12_formula, white_data_format)
            
            # Write total detected row in row 13 (matching sample)
            total_detected_format = workbook.add_format({
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#FFE699',
                'border': 1,
                'text_wrap': True
            })
            # Set row height for row 13 to accommodate wrapped text
            metrics_ws.set_row(12, 30)  # Row 13 (0-indexed as 12)
            
            metrics_ws.merge_range('A13:B13', 'Total Number of defects found and removed in phase', total_detected_format)
            for j, detected_phase in enumerate(phases):
                col_letter = chr(ord('C') + j)  # C, D, E, F, G, H, I
                # Only sum active (non-grey) cells for each column
                # Based on column, determine which rows are active
                if j == 0:  # Column C (Requirements)
                    formula = '=SUM(C6)'  # Only C6 active
                elif j == 1:  # Column D (Design)
                    formula = '=SUM(D6:D7)'  # D6:D7 active
                elif j == 2:  # Column E (Implementation)
                    formula = '=SUM(E6:E8)'  # E6:E8 active
                elif j == 3:  # Column F (Unit testing)
                    formula = '=SUM(F6:F9)'  # F6:F9 active
                elif j == 4:  # Column G (QA)
                    formula = '=SUM(G6:G10)'  # G6:G10 active
                elif j == 5:  # Column H (UAT)
                    formula = '=SUM(H6:H11)'  # H6:H11 active
                elif j == 6:  # Column I (Production)
                    formula = '=SUM(I6:I12)'  # I6:I12 active
                else:
                    formula = f'=SUM({col_letter}6:{col_letter}12)'  # Fallback
                
                cell_ref = f'{col_letter}13'
                metrics_ws.write_formula(cell_ref, formula, white_data_format)
            
            # Write total validation message in J13 using formula
            # Show "Total = [value]" and indicate if matching or not
            matching_formula = '=IF(SUM(J6:J12)=SUM(C13:I13), "Total = " & SUM(J6:J12) & " (Matching)", "Total = " & SUM(J6:J12) & " (Mismatch)")'
            
            # Create formats for matching and non-matching
            matching_format = workbook.add_format({
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#D5E8D4',  # Light green for matching
                'border': 1,
                'text_wrap': True
            })
            
            mismatch_format = workbook.add_format({
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#FFB6C1',  # Light red for mismatch
                'border': 1,
                'text_wrap': True
            })
            
            # Write the formula with matching format (green)
            # The formula will show the status in the text
            metrics_ws.write_formula('J13', matching_formula, matching_format)
            
            # Add conditional formatting to J13
            # Green background if SUM(J6:J12) = SUM(C13:I13), red if not
            metrics_ws.conditional_format('J13', {
                'type': 'formula',
                'criteria': '=SUM(J6:J12)=SUM(C13:I13)',
                'format': matching_format
            })
            
            metrics_ws.conditional_format('J13', {
                'type': 'formula', 
                'criteria': '=SUM(J6:J12)<>SUM(C13:I13)',
                'format': mismatch_format
            })
            
            # Calculate percentage formulas (hidden, used only for charts)
            # These formulas are needed for the donut charts but the labels are removed
            review_formula = '=IF(SUM(J6:J12)>0, (SUM(C13:E13)/SUM(J6:J12))*100, 0)'
            metrics_ws.write_formula('D15', review_formula, workbook.add_format({
                'num_format': '0.0"%"'
            }))
            
            testing_formula = '=IF(SUM(J6:J12)>0, (SUM(F13:I13)/SUM(J6:J12))*100, 0)'
            metrics_ws.write_formula('D16', testing_formula, workbook.add_format({
                'num_format': '0.0"%"'
            }))
            
            pre_production_formula = '=IF(SUM(J6:J12)>0, (SUM(C13:H13)/SUM(J6:J12))*100, 0)'
            metrics_ws.write_formula('D18', pre_production_formula, workbook.add_format({
                'num_format': '0.0"%"'
            }))
            
            post_production_formula = '=IF(SUM(J6:J12)>0, (I13/SUM(J6:J12))*100, 0)'
            metrics_ws.write_formula('D19', post_production_formula, workbook.add_format({
                'num_format': '0.0"%"'
            }))
            
            # Create chart data table for donut charts (off-screen at row 200, not hidden)
            # Chart 1: Review vs Testing (using D15 and D16)
            # Write labels in column F (row 201-202)
            metrics_ws.write(200, 5, 'Reviews', workbook.add_format())  # F201
            metrics_ws.write(201, 5, 'Testing', workbook.add_format())  # F202
            # Write values directly from D15 and D16 formulas (values are already percentages like 40.0, 60.0)
            metrics_ws.write_formula(200, 6, '=D15', workbook.add_format({'num_format': '0'}))  # G201
            metrics_ws.write_formula(201, 6, '=D16', workbook.add_format({'num_format': '0'}))  # G202
            
            # Chart 2: Pre-production vs Post-production (using D18 and D19)
            # Write labels in column I (row 201-202)
            metrics_ws.write(200, 8, 'Pre-production', workbook.add_format())  # I201
            metrics_ws.write(201, 8, 'Post-production', workbook.add_format())  # I202
            # Write values directly from D18 and D19 formulas (values are already percentages like 90.0, 10.0)
            metrics_ws.write_formula(200, 9, '=D18', workbook.add_format({'num_format': '0'}))  # J201
            metrics_ws.write_formula(201, 9, '=D19', workbook.add_format({'num_format': '0'}))  # J202
            
            # Create Donut Chart 1: Review vs Testing (positioned at F15)
            # Comparing defects detected in Reviews (D15) vs Testing (D16)
            chart1 = workbook.add_chart({'type': 'doughnut'})
            chart1.add_series({
                'name': 'Defects detected in Reviews vs Testing',
                'categories': ['Metrics Board', 200, 5, 201, 5],  # F201:F202 (row 200-201, col 5)
                'values': ['Metrics Board', 200, 6, 201, 6],      # G201:G202 (row 200-201, col 6)
                'data_labels': {
                    'percentage': True,
                    'category': False,
                    'value': False
                },
                'points': [
                    {'fill': {'color': '#6b8e23'}},  # Blue for Reviews
                    {'fill': {'color': '#c93435'}}   # Red/Orange for Testing
                ]
            })
            chart1.set_title({'name': 'Defect detected in Review vs testing phase', 'name_font': {'size': 11, 'bold': True}})
            chart1.set_legend({
                'position': 'bottom',
                'font': {'size': 9},
                'delete_series': False
            })
            chart1.set_size({'width': 320, 'height': 240})
            chart1.set_style(10)  # Use a predefined style
            metrics_ws.insert_chart('C15', chart1)
            
            # Create Donut Chart 2: Pre-production vs Post-production (positioned at G15)
            # Comparing defects detected in pre-production (D18) vs post-production (D19)
            chart2 = workbook.add_chart({'type': 'doughnut'})
            chart2.add_series({
                'name': 'Defects detected in Pre vs post production testing',
                'categories': ['Metrics Board', 200, 8, 201, 8],  # I201:I202 (row 200-201, col 8)
                'values': ['Metrics Board', 200, 9, 201, 9],      # J201:J202 (row 200-201, col 9)
                'data_labels': {
                    'percentage': True,
                    'category': False,
                    'value': False
                },
                'points': [
                    {'fill': {'color': '#6b8e23'}},  # Blue for Pre-production
                    {'fill': {'color': '#c93435'}}   # Red/Orange for Post-production
                ]
            })
            chart2.set_title({'name': 'Defect detected in Pre vs post production testing', 'name_font': {'size': 11, 'bold': True}})
            chart2.set_legend({
                'position': 'bottom',
                'font': {'size': 9},
                'delete_series': False
            })
            chart2.set_size({'width': 320, 'height': 240})
            chart2.set_style(10)  # Use a predefined style
            metrics_ws.insert_chart('G15', chart2)
            
            # Add Phase Activities table starting at L5
            # Define formats for the phase activities table to match existing metrics table
            table_header_format = workbook.add_format({
                'bold': True,
                'border': 1,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#D9D9D9',  # Grey background to match "Phase detected" and "Phase injected"
                'text_wrap': True
            })
            
            phase_cell_format = workbook.add_format({
                'bold': True,
                'border': 1,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': 'white',  # White background
                'text_wrap': True
            })
            
            activity_format = workbook.add_format({
                'border': 1,
                'align': 'left',
                'valign': 'vcenter',
                'text_wrap': False  # Remove text wrapping
            })
            
            # Phase activities data
            phase_activities = {
                'Requirements': 'BRD Reviews, BRD Updates, Requirement Analysis',
                'Design': 'Source table structure designs, Forms',
                'Implementation': 'Code, Data load, API, Portal, CMS, Integration',
                'Unit testing': 'Backend Validations',
                'QA': 'Test plan, Test data, Test cases.',
                'UAT': 'Test plan, Test data, Test cases, AM test scenarios',
                'Production': 'Data validations, High level functionality validations'
            }
            
            # Write table headers
            metrics_ws.write('L5', 'Phases', table_header_format)
            metrics_ws.merge_range('M5:O5', 'Activities', table_header_format)
            
            # Write phase activities data starting at L6 with merged cells
            row = 6
            for phase, activities in phase_activities.items():
                metrics_ws.write(f'L{row}', phase, phase_cell_format)
                metrics_ws.merge_range(f'M{row}:O{row}', activities, activity_format)
                row += 1
            
            # Auto-adjust column widths - reduced for better page fit
            metrics_ws.set_column('A:A', 5)
            metrics_ws.set_column('B:B', 22)  # Percentage labels column (reduced)
            metrics_ws.set_column('C:C', 14)  # Requirements (reduced)
            metrics_ws.set_column('D:D', 12)  # Design (reduced)
            metrics_ws.set_column('E:E', 14)  # Implementation (reduced)
            metrics_ws.set_column('F:F', 12)  # Unit testing (reduced)
            metrics_ws.set_column('G:G', 12)  # QA (reduced)
            metrics_ws.set_column('H:H', 12)  # UAT (reduced)
            metrics_ws.set_column('I:I', 12)  # Production (reduced)
            metrics_ws.set_column('J:J', 13)  # Total injected (reduced)
            metrics_ws.set_column('L:L', 14)  # Phases column for activities table (reduced)
            metrics_ws.set_column('M:M', 14)  # Activities column start (reduced)
            metrics_ws.set_column('N:N', 14)  # Activities column middle (reduced)
            metrics_ws.set_column('O:O', 14)  # Activities column end (reduced)
            
        except Exception as e:
            print(f"Error creating Metrics Board worksheet: {str(e)}")
            return

    def _create_visual_board_worksheet(self, writer, df, project_key, metrics_data=None):
        """Create Visual Board worksheet with charts for defect data visualization"""
        try:
            # Get the workbook object
            workbook = writer.book
            
            # Create the worksheet
            visual_ws = workbook.add_worksheet('Visual Board')
            
            # Set landscape orientation and fit to one page
            visual_ws.set_landscape()
            visual_ws.fit_to_pages(1, 1)  # 1 page wide, 1 page tall
            
            # Hide gridlines for cleaner look
            visual_ws.hide_gridlines(2)  # Hide both screen and printed gridlines
            
            
            # Create format for clean background
            clean_bg_format = workbook.add_format({
                'border': 0,
                'bg_color': 'white'
            })
            
            # Apply white background to entire worksheet (A1:Z100)
            for row in range(1, 101):
                for col in range(0, 26):  # A to Z
                    visual_ws.write(row, col, '', clean_bg_format)
            
            # Use processed metrics data if available, otherwise calculate from DataFrame
            if metrics_data:
                status_counts = metrics_data.get('status_breakdown', {})
                priority_counts = metrics_data.get('priority_distribution', {})
                environment_counts = metrics_data.get('environment_analysis', {})
                assignee_counts = metrics_data.get('assignee_workload', {})
            else:
                # Fallback to calculating from DataFrame
                status_counts = self._calculate_status_counts(df)
                priority_counts = self._calculate_priority_counts(df)
                environment_counts = self._calculate_environment_counts(df)
                assignee_counts = self._calculate_assignee_counts(df)
            
            # Create data tables in Visual Board sheet for charts
            self._create_chart_data_tables_in_visual_board(workbook, visual_ws, status_counts, priority_counts, environment_counts, assignee_counts)
            
            # Create charts as specified
            self._create_status_doughnut_chart(workbook, visual_ws, status_counts)  # A1
            self._create_priority_bar_chart(workbook, visual_ws, priority_counts)   # G1
            self._create_environment_doughnut_chart(workbook, visual_ws, environment_counts)  # A15
            self._create_assignee_bar_chart(workbook, visual_ws, assignee_counts)   # H15
            
            # Create Time Metrics card at R3
            self._create_time_metrics_card(workbook, visual_ws, df)
            
            # Create Overall Statistics chart
            self._create_overall_statistics_chart(workbook, visual_ws, df)
            
        except Exception as e:
            print(f"Error creating Visual Board worksheet: {str(e)}")
            return

    def _calculate_status_counts(self, df):
        """Calculate status counts from DataFrame with consistent ordering"""
        status_counts = {}
        if 'Status' in df.columns:
            for status in df['Status']:
                status_str = str(status).strip()
                if status_str and status_str.lower() not in ['nan', 'none', '']:
                    status_counts[status_str] = status_counts.get(status_str, 0) + 1
                else:
                    # Count missing/empty status as "Not Specified"
                    status_counts['Not Specified'] = status_counts.get('Not Specified', 0) + 1
        
        # Define preferred order for status display (matching Metrics Board style)
        preferred_order = ['TO DO', 'DEV IN PROGRESS', 'IN REVIEW', 'READY FOR QA', 'QA IN PROGRESS', 'UAT', 'DONE']
        
        # Create ordered dictionary
        ordered_counts = {}
        for status in preferred_order:
            if status in status_counts:
                ordered_counts[status] = status_counts[status]
        
        # Add any remaining statuses not in preferred order
        for status, count in status_counts.items():
            if status not in ordered_counts:
                ordered_counts[status] = count
                
        return ordered_counts

    def _calculate_priority_counts(self, df):
        """Calculate priority counts from DataFrame with consistent ordering"""
        priority_counts = {}
        if 'Priority' in df.columns:
            for priority in df['Priority']:
                priority_str = str(priority).strip()
                if priority_str and priority_str.lower() not in ['nan', 'none', '']:
                    priority_counts[priority_str] = priority_counts.get(priority_str, 0) + 1
                else:
                    # Count missing/empty priority as "Not Specified"
                    priority_counts['Not Specified'] = priority_counts.get('Not Specified', 0) + 1
        
        # Define preferred order for priority display (matching Metrics Board style)
        preferred_order = ['Highest', 'High', 'Medium', 'Low', 'Lowest']
        
        # Create ordered dictionary
        ordered_counts = {}
        for priority in preferred_order:
            if priority in priority_counts:
                ordered_counts[priority] = priority_counts[priority]
        
        # Add any remaining priorities not in preferred order
        for priority, count in priority_counts.items():
            if priority not in ordered_counts:
                ordered_counts[priority] = count
                
        return ordered_counts

    def _calculate_environment_counts(self, df):
        """Calculate environment counts from DataFrame with consistent ordering"""
        environment_counts = {}
        if 'Environment' in df.columns:
            for environment in df['Environment']:
                env_str = str(environment).strip()
                if env_str and env_str.lower() not in ['nan', 'none', '']:
                    environment_counts[env_str] = environment_counts.get(env_str, 0) + 1
                else:
                    # Count missing/empty environment as "Not Specified"
                    environment_counts['Not Specified'] = environment_counts.get('Not Specified', 0) + 1
        
        # Define preferred order for environment display (matching Metrics Board style)
        preferred_order = ['PROD', 'UAT', 'QA', 'DEV', 'Not Specified']
        
        # Create ordered dictionary
        ordered_counts = {}
        for environment in preferred_order:
            if environment in environment_counts:
                ordered_counts[environment] = environment_counts[environment]
        
        # Add any remaining environments not in preferred order
        for environment, count in environment_counts.items():
            if environment not in ordered_counts:
                ordered_counts[environment] = count
                
        return ordered_counts

    def _calculate_assignee_counts(self, df):
        """Calculate assignee counts from DataFrame with consistent ordering"""
        assignee_counts = {}
        if 'Assignee' in df.columns:
            for assignee in df['Assignee']:
                assignee_str = str(assignee).strip()
                if assignee_str and assignee_str.lower() not in ['nan', 'none', '']:
                    assignee_counts[assignee_str] = assignee_counts.get(assignee_str, 0) + 1
                else:
                    # Count missing/empty assignee as "Unassigned"
                    assignee_counts['Unassigned'] = assignee_counts.get('Unassigned', 0) + 1
        
        # Sort assignees alphabetically for consistent display (matching Metrics Board style)
        ordered_counts = {}
        sorted_assignees = sorted(assignee_counts.keys())
        
        # Put "Unassigned" at the end if it exists
        if 'Unassigned' in sorted_assignees:
            sorted_assignees.remove('Unassigned')
            sorted_assignees.append('Unassigned')
        
        for assignee in sorted_assignees:
            ordered_counts[assignee] = assignee_counts[assignee]
                
        return ordered_counts

    def _create_chart_data_tables_in_visual_board(self, workbook, worksheet, status_counts, priority_counts, environment_counts, assignee_counts):
        """Create data tables for charts in Visual Board sheet (positioned off-screen)"""
        
        # Status data table (off-screen at A100:B120)
        row = 100
        worksheet.write(f'A{row}', 'Status', workbook.add_format({'bold': True}))
        worksheet.write(f'B{row}', 'Count', workbook.add_format({'bold': True}))
        row += 1
        # Sort status counts to ensure consistent ordering
        sorted_status_counts = sorted(status_counts.items())
        for i, (status, count) in enumerate(sorted_status_counts):
            worksheet.write(f'A{row}', status)
            worksheet.write(f'B{row}', count)
            row += 1
        
        # Priority data table (off-screen at D100:E120)
        row = 100
        worksheet.write(f'D{row}', 'Priority', workbook.add_format({'bold': True}))
        worksheet.write(f'E{row}', 'Count', workbook.add_format({'bold': True}))
        row += 1
        for priority, count in priority_counts.items():
            worksheet.write(f'D{row}', priority)
            worksheet.write(f'E{row}', count)
            row += 1
        
        # Environment data table (off-screen at G100:H120)
        row = 100
        worksheet.write(f'G{row}', 'Environment', workbook.add_format({'bold': True}))
        worksheet.write(f'H{row}', 'Count', workbook.add_format({'bold': True}))
        row += 1
        for environment, count in environment_counts.items():
            worksheet.write(f'G{row}', environment)
            worksheet.write(f'H{row}', count)
            row += 1
        
        # Assignee data table (off-screen at J100:K120)
        row = 100
        worksheet.write(f'J{row}', 'Assignee', workbook.add_format({'bold': True}))
        worksheet.write(f'K{row}', 'Count', workbook.add_format({'bold': True}))
        row += 1
        for assignee, count in assignee_counts.items():
            worksheet.write(f'J{row}', assignee)
            worksheet.write(f'K{row}', count)
            row += 1

    def _create_chart_data_tables(self, workbook, worksheet, status_counts, priority_counts, environment_counts, assignee_counts):
        """Create data tables for charts (positioned off-screen to avoid chart overlap)"""
        
        # Status data table (off-screen at A100:B120)
        row = 100
        worksheet.write(f'A{row}', 'Status', workbook.add_format({'bold': True}))
        worksheet.write(f'B{row}', 'Count', workbook.add_format({'bold': True}))
        row += 1
        # Sort status counts to ensure consistent ordering
        sorted_status_counts = sorted(status_counts.items())
        for i, (status, count) in enumerate(sorted_status_counts):
            worksheet.write(f'A{row}', status)
            worksheet.write(f'B{row}', count)
            row += 1
        
        # Priority data table (off-screen at D100:E120)
        row = 100
        worksheet.write(f'D{row}', 'Priority', workbook.add_format({'bold': True}))
        worksheet.write(f'E{row}', 'Count', workbook.add_format({'bold': True}))
        row += 1
        for priority, count in priority_counts.items():
            worksheet.write(f'D{row}', priority)
            worksheet.write(f'E{row}', count)
            row += 1
        
        # Environment data table (off-screen at G100:H120)
        row = 100
        worksheet.write(f'G{row}', 'Environment', workbook.add_format({'bold': True}))
        worksheet.write(f'H{row}', 'Count', workbook.add_format({'bold': True}))
        row += 1
        for environment, count in environment_counts.items():
            worksheet.write(f'G{row}', environment)
            worksheet.write(f'H{row}', count)
            row += 1
        
        # Assignee data table (off-screen at J100:K120)
        row = 100
        worksheet.write(f'J{row}', 'Assignee', workbook.add_format({'bold': True}))
        worksheet.write(f'K{row}', 'Count', workbook.add_format({'bold': True}))
        row += 1
        for assignee, count in assignee_counts.items():
            worksheet.write(f'J{row}', assignee)
            worksheet.write(f'K{row}', count)
            row += 1

    def _create_status_doughnut_chart(self, workbook, worksheet, status_counts):
        """Create pie chart for status distribution - A1"""
        if not status_counts:
            return
            
        chart = workbook.add_chart({'type': 'pie'})
        
        # Calculate data range (Status data table is off-screen at A100:B120)
        data_start_row = 100
        data_end_row = data_start_row + len(status_counts) - 1
        
        # Add data series
        chart.add_series({
            'name': 'Defect Status',
            'categories': ['Visual Board', data_start_row, 0, data_end_row, 0],  # A column
            'values': ['Visual Board', data_start_row, 1, data_end_row, 1],      # B column
            'data_labels': {'value': True},
        })
        
        # Configure chart
        chart.set_title({'name': 'Phase', 'name_font': {'size': 12}})
        chart.set_style(10)
        chart.set_legend({'position': 'right'})
        chart.set_size({'width': 480, 'height': 360})
        
        # Insert chart at A1 with specified size
        worksheet.insert_chart('A1', chart, {'x_scale': 0.9, 'y_scale': 0.7})

    def _create_priority_bar_chart(self, workbook, worksheet, priority_counts):
        """Create bar chart for priority distribution - G1"""
        if not priority_counts:
            return
            
        chart = workbook.add_chart({'type': 'column'})
        
        # Calculate data range (Priority data table is off-screen at D100:E120)
        data_start_row = 100
        data_end_row = data_start_row + len(priority_counts) - 1
        
        # Add data series
        chart.add_series({
            'name': 'Priority Distribution',
            'categories': ['Visual Board', data_start_row, 3, data_end_row, 3],  # D column
            'values': ['Visual Board', data_start_row, 4, data_end_row, 4],      # E column
            'data_labels': {'value': True},
        })
        
        # Configure chart
        chart.set_title({'name': 'Priority', 'name_font': {'size': 12}})
        chart.set_x_axis({'name': 'Priority Level'})
        chart.set_y_axis({'name': 'Count'})
        chart.set_style(10)
        chart.set_legend({'position': 'right'})
        chart.set_size({'width': 480, 'height': 360})
        
        # Insert chart at G1 with specified size
        worksheet.insert_chart('H1', chart, {'x_scale': 0.9, 'y_scale': 0.7})

    def _create_environment_doughnut_chart(self, workbook, worksheet, environment_counts):
        """Create doughnut chart for environment analysis - A15"""
        if not environment_counts:
            return
            
        chart = workbook.add_chart({'type': 'doughnut'})
        
        # Calculate data range (Environment data table is off-screen at G100:H120)
        data_start_row = 100
        data_end_row = data_start_row + len(environment_counts) - 1
        
        # Add data series
        chart.add_series({
            'name': 'Environment Analysis',
            'categories': ['Visual Board', data_start_row, 6, data_end_row, 6],  # G column
            'values': ['Visual Board', data_start_row, 7, data_end_row, 7],      # H column
            'data_labels': {'value': True},
        })
        
        # Configure chart
        chart.set_title({'name': 'Environment', 'name_font': {'size': 12}})
        chart.set_style(10)
        chart.set_legend({'position': 'right'})
        chart.set_size({'width': 480, 'height': 360})
        
        # Insert chart at A15 with specified size
        worksheet.insert_chart('A15', chart, {'x_scale': 0.9, 'y_scale': 0.7})

    def _create_assignee_bar_chart(self, workbook, worksheet, assignee_counts):
        """Create pie chart for assignee workload - H15"""
        if not assignee_counts:
            return
            
        chart = workbook.add_chart({'type': 'pie'})
        
        # Calculate data range (Assignee data table is off-screen at J100:K120)
        data_start_row = 100
        data_end_row = data_start_row + len(assignee_counts) - 1
        
        
        # Add data series
        chart.add_series({
            'name': 'Assignee Workload',
            'categories': ['Visual Board', data_start_row, 9, data_end_row, 9],  # J column
            'values': ['Visual Board', data_start_row, 10, data_end_row, 10],    # K column
            'data_labels': {'value': True},
        })
        
        # Configure chart
        chart.set_title({'name': 'Assignee', 'name_font': {'size': 12}})
        chart.set_style(10)
        chart.set_legend({'position': 'right'})
        chart.set_size({'width': 480, 'height': 360})
        
        # Insert chart at H15 with specified size
        worksheet.insert_chart('H15', chart, {'x_scale': 0.9, 'y_scale': 0.7})

    def _create_time_metrics_card(self, workbook, worksheet, df):
        """Create Time Metrics chart at R3"""
        try:
            # Calculate resolution times
            resolution_times = self._calculate_resolution_times(df)
            
            # Calculate metrics
            if resolution_times:
                avg_resolution = sum(resolution_times) / len(resolution_times)
                # Round all values to 1 decimal place
                avg_resolution = round(avg_resolution, 1)
                min_resolution = round(min(resolution_times), 1)
                max_resolution = round(max(resolution_times), 1)
                resolved_count = len(resolution_times)
            else:
                avg_resolution = 0
                min_resolution = 0
                max_resolution = 0
                resolved_count = 0
            
            # Create data table for the chart (off-screen)
            data_start_row = 120  # Use row 120+ for chart data
            data_start_col = 17   # Column R
            
            # Write chart data
            chart_data = [
                ('Minimum', min_resolution),
                ('Maximum', max_resolution),
                ('Average', avg_resolution)
            ]
            
            row = data_start_row
            for metric_name, value in chart_data:
                worksheet.write(row, data_start_col, metric_name)      # R column - categories
                worksheet.write(row, data_start_col + 1, value)       # S column - values
                row += 1
            
            # Create bar chart
            chart = workbook.add_chart({'type': 'column'})
            
            # Add data series with formatted data labels (1 decimal place)
            chart.add_series({
                'name': 'Time Metrics',
                'categories': ['Visual Board', data_start_row, data_start_col, data_start_row + len(chart_data) - 1, data_start_col],
                'values': ['Visual Board', data_start_row, data_start_col + 1, data_start_row + len(chart_data) - 1, data_start_col + 1],
                'data_labels': {
                    'value': True,
                    'num_format': '0.0',  # Format to 1 decimal place
                },
            })
            
            # Configure chart
            chart.set_title({'name': 'Time', 'name_font': {'size': 12}})
            chart.set_x_axis({'name': 'Resolution time'})
            chart.set_y_axis({'name': 'Hours'})
            chart.set_style(10)
            chart.set_legend({'none': True})  # No legend needed for single series
            chart.set_size({'width': 480, 'height': 360})
            
            # Insert chart at R3
            worksheet.insert_chart('O1', chart, {'x_scale': 0.9, 'y_scale': 0.7})
            
        except Exception as e:
            print(f"Error creating Time Metrics chart: {str(e)}")
            return

    def _create_overall_statistics_chart(self, workbook, worksheet, df):
        """Create Overall Statistics horizontal bar chart"""
        try:
            # Calculate overall statistics
            total_defects = len(df)
            open_count, closed_count = self._calculate_open_closed_counts(df)
            
            # Create data table for the chart (off-screen)
            data_start_row = 130  # Use row 130+ for chart data
            data_start_col = 20   # Column U
            
            # Write chart data
            chart_data = [
                ('Closed', closed_count),
                ('Open', open_count),                
                ('Total', total_defects)
            ]
            
            row = data_start_row
            for metric_name, value in chart_data:
                worksheet.write(row, data_start_col, metric_name)      # U column - categories
                worksheet.write(row, data_start_col + 1, value)       # V column - values
                row += 1
            
            # Create horizontal bar chart
            chart = workbook.add_chart({'type': 'bar'})
            
            # Add data series
            chart.add_series({
                'name': 'Overall Statistics',
                'categories': ['Visual Board', data_start_row, data_start_col, data_start_row + len(chart_data) - 1, data_start_col],
                'values': ['Visual Board', data_start_row, data_start_col + 1, data_start_row + len(chart_data) - 1, data_start_col + 1],
                'data_labels': {
                    'value': True,
                    'num_format': '0.0',  # Format to 1 decimal place
                },
            })
            
            # Configure chart
            chart.set_title({'name': 'Status', 'name_font': {'size': 12}})
            chart.set_x_axis({'name': 'Count'})
            chart.set_y_axis({'name': 'Defect Types'})
            chart.set_style(10)
            chart.set_legend({'none': True})  # No legend needed for single series
            chart.set_size({'width': 480, 'height': 360})
            
            # Insert chart at R15 (below the Time Metrics chart)
            worksheet.insert_chart('O15', chart, {'x_scale': 0.9, 'y_scale': 0.7})
            
        except Exception as e:
            print(f"Error creating Defect Statistics chart: {str(e)}")
            return

    def _create_powerbi_connection_info(self, excel_file_path, project_key):
        """Create Power BI connection information and configuration files"""
        try:
            from datetime import datetime
            import json
            import os
            
            # Create Power BI connection configuration
            connection_config = {
                "project": project_key,
                "data_source": {
                    "type": "Excel",
                    "file_path": excel_file_path,
                    "last_updated": datetime.now().isoformat()
                },
                "sheets": {
                    "jira_report": {
                        "name": "Jira Report",
                        "description": "Main defect data with all fields",
                        "primary_key": "Key"
                    },
                    "qa_report": {
                        "name": "QA Report", 
                        "description": "QA summary and defect data",
                        "primary_key": "Key"
                    },
                    "metrics_board": {
                        "name": "Metrics Board",
                        "description": "Defect leakage matrix and metrics",
                        "primary_key": "Phase"
                    },
                    "visual_board": {
                        "name": "Visual Board",
                        "description": "Charts and visualizations",
                        "primary_key": "Chart_Type"
                    }
                },
                "recommended_visualizations": [
                    {
                        "type": "pie_chart",
                        "title": "Defect Status Distribution",
                        "data_source": "jira_report",
                        "x_axis": "Status",
                        "y_axis": "Count"
                    },
                    {
                        "type": "bar_chart",
                        "title": "Priority Distribution", 
                        "data_source": "jira_report",
                        "x_axis": "Priority",
                        "y_axis": "Count"
                    },
                    {
                        "type": "doughnut_chart",
                        "title": "Environment Analysis",
                        "data_source": "jira_report", 
                        "x_axis": "Environment",
                        "y_axis": "Count"
                    },
                    {
                        "type": "column_chart",
                        "title": "Assignee Workload",
                        "data_source": "jira_report",
                        "x_axis": "Assignee", 
                        "y_axis": "Count"
                    },
                    {
                        "type": "matrix",
                        "title": "Defect Leakage Matrix",
                        "data_source": "metrics_board",
                        "rows": "Phase Injected",
                        "columns": "Phase Detected",
                        "values": "Count"
                    }
                ],
                "refresh_settings": {
                    "auto_refresh": True,
                    "refresh_interval": "1 hour",
                    "last_refresh": datetime.now().isoformat()
                }
            }
            
            # Save connection configuration
            config_file = excel_file_path.replace('.xlsx', '_PowerBI_Config.json')
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(connection_config, f, indent=2, ensure_ascii=False)
            
            print(f"✅ Power BI configuration saved: {config_file}")
            return config_file
            
        except Exception as e:
            print(f"❌ Error creating Power BI connection info: {str(e)}")
            return None

    def _create_powerbi_import_script(self, excel_file_path, project_key):
        """Create Power BI import script for automated report generation"""
        try:
            import os
            
            # Create Power BI import script
            import_script = f'''# Power BI Import Script for {project_key}
# Generated on {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

def create_powerbi_report():
    """Create comprehensive Power BI-style report from Excel data"""
    
    # File paths
    excel_file = r"{excel_file_path}"
    project_name = "{project_key}"
    
    
    try:
        # Load data from Excel sheets with proper column handling
        
        # Read Jira Report with proper column names
        jira_raw = pd.read_excel(excel_file, sheet_name='Jira Report', header=None)
        # Find the header row (contains 'Key', 'Status', etc.)
        header_row = None
        for i, row in jira_raw.iterrows():
            if 'Key' in str(row.values) and 'Status' in str(row.values):
                header_row = i
                break
        
        if header_row is not None:
            jira_data = pd.read_excel(excel_file, sheet_name='Jira Report', header=header_row)
            # Remove any rows that are all NaN
            jira_data = jira_data.dropna(how='all')
        else:
            print("❌ Could not find proper header row in Jira Report")
            return False
        
        
        # Create comprehensive report
        fig = plt.figure(figsize=(20, 16))
        fig.suptitle(f'Power BI Report - {{project_name}}', fontsize=20, fontweight='bold')
        
        # 1. Defect Status Distribution (Pie Chart)
        plt.subplot(3, 4, 1)
        status_counts = jira_data['Status'].value_counts()
        colors = plt.cm.Set3(range(len(status_counts)))
        plt.pie(status_counts.values, labels=status_counts.index, autopct='%1.1f%%', 
                colors=colors, startangle=90)
        plt.title('Defect Status Distribution', fontweight='bold', fontsize=12)
        
        # 2. Priority Distribution (Treemap Chart)
        plt.subplot(3, 4, 2)
        priority_counts = jira_data['Priority'].value_counts()
        
        # Create Treemap Chart for Defect by Priority
        from matplotlib.patches import Rectangle
        ax_treemap = plt.gca()
        ax_treemap.clear()
        ax_treemap.set_xlim(0, 1)
        ax_treemap.set_ylim(0, 1)
        ax_treemap.axis('off')
        
        # Sort priorities by count (descending)
        priorities_sorted = priority_counts.sort_values(ascending=False)
        priorities = priorities_sorted.index.tolist()
        counts = priorities_sorted.values.tolist()
        total_count = sum(counts)
        
        # Define colors for priorities
        priority_colors = {{}}
        for priority in priorities:
            priority_lower = str(priority).lower()
            if 'critical' in priority_lower or 'blocker' in priority_lower:
                priority_colors[priority] = '#F44336'  # Red
            elif 'highest' in priority_lower:
                priority_colors[priority] = '#F44336'  # Red for highest
            elif 'high' in priority_lower or 'major' in priority_lower:
                priority_colors[priority] = '#FF9800'  # Orange
            elif 'medium' in priority_lower or 'normal' in priority_lower:
                priority_colors[priority] = '#FFC107'  # Yellow
            else:
                priority_colors[priority] = '#4CAF50'  # Green
        
        # Simple treemap layout
        x, y = 0.05, 0.88
        width, height = 0.95, 0.90
        rectangles = []
        current_x = x
        current_y = y
        row_items = []
        row_total = 0
        remaining_width = width
        remaining_height = height
        
        for i, (priority, count) in enumerate(zip(priorities, counts)):
            percentage = count / total_count if total_count > 0 else 0
            row_items.append((priority, count, percentage))
            row_total += count
            
            if len(row_items) > 1:
                row_height = (row_total / total_count) * remaining_height if total_count > 0 else remaining_height / len(priorities)
                avg_width = remaining_width / len(row_items)
                worst_aspect = max(avg_width / row_height, row_height / avg_width) if row_height > 0 else 1
                
                if i < len(priorities) - 1:
                    next_count = counts[i + 1] if i + 1 < len(counts) else 0
                    test_row_total = row_total + next_count
                    test_row_height = (test_row_total / total_count) * remaining_height if total_count > 0 else remaining_height / (len(priorities) - i)
                    test_avg_width = remaining_width / (len(row_items) + 1)
                    test_worst_aspect = max(test_avg_width / test_row_height, test_row_height / test_avg_width) if test_row_height > 0 else 1
                    
                    if test_worst_aspect > worst_aspect * 1.1:
                        row_height = (row_total / total_count) * remaining_height if total_count > 0 else remaining_height / len(priorities)
                        current_x_row = current_x
                        for j, (p, c, pct) in enumerate(row_items):
                            item_width = (c / row_total) * remaining_width if row_total > 0 else remaining_width / len(row_items)
                            rect_y = current_y - row_height
                            rectangles.append((current_x_row, rect_y, item_width, row_height, p, c, priority_colors[p]))
                            current_x_row += item_width
                        current_y -= row_height
                        remaining_height -= row_height
                        row_items = []
                        row_total = 0
                        continue
            
            if i == len(priorities) - 1:
                row_height = (row_total / total_count) * remaining_height if total_count > 0 else remaining_height
                current_x_row = current_x
                for j, (p, c, pct) in enumerate(row_items):
                    item_width = (c / row_total) * remaining_width if row_total > 0 else remaining_width / len(row_items)
                    rect_y = current_y - row_height
                    rectangles.append((current_x_row, rect_y, item_width, row_height, p, c, priority_colors[p]))
                    current_x_row += item_width
        
        # Draw rectangles
        for rect_x, rect_y, rect_width, rect_height, priority, count, color in rectangles:
            rect = Rectangle((rect_x, rect_y), rect_width, rect_height,
                           facecolor=color, edgecolor='white', linewidth=2)
            ax_treemap.add_patch(rect)
            label_x = rect_x + rect_width / 2
            label_y = rect_y + rect_height / 2 + 0.02
            ax_treemap.text(label_x, label_y, str(priority), 
                          ha='center', va='center', fontsize=8, fontweight='bold', color='white')
            count_y = rect_y + rect_height / 2 - 0.02
            ax_treemap.text(label_x, count_y, str(int(count)), 
                          ha='center', va='center', fontsize=9, fontweight='bold', color='white')
        
        plt.title('Priority Distribution', fontweight='bold', fontsize=12)
        
        # 3. Environment Analysis (Doughnut Chart)
        plt.subplot(3, 4, 3)
        env_counts = jira_data['Environment'].value_counts()
        colors_env = plt.cm.Pastel1(range(len(env_counts)))
        plt.pie(env_counts.values, labels=env_counts.index, autopct='%1.1f%%',
                colors=colors_env, startangle=90, pctdistance=0.85)
        centre_circle = plt.Circle((0,0), 0.70, fc='white')
        plt.gca().add_artist(centre_circle)
        plt.title('Environment Analysis', fontweight='bold', fontsize=12)
        
        # 4. Assignee Workload (Horizontal Bar Chart)
        plt.subplot(3, 4, 4)
        assignee_counts = jira_data['Assignee'].value_counts()
        bars = plt.barh(assignee_counts.index, assignee_counts.values,
                       color=plt.cm.viridis(range(len(assignee_counts))))
        plt.title('Assignee Workload', fontweight='bold', fontsize=12)
        for i, bar in enumerate(bars):
            width = bar.get_width()
            plt.text(width + 0.1, bar.get_y() + bar.get_height()/2,
                    f'{{int(width)}}', ha='left', va='center')
        
        # 5. Defect Trend Over Time (Line Chart)
        plt.subplot(3, 4, 5)
        jira_data['Created Date'] = pd.to_datetime(jira_data['Created Date'])
        daily_counts = jira_data.groupby(jira_data['Created Date'].dt.date).size()
        plt.plot(daily_counts.index, daily_counts.values, marker='o', linewidth=2, markersize=6)
        plt.title('Defect Trend Over Time', fontweight='bold', fontsize=12)
        plt.xticks(rotation=45)
        plt.grid(True, alpha=0.3)
        
        # 6. Resolution Time Analysis (Box Plot)
        plt.subplot(3, 4, 6)
        # Calculate resolution time in days
        jira_data['Resolution Date'] = pd.to_datetime(jira_data['Resolution Date'])
        jira_data['Resolution Time (Days)'] = (jira_data['Resolution Date'] - jira_data['Created Date']).dt.days
        resolution_times = jira_data['Resolution Time (Days)'].dropna()
        if len(resolution_times) > 0:
            plt.boxplot(resolution_times, patch_artist=True, 
                       boxprops=dict(facecolor='lightblue', alpha=0.7))
            plt.title('Resolution Time Distribution', fontweight='bold', fontsize=12)
            plt.ylabel('Days')
        else:
            plt.text(0.5, 0.5, 'No Resolution Data', ha='center', va='center', transform=plt.gca().transAxes)
            plt.title('Resolution Time Distribution', fontweight='bold', fontsize=12)
        
        # 7. Phase Detection Analysis (Stacked Bar Chart)
        plt.subplot(3, 4, 7)
        phase_detected = jira_data['Phase Detected'].value_counts()
        bars = plt.bar(phase_detected.index, phase_detected.values,
                      color=plt.cm.Set2(range(len(phase_detected))))
        plt.title('Phase Detection Analysis', fontweight='bold', fontsize=12)
        plt.xticks(rotation=45)
        for bar in bars:
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                    f'{{int(height)}}', ha='center', va='bottom')
        
        # 8. Phase Injection Analysis (Stacked Bar Chart)
        plt.subplot(3, 4, 8)
        phase_injected = jira_data['Phase Injected'].value_counts()
        bars = plt.bar(phase_injected.index, phase_injected.values,
                      color=plt.cm.Set1(range(len(phase_injected))))
        plt.title('Phase Injection Analysis', fontweight='bold', fontsize=12)
        plt.xticks(rotation=45)
        for bar in bars:
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                    f'{{int(height)}}', ha='center', va='bottom')
        
        # 9. Status vs Priority Heatmap
        plt.subplot(3, 4, 9)
        status_priority = pd.crosstab(jira_data['Status'], jira_data['Priority'])
        sns.heatmap(status_priority, annot=True, fmt='d', cmap='YlOrRd', cbar_kws={{'label': 'Count'}})
        plt.title('Status vs Priority Heatmap', fontweight='bold', fontsize=12)
        plt.xticks(rotation=45)
        plt.yticks(rotation=0)
        
        # 10. Environment vs Status Analysis
        plt.subplot(3, 4, 10)
        env_status = pd.crosstab(jira_data['Environment'], jira_data['Status'])
        env_status.plot(kind='bar', stacked=True, ax=plt.gca(), colormap='tab10')
        plt.title('Environment vs Status', fontweight='bold', fontsize=12)
        plt.xticks(rotation=45)
        plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
        
        # 11. Summary Statistics Table
        plt.subplot(3, 4, 11)
        plt.axis('off')
        summary_stats = {{
            'Total Defects': len(jira_data),
            'Open Defects': len(jira_data[jira_data['Status'].isin(['TO DO', 'IN PROGRESS', 'QA', 'UAT'])]),
            'Closed Defects': len(jira_data[jira_data['Status'].isin(['DONE', 'CLOSED'])]),
            'Avg Resolution Time': f"{{resolution_times.mean():.1f}} days" if len(resolution_times) > 0 else "N/A",
            'High Priority': len(jira_data[jira_data['Priority'] == 'High']),
            'Medium Priority': len(jira_data[jira_data['Priority'] == 'Medium']),
            'Low Priority': len(jira_data[jira_data['Priority'] == 'Low'])
        }}
        
        table_data = [[k, v] for k, v in summary_stats.items()]
        table = plt.table(cellText=table_data, colLabels=['Metric', 'Value'],
                         cellLoc='left', loc='center', bbox=[0, 0, 1, 1])
        table.auto_set_font_size(False)
        table.set_fontsize(10)
        table.scale(1, 2)
        plt.title('Summary Statistics', fontweight='bold', fontsize=12, pad=20)
        
        # 12. Defect Leakage Matrix (if available)
        plt.subplot(3, 4, 12)
        try:
            # Try to create a simple leakage visualization
            leakage_data = jira_data.groupby(['Phase Injected', 'Phase Detected']).size().unstack(fill_value=0)
            if not leakage_data.empty:
                sns.heatmap(leakage_data, annot=True, fmt='d', cmap='Blues', cbar_kws={{'label': 'Count'}})
                plt.title('Defect Leakage Matrix', fontweight='bold', fontsize=12)
            else:
                plt.text(0.5, 0.5, 'No Leakage Data Available', ha='center', va='center', 
                        transform=plt.gca().transAxes, fontsize=12)
                plt.title('Defect Leakage Matrix', fontweight='bold', fontsize=12)
        except:
            plt.text(0.5, 0.5, 'Leakage Matrix Not Available', ha='center', va='center', 
                    transform=plt.gca().transAxes, fontsize=12)
            plt.title('Defect Leakage Matrix', fontweight='bold', fontsize=12)
        
        plt.tight_layout()
        
        # Save the report
        report_file = excel_file_path.replace('.xlsx', '_PowerBI_Report.png')
        plt.savefig(report_file, dpi=300, bbox_inches='tight', facecolor='white')
        print(f"✅ Power BI report saved: {{report_file}}")
        
        # Also save as PDF for better quality
        pdf_file = excel_file_path.replace('.xlsx', '_PowerBI_Report.pdf')
        plt.savefig(pdf_file, bbox_inches='tight', facecolor='white')
        print(f"✅ Power BI report (PDF) saved: {{pdf_file}}")
        
        plt.show()
        
        # Create data summary for Power BI
        self._create_powerbi_data_summary(jira_data, excel_file_path)
        
        print("🎉 Power BI report generation completed successfully!")
        
    except Exception as e:
        print(f"❌ Error generating Power BI report: {{str(e)}}")
        return False
    
    return True

def _create_powerbi_data_summary(data, excel_file_path):
    """Create data summary for Power BI consumption"""
    try:
        summary = {{
            'total_defects': len(data),
            'status_breakdown': data['Status'].value_counts().to_dict(),
            'priority_breakdown': data['Priority'].value_counts().to_dict(),
            'environment_breakdown': data['Environment'].value_counts().to_dict(),
            'assignee_breakdown': data['Assignee'].value_counts().to_dict(),
            'phase_detected_breakdown': data['Phase Detected'].value_counts().to_dict(),
            'phase_injected_breakdown': data['Phase Injected'].value_counts().to_dict(),
            'generated_at': datetime.now().isoformat()
        }}
        
        import json
        summary_file = excel_file_path.replace('.xlsx', '_PowerBI_Summary.json')
        with open(summary_file, 'w', encoding='utf-8') as f:
            json.dump(summary, f, indent=2, ensure_ascii=False)
        
        print(f"✅ Power BI data summary saved: {{summary_file}}")
        
    except Exception as e:
        print(f"❌ Error creating data summary: {{str(e)}}")

if __name__ == "__main__":
    create_powerbi_report()
'''
            
            # Save the import script
            script_file = excel_file_path.replace('.xlsx', '_PowerBI_Import.py')
            with open(script_file, 'w', encoding='utf-8') as f:
                f.write(import_script)
            
            print(f"✅ Power BI import script created: {script_file}")
            return script_file
            
        except Exception as e:
            print(f"❌ Error creating Power BI import script: {str(e)}")
            return None

    def _create_powerbi_desktop_instructions(self, excel_file_path, project_key):
        """Create instructions for Power BI Desktop integration"""
        try:
            instructions = f"""# Power BI Desktop Integration Instructions
# Project: {project_key}
# Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

## Step 1: Open Power BI Desktop
1. Launch Power BI Desktop
2. Click "Get Data" or "Home" > "Get Data"

## Step 2: Connect to Excel File
1. Select "Excel workbook" from the data sources
2. Browse to: {excel_file_path}
3. Click "Open"

## Step 3: Select Sheets to Import
Select the following sheets:
- ✅ Jira Report (Main defect data)
- ✅ QA Report (QA summary and data)
- ✅ Metrics Board (Defect leakage matrix)
- ✅ Visual Board (Charts and visualizations)

## Step 4: Data Transformation (Optional)
1. In Power Query Editor, you can:
   - Clean data types
   - Add calculated columns
   - Filter data as needed
2. Click "Close & Apply" when done

## Step 5: Create Visualizations
### Recommended Visualizations:

#### 1. Defect Status Distribution (Pie Chart)
- Visual: Pie chart
- Legend: Status
- Values: Count of Status

#### 2. Priority Distribution (Bar Chart)
- Visual: Clustered column chart
- Axis: Priority
- Values: Count of Priority

#### 3. Environment Analysis (Doughnut Chart)
- Visual: Doughnut chart
- Legend: Environment
- Values: Count of Environment

#### 4. Assignee Workload (Bar Chart)
- Visual: Clustered column chart
- Axis: Assignee
- Values: Count of Assignee

#### 5. Defect Leakage Matrix (Matrix)
- Visual: Matrix
- Rows: Phase Injected
- Columns: Phase Detected
- Values: Count

#### 6. Resolution Time Analysis (Line Chart)
- Visual: Line chart
- Axis: Created Date
- Values: Count of Key

## Step 6: Create Dashboard
1. Create a new dashboard
2. Pin the visualizations to the dashboard
3. Arrange them in a logical layout
4. Add filters and slicers as needed

## Step 7: Set Up Auto-Refresh (Optional)
1. Go to "Home" > "Refresh"
2. Set up scheduled refresh if using Power BI Service
3. Configure refresh frequency (hourly, daily, etc.)

## Data Model Relationships
- Primary Key: Key (from Jira Report)
- Foreign Keys: Use Key to link between sheets
- Date Fields: Created Date, Resolution Date, Last Status Change

## Tips for Better Visualizations
1. Use consistent color schemes
2. Add data labels where appropriate
3. Use appropriate chart types for data
4. Add titles and descriptions
5. Consider mobile-friendly layouts

## Troubleshooting
- If data doesn't refresh, check file permissions
- Ensure Excel file is not open in another application
- Verify data types are correct in Power Query Editor
- Check for missing or null values

## File Locations
- Excel File: {excel_file_path}
- Configuration: {excel_file_path.replace('.xlsx', '_PowerBI_Config.json')}
- Import Script: {excel_file_path.replace('.xlsx', '_PowerBI_Import.py')}
- Instructions: {excel_file_path.replace('.xlsx', '_PowerBI_Instructions.md')}

---
Generated by Jira Automation Tool
For support, contact your system administrator.
"""
            
            # Save instructions
            instructions_file = excel_file_path.replace('.xlsx', '_PowerBI_Instructions.md')
            with open(instructions_file, 'w', encoding='utf-8') as f:
                f.write(instructions)
            
            print(f"✅ Power BI instructions created: {instructions_file}")
            return instructions_file
            
        except Exception as e:
            print(f"❌ Error creating Power BI instructions: {str(e)}")
            return None

    def _launch_powerbi_automatically(self, excel_file_path, project_key):
        """Automatically launch Power BI Desktop with the Excel data loaded"""
        try:
            import subprocess
            import os
            import json
            from datetime import datetime
            
            print("🚀 Launching Power BI Desktop automatically...")
            
            # Create Power BI template file (.pbit) for automatic loading
            template_data = {
                "version": "1.0",
                "name": f"{project_key}_Defect_Analysis",
                "description": f"Automated defect analysis report for {project_key}",
                "data_sources": [
                    {
                        "name": "JiraDefects",
                        "type": "Excel",
                        "path": excel_file_path,
                        "sheets": ["Jira Report", "QA Report", "Metrics Board", "Visual Board"]
                    }
                ],
                "visualizations": [
                    {
                        "type": "pie_chart",
                        "title": "Defect Status Distribution",
                        "data_source": "Jira Report",
                        "x_axis": "Status",
                        "y_axis": "Count"
                    },
                    {
                        "type": "bar_chart", 
                        "title": "Priority Distribution",
                        "data_source": "Jira Report",
                        "x_axis": "Priority",
                        "y_axis": "Count"
                    },
                    {
                        "type": "doughnut_chart",
                        "title": "Environment Analysis", 
                        "data_source": "Jira Report",
                        "x_axis": "Environment",
                        "y_axis": "Count"
                    },
                    {
                        "type": "column_chart",
                        "title": "Assignee Workload",
                        "data_source": "Jira Report", 
                        "x_axis": "Assignee",
                        "y_axis": "Count"
                    }
                ],
                "created_at": datetime.now().isoformat()
            }
            
            # Save template configuration
            template_file = excel_file_path.replace('.xlsx', '_PowerBI_Template.json')
            with open(template_file, 'w', encoding='utf-8') as f:
                json.dump(template_data, f, indent=2, ensure_ascii=False)
            
            # Create Power BI Desktop launch script
            launch_script = f'''# Power BI Desktop Auto-Launch Script
import subprocess
import os
import time
import json

def launch_powerbi_with_data():
    """Launch Power BI Desktop and load Excel data automatically"""
    
    excel_file = r"{excel_file_path}"
    project_name = "{project_key}"
    
    print("🚀 Starting Power BI Desktop...")
    
    try:
        # Try to launch Power BI Desktop
        powerbi_paths = [
            r"C:\\Program Files\\Microsoft Power BI Desktop\\bin\\PBIDesktop.exe",
            r"C:\\Program Files (x86)\\Microsoft Power BI Desktop\\bin\\PBIDesktop.exe",
            r"C:\\Users\\{os.getenv('USERNAME')}\\AppData\\Local\\Microsoft\\WindowsApps\\PBIDesktop.exe"
        ]
        
        powerbi_exe = None
        for path in powerbi_paths:
            if os.path.exists(path):
                powerbi_exe = path
                break
        
        if powerbi_exe:
            print(f"✅ Found Power BI Desktop at: {{powerbi_exe}}")
            
            # Launch Power BI Desktop
            subprocess.Popen([powerbi_exe])
            print("🚀 Power BI Desktop launched successfully!")
            
            # Wait a moment for Power BI to start
            time.sleep(3)
            
            # Create a simple instruction file for the user
            instructions = f"""
# Power BI Desktop Auto-Launch Instructions

Power BI Desktop has been launched automatically!

## Next Steps:
1. In Power BI Desktop, click "Get Data"
2. Select "Excel workbook"
3. Browse to: {excel_file_path}
4. Click "Open"
5. Select all sheets: Jira Report, QA Report, Metrics Board, Visual Board
6. Click "Load"

## Recommended Visualizations:
- Defect Status Distribution (Pie Chart)
- Priority Distribution (Bar Chart) 
- Environment Analysis (Doughnut Chart)
- Assignee Workload (Column Chart)
- Defect Leakage Matrix (Matrix)

## Data Source: {excel_file_path}
## Project: {project_key}
## Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
"""
            
            instruction_file = excel_file_path.replace('.xlsx', '_PowerBI_AutoLaunch_Instructions.txt')
            with open(instruction_file, 'w', encoding='utf-8') as f:
                f.write(instructions)
            
            print(f"📋 Instructions saved to: {{instruction_file}}")
            print("✅ Power BI Desktop launched! Please follow the instructions to load your data.")
            
        else:
            print("❌ Power BI Desktop not found. Please install Power BI Desktop first.")
            print("📥 Download from: https://powerbi.microsoft.com/desktop/")
            return False
            
    except Exception as e:
        print(f"❌ Error launching Power BI Desktop: {{str(e)}}")
        return False
    
    return True

if __name__ == "__main__":
    launch_powerbi_with_data()
'''
            
            # Save and execute the launch script
            script_file = excel_file_path.replace('.xlsx', '_PowerBI_AutoLaunch.py')
            with open(script_file, 'w', encoding='utf-8') as f:
                f.write(launch_script)
            
            # Execute the launch script
            result = subprocess.run(['python', script_file], capture_output=True, text=True)
            print(result.stdout)
            if result.stderr:
                print(f"⚠️ Warning: {{result.stderr}}")
            
            print(f"✅ Power BI auto-launch script created: {script_file}")
            return script_file
            
        except Exception as e:
            print(f"❌ Error creating Power BI auto-launch: {str(e)}")
            return None

    def _generate_powerbi_report_directly(self, excel_file_path, project_key):
        """Generate Power BI-style report using existing charts from Visual Board and Metrics Board"""
        try:
            import subprocess
            import os
            from datetime import datetime
            
            
            # Create Power BI dashboard script that uses existing Excel charts
            report_script = f'''# Power BI Dashboard using Existing Charts for {project_key}
# Generated on {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.patches import Rectangle
from datetime import datetime
import warnings
import os
import subprocess
import psutil
try:
    import squarify
    HAS_SQUARIFY = True
except ImportError:
    HAS_SQUARIFY = False
warnings.filterwarnings('ignore')

def close_existing_files():
    """Close any existing Excel and PDF viewer processes to prevent multiple files from opening"""
    try:
        # Close Excel processes
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info['name'] and 'excel' in proc.info['name'].lower():
                    proc.terminate()
                    print("Closed existing Excel process:", proc.info['name'], "(PID:", proc.info['pid'], ")")
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
        
        # Close PDF viewer processes (common PDF viewers)
        pdf_viewers = ['acrobat', 'adobe', 'foxit', 'sumatra', 'chrome', 'firefox', 'edge']
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info['name']:
                    proc_name = proc.info['name'].lower()
                    if any(viewer in proc_name for viewer in pdf_viewers):
                        # Only close if it's likely a PDF viewer (not the main browser)
                        if 'acrobat' in proc_name or 'adobe' in proc_name or 'foxit' in proc_name or 'sumatra' in proc_name:
                            proc.terminate()
                            print("Closed existing PDF viewer:", proc.info['name'], "(PID:", proc.info['pid'], ")")
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
                
        # Give processes a moment to close
        import time
        time.sleep(1)
        
    except Exception as e:
        print("Warning: Could not close existing processes:", str(e))

def generate_powerbi_dashboard():
    """Generate Power BI dashboard using existing charts from Visual Board and Metrics Board"""
    
    excel_file = r"{excel_file_path}"
    project_name = "{project_key}"
    
    
    try:
        # Load data from all sheets
        
        # Read Jira Report for data
        jira_raw = pd.read_excel(excel_file, sheet_name='Jira Report', header=None, engine='openpyxl')
        header_row = None
        for i, row in jira_raw.iterrows():
            if 'Key' in str(row.values) and 'Status' in str(row.values):
                header_row = i
                break
        
        if header_row is not None:
            jira_data = pd.read_excel(excel_file, sheet_name='Jira Report', header=header_row, engine='openpyxl')
            jira_data = jira_data.dropna(how='all')
        else:
            print("Could not find proper header row in Jira Report")
            return False
        
        # Read Visual Board data (charts data)
        try:
            visual_data = pd.read_excel(excel_file, sheet_name='Visual Board', header=None, engine='openpyxl')
        except:
            print("Could not load Visual Board data")
            visual_data = None
        
        # Read Metrics Board data
        try:
            metrics_data = pd.read_excel(excel_file, sheet_name='Metrics Board', header=None, engine='openpyxl')
        except:
            print("Could not load Metrics Board data")
            metrics_data = None
        
        
        # Create Power BI Dashboard with light and simple colors
        fig = plt.figure(figsize=(20, 12))
        fig.suptitle(f'Power BI Dashboard - {{project_name}}', fontsize=20, fontweight='bold', y=0.95)
        
        # Define light and simple colors
        light_colors = ['#E8F4FD', '#F0F8E8', '#FFF8E1', '#F3E5F5', '#E0F2F1', '#FCE4EC', '#E3F2FD', '#F1F8E9']
        simple_colors = ['#81C784', '#64B5F6', '#FFB74D', '#F06292', '#4DB6AC', '#FF8A65', '#9575CD', '#A5D6A7']
        
        # Chart 1: Phase (Line Graph) - top-left
        plt.subplot(2, 3, 1)
        status_counts = jira_data['Status'].value_counts().sort_index()
        x_pos = range(len(status_counts))
        plt.plot(x_pos, status_counts.values, marker='o', linewidth=2, markersize=8, 
                color=simple_colors[0])
        plt.fill_between(x_pos, status_counts.values, alpha=0.3, color=simple_colors[0])
        plt.xticks(x_pos, status_counts.index, rotation=45, fontsize=9, ha='right')
        plt.ylabel('Count', fontsize=9)
        plt.grid(True, alpha=0.3)
        for i, value in enumerate(status_counts.values):
            plt.text(i, value + 0.1, str(int(value)), ha='center', va='bottom', fontweight='bold', fontsize=9)
        plt.title('Phase', fontweight='bold', fontsize=12, pad=15)
        
        # Chart 2: Priority Distribution - Treemap Chart
        plt.subplot(2, 3, 2)
        priority_counts = jira_data['Priority'].value_counts()
        
        # Create Treemap Chart for Defect by Priority
        ax_treemap = plt.gca()
        ax_treemap.clear()
        ax_treemap.set_xlim(0, 1)
        ax_treemap.set_ylim(0, 1)
        ax_treemap.axis('off')
        
        # Sort priorities by count (descending) for better visualization
        priorities_sorted = priority_counts.sort_values(ascending=False)
        priorities = priorities_sorted.index.tolist()
        counts = priorities_sorted.values.tolist()
        total_count = sum(counts)
        
        # Define colors for priorities (red for critical, orange for high, yellow for medium, green for low)
        priority_colors = {{}}
        for priority in priorities:
            priority_lower = str(priority).lower()
            if 'critical' in priority_lower or 'blocker' in priority_lower:
                priority_colors[priority] = '#F44336'  # Red
            elif 'highest' in priority_lower:
                priority_colors[priority] = '#F44336'  # Red for highest
            elif 'high' in priority_lower or 'major' in priority_lower:
                priority_colors[priority] = '#FF9800'  # Orange
            elif 'medium' in priority_lower or 'normal' in priority_lower:
                priority_colors[priority] = '#FFC107'  # Yellow
            else:
                priority_colors[priority] = '#4CAF50'  # Green
        
        # Simple treemap layout algorithm (squarified)
        # Start from top-left corner
        x, y = 0.05, 0.88
        width, height = 0.95, 0.90
        
        # Calculate rectangle positions for treemap
        rectangles = []
        current_x = x
        current_y = y
        
        # Simple row-based layout
        row_items = []
        row_total = 0
        remaining_width = width
        remaining_height = height
        
        for i, (priority, count) in enumerate(zip(priorities, counts)):
            percentage = count / total_count if total_count > 0 else 0
            
            # Add to current row
            row_items.append((priority, count, percentage))
            row_total += count
            
            # Calculate if we should start a new row (squarified algorithm)
            if len(row_items) > 1:
                # Calculate row height based on total count in row
                row_height = (row_total / total_count) * remaining_height if total_count > 0 else remaining_height / len(priorities)
                
                # Calculate worst aspect ratio if we add this item
                max_count_in_row = max(c for _, c, _ in row_items)
                min_count_in_row = min(c for _, c, _ in row_items)
                
                # Estimate widths
                avg_width = remaining_width / len(row_items)
                worst_aspect = max(avg_width / row_height, row_height / avg_width) if row_height > 0 else 1
                
                # If adding next item would make aspect ratio worse, finalize current row
                if i < len(priorities) - 1:  # Not the last item
                    next_count = counts[i + 1] if i + 1 < len(counts) else 0
                    test_row_total = row_total + next_count
                    test_row_height = (test_row_total / total_count) * remaining_height if total_count > 0 else remaining_height / (len(priorities) - i)
                    test_avg_width = remaining_width / (len(row_items) + 1)
                    test_worst_aspect = max(test_avg_width / test_row_height, test_row_height / test_avg_width) if test_row_height > 0 else 1
                    
                    # If adding next item makes it worse, finalize current row
                    if test_worst_aspect > worst_aspect * 1.1:  # 10% threshold
                        # Draw current row with proportional widths
                        row_height = (row_total / total_count) * remaining_height if total_count > 0 else remaining_height / len(priorities)
                        current_x_row = current_x
                        for j, (p, c, pct) in enumerate(row_items):
                            # Width proportional to count
                            item_width = (c / row_total) * remaining_width if row_total > 0 else remaining_width / len(row_items)
                            rect_y = current_y - row_height
                            rectangles.append((current_x_row, rect_y, item_width, row_height, p, c, priority_colors[p]))
                            current_x_row += item_width
                        
                        # Move to next row
                        current_y -= row_height
                        remaining_height -= row_height
                        row_items = []
                        row_total = 0
                        continue
            
            # If last item, draw the final row
            if i == len(priorities) - 1:
                row_height = (row_total / total_count) * remaining_height if total_count > 0 else remaining_height
                current_x_row = current_x
                for j, (p, c, pct) in enumerate(row_items):
                    # Width proportional to count
                    item_width = (c / row_total) * remaining_width if row_total > 0 else remaining_width / len(row_items)
                    rect_y = current_y - row_height
                    rectangles.append((current_x_row, rect_y, item_width, row_height, p, c, priority_colors[p]))
                    current_x_row += item_width
        
        # Draw rectangles
        for rect_x, rect_y, rect_width, rect_height, priority, count, color in rectangles:
            # Draw rectangle
            rect = Rectangle((rect_x, rect_y), rect_width, rect_height,
                           facecolor=color, edgecolor='white', linewidth=2)
            ax_treemap.add_patch(rect)
            
            # Add label (priority name)
            label_x = rect_x + rect_width / 2
            label_y = rect_y + rect_height / 2 + 0.02
            ax_treemap.text(label_x, label_y, str(priority), 
                          ha='center', va='center', fontsize=9, fontweight='bold', color='white')
            
            # Add count
            count_y = rect_y + rect_height / 2 - 0.02
            ax_treemap.text(label_x, count_y, str(int(count)), 
                          ha='center', va='center', fontsize=10, fontweight='bold', color='white')
        
        plt.title('Priority', fontweight='bold', fontsize=12, pad=15)
 
        
        # Chart 3: Environment Analysis (Doughnut Chart) - top-right, show count
        plt.subplot(2, 3, 3)
        env_counts = jira_data['Environment'].value_counts()
        # Format labels to show count instead of percentage
        def make_autopct(values):
            def my_autopct(pct):
                total = sum(values)
                val = int(round(pct*total/100.0))
                return f'{{val}}'
            return my_autopct
        wedges, texts, autotexts = plt.pie(env_counts.values, labels=env_counts.index,
                                          autopct=make_autopct(env_counts.values),
                                          colors=simple_colors[:len(env_counts)],
                                          startangle=90, pctdistance=0.85, textprops={{'fontsize': 9}})
        centre_circle = plt.Circle((0,0), 0.70, fc='white')
        plt.gca().add_artist(centre_circle)
        plt.title('Environment', fontweight='bold', fontsize=12, pad=15)
        
        # Chart 4: Assignee Workload (3D Pie Chart)
        plt.subplot(2, 3, 4)
        assignee_counts = jira_data['Assignee'].value_counts()
        colors = simple_colors[:len(assignee_counts)]
        # Format labels to show count instead of percentage
        def make_autopct_assignee_preview(values):
            def my_autopct(pct):
                total = sum(values)
                val = int(round(pct*total/100.0))
                return f'{{val}}'
            return my_autopct
        # Create 3D pie chart effect with explode and shadow
        explode = [0.1] * len(assignee_counts)  # Explode for 3D effect
        wedges, texts, autotexts = plt.pie(assignee_counts.values, labels=assignee_counts.index,
                                          autopct=make_autopct_assignee_preview(assignee_counts.values),
                                          colors=colors,
                                          startangle=90, explode=explode,
                                          textprops={{'fontsize': 9}}, shadow=True, 
                                          wedgeprops={{'edgecolor': 'white', 'linewidth': 2}})
        # Enhance 3D effect by making text bold
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontweight('bold')
            autotext.set_fontsize(9)
        plt.title('Assignee', fontweight='bold', fontsize=12, pad=15)
        
        # Chart 5: Time Metrics (from Visual Board - Column Chart)
        plt.subplot(2, 3, 5)
        # Calculate resolution times in hours
        jira_data['Created Date'] = pd.to_datetime(jira_data['Created Date'])
        jira_data['Resolution Date'] = pd.to_datetime(jira_data['Resolution Date'])
        jira_data['Resolution Time (Hours)'] = (jira_data['Resolution Date'] - jira_data['Created Date']).dt.total_seconds() / 3600
        resolution_times = jira_data['Resolution Time (Hours)'].dropna()
        
        if len(resolution_times) > 0:
            time_metrics = ['Minimum', 'Maximum', 'Average']
            min_val = round(resolution_times.min(), 1)  # Round to 1 decimal place
            max_val = round(resolution_times.max(), 1)  # Round to 1 decimal place
            avg_val = round(resolution_times.mean(), 1)  # Round to 1 decimal place
            time_values = [min_val, max_val, avg_val]
            bars = plt.bar(time_metrics, time_values, color=simple_colors[:3])
            plt.title('Time', fontweight='bold', fontsize=12, pad=15)
            plt.ylabel('Hours')
            for bar, value in zip(bars, time_values):
                # Format all values with 1 decimal place (xx.x)
                plt.text(bar.get_x() + bar.get_width()/2., bar.get_height() + 0.1,
                        f'{{value:.1f}}'.format(value=value), ha='center', va='bottom', fontweight='bold', fontsize=9)
        else:
            plt.text(0.5, 0.5, 'No Resolution Data', ha='center', va='center', 
                    transform=plt.gca().transAxes, fontsize=12)
            plt.title('Time', fontweight='bold', fontsize=12, pad=15)
        
        # Chart 6: Overall Statistics (from Visual Board - Horizontal Bar Chart)
        plt.subplot(2, 3, 6)
        total_defects = len(jira_data)
        # Count closed defects (only DONE, CLOSED, COMPLETED, RESOLVED)
        closed_statuses = ['DONE', 'CLOSED', 'COMPLETED', 'RESOLVED']
        closed_defects = len(jira_data[jira_data['Status'].isin(closed_statuses)])
        # All other defects are considered open (ensures Open + Closed = Total)
        open_defects = total_defects - closed_defects
        
        stats_labels = ['Total Defects', 'Open Defects', 'Closed Defects']
        stats_values = [total_defects, open_defects, closed_defects]
        bars = plt.barh(stats_labels, stats_values, color=simple_colors[:3])
        plt.title('Status', fontweight='bold', fontsize=12, pad=15)
        for bar, value in zip(bars, stats_values):
            plt.text(bar.get_width() + 0.1, bar.get_y() + bar.get_height()/2,
                    f'{{int(value)}}', ha='left', va='center', fontweight='bold', fontsize=9)
        
        plt.tight_layout()
        
        # Create simplified PDF with both pages (no PNG files)
        base_name = excel_file.replace('.xlsx', '')
        
        # Create Page 2: Simplified - Only Defect Leakage Matrix
        plt.figure(figsize=(20, 12))
        plt.suptitle(f'Power BI Dashboard - {{project_name}} (Defect Leakage Matrix)', fontsize=20, fontweight='bold', y=0.95)
        
        # Create defect leakage matrix from the data
        try:
            leakage_matrix = jira_data.groupby(['Phase Injected', 'Phase Detected']).size().unstack(fill_value=0)
            
            # Large Defect Leakage Matrix (Full page)
            plt.subplot(1, 1, 1)
            if not leakage_matrix.empty:
                sns.heatmap(leakage_matrix, annot=True, fmt='d', cmap='Blues', 
                           cbar_kws={{'label': 'Count'}}, linewidths=0.5,
                           square=True, annot_kws={{'size': 12}})
                plt.title('Defect Leakage Matrix', fontweight='bold', fontsize=18, pad=30)
                plt.xlabel('Phase Detected', fontweight='bold', fontsize=14)
                plt.ylabel('Phase Injected', fontweight='bold', fontsize=14)
            else:
                plt.text(0.5, 0.5, 'No Leakage Data Available', ha='center', va='center', 
                        transform=plt.gca().transAxes, fontsize=16)
                plt.title('Defect Leakage Matrix', fontweight='bold', fontsize=18, pad=30)
        except Exception as e:
            plt.subplot(1, 1, 1)
            plt.text(0.5, 0.5, 'Leakage Matrix Not Available', ha='center', va='center', 
                    transform=plt.gca().transAxes, fontsize=16)
            plt.title('Defect Leakage Matrix', fontweight='bold', fontsize=18, pad=30)
        
        plt.tight_layout()
        
        # Create PDF with 2 pages: First 6 charts on Page 1, Defect Leakage Matrix on Page 2
        from matplotlib.backends.backend_pdf import PdfPages
        pdf_file = f"{{base_name}}.pdf"
        with PdfPages(pdf_file) as pdf:
            # Page 1: All 6 Charts (2x3 grid layout)
            plt.figure(figsize=(20, 12))  # Standard size for all pages
            # Remove header - no suptitle
            
            # 2x3 grid for all 6 charts
            # Top row: Phase, Priority, Environment
            plt.subplot(2, 3, 1)
            status_counts = jira_data['Status'].value_counts().sort_index()
            x_pos = range(len(status_counts))
            plt.plot(x_pos, status_counts.values, marker='o', linewidth=2, markersize=8, 
                    color=simple_colors[0])
            plt.fill_between(x_pos, status_counts.values, alpha=0.3, color=simple_colors[0])
            plt.xticks(x_pos, status_counts.index, rotation=45, fontsize=11, ha='right')
            plt.ylabel('Count', fontsize=11)
            plt.grid(True, alpha=0.3)
            for i, value in enumerate(status_counts.values):
                plt.text(i, value + 0.1, str(int(value)), ha='center', va='bottom', fontweight='bold', fontsize=10)
            plt.title('Phase', fontweight='bold', fontsize=14, pad=15)  # Adjusted title
            
            plt.subplot(2, 3, 2)
            priority_counts = jira_data['Priority'].value_counts()
            
            # Create Treemap Chart for Defect by Priority
            ax_treemap = plt.gca()
            ax_treemap.clear()
            ax_treemap.set_xlim(0, 1)
            ax_treemap.set_ylim(0, 1)
            ax_treemap.axis('off')
            
            # Sort priorities by count (descending) for better visualization
            priorities_sorted = priority_counts.sort_values(ascending=False)
            priorities = priorities_sorted.index.tolist()
            counts = priorities_sorted.values.tolist()
            total_count = sum(counts)
            
            # Define colors for priorities (red for critical, orange for high, yellow for medium, green for low)
            priority_colors = {{}}
            for priority in priorities:
                priority_lower = str(priority).lower()
                if 'critical' in priority_lower or 'blocker' in priority_lower:
                    priority_colors[priority] = '#F44336'  # Red
                elif 'highest' in priority_lower:
                    priority_colors[priority] = '#F44336'  # Red for highest
                elif 'high' in priority_lower or 'major' in priority_lower:
                    priority_colors[priority] = '#FF9800'  # Orange
                elif 'medium' in priority_lower or 'normal' in priority_lower:
                    priority_colors[priority] = '#FFC107'  # Yellow
                else:
                    priority_colors[priority] = '#4CAF50'  # Green
            
            # Simple treemap layout algorithm (squarified)
            # Start from top-left corner
            x, y = 0.05, 0.88
            width, height = 0.95, 0.90
            
            # Calculate rectangle positions for treemap
            rectangles = []
            current_x = x
            current_y = y
            
            # Simple row-based layout
            row_items = []
            row_total = 0
            remaining_width = width
            remaining_height = height
            
            for i, (priority, count) in enumerate(zip(priorities, counts)):
                percentage = count / total_count if total_count > 0 else 0
                
                # Add to current row
                row_items.append((priority, count, percentage))
                row_total += count
                
                # Calculate if we should start a new row (squarified algorithm)
                if len(row_items) > 1:
                    # Calculate row height based on total count in row
                    row_height = (row_total / total_count) * remaining_height if total_count > 0 else remaining_height / len(priorities)
                    
                    # Calculate worst aspect ratio if we add this item
                    max_count_in_row = max(c for _, c, _ in row_items)
                    min_count_in_row = min(c for _, c, _ in row_items)
                    
                    # Estimate widths
                    avg_width = remaining_width / len(row_items)
                    worst_aspect = max(avg_width / row_height, row_height / avg_width) if row_height > 0 else 1
                    
                    # If adding next item would make aspect ratio worse, finalize current row
                    if i < len(priorities) - 1:  # Not the last item
                        next_count = counts[i + 1] if i + 1 < len(counts) else 0
                        test_row_total = row_total + next_count
                        test_row_height = (test_row_total / total_count) * remaining_height if total_count > 0 else remaining_height / (len(priorities) - i)
                        test_avg_width = remaining_width / (len(row_items) + 1)
                        test_worst_aspect = max(test_avg_width / test_row_height, test_row_height / test_avg_width) if test_row_height > 0 else 1
                        
                        # If adding next item makes it worse, finalize current row
                        if test_worst_aspect > worst_aspect * 1.1:  # 10% threshold
                            # Draw current row with proportional widths
                            row_height = (row_total / total_count) * remaining_height if total_count > 0 else remaining_height / len(priorities)
                            current_x_row = current_x
                            for j, (p, c, pct) in enumerate(row_items):
                                # Width proportional to count
                                item_width = (c / row_total) * remaining_width if row_total > 0 else remaining_width / len(row_items)
                                rect_y = current_y - row_height
                                rectangles.append((current_x_row, rect_y, item_width, row_height, p, c, priority_colors[p]))
                                current_x_row += item_width
                            
                            # Move to next row
                            current_y -= row_height
                            remaining_height -= row_height
                            row_items = []
                            row_total = 0
                            continue
                
                # If last item, draw the final row
                if i == len(priorities) - 1:
                    row_height = (row_total / total_count) * remaining_height if total_count > 0 else remaining_height
                    current_x_row = current_x
                    for j, (p, c, pct) in enumerate(row_items):
                        # Width proportional to count
                        item_width = (c / row_total) * remaining_width if row_total > 0 else remaining_width / len(row_items)
                        rect_y = current_y - row_height
                        rectangles.append((current_x_row, rect_y, item_width, row_height, p, c, priority_colors[p]))
                        current_x_row += item_width
            
            # Draw rectangles
            for rect_x, rect_y, rect_width, rect_height, priority, count, color in rectangles:
                # Draw rectangle
                rect = Rectangle((rect_x, rect_y), rect_width, rect_height,
                               facecolor=color, edgecolor='white', linewidth=2)
                ax_treemap.add_patch(rect)
                
                # Add label (priority name)
                label_x = rect_x + rect_width / 2
                label_y = rect_y + rect_height / 2 + 0.02
                ax_treemap.text(label_x, label_y, str(priority), 
                              ha='center', va='center', fontsize=10, fontweight='bold', color='white')
                
                # Add count
                count_y = rect_y + rect_height / 2 - 0.02
                ax_treemap.text(label_x, count_y, str(int(count)), 
                              ha='center', va='center', fontsize=11, fontweight='bold', color='white')
            
            plt.title('Priority', fontweight='bold', fontsize=14, pad=15)  # Adjusted title
            
            plt.subplot(2, 3, 3)
            env_counts = jira_data['Environment'].value_counts()
            # Format labels to show count instead of percentage
            def make_autopct_env(values):
                def my_autopct(pct):
                    total = sum(values)
                    val = int(round(pct*total/100.0))
                    return f'{{val}}'
                return my_autopct
            wedges, texts, autotexts = plt.pie(env_counts.values, labels=env_counts.index,
                                              autopct=make_autopct_env(env_counts.values),
                                              colors=simple_colors[:len(env_counts)],
                                              startangle=90, pctdistance=0.85, textprops={{'fontsize': 12}})  # Adjusted text
            centre_circle = plt.Circle((0,0), 0.70, fc='white')
            plt.gca().add_artist(centre_circle)
            plt.title('Environment', fontweight='bold', fontsize=14, pad=15)  # Adjusted title
            
            # Bottom row: Assignee, Time Metrics, Overall Statistics
            plt.subplot(2, 3, 4)
            assignee_counts = jira_data['Assignee'].value_counts()
            colors = simple_colors[:len(assignee_counts)]
            # Format labels to show count instead of percentage
            def make_autopct_assignee_pdf(values):
                def my_autopct(pct):
                    total = sum(values)
                    val = int(round(pct*total/100.0))
                    return f'{{val}}'
                return my_autopct
            # Create 3D pie chart effect with explode and shadow
            explode = [0.1] * len(assignee_counts)  # Explode for 3D effect
            wedges, texts, autotexts = plt.pie(assignee_counts.values, labels=assignee_counts.index,
                                              autopct=make_autopct_assignee_pdf(assignee_counts.values),
                                              colors=colors,
                                              startangle=90, explode=explode,
                                              textprops={{'fontsize': 11}}, shadow=True, 
                                              wedgeprops={{'edgecolor': 'white', 'linewidth': 2}})
            # Enhance 3D effect by making text bold and adding depth
            for autotext in autotexts:
                autotext.set_color('white')
                autotext.set_fontweight('bold')
                autotext.set_fontsize(11)
            plt.title('Assignee', fontweight='bold', fontsize=14, pad=15)  # Adjusted title
            
            plt.subplot(2, 3, 5)
            # Calculate resolution times in hours for PDF
            jira_data['Created Date'] = pd.to_datetime(jira_data['Created Date'])
            jira_data['Resolution Date'] = pd.to_datetime(jira_data['Resolution Date'])
            jira_data['Resolution Time (Hours)'] = (jira_data['Resolution Date'] - jira_data['Created Date']).dt.total_seconds() / 3600
            resolution_times = jira_data['Resolution Time (Hours)'].dropna()
            
            if len(resolution_times) > 0:
                time_metrics = ['Minimum', 'Maximum', 'Average']
                min_val = round(resolution_times.min(), 1)  # Round to 1 decimal place
                max_val = round(resolution_times.max(), 1)  # Round to 1 decimal place
                avg_val = round(resolution_times.mean(), 1)  # Round to 1 decimal place
                time_values = [min_val, max_val, avg_val]
                bars = plt.bar(time_metrics, time_values, color=simple_colors[:3])
                plt.title('Time', fontweight='bold', fontsize=14, pad=15)  # Adjusted title
                plt.ylabel('Hours', fontsize=12)  # Adjusted axis label
                plt.xticks(fontsize=11)  # Adjusted axis text
                for bar, value in zip(bars, time_values):
                    # Format all values with 1 decimal place (xx.x)
                    plt.text(bar.get_x() + bar.get_width()/2., bar.get_height() + 0.1,
                            f'{{value:.1f}}'.format(value=value), ha='center', va='bottom', fontweight='bold', fontsize=11)  # Adjusted text
            else:
                plt.text(0.5, 0.5, 'No Resolution Data', ha='center', va='center', 
                        transform=plt.gca().transAxes, fontsize=14)  # Adjusted text
                plt.title('Time', fontweight='bold', fontsize=14, pad=15)  # Adjusted title
            
            plt.subplot(2, 3, 6)
            # Calculate overall statistics
            total_defects = len(jira_data)
            # Count closed defects (only DONE, CLOSED, COMPLETED, RESOLVED)
            closed_statuses = ['DONE', 'CLOSED', 'COMPLETED', 'RESOLVED']
            closed_defects = len(jira_data[jira_data['Status'].isin(closed_statuses)])
            # All other defects are considered open (ensures Open + Closed = Total)
            open_defects = total_defects - closed_defects
            
            stats_labels = ['Total Defects', 'Open Defects', 'Closed Defects']
            stats_values = [total_defects, open_defects, closed_defects]
            bars = plt.barh(stats_labels, stats_values, color=simple_colors[:3])
            plt.title('Status', fontweight='bold', fontsize=14, pad=15)  # Adjusted title
            plt.yticks(fontsize=11)  # Adjusted axis text
            for bar, value in zip(bars, stats_values):
                plt.text(bar.get_width() + 0.1, bar.get_y() + bar.get_height()/2,
                        f'{{int(value)}}', ha='left', va='center', fontweight='bold', fontsize=11)  # Adjusted text
            
            plt.tight_layout()
            pdf.savefig(bbox_inches='tight', facecolor='white', pad_inches=0.3)  # 70% layout
            plt.close()
            
            # Page 2: Defect Leakage Matrix (horizontal layout, same size as Page 1)
            plt.figure(figsize=(20, 12))  # Same size as Page 1
            # Remove header - no suptitle
            
            # Full page Defect Leakage Matrix
            ax = plt.subplot(1, 1, 1)
            if not leakage_matrix.empty:
                # Transpose matrix to make it horizontal
                horizontal_matrix = leakage_matrix.T
                sns.heatmap(horizontal_matrix, annot=True, fmt='d', cmap='Blues', 
                           cbar_kws={{'label': 'Count'}}, linewidths=0.5,
                           square=True, annot_kws={{'size': 16}}, ax=ax)  # Larger annotation text
                ax.set_title('Defect Leakage Matrix', fontweight='bold', fontsize=24, pad=40)  # Larger title
                ax.set_xlabel('Phase Injected', fontweight='bold', fontsize=18)  # Larger axis label
                ax.set_ylabel('Phase Detected', fontweight='bold', fontsize=18)  # Larger axis label
                ax.tick_params(axis='x', labelsize=14)  # Larger axis text
                ax.tick_params(axis='y', labelsize=14)  # Larger axis text
            else:
                ax.text(0.5, 0.5, 'No Leakage Data Available', ha='center', va='center', 
                       transform=ax.transAxes, fontsize=20)  # Larger text
                ax.set_title('Defect Leakage Matrix', fontweight='bold', fontsize=24, pad=40)  # Larger title
            
            plt.tight_layout()
            pdf.savefig(bbox_inches='tight', facecolor='white', pad_inches=0.5)  # Same 70% layout
            plt.close()
        
        # Set PDF to open at 54% zoom
        try:
            try:
                from PyPDF2 import PdfReader, PdfWriter
                from PyPDF2.generic import NameObject, NumberObject, DictionaryObject, ArrayObject
            except ImportError:
                from pypdf import PdfReader, PdfWriter
                from pypdf.generic import NameObject, NumberObject, DictionaryObject, ArrayObject
            
            # Read the PDF
            reader = PdfReader(pdf_file)
            writer = PdfWriter()
            
            # Copy all pages
            for page in reader.pages:
                writer.add_page(page)
            
            # Copy metadata if available
            if reader.metadata:
                writer.add_metadata(reader.metadata)
            
            # Set zoom to 54% (0.54)
            zoom_factor = 0.54
            
            # Get first page reference
            first_page_ref = writer._pages[0].indirect_reference
            
            # Create destination array: [page, /XYZ, left, top, zoom]
            # For 54% zoom, we use 0.54
            dest_array = ArrayObject([
                first_page_ref,
                NameObject('/XYZ'),
                NumberObject(0),
                NumberObject(0),  # Use 0 for top to show from top
                NumberObject(zoom_factor)
            ])
            
            # Create OpenAction with destination
            open_action = DictionaryObject()
            open_action[NameObject('/S')] = NameObject('/GoTo')
            open_action[NameObject('/D')] = dest_array
            
            # Set OpenAction in the root catalog
            writer._root_object[NameObject('/OpenAction')] = open_action
            
            # Write the modified PDF
            with open(pdf_file, 'wb') as output_file:
                writer.write(output_file)
                
        except (ImportError, Exception) as e:
            # If setting zoom fails (library not available or other error), continue without it
            pass
        
        print(f"Power BI Dashboard (PDF) saved: {{pdf_file}}")
        
        # Don't launch files here - they will be launched together at the end
        # Return PDF file path for batch launching
        
        # Remove dashboard auto-popup - no plt.show()
        
        
        return pdf_file
        
    except Exception as e:
        print(f"Error generating Power BI dashboard: {{str(e)}}")
        return False

if __name__ == "__main__":
    generate_powerbi_dashboard()
'''
            
            # Execute the script directly without saving it
            
            # Create a temporary script and execute it
            import tempfile
            with tempfile.NamedTemporaryFile(mode='w', suffix='.py', delete=False) as temp_script:
                temp_script.write(report_script)
                temp_script_path = temp_script.name
            
            # Construct the expected PDF file path (same pattern as in the script)
            # Ensure we use the same directory as the Excel file
            excel_dir = os.path.dirname(os.path.abspath(excel_file_path))
            excel_basename = os.path.basename(excel_file_path)
            pdf_basename = excel_basename.replace('.xlsx', '.pdf')
            pdf_file_path = os.path.join(excel_dir, pdf_basename)
            
            try:
                result = subprocess.run(['python', temp_script_path], capture_output=True, text=True)
                print(result.stdout)
                if result.stderr:
                    print(f"⚠️ Warning: {result.stderr}")
                
                # Clean up temporary file
                import os
                os.unlink(temp_script_path)
                
                # Check if PDF file was created
                if os.path.exists(pdf_file_path):
                    return pdf_file_path
                else:
                    print(f"⚠️ Warning: PDF file not found at expected path: {pdf_file_path}")
                    return None
            except Exception as e:
                print(f"Error executing Power BI script: {str(e)}")
                # Clean up temporary file even on error
                try:
                    import os
                    os.unlink(temp_script_path)
                except:
                    pass
                return None
            
        except Exception as e:
            print(f"❌ Error creating Power BI direct report: {str(e)}")
            return None
