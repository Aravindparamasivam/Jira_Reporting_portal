from datetime import datetime

class JiraExporter:
    
    def __init__(self, issues):
        self.issues = issues

    def _format_date(self, date_str):
        if not date_str:
            return None
        try:
            dt = datetime.fromisoformat(date_str.replace('Z', '+00:00'))
            return dt.strftime('%Y-%m-%d %H:%M:%S')
        except:
            return date_str

    def _calculate_duration_hours(self, created_date_str, resolution_date_str, issue=None):
        """Calculate duration between created date and resolution date in HH:MM format"""
        if not created_date_str:
            return 'N/A'
        
        # If no resolution date, return 'N/A' or '00:00' based on requirement
        if not resolution_date_str:
            return '00:00'
        
        try:
            # Calculate duration as Resolution Date - Created Date
            return self._calculate_simple_duration(created_date_str, resolution_date_str)
            
        except Exception as e:
            return 'N/A'
    
    def _calculate_simple_duration(self, created_date_str, resolution_date_str):
        """Calculate duration between created date and resolution date"""
        try:
            # Parse both dates
            created_dt = datetime.fromisoformat(created_date_str.replace('Z', '+00:00'))
            resolution_dt = datetime.fromisoformat(resolution_date_str.replace('Z', '+00:00'))
            
            # Calculate the difference: Resolution Date - Created Date
            duration = resolution_dt - created_dt
            
            # Get total seconds and convert to HH:MM
            total_seconds = int(duration.total_seconds())
            
            # Handle negative durations (if resolution date is before created date)
            if total_seconds < 0:
                return '00:00'
            
            # Calculate hours, minutes
            hours = total_seconds // 3600
            minutes = (total_seconds % 3600) // 60
            
            # Format as HH:MM
            return f"{hours:02d}:{minutes:02d}"
            
        except Exception as e:
            return 'N/A'
    
    def _calculate_smart_duration(self, created_date_str, last_status_change_str, issue):
        """Calculate duration considering status history - reset when moving from DONE to open state"""
        try:
            # Parse created date
            created_dt = datetime.fromisoformat(created_date_str.replace('Z', '+00:00'))
            
            # Get status history from changelog
            status_changes = []
            if 'changelog' in issue:
                for history in issue['changelog']['histories']:
                    for item in history['items']:
                        if item['field'] == 'status':
                            from_status = item.get('fromString', '').upper().strip()
                            to_status = item.get('toString', '').upper().strip()
                            change_date = datetime.fromisoformat(history['created'].replace('Z', '+00:00'))
                            status_changes.append((change_date, from_status, to_status))
            
            # Sort status changes by date (oldest first)
            status_changes.sort(key=lambda x: x[0])
            
            # Find the most recent transition from DONE to an open state
            last_done_to_open_transition = None
            for change_date, from_status, to_status in reversed(status_changes):
                if from_status == 'DONE' and to_status not in ['DONE']:
                    last_done_to_open_transition = change_date
                    break
            
            # Determine the start date for duration calculation
            if last_done_to_open_transition:
                # Use the most recent transition from DONE to open state
                start_date = last_done_to_open_transition
            else:
                # Use the original created date
                start_date = created_dt
            
            # Parse last status change date
            last_change_dt = datetime.fromisoformat(last_status_change_str.replace('Z', '+00:00'))
            
            # Calculate the difference
            duration = last_change_dt - start_date
            
            # Get total seconds and convert to HH:MM
            total_seconds = int(duration.total_seconds())
            
            # Handle negative durations
            if total_seconds < 0:
                return '00:00'
            
            # Calculate hours, minutes
            hours = total_seconds // 3600
            minutes = (total_seconds % 3600) // 60
            
            # Format as HH:MM
            return f"{hours:02d}:{minutes:02d}"
            
        except Exception as e:
            # Fallback to simple calculation if smart calculation fails
            return self._calculate_simple_duration(created_date_str, last_status_change_str)

    def _get_assignee_history(self, issue):
        assignee_history = []
        if 'changelog' in issue:
            # Create a list of all assignee changes with their timestamps
            changes = []
            for history in issue['changelog']['histories']:
                for item in history['items']:
                    if item['field'] == 'assignee':
                        from_name = item.get('fromString', 'Unassigned')
                        to_name = item.get('toString', 'Unassigned')
                        date = self._format_date(history['created'])
                        changes.append((date, f"{from_name} → {to_name} ({date})"))
            
            # Sort changes by date (oldest first)
            changes.sort(key=lambda x: x[0])
            
            # Extract just the formatted strings in chronological order
            assignee_history = [change[1] for change in changes]
            
        return ' | '.join(assignee_history) if assignee_history else 'No changes'

    def _get_parent_story_prefix(self, issue):
        """Extract prefix from parent story ticket summary"""
        try:
            # Get parent information from the issue
            parent = issue.get('fields', {}).get('parent')
            if parent:
                # The parent summary is nested inside parent['fields']['summary']
                parent_fields = parent.get('fields', {})
                parent_summary = parent_fields.get('summary', '')
                
                if parent_summary:
                    # Extract prefix before the first separator (space, dash, colon, etc.)
                    import re
                    # Look for patterns like "Rates - Sample 1" -> extract "Rates"
                    # Also handle patterns like "Forms - Sample 2" -> extract "Forms"
                    match = re.match(r'^([^-:\s]+)', parent_summary.strip())
                    if match:
                        return match.group(1).strip()
                    
                    # If no match with regex, try to split by common separators
                    for separator in [' - ', ':', ' ']:
                        if separator in parent_summary:
                            parts = parent_summary.split(separator)
                            if parts and parts[0].strip():
                                return parts[0].strip()
            
            # If no parent found, return 'NA' instead of trying to extract from issue summary
            # This ensures we only get phase from parent stories, not from the issue itself
            return 'NA'
        except Exception as e:
            return 'NA'

    def _get_environment_field(self, fields):
        """Extract environment information from the Environment field"""
        try:
            # First, check the standard environment field
            environment_value = fields.get('environment')
            if environment_value:
                # Handle Jira Document Format (Atlassian Document Format)
                if isinstance(environment_value, dict) and 'content' in environment_value:
                    # Extract text from the document format
                    text_content = self._extract_text_from_doc(environment_value)
                    if text_content and text_content.strip():
                        return text_content.strip()
                # Handle simple string
                elif isinstance(environment_value, str) and environment_value.strip():
                    return environment_value.strip()
            
            # Look for the Environment field in custom fields with 'environment' in the name
            for field_key, field_value in fields.items():
                if isinstance(field_key, str) and 'environment' in field_key.lower():
                    if field_value:
                        # Handle Jira Document Format
                        if isinstance(field_value, dict) and 'content' in field_value:
                            text_content = self._extract_text_from_doc(field_value)
                            if text_content and text_content.strip():
                                return text_content.strip()
                        # Handle simple string
                        elif isinstance(field_value, str) and field_value.strip():
                            return field_value.strip()
            
            # Also check for common environment field names
            environment_field_names = [
                'environment',
                'Environment',
                'ENVIRONMENT',
                'qa_environment',
                'uat_environment',
                'deployment_environment'
            ]
            
            for field_name in environment_field_names:
                if field_name in fields:
                    field_value = fields[field_name]
                    if field_value:
                        # Handle Jira Document Format
                        if isinstance(field_value, dict) and 'content' in field_value:
                            text_content = self._extract_text_from_doc(field_value)
                            if text_content and text_content.strip():
                                return text_content.strip()
                        # Handle simple string
                        elif isinstance(field_value, str) and field_value.strip():
                            return field_value.strip()
            
            # Try to infer environment from status (only if it's a clear environment status)
            status_field = fields.get('status')
            if status_field and isinstance(status_field, dict):
                status_name = status_field.get('name', '').upper()
                if status_name in ['DEV', 'QA', 'UAT', 'PROD']:
                    return status_name
            
            # Try to infer from labels (only if they are clear environment labels)
            labels = fields.get('labels', [])
            if labels:
                for label in labels:
                    label_upper = label.upper()
                    if label_upper in ['DEV', 'QA', 'UAT', 'PROD', 'TEST', 'STAGING']:
                        return label
            
            # Try to infer from components (only if they are clear environment components)
            components = fields.get('components', [])
            if components:
                for component in components:
                    if isinstance(component, dict):
                        component_name = component.get('name', '').upper()
                        if component_name in ['DEV', 'QA', 'UAT', 'PROD', 'TEST', 'STAGING']:
                            return component_name
            
            # Default fallback - return a more descriptive value instead of 'NA'
            return 'Not Specified'
        except Exception as e:
            return 'Not Specified'

    def _extract_text_from_doc(self, doc_obj):
        """Extract text content from Jira Document Format object"""
        try:
            if not isinstance(doc_obj, dict) or 'content' not in doc_obj:
                return None
            
            text_parts = []
            
            def extract_text_recursive(content_list):
                for item in content_list:
                    if isinstance(item, dict):
                        if item.get('type') == 'text' and 'text' in item:
                            text_parts.append(item['text'])
                        elif 'content' in item:
                            extract_text_recursive(item['content'])
            
            extract_text_recursive(doc_obj['content'])
            return ' '.join(text_parts)
            
        except Exception as e:
            return None

    def _get_odc_defect_type(self, fields):
        """Extract ODC Defect Type field from Jira"""
        try:
            # Check for specific known field IDs first
            if 'customfield_10124' in fields:
                field_value = fields['customfield_10124']
                if isinstance(field_value, dict) and 'value' in field_value:
                    value = field_value['value']
                    # Extract only the prefix before the arrow
                    if '→' in value:
                        return value.split('→')[0].strip()
                    return value
            
            # Look for custom fields that might contain ODC Defect Type
            for field_key, field_value in fields.items():
                if field_key.startswith('customfield_'):
                    if isinstance(field_value, dict) and 'value' in field_value:
                        field_name = str(field_value.get('name', '')).lower()
                        if 'defect type' in field_name or ('type' in field_name and 'defect' in field_name):
                            value = field_value['value']
                            # Extract only the prefix before the arrow
                            if '→' in value:
                                return value.split('→')[0].strip()
                            return value
            
            return 'Not Specified'
        except Exception as e:
            return 'Not Specified'

    def _get_odc_trigger(self, fields):
        """Extract ODC Trigger field from Jira"""
        try:
            # Check for specific known field IDs first
            if 'customfield_10125' in fields:
                field_value = fields['customfield_10125']
                if isinstance(field_value, dict) and 'value' in field_value:
                    value = field_value['value']
                    # Extract only the prefix before the arrow
                    if '→' in value:
                        return value.split('→')[0].strip()
                    return value
            
            # Look for custom fields that might contain ODC Trigger
            for field_key, field_value in fields.items():
                if field_key.startswith('customfield_'):
                    if isinstance(field_value, dict) and 'value' in field_value:
                        field_name = str(field_value.get('name', '')).lower()
                        if 'defect trigger' in field_name or ('trigger' in field_name and 'defect' in field_name):
                            value = field_value['value']
                            # Extract only the prefix before the arrow
                            if '→' in value:
                                return value.split('→')[0].strip()
                            return value
            
            return 'Not Specified'
        except Exception as e:
            return 'Not Specified'

    def _get_odc_impact(self, fields):
        """Extract ODC Impact field from Jira"""
        try:
            # Check for specific known field IDs first
            if 'customfield_10126' in fields:
                field_value = fields['customfield_10126']
                if isinstance(field_value, dict) and 'value' in field_value:
                    value = field_value['value']
                    # Extract only the prefix before the arrow
                    if '→' in value:
                        return value.split('→')[0].strip()
                    return value
            
            # Look for custom fields that might contain ODC Impact
            for field_key, field_value in fields.items():
                if field_key.startswith('customfield_'):
                    if isinstance(field_value, dict) and 'value' in field_value:
                        field_name = str(field_value.get('name', '')).lower()
                        if 'defect impact' in field_name or ('impact' in field_name and 'defect' in field_name):
                            value = field_value['value']
                            # Extract only the prefix before the arrow
                            if '→' in value:
                                return value.split('→')[0].strip()
                            return value
            
            return 'Not Specified'
        except Exception as e:
            return 'Not Specified'

    def _extract_abbreviated_phase(self, value):
        """Extract abbreviated phase name from full phase string (e.g., 'UAT' from 'UAT (User Acceptance Test)')"""
        if not value or value == 'Not Specified':
            return value
        
        # Extract only the prefix before the arrow if present
        if '→' in value:
            value = value.split('→')[0].strip()
        
        # Extract only the part before parentheses (e.g., "UAT" from "UAT (User Acceptance Test)")
        if '(' in value:
            value = value.split('(')[0].strip()
        
        return value

    def _get_odc_phase_detected(self, fields):
        """Extract ODC Phase Detected field from Jira"""
        try:
            # Check for specific known field IDs first
            if 'customfield_10127' in fields:
                field_value = fields['customfield_10127']
                if isinstance(field_value, dict) and 'value' in field_value:
                    value = field_value['value']
                    return self._extract_abbreviated_phase(value)
            
            # Look for custom fields that might contain ODC Phase Detected
            for field_key, field_value in fields.items():
                if field_key.startswith('customfield_'):
                    if isinstance(field_value, dict) and 'value' in field_value:
                        field_name = str(field_value.get('name', '')).lower()
                        if 'phase detected' in field_name or ('detected' in field_name and 'phase' in field_name):
                            value = field_value['value']
                            return self._extract_abbreviated_phase(value)
            
            return 'Not Specified'
        except Exception as e:
            return 'Not Specified'

    def _get_odc_phase_injected(self, fields):
        """Extract ODC Phase Injected field from Jira"""
        try:
            # Check for specific known field IDs first
            if 'customfield_10128' in fields:
                field_value = fields['customfield_10128']
                if isinstance(field_value, dict) and 'value' in field_value:
                    value = field_value['value']
                    return self._extract_abbreviated_phase(value)
            
            # Look for custom fields that might contain ODC Phase Injected
            for field_key, field_value in fields.items():
                if field_key.startswith('customfield_'):
                    if isinstance(field_value, dict) and 'value' in field_value:
                        field_name = str(field_value.get('name', '')).lower()
                        if 'phase injected' in field_name or ('injected' in field_name and 'phase' in field_name):
                            value = field_value['value']
                            return self._extract_abbreviated_phase(value)
            
            return 'Not Specified'
        except Exception as e:
            return 'Not Specified'

    def _get_odc_severity(self, fields):
        """Extract ODC Severity field from Jira"""
        try:
            # Check for specific known field IDs first
            if 'customfield_10129' in fields:
                field_value = fields['customfield_10129']
                if isinstance(field_value, dict) and 'value' in field_value:
                    value = field_value['value']
                    # Extract only the prefix before the arrow
                    if '→' in value:
                        return value.split('→')[0].strip()
                    return value
            
            # Look for custom fields that might contain ODC Severity
            for field_key, field_value in fields.items():
                if field_key.startswith('customfield_'):
                    if isinstance(field_value, dict) and 'value' in field_value:
                        field_name = str(field_value.get('name', '')).lower()
                        if 'severity' in field_name:
                            value = field_value['value']
                            # Extract only the prefix before the arrow
                            if '→' in value:
                                return value.split('→')[0].strip()
                            return value
            
            return 'Not Specified'
        except Exception as e:
            return 'Not Specified'

    def _get_odc_resolution(self, fields):
        """Extract ODC Resolution field from Jira"""
        try:
            # Check for specific known field IDs first
            if 'customfield_10130' in fields:
                field_value = fields['customfield_10130']
                if isinstance(field_value, dict) and 'value' in field_value:
                    value = field_value['value']
                    # Extract only the prefix before the arrow
                    if '→' in value:
                        return value.split('→')[0].strip()
                    return value
            
            # Look for custom fields that might contain ODC Resolution
            for field_key, field_value in fields.items():
                if field_key.startswith('customfield_'):
                    if isinstance(field_value, dict) and 'value' in field_value:
                        field_name = str(field_value.get('name', '')).lower()
                        if 'resolution' in field_name:
                            value = field_value['value']
                            # Extract only the prefix before the arrow
                            if '→' in value:
                                return value.split('→')[0].strip()
                            return value
            
            return 'Not Specified'
        except Exception as e:
            return 'Not Specified'

    def calculate_defect_metrics(self, defects_data):
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
                # Try both field name variations
                phase_injected = defect.get('Phase Injected') or defect.get('ODC Phase Injected', 'Not Specified')
                phase_detected = defect.get('Phase Detected') or defect.get('ODC Phase Detected', 'Not Specified')
                
                # Skip if either phase is not specified
                if not phase_injected or phase_injected == 'Not Specified' or not phase_detected or phase_detected == 'Not Specified':
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
                    'Unit Test Phase': 'Unit testing',  # Handle "Unit Test Phase" variant
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
            
            # Calculate review vs testing based on DETECTED phases
            # Review: Defects detected in Requirements, Design, Implementation
            review_defects = totals_detected.get('Requirements', 0) + totals_detected.get('Design', 0) + totals_detected.get('Implementation', 0)
            
            # Testing: Defects detected in Unit testing, QA, UAT, Production
            testing_defects = totals_detected.get('Unit testing', 0) + totals_detected.get('QA', 0) + totals_detected.get('UAT', 0) + totals_detected.get('Production', 0)
            
            # Calculate pre vs post production based on DETECTED phases
            # Pre-production: Defects detected in Requirements, Design, Implementation, Unit testing, QA, UAT
            pre_production_defects = (totals_detected.get('Requirements', 0) + 
                                     totals_detected.get('Design', 0) + 
                                     totals_detected.get('Implementation', 0) + 
                                     totals_detected.get('Unit testing', 0) + 
                                     totals_detected.get('QA', 0) + 
                                     totals_detected.get('UAT', 0))
            
            # Post-production: Defects detected in Production
            post_production_defects = totals_detected.get('Production', 0)
            
            # Calculate percentages (return None for division by zero to show #DIV/0! in Excel)
            total_detected_for_review_testing = review_defects + testing_defects
            percentage_review = (review_defects / total_detected_for_review_testing * 100) if total_detected_for_review_testing > 0 else None
            percentage_testing = (testing_defects / total_detected_for_review_testing * 100) if total_detected_for_review_testing > 0 else None
            
            total_detected_for_pre_post = pre_production_defects + post_production_defects
            percentage_pre_production = (pre_production_defects / total_detected_for_pre_post * 100) if total_detected_for_pre_post > 0 else None
            percentage_post_production = (post_production_defects / total_detected_for_pre_post * 100) if total_detected_for_pre_post > 0 else None
            
            return {
                'matrix': matrix,
                'totals_injected': totals_injected,
                'totals_detected': totals_detected,
                'total_defects': total_defects,
                'percentage_review': percentage_review,
                'percentage_testing': percentage_testing,
                'percentage_pre_production': percentage_pre_production,
                'percentage_post_production': percentage_post_production,
                'phases': phases
            }
            
        except Exception as e:
            print(f"Error calculating defect metrics: {e}")
            return None

    def calculate_defect_metrics_from_dates(self, defects_data):
        """Calculate defect leakage metrics using date-based analysis from Jira data"""
        try:
            # Initialize the defect leakage matrix
            phases = ['Requirements', 'Design', 'Implementation', 'Unit testing', 'QA', 'UAT', 'Production']
            matrix = {}
            
            # Initialize matrix with zeros
            for introduced in phases:
                matrix[introduced] = {}
                for detected in phases:
                    matrix[introduced][detected] = 0
            
            # Analyze each defect based on dates and status transitions
            for defect in defects_data:
                created_date = defect.get('Created Date')
                resolution_date = defect.get('Resolution Date')
                last_status_change = defect.get('Last Status Change')
                status = defect.get('Status', '').upper()
                assignee_history = defect.get('Assignee History', '')
                
                # Determine phase introduced based on creation date and initial status
                phase_injected = self._determine_phase_injected(created_date, status, assignee_history)
                
                # Determine phase detected based on resolution date and final status
                phase_detected = self._determine_phase_detected(resolution_date, last_status_change, status, assignee_history)
                
                # Only count if both phases are valid
                if phase_injected in phases and phase_detected in phases:
                    matrix[phase_injected][phase_detected] += 1
            
            # Calculate totals
            totals_injected = {}
            totals_detected = {}
            
            for phase in phases:
                totals_injected[phase] = sum(matrix[phase].values())
                totals_detected[phase] = sum(matrix[p][phase] for p in phases)
            
            # Calculate percentages
            total_defects = sum(totals_injected.values())
            
            # Calculate review vs testing based on DETECTED phases
            # Review: Defects detected in Requirements, Design, Implementation
            review_defects = totals_detected.get('Requirements', 0) + totals_detected.get('Design', 0) + totals_detected.get('Implementation', 0)
            
            # Testing: Defects detected in Unit testing, QA, UAT, Production
            testing_defects = totals_detected.get('Unit testing', 0) + totals_detected.get('QA', 0) + totals_detected.get('UAT', 0) + totals_detected.get('Production', 0)
            
            # Calculate pre vs post production based on DETECTED phases
            # Pre-production: Defects detected in Requirements, Design, Implementation, Unit testing, QA, UAT
            pre_production_defects = (totals_detected.get('Requirements', 0) + 
                                     totals_detected.get('Design', 0) + 
                                     totals_detected.get('Implementation', 0) + 
                                     totals_detected.get('Unit testing', 0) + 
                                     totals_detected.get('QA', 0) + 
                                     totals_detected.get('UAT', 0))
            
            # Post-production: Defects detected in Production
            post_production_defects = totals_detected.get('Production', 0)
            
            # Calculate percentages (return None for division by zero to show #DIV/0! in Excel)
            total_detected_for_review_testing = review_defects + testing_defects
            percentage_review = (review_defects / total_detected_for_review_testing * 100) if total_detected_for_review_testing > 0 else None
            percentage_testing = (testing_defects / total_detected_for_review_testing * 100) if total_detected_for_review_testing > 0 else None
            
            total_detected_for_pre_post = pre_production_defects + post_production_defects
            percentage_pre_production = (pre_production_defects / total_detected_for_pre_post * 100) if total_detected_for_pre_post > 0 else None
            percentage_post_production = (post_production_defects / total_detected_for_pre_post * 100) if total_detected_for_pre_post > 0 else None
            
            return {
                'matrix': matrix,
                'totals_injected': totals_injected,
                'totals_detected': totals_detected,
                'total_defects': total_defects,
                'percentage_review': percentage_review,
                'percentage_testing': percentage_testing,
                'percentage_pre_production': percentage_pre_production,
                'percentage_post_production': percentage_post_production,
                'phases': phases
            }
            
        except Exception as e:
            print(f"Error calculating defect metrics from dates: {e}")
            return None

    def _determine_phase_injected(self, created_date, status, assignee_history):
        """Determine the phase where the defect was introduced based on creation date and status"""
        try:
            if not created_date:
                return 'Requirements'  # Default assumption
            
            # Parse the created date
            from datetime import datetime
            created_dt = datetime.strptime(created_date, '%Y-%m-%d %H:%M:%S')
            
            # Analyze status and assignee history to determine phase
            status_upper = status.upper()
            
            # If status indicates a specific phase, use that
            if 'REQUIREMENTS' in status_upper or 'ANALYSIS' in status_upper:
                return 'Requirements'
            elif 'DESIGN' in status_upper:
                return 'Design'
            elif 'CODE' in status_upper or 'DEVELOPMENT' in status_upper:
                return 'Code'
            elif 'UNIT TEST' in status_upper or 'UNIT TESTING' in status_upper:
                return 'Unit testing'
            elif 'QA' in status_upper or 'TESTING' in status_upper:
                return 'QA'
            elif 'UAT' in status_upper:
                return 'UAT'
            elif 'PRODUCTION' in status_upper or 'PROD' in status_upper:
                return 'Production'
            
            # Analyze assignee history for clues
            if assignee_history:
                history_lower = assignee_history.lower()
                if 'dev' in history_lower or 'developer' in history_lower:
                    return 'Code'
                elif 'qa' in history_lower or 'tester' in history_lower:
                    return 'QA'
                elif 'analyst' in history_lower or 'business' in history_lower:
                    return 'Requirements'
            
            # Default based on common patterns
            return 'Code'  # Most defects are typically introduced during coding
            
        except Exception as e:
            print(f"Error determining phase injected: {e}")
            return 'Requirements'

    def _determine_phase_detected(self, resolution_date, last_status_change, status, assignee_history):
        """Determine the phase where the defect was detected based on resolution date and status"""
        try:
            # Use resolution date or last status change
            detection_date = resolution_date if resolution_date else last_status_change
            
            if not detection_date:
                return 'QA'  # Default assumption
            
            # Parse the detection date
            from datetime import datetime
            detection_dt = datetime.strptime(detection_date, '%Y-%m-%d %H:%M:%S')
            
            # Analyze status to determine detection phase
            status_upper = status.upper()
            
            # If status indicates a specific phase, use that
            if 'REQUIREMENTS' in status_upper or 'ANALYSIS' in status_upper:
                return 'Requirements'
            elif 'DESIGN' in status_upper:
                return 'Design'
            elif 'CODE' in status_upper or 'DEVELOPMENT' in status_upper:
                return 'Code'
            elif 'UNIT TEST' in status_upper or 'UNIT TESTING' in status_upper:
                return 'Unit testing'
            elif 'QA' in status_upper or 'TESTING' in status_upper:
                return 'QA'
            elif 'UAT' in status_upper:
                return 'UAT'
            elif 'PRODUCTION' in status_upper or 'PROD' in status_upper:
                return 'Production'
            
            # Analyze assignee history for detection clues
            if assignee_history:
                history_lower = assignee_history.lower()
                if 'qa' in history_lower or 'tester' in history_lower:
                    return 'QA'
                elif 'dev' in history_lower or 'developer' in history_lower:
                    return 'Code'
                elif 'analyst' in history_lower or 'business' in history_lower:
                    return 'Requirements'
            
            # Default based on common patterns
            return 'QA'  # Most defects are typically detected during QA
            
        except Exception as e:
            print(f"Error determining phase detected: {e}")
            return 'QA'

    def extract_data(self, header_value=None, parent_progress=None, parent_status=None):
        extracted_data = []
        main_ticket_summary = header_value
        
        # Check if issues list is empty
        if not self.issues:
            # Return a minimal structure to prevent errors
            return [], main_ticket_summary, parent_progress, parent_status
        
        # Extract all data as before
        for issue in self.issues:
            # Skip if issue is None
            if not issue:
                continue
                
            fields = issue.get('fields', {})
            if not fields:
                continue
            
            # Safely extract status with fallback
            status_field = fields.get('status')
            if status_field and isinstance(status_field, dict):
                status_value = status_field.get('name', 'Unknown')
            else:
                status_value = 'Unknown'
            
            # Convert status to uppercase
            status_value = status_value.upper()
            
            # Safe extraction with null checks
            priority = fields.get('priority')
            priority_name = priority.get('name') if priority else 'Unknown'
            
            issuetype = fields.get('issuetype')
            issuetype_name = issuetype.get('name') if issuetype else 'Unknown'
            
            reporter = fields.get('reporter')
            reporter_name = reporter.get('displayName') if reporter else 'Unknown'
            
            assignee = fields.get('assignee')
            assignee_name = assignee.get('displayName') if assignee else 'Unassigned'
            
            # Get raw date strings for duration calculation
            created_date_raw = fields.get('created')
            last_status_change_raw = fields.get('statuscategorychangedate')
            
            # Get Environment
            environment_value = self._get_environment_field(fields)
            
            # Get Phase Detected and Phase Injected
            phase_detected = self._get_odc_phase_detected(fields)
            phase_injected = self._get_odc_phase_injected(fields)
            
            # For DONE status, use last status change date as resolution date if resolution date is not available
            resolution_date = fields.get('resolutiondate')
            if status_value == 'DONE' and not resolution_date:
                resolution_date = last_status_change_raw
            
            data = {
                'Key': issue.get('key'),
                'Environment': environment_value,
                'Summary': fields.get('summary'),
                'Priority': priority_name,
                'Issue Type': issuetype_name,
                'Reporter': reporter_name,
                'Assignee': assignee_name,
                'Status': status_value,
                'Created Date': self._format_date(created_date_raw),
                # Use resolution date or last status change for DONE status
                'Resolution Date': self._format_date(resolution_date),
                'Last Status Change': self._format_date(last_status_change_raw),
                'Assignee History': self._get_assignee_history(issue),
                'Duration (HH:MM)': self._calculate_duration_hours(created_date_raw, resolution_date),
                'Phase Detected': phase_detected,
                'Phase Injected': phase_injected,
                '_created_date_raw': created_date_raw,  # Keep raw date for sorting
            }
            extracted_data.append(data)
        
        # Sort tickets by created date (oldest to newest)
        extracted_data.sort(key=lambda x: x['_created_date_raw'] or '')
        
        # Remove the temporary sorting field
        for data in extracted_data:
            data.pop('_created_date_raw', None)
        
        return extracted_data, main_ticket_summary, parent_progress, parent_status

    def extract_odc_data(self):
        """Extract ODC classification data for the ODC Classification tab"""
        odc_data = []
        
        for issue in self.issues:
            if not issue:
                continue
                
            fields = issue.get('fields', {})
            if not fields:
                continue
            
            # Get ODC classification fields directly from Jira
            odc_defect_type = self._get_odc_defect_type(fields)
            odc_trigger = self._get_odc_trigger(fields)
            odc_impact = self._get_odc_impact(fields)
            odc_phase_detected = self._get_odc_phase_detected(fields)
            odc_phase_injected = self._get_odc_phase_injected(fields)
            odc_severity = self._get_odc_severity(fields)
            odc_resolution = self._get_odc_resolution(fields)
            
            odc_data.append({
                'Key': issue.get('key'),
                'Summary': fields.get('summary'),
                'ODC Defect Type': odc_defect_type,
                'ODC Trigger': odc_trigger,
                'ODC Impact': odc_impact,
                'ODC Phase Detected': odc_phase_detected,
                'ODC Phase Injected': odc_phase_injected,
                'ODC Severity': odc_severity,
                'ODC Resolution': odc_resolution,
            })
        
        return odc_data