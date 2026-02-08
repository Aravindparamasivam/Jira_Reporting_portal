# Jira Automation SDA

Automated tool to extract defects from Jira EPIC tickets and generate comprehensive Excel reports with defect leakage analysis, visual dashboards, and Power BI integration.

## Overview

This tool provides end-to-end automation for Jira defect tracking, analysis, and reporting. It generates multi-sheet Excel workbooks with detailed defect information, defect leakage matrices, interactive charts, and automated Power BI PDF dashboards.

## Workflow

This tool follows a streamlined interactive workflow:

1. **Archive existing files** → Automatically closes Excel processes and moves old reports to Archive folder
2. **Interactive Project Key input** → Enter project key(s) dynamically when prompted
3. **Process each project** → Fetch defects from EPIC child work items and subtasks
4. **Extract defect details** → Gather comprehensive information including Phase Detected and Phase Injected
5. **Generate Excel reports** → Create professionally formatted multi-sheet reports with charts
6. **Generate Power BI Dashboard** → Create PDF dashboard with visualizations
7. **Auto-launch files** → Opens Excel and PDF files with maximized windows

## Features

### Core Functionality
- ✅ **Interactive Input**: Enter project keys dynamically when prompted
- ✅ **Multi-Project Processing**: Process multiple comma-separated project keys in one run
- ✅ **Professional Excel Reports**: Multi-sheet workbooks with advanced formatting
- ✅ **File Archiving**: Automatic archival of old reports with retention management (keeps 10 most recent)
- ✅ **Process Management**: Automatically closes Excel and PDF processes before generating new files
- ✅ **Window Management**: Automatically maximizes Excel window on launch

### Excel Report Features
- ✅ **Multi-Sheet Structure**: Jira Report, QA Report, Metrics Board, Visual Board
- ✅ **Defect Leakage Matrix**: 7×7 matrix showing phase injection vs. detection with inactive cell logic
- ✅ **Interactive Charts**: Donut charts for Review vs Testing and Pre-production vs Post-production
- ✅ **Visual Dashboards**: Status, Priority, Environment, and Assignee workload charts
- ✅ **Phase Tracking**: Requirements, Design, Implementation, Unit testing, QA, UAT, Production
- ✅ **Formula-Based Calculations**: Dynamic SUM formulas with validation
- ✅ **Conditional Formatting**: Color-coded status indicators and validation messages
- ✅ **Smart Alignment**: Text fields left-aligned, data fields center-aligned

### Power BI Integration
- ✅ **Automated PDF Generation**: Creates Power BI-style PDF dashboards
- ✅ **Chart Visualizations**: Pie, bar, and doughnut charts
- ✅ **Single-Page Layout**: Compact dashboard design without scrolling
- ✅ **Automatic Launch**: Opens PDF in browser and Excel in maximized window

## Quick Start

### 1. Setup
```bash
# Install dependencies
pip install -r requirements.txt

# Configure Jira credentials in config.py
# Update JIRA_URL, EMAIL, and API_TOKEN
```

### 2. Usage

```bash
# Run the tool
python main.py

# Enter project key(s) when prompted:
# Single project: SDAA-1
# Multiple projects: SDAA-1, SDAA-2, SDAA-3
```

## Configuration

Update the credentials in `config.py`:

```python
JIRA_URL = "https://your-jira-instance.atlassian.net"
EMAIL = "your-email@company.com"
API_TOKEN = "your_api_token"
EXPORT_FOLDER = "Files"
ARCHIVE_FOLDER = "Archive"
```

## Excel Report Structure

### Sheet 1: "Jira Report" (Main detailed view)
Complete defect information with all fields:
- **Key**
- **Environment**
- **Summary**
- **Priority**
- **Issue Type**
- **Reporter**
- **Assignee**
- **Status** (with color coding)
- **Created Date**
- **Resolution Date**
- **Last Status Change**
- **Assignee History**
- **Duration (HH:MM:SS)**
- **Phase Detected**
- **Phase Injected**

### Sheet 2: "QA Report" (Simplified view)
Streamlined QA-focused information:
- **S.No**
- **Defect Key**
- **Defect Description**
- **Status**

### Sheet 3: "Metrics Board" (Defect Leakage Analysis)
Comprehensive defect leakage matrix and analytics:
- **Defect Leakage Matrix**: 7×7 matrix (Phase Injected × Phase Detected)
  - Review phases: Requirements, Design, Implementation
  - Testing phases: Unit testing, QA, UAT, Production
  - Inactive cells (greyed out) for impossible combinations
  - Row totals (Total injected) and column totals (Total detected)
  - Validation formula showing matching/mismatch status
- **Donut Charts**:
  - Review vs Testing phase defects (positioned at C15)
  - Pre-production vs Post-production defects (positioned at G15)
- **Phase Activities Table**: Description of activities for each phase

### Sheet 4: "Visual Board" (Dashboard)
Interactive charts and visualizations:
- **Defect Status Distribution**: Pie chart
- **Priority Distribution**: Column chart
- **Environment Analysis**: Doughnut chart
- **Assignee Workload Management**: Column chart
- **Landscape orientation**: Single-page layout without scrolling

## Power BI Dashboard

The tool automatically generates a PDF dashboard with:
- **Status Distribution Chart**: Visual breakdown of defect statuses
- **Priority Analysis**: Priority level distribution
- **Environment Breakdown**: Defect distribution across environments
- **Assignee Workload**: Team member workload visualization
- **Defect Leakage Matrix**: Visual representation of phase relationships

## File Management

### File Naming Convention
- **Format**: `Project_key={project_key},Date={date},Time={time}.xlsx`
- **Example**: `Project_key=SDAA-1,Date=2025-11-04,Time=12-37-04.xlsx`
- **PDF Dashboard**: `Project_key={project_key},Date={date},Time={time}_Dashboard.pdf`

### Auto-Archiving
- **Automatic**: Existing files moved to Archive folder before new generation
- **Process Management**: Closes Excel and PDF processes to prevent file conflicts
- **Retention**: Only keeps 10 most recent files in Archive
- **Organization**: Clean separation between current (Files) and archived reports

### Output Directories
- **Files/**: Active reports folder (newly generated files)
- **Archive/**: Historical reports folder (automatically managed)

## Visual Features

### Status Color Coding
- 🟢 **Green**: Done/Completed/Closed status
- 🔴 **Red**: To Do/Open status
- 🟡 **Yellow**: In Progress status

### Chart Colors
- **Blue** (#6b8e23): First segment (Reviews, Pre-production)
- **Red** (#c93435): Second segment (Testing, Post-production)

## Sample Output

```
Archiving existing files...
Closing any open Excel processes...
Closed Excel process: EXCEL.EXE (PID: 12345)
Archived: Project_key=SDAA-1,Date=2025-11-04,Time=12-30-00.xlsx
Successfully archived 1 file(s)
Enter project key(s) (comma-separated for multiple): SDAA-1

Processing project: SDAA-1
Fetching issues from Jira...
Extracting relevant fields...
Exporting to Excel...

📊 Generating Power BI dashboard...
Power BI Dashboard (PDF) saved: Files\Project_key=SDAA-1,Date=2025-11-04,Time=12-37-04_Dashboard.pdf
PDF dashboard launched in browser: Files\Project_key=SDAA-1,Date=2025-11-04,Time=12-37-04_Dashboard.pdf
Excel file launched: Files\Project_key=SDAA-1,Date=2025-11-04,Time=12-37-04.xlsx

Exported for SDAA-1:
Excel: Files\Project_key=SDAA-1,Date=2025-11-04,Time=12-37-04.xlsx
```

## Defect Leakage Matrix

The Metrics Board includes a comprehensive defect leakage matrix that shows:
- **Phase Injected**: Where defects were introduced (Requirements, Design, Implementation, Unit testing, QA, UAT, Production)
- **Phase Detected**: Where defects were found
- **Matrix Logic**: Grey cells indicate inactive/impossible combinations
- **Row Totals**: Total defects injected in each phase
- **Column Totals**: Total defects detected in each phase
- **Validation**: Automatic validation showing if totals match

### Phase Definitions
- **Requirements**: BRD Reviews, BRD Updates, Requirement Analysis
- **Design**: Source table structure designs, Forms
- **Implementation**: Code, Data load, API, Portal, CMS, Integration
- **Unit testing**: Backend Validations
- **QA**: Test plan, Test data, Test cases
- **UAT**: Test plan, Test data, Test cases, AM test scenarios
- **Production**: Data validations, High level functionality validations

## File Organization

```
JiraAutomation_3.0_SDA/
├── main.py                    # Main execution script with interactive flow
├── config.py                  # Configuration with Jira credentials
├── jira_api.py               # Direct Jira REST API client
├── jira_exporter.py          # Data transformation and extraction
├── exporter.py               # Excel report generation with formatting
├── file_handler.py           # File archiving and process management
├── requirements.txt          # Python dependencies
├── README.md                 # Project documentation
├── Files/                    # Generated reports folder (auto-created)
└── Archive/                  # Archived reports folder (auto-created)
```

## Technical Architecture

### Data Processing Pipeline
1. **Fetch**: Get EPIC child work items and their subtasks via Jira REST API
2. **Filter**: Identify Bug/Defect type issues
3. **Extract**: Parse all relevant fields including Phase Detected and Phase Injected
4. **Transform**: Format dates, calculate durations, map phases to standardized names
5. **Process**: Build defect leakage matrix and calculate metrics
6. **Export**: Generate multi-sheet Excel with advanced formatting and charts
7. **Visualize**: Create Power BI PDF dashboard with charts

### Technology Stack
- **Python 3.7+**
- **pandas**: Data manipulation and processing
- **xlsxwriter**: Excel file generation and formatting
- **openpyxl**: Excel file reading
- **matplotlib/seaborn**: Chart generation for Power BI
- **requests**: Jira REST API communication
- **psutil**: Process management

## Requirements

- Python 3.7+
- Valid Jira API Token
- Network access to Jira instance
- Required packages (see `requirements.txt`):
  - `requests==2.31.0`
  - `pandas==2.1.3`
  - `openpyxl==3.1.2`
  - `xlsxwriter==3.1.9`
  - `matplotlib==3.8.2`
  - `seaborn==0.13.0`
  - `psutil==5.9.6`

## Troubleshooting

### Common Issues
- **401 Unauthorized**: Check API token in config.py
- **No data found**: Verify EPIC exists and has Bug/Defect subtasks
- **Missing columns**: Ensure Jira fields are accessible with your permissions
- **Charts not displaying**: Ensure Excel file is saved and reopened
- **File access errors**: Old Excel processes are automatically closed before archiving

### Process Management
- The tool automatically closes existing Excel and PDF processes
- If files are still locked, manually close Excel before running the automation
- Archive folder maintains only the 10 most recent files

## Key Metrics Tracked

### Defect Metrics
- Total defects by status, priority, environment, and assignee
- Defect leakage rates (phase injection vs. detection)
- Pre-production vs. post-production defect ratios
- Review vs. testing phase effectiveness
- Completion percentages and resolution times

### Quality Indicators
- Percentage of defects detected in reviews
- Percentage of defects detected in testing
- Percentage of defects in pre-production
- Percentage of defects in post-production

## Benefits

- **Time Savings**: 95% reduction in manual report generation time
- **Automated Analysis**: Complete defect leakage analysis in seconds
- **Visual Insights**: Interactive charts and dashboards for quick understanding
- **Process Improvement**: Identify where defects are introduced vs. caught
- **Standardized Reporting**: Consistent metrics across all projects
- **Historical Tracking**: Automatic archiving for trend analysis

The tool provides **professional-grade Excel reports** with **interactive charts**, **defect leakage analysis**, **automated Power BI dashboards**, and **comprehensive file management** - perfect for defect tracking, quality analysis, and process optimization!
