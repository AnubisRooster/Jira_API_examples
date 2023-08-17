from jira import JIRA
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Alignment

# Jira credentials and connection
jira = JIRA(server='https://memsql.atlassian.net', basic_auth=('it-automation+jira_reporting@singlestore.com', 'ATATT3xFfGF0JgsZms4GLYlWArKOSeTOnIOZx4pz0YQ0WkLnN3rQg8777YpW-m39gI8oS463by7hv_boSRiTNW9pYPadtDnXlDZiSTvLtRns8Xrh4CUQLRQsdnluREK84CoedSP0PN2Gk5H58BXI5EszVdjb6ht-ozqEFTiEesQhD5aPgH1Vbc0=FECBFB84'))

# Calculate the date 30 days ago
end_date = datetime.now()
start_date = end_date - timedelta(days=30)

# Fetch all issues from the project
jql_query = f'project = PMZ'
all_issues = jira.search_issues(jql_query, maxResults=False)

# Create a new Excel workbook and add a worksheet
workbook = Workbook()
worksheet = workbook.active
worksheet.title = 'Due Date Changes'

# Add headers to the worksheet
headers = ['Issue Key', 'Summary', 'Status', 'Created Date', 'Previous Due Date', 'Current Due Date', 'Components']
worksheet.append(headers)

# Iterate through the fetched issues
for issue in all_issues:
    issue_key = issue.key
    summary = issue.fields.summary
    created_date = issue.fields.created[:10]  # Extract the first 10 characters (YYYY-MM-DD)

    # Fetch issue changelogs
    changelogs = jira.issue(issue_key, expand='changelog').changelog.histories

    # Initialize variables to store previous values
    prev_due_date = None
    current_due_date = issue.fields.duedate

    components = issue.fields.components

    # Extract change details
    for change in changelogs:
        for item in change.items:
            if item.field == 'duedate':
                prev_due_date = item.fromString[:10] if item.fromString else None

    # Check if there are changes to the due date and if prev_due_date is not None
    if prev_due_date is not None and prev_due_date != current_due_date:
        # Convert components objects to a list of component names
        component_names = [component.name for component in components] if components else []

        # Add a row to the worksheet
        row_data = [issue_key, summary, issue.fields.status.name, created_date, prev_due_date, current_due_date, ', '.join(component_names)]
        worksheet.append(row_data)

# Apply alignment to headers and cells
for row in worksheet.iter_rows(min_row=1, max_row=1):
    for cell in row:
        cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
for column_cells in worksheet.columns:
    max_length = 0
    for cell in column_cells:
        if cell.value:
            cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            cell_length = len(str(cell.value))
            if cell_length > max_length:
                max_length = cell_length
    adjusted_width = (max_length + 2)
    worksheet.column_dimensions[column_cells[0].column_letter].width = adjusted_width

# Save the Excel workbook
workbook.save('due_date_changes.xlsx')
