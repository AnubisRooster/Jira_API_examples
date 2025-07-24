from datetime import datetime, timedelta
from jira import JIRA
import pandas as pd

# Jira credentials and connection
jira = JIRA(server='[jira_domain]', basic_auth=('[emailaddress]', '[auth-token]'))

# JQL query for issues by Assignee within the last six months
jql_query = 'project in ([specify_projects) AND issuetype in (Bug, Story, Task, Sub-task, Improvement, Investigation, "New Feature", Request) and updated > -7d'

# Calculate the date range
end_date = datetime.now().date()
start_date = end_date - timedelta(days=7)

# Retrieve the updated tasks
issues = jira.search_issues(jql_query, maxResults=False)

# Prepare the data for the Excel spreadsheet
data = []
for issue in issues:
    assignee = issue.fields.assignee.displayName if issue.fields.assignee else "Unassigned"
    summary = issue.fields.summary
    issue_key = issue.key
    status = issue.fields.status.name
    parent_issue_summary = issue.fields.parent.fields.summary if hasattr(issue.fields, 'parent') and issue.fields.parent else ""
    data.append([assignee, summary, issue_key, status, parent_issue_summary])

# Create a DataFrame from the data
df = pd.DataFrame(data, columns=["Assignee", "Summary", "Issue Key", "Status", "Parent Issue Summary"])

# Sort the DataFrame by assignee
df = df.sort_values(by="Assignee")

# Get the current date for the output file name
current_date = datetime.now().strftime("%Y-%m-%d")

# Set the output file name based on the current date
output_file = f"task_data_{current_date}.xlsx"

# Export the DataFrame to an Excel spreadsheet
df.to_excel(output_file, index=False)

print(f"Task data exported to {output_file} successfully.")
