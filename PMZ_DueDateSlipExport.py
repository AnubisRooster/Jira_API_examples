from jira import JIRA
from datetime import datetime, timedelta
import pandas as pd

# Set up your Jira server URL and credentials
JIRA_SERVER = 'https://[domain].atlassian.net'
USERNAME = '[email]'
PASSWORD = '[auth_token]'

# Connect to Jira
jira = JIRA(server=JIRA_SERVER, basic_auth=(USERNAME, PASSWORD))

# Calculate the date 90 days ago from today
ninety_days_ago = (datetime.now() - timedelta(days=90)).strftime('%Y-%m-%d')

# JQL query to find issues updated in the last 90 days
jql_query = f'project = PMZ AND issuetype = Features AND updatedDate >= "{ninety_days_ago}"'

# Initialize variables for pagination
start_at = 0
max_results = 50  # Adjust as needed

# Create lists to store data
issue_keys = []
summaries = []
statuses = []
old_due_dates = []
new_due_dates = []
components = []

# Retrieve issues and handle pagination
while True:
    issues = jira.search_issues(jql_query, expand='changelog', startAt=start_at, maxResults=max_results)

    if not issues:
        break

    for issue in issues:
        has_due_date_changes = False
        for history in issue.changelog.histories:
            for item in history.items:
                if item.field == 'duedate':
                    has_due_date_changes = True
                    old_due_date = item.fromString
                    new_due_date = item.toString
                    if old_due_date:
                        issue_keys.append(issue.key)
                        summaries.append(issue.fields.summary)
                        statuses.append(issue.fields.status.name)
                        old_due_dates.append(old_due_date[:10] if old_due_date else '')
                        new_due_dates.append(new_due_date[:10] if new_due_date else '')
                        components.append(', '.join([component.name for component in issue.fields.components]))
                    break  # Exit the inner loop once a due date change is found

        if has_due_date_changes:
            pass  # Continue to next issue

    start_at += max_results

# Create a DataFrame
data = {
    'Issue Key': issue_keys,
    'Summary': summaries,
    'Status': statuses,  # Add Status column
    'Old Due Date': old_due_dates,
    'New Due Date': new_due_dates,
    'Components': components
}
df = pd.DataFrame(data)

# Save DataFrame to Excel
excel_file = 'realest_due_date_changes.xlsx'
df.to_excel(excel_file, index=False)

print(f'Data saved to {excel_file}')
