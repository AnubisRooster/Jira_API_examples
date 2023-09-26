from datetime import datetime, timedelta
from jira import JIRA
import pandas as pd

# Jira credentials and connection
jira = JIRA(server='https://memsql.atlassian.net', basic_auth=('it-automation+jira_reporting@singlestore.com', 'ATATT3xFfGF0JgsZms4GLYlWArKOSeTOnIOZx4pz0YQ0WkLnN3rQg8777YpW-m39gI8oS463by7hv_boSRiTNW9pYPadtDnXlDZiSTvLtRns8Xrh4CUQLRQsdnluREK84CoedSP0PN2Gk5H58BXI5EszVdjb6ht-ozqEFTiEesQhD5aPgH1Vbc0=FECBFB84'))

# JQL query for issues by Assignee within the last six months
jql_query = 'project in (INFRA, PSY) AND issuetype in (Bug, Story, Task, Sub-task, Improvement, Investigation, "New Feature", Request) AND Sprint in closedSprints() AND updated > startOfDay(-90d)'

# Calculate the date range
end_date = datetime.now().date()
start_date = end_date - timedelta(days=7)

# Retrieve the updated tasks
issues = jira.search_issues(jql_query, maxResults=False)

# Prepare the data for the CSV file
data = []
for issue in issues:
    assignee = issue.fields.assignee.displayName if issue.fields.assignee else "Unassigned"
    summary = issue.fields.summary
    issue_key = issue.key
    status = issue.fields.status.name
    parent_issue_summary = issue.fields.parent.fields.summary if hasattr(issue.fields,
                                                                         'parent') and issue.fields.parent else ""

    # Extract sprint information
    sprints = [sprint.name for sprint in issue.fields.customfield_10021] if hasattr(issue.fields,
                                                                                    'customfield_10021') else []
    sprint_value = ', '.join(sprints)

    data.append([assignee, summary, issue_key, status, parent_issue_summary, sprint_value])

# Create a DataFrame from the data
df = pd.DataFrame(data, columns=["Assignee", "Summary", "Issue Key", "Status", "Parent Issue Summary", "Sprint"])

# Sort the DataFrame by assignee
df = df.sort_values(by="Assignee")

# Get the current date for the output file name
current_date = datetime.now().strftime("%Y-%m-%d")

# Set the output file name based on the current date with .csv extension
output_file = f"sprint_export_{current_date}.csv"

# Export the DataFrame to a CSV file
df.to_csv(output_file, index=False)

print(f"Task data exported to {output_file} successfully.")
