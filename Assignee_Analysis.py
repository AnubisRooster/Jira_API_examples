import pandas as pd
from jira import JIRA
from pytz import timezone
import matplotlib.pyplot as plt
import numpy as np

# Jira credentials and connection
jira = JIRA(server='[jira_domain]', basic_auth=('[emailaddress]', '[auth_token]'))

# JQL query for retrieving incidents involving PagerDuty within the last thirty days
jql_query = 'project in ([projects]) AND issuetype in (Bug, Story, Task, Sub-task, Improvement, Investigation, "New Feature", Request) AND status = Done AND created > -180d AND assignee != EMPTY'

# Retrieve issues based on the JQL query
issues = jira.search_issues(jql_query, maxResults=False, expand='changelog')
print(f"Number of issues retrieved: {len(issues)}")

# Extract necessary information and store in a DataFrame
data = []
for i, issue in enumerate(issues, 1):
    print(f"Processing issue {i}/{len(issues)}")
    incident_id = issue.key
    created_datetime = pd.to_datetime(issue.fields.created)

    # Check if the issue has transitioned to "In Progress"
    time_spent_in_progress = None
    in_progress_datetime = None
    for history in issue.changelog.histories:
        for item in history.items:
            if item.field == 'status' and item.toString == 'In Progress':
                in_progress_datetime = pd.to_datetime(history.created)
                break
        if in_progress_datetime:
            break

    # Check if the issue is closed
    if issue.fields.resolutiondate:
        closed_datetime = pd.to_datetime(issue.fields.resolutiondate)

        # Exclude weekends from the time in progress calculation
        if in_progress_datetime:
            start_date = in_progress_datetime.tz_convert(timezone('UTC'))
            end_date = closed_datetime.tz_convert(timezone('UTC'))
            business_days_in_progress = pd.bdate_range(start_date, end_date)
            if len(business_days_in_progress) > 0:
                time_spent_in_progress = np.busday_count(business_days_in_progress[0].date(), business_days_in_progress[-1].date())

        # Exclude weekends from the time to close calculation
        start_date = created_datetime.tz_convert(timezone('UTC'))
        end_date = closed_datetime.tz_convert(timezone('UTC'))
        business_days_to_close = pd.bdate_range(start_date, end_date)
        if len(business_days_to_close) > 0:
            time_to_close = np.busday_count(business_days_to_close[0].date(), business_days_to_close[-1].date())
        else:
            time_to_close = None
    else:
        time_to_close = None

    assignee = issue.fields.assignee.displayName if issue.fields.assignee else 'Unassigned'

    data.append([incident_id, created_datetime, time_spent_in_progress, time_to_close, assignee])

df = pd.DataFrame(data, columns=['Incident ID', 'Created Date', 'Time Spent in Progress (business days)', 'Time to Close (business days)', 'Assignee'])

# Calculate average time spent in progress and average time to close per assignee
assignee_metrics = df.groupby('Assignee').agg({'Time Spent in Progress (business days)': 'mean', 'Time to Close (business days)': 'mean', 'Incident ID': 'count'})
assignee_metrics.reset_index(inplace=True)
assignee_metrics.columns = ['Assignee', 'Average Time Spent in Progress (business days)', 'Average Time to Close (business days)', 'Total Issues Assigned']

# Export the table to an Excel file
assignee_metrics.to_excel('assignee_metrics.xlsx', index=False)

print("Process completed successfully!")
