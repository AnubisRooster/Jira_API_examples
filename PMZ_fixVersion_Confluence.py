from jira import JIRA
from atlassian import Confluence
from openpyxl import Workbook
from datetime import datetime, timedelta

# Jira Configuration
jira_url = 'https://memsql.atlassian.net'
jira_username = 'it-automation+jira_reporting@singlestore.com'
jira_api_token = 'ATATT3xFfGF0JgsZms4GLYlWArKOSeTOnIOZx4pz0YQ0WkLnN3rQg8777YpW-m39gI8oS463by7hv_boSRiTNW9pYPadtDnXlDZiSTvLtRns8Xrh4CUQLRQsdnluREK84CoedSP0PN2Gk5H58BXI5EszVdjb6ht-ozqEFTiEesQhD5aPgH1Vbc0=FECBFB84'
jira_project_key = 'PMZ'

# Confluence Configuration
confluence_url = 'https://memsql.atlassian.net'
confluence_username = 'it-automation+jira_reporting@singlestore.com'
confluence_api_token = 'ATATT3xFfGF0JgsZms4GLYlWArKOSeTOnIOZx4pz0YQ0WkLnN3rQg8777YpW-m39gI8oS463by7hv_boSRiTNW9pYPadtDnXlDZiSTvLtRns8Xrh4CUQLRQsdnluREK84CoedSP0PN2Gk5H58BXI5EszVdjb6ht-ozqEFTiEesQhD5aPgH1Vbc0=FECBFB84'
confluence_page_id = '2347892815'

# Initialize Jira
jira = JIRA(server=jira_url, basic_auth=(jira_username, jira_api_token))

# Query Jira for issues updated in the past 7 days
query = f'project = {jira_project_key} AND issuetype = Features AND fixVersion CHANGED AFTER "2023-08-10"'
issues = jira.search_issues(query, fields="summary,status,fixVersions,created,components")

# Calculate date seven days ago
seven_days_ago = ('August 10, 2023')

# Initialize a list to store extracted results
results = []

# Iterate through the fetched issues
for issue in issues:
    issue_key = issue.key
    summary = issue.fields.summary
    created_date = issue.fields.created[:10]

    # Fetch issue changelogs
    changelogs = jira.issue(issue_key, expand='changelog').changelog.histories

    # Initialize variables to store previous values
    prev_fix_versions = []

    # Extract change details
    for change in changelogs:
        for item in change.items:
            if item.field == 'Fix Version':
                # Check if the 'fromString' attribute exists in the item and it's not None
                if hasattr(item, 'fromString') and item.fromString is not None:
                    prev_fix_versions.append(item.fromString)

    # Get the most recent Fix Version value from the changelog
    most_recent_prev_fix_version = prev_fix_versions[-1] if prev_fix_versions else None

    # Get the current "fixVersions" associated with the Jira issue
    current_fix_versions = [fix_version.name for fix_version in issue.fields.fixVersions] if issue.fields.fixVersions else []

    # Check if there are changes to the fix version
    if most_recent_prev_fix_version is not None:
        # Add the relevant information to the results list
        results.append({
            'key': issue_key,
            'summary': summary,
            'status': issue.fields.status.name,
            'created_date': created_date,
            'previous_fix_versions': most_recent_prev_fix_version,
            'current_fix_versions': ', '.join(current_fix_versions),
        })

# Initialize Confluence
confluence = Confluence(
    url=confluence_url,
    username=confluence_username,
    password=confluence_api_token)

# Construct the content for the Confluence page
page_content = f"Updated Issue Fix Versions Since {seven_days_ago}\n\n"

# Add a table with headers
page_content += "|| Issue Key || Summary || Status || Created Date || Previous Fix Versions || Current Fix Versions ||\n"

# Add rows to the table with issue details
for result in results:
    row = (
        f"| {result['key']} | {result['summary']} | {result['status']} | "
        f"{result['created_date']} | {result['previous_fix_versions']} | {result['current_fix_versions']} |\n"
    )
    page_content += row

# Update the Confluence page
confluence.update_page(
    page_id=confluence_page_id,
    title=f"Updated Issue Fix Versions Since August 10, 2023",
    body=page_content,
    representation='wiki'
)
