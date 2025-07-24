import os
import sys
from jira import JIRA
import matplotlib.pyplot as plt
from configparser import ConfigParser
from atlassian import Confluence

# Jira credentials
JIRA_SERVER = os.getenv("JIRA_SERVER", 'https://[domain].atlassian.net')
CONFLUENCE_SERVER = os.getenv("CONFLUENCE_SERVER", 'https://[domain].atlassian.net')
DRY_RUN = True if os.getenv("DRY_RUN", "false") == "true" else False

# Load credentials from config file - if this file exists, otherwise use dev-jira-auth.sh from Francisco
config = ConfigParser()
config.read('config.ini')

JIRA_USERNAME = config.get('Jira', 'username')
JIRA_PASSWORD = config.get('Jira', 'password')
JIRA_BOARD_ID = config.get('Board', 'id')

try:
    # Connect to Jira
    jira = JIRA(server=JIRA_SERVER, basic_auth=(JIRA_USERNAME, JIRA_PASSWORD))

    closed_sprints = [s for s in jira.sprints(JIRA_BOARD_ID) if s.state == 'closed']

    if not closed_sprints:
        print("No closed sprints found for the scrum board.")
        sys.exit(1)
    else:
        # Find the last closed sprint
        sprint = max(closed_sprints, key=lambda s: s.id)
        sprint_id = sprint.id
        print(f"Last closed sprint ID: {sprint_id} {sprint.name}")

    # Get the sprint issues
    sprint_issues = jira.search_issues(
        f'Sprint = {sprint.id}',
        maxResults=False,
    )

    for issue in sprint_issues:
        print(f"Issue {issue.key}")

    # Prepare the table content
    table_rows = []
    assignee_counts = {}
    for issue in sprint_issues:
        table_row = "| {key} | {summary} | {status} | {assignee} | {parent_summary} |\n".format(
            key=issue.key,
            summary=issue.fields.summary,
            status=issue.fields.status.name,
            assignee=issue.fields.assignee.displayName if issue.fields.assignee else "Unassigned",
            parent_summary=parent_summary,
        )
        table_rows.append(table_row)
        if issue.fields.assignee:
            assignee = issue.fields.assignee.displayName
            if assignee in assignee_counts:
                assignee_counts[assignee] += 1
            else:
                assignee_counts[assignee] = 1
        else:
            if "Unassigned" in assignee_counts:
                assignee_counts["Unassigned"] += 1
            else:
                assignee_counts["Unassigned"] = 1
        
        if parent_summary:
            if parent_summary in parent_task_summaries:
                parent_task_summaries[parent_summary] += 1
            else:
                parent_task_summaries[parent_summary] = 1

    # Create the bar chart for issue status
    status_counts = {}
    for issue in sprint_issues:
        status = issue.fields.status.name
        if status in status_counts:
            status_counts[status] += 1
        else:
            status_counts[status] = 1

    labels = list(status_counts.keys())
    values = list(status_counts.values())

    plt.figure(figsize=(6, 4))
    plt.bar(labels, values)
    plt.xlabel('Status')
    plt.ylabel('Count')
    plt.title(f'Sprint {sprint.name} Issue Status')
    plt.xticks(rotation=45)
    plt.tight_layout()

    # Save the bar chart as an image
    chart_file = "status_bar_chart.png"
    plt.savefig(chart_file)
    plt.close()

    print(f"Status bar chart saved successfully: {chart_file}")

    # Create the pie chart for issues by assignee
    assignee_labels = list(assignee_counts.keys())
    assignee_values = list(assignee_counts.values())

    plt.figure(figsize=(6, 4))
    plt.pie(assignee_values, labels=assignee_labels, autopct='%1.1f%%')
    plt.title(f'Sprint {sprint.name} Issues by Assignee')
    plt.tight_layout()

    # Save the pie chart as an image
    assignee_chart_file = "assignee_pie_chart.png"
    plt.savefig(assignee_chart_file)
    plt.close()

    print(f"Assignee pie chart saved successfully: {assignee_chart_file}")

    # Create the Confluence page content
    page_title = f'Sprint {sprint.name} Summary'
    page_content = f"""
    h2. Sprint Summary - {sprint.name}

    ||Issue Key||Summary||Status||Assignee||Epic||
    |{"".join(table_rows)}|

    h3. Bar Chart: Sprint {sprint.name} Issue Status
    !status_bar_chart.png!

    h3. Pie Chart: Issues by Assignee
    !assignee_pie_chart.png!
    """

    print(page_content)

    if DRY_RUN:
        print("Skipping page creation DRY_RUN=true")
    else:
        # Connect to Confluence
        CONFLUENCE_USERNAME = config.get('Confluence', 'username')
        CONFLUENCE_PASSWORD = config.get('Confluence', 'password')
        CONFLUENCE_SPACE_KEY = config.get('Confluence', 'space_key')
        CONFLUENCE_PARENT_PAGE_ID = config.get('Confluence', 'parent_page_id')

        confluence = Confluence(
            url=CONFLUENCE_SERVER,
            username=CONFLUENCE_USERNAME,
            password=CONFLUENCE_PASSWORD,
        )

        # Create the page in Confluence
        confluence.create_page(
            space=CONFLUENCE_SPACE_KEY,
            title=page_title,
            body=page_content,
            parent_id=CONFLUENCE_PARENT_PAGE_ID,
            representation='wiki',
        )

        print("Page created successfully in Confluence.")

except Exception as e:
    print(f"Error: {e}")
    if hasattr(e, "response") and hasattr(e.response, "text"):
        print(f"Response content: {e.response.text}")
