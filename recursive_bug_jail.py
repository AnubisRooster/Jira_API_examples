import requests

# Jira API endpoint
api_url = "https://memsql.atlassian.net/rest/api/2"


# Jira project key
project_key = "INFRA"

# Jira query to find assignees with more than nine tasks
jql_query = f'project = {project_key} AND assignee is <assignee> GROUP BY assignee HAVING count(assignee) > 1'

# Jira authentication credentials
username = "it-automation+jira_reporting@singlestore.com"
api_token = ""

# Email configuration
sender_email = "jira@memsql.atlassian.net"
email_subject = "Bug Jail Notification"
email_body_template = """
Dear {assignee_name},

You have more than one bugs assigned to you. You are currently in Bug Jail until you resolve the following:

{task_summaries}

Please review your workload and take necessary actions.

Best regards,
Your Project Team
"""

def send_email(recipient, subject, body):
    # Use your preferred method or library to send emails
    # Example: Using the built-in `smtplib` library
    import smtplib
    from email.mime.text import MIMEText

    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = recipient

    smtp_server = "smtp.example.com"
    smtp_port = 587
    smtp_username = "smtp_username"
    smtp_password = "smtp_password"

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.send_message(msg)

# Function to get Jira issues based on JQL query
def get_issues(jql):
    url = f"{api_url}/search"
    headers = {
        "Content-Type": "application/json"
    }
    auth = (username, api_token)
    payload = {
        "jql": jql,
        "maxResults": 1000
    }

    response = requests.post(url, headers=headers, auth=auth, json=payload)

    if response.status_code == 200:
        return response.json()["issues"]
    else:
        print(f"Error: {response.status_code} - {response.text}")
        return []

# Get assignees with more than nine tasks
issues = get_issues(jql_query)

for issue in issues:
    assignee = issue["fields"]["assignee"]
    assignee_name = assignee["displayName"]
    assignee_email = assignee["emailAddress"]

    # Get task summaries for the assignee
    jql_query = f'assignee = "{assignee_name}"'
    assigned_tasks = get_issues(jql_query)

    task_summaries = "\n".join([task["fields"]["summary"] for task in assigned_tasks])

    # Prepare email body
    email_body = email_body_template.format(assignee_name=assignee_name, task_summaries=task_summaries)

    # Send email
    send_email(assignee_email, email_subject, email_body)

print("Emails sent successfully!")
