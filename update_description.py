import requests
import base64

# Set up Jira API credentials and base URL
username = '[email@address.com]'
api_token = '[auth_token]'
base_url = 'https://[domain].atlassian.net'

# Encode credentials as a base64 string
credentials = base64.b64encode(f'{username}:{api_token}'.encode('utf-8')).decode('utf-8')
headers = {
    'Authorization': f'Basic {credentials}',
    'Content-Type': 'application/json'
}

# Define JQL query to find CLOSED issues from the most recent version release
jql_query = 'project = [project] AND status in (Closed, Done) AND fixVersion = "Test Release 0.1" ORDER BY resolutiondate DESC'

# Get issue keys of matching issues using JQL search
search_url = f'{base_url}/rest/api/2/search'
payload = {
    'jql': jql_query,
    'maxResults': 1000,
    'fields': 'key'
}
response = requests.get(search_url, headers=headers, params=payload)
search_results = response.json()
issue_keys = [issue['key'] for issue in search_results['issues']]

# Define new value for the custom field
new_description_value = 'We released this thing.'

# Update custom field for each issue
for issue_key in issue_keys:
    issue_url = f'{base_url}/rest/api/2/issue/{issue_key}'
    response = requests.get(issue_url, headers=headers)
    issue_data = response.json()

    issue_data['fields']['customfield_123'] = new_custom_field_value
    update_url = f'{base_url}/rest/api/2/issue/{issue_key}'
    response = requests.put(update_url, headers=headers, json=issue_data)

    if response.status_code == 204:
        print(f'Custom field updated successfully for issue {issue_key}')
    else:
        print(f'Failed to update custom field for issue {issue_key}. Status code: {response.status_code}')

