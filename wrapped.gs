function getEpicsFromJira() {
    // Jira credentials and URL
    var jiraDomain = 'memsql.atlassian.net';  // Jira domain
    var jiraEmail = 'mfink@singlestore.com'; // Your Jira email
    var jiraToken = 'ATATT3xFfGF0ZYV987OGmCENMnHBz1tD_DHylzKuCgFiOV1fSdULdiLheg2Pmr14ucy1giqLP3AQHfIWHxPDH20QnrkMCt0VYgNBhxmodgoDaRGs8r-iHs2srBmTJBDZq2VR17sUKNacDZ-VHkjTSf74cw0E1Oq588Y4YAjF17DhSaYQL1HianU=C24D52E4';         // 
    var projectKeys = ['PSY', 'INFRA', 'DEVPLAT','DATA']; // List of project keys

    // Jira API authorization
    var headers = {
        "Authorization": "Basic " + Utilities.base64Encode(jiraEmail + ":" + jiraToken),
        "Content-Type": "application/json"
    };

    // Header row for consolidated epics data
    var consolidatedEpics = [["Summary", "Team", "Sub-Team", "Project", "Assignee", "Story Point Estimate", "Fix Version", "Epic Key (URL)", "Due Date", "Status"]]; 

    // Loop through each project and fetch epics
    projectKeys.forEach(function(projectKey) {
        var startAt = 0;
        var maxResults = 50;  // Maximum number of epics Jira returns per request

        // Loop through Jira API pages until all epics are retrieved
        while (true) {
            // Jira API URL for searching epics using JQL with pagination and filtering by Fix Version
            var url = `https://${jiraDomain}/rest/api/3/search?jql=project=${projectKey} AND issuetype=Epic AND fixVersion IS NOT EMPTY&startAt=${startAt}&maxResults=${maxResults}`;

            // Send GET request to Jira
            var response = UrlFetchApp.fetch(url, { "method": "GET", "headers": headers });
            var data = JSON.parse(response.getContentText());

            // Append the fetched issues to the consolidatedEpics array
            data.issues.forEach(function(issue) {
                var summary = issue.fields.summary;
                var status = issue.fields.status.name;
                var dueDate = issue.fields.duedate ? issue.fields.duedate : "N/A"; // Handle missing due date
                var assignee = issue.fields.assignee ? issue.fields.assignee.displayName : "Unassigned"; // Handle missing assignee
                var storyPoints = issue.fields.customfield_12705 ? issue.fields.customfield_12705 : "N/A"; // Assuming Story Points is in customfield_12705
                var fixVersion = issue.fields.fixVersions && issue.fields.fixVersions.length > 0 ? issue.fields.fixVersions[0].name : "N/A"; // Take first Fix Version if available
                var subTeam = issue.fields.customfield_10900.name || "N/A"; // Retrieve Sub-Team from customfield_10900

                // Create the clickable URL for the Epic Key
                var epicUrl = `https://${jiraDomain}/browse/${issue.key}`;
                var clickableKey = `=HYPERLINK("${epicUrl}", "${issue.key}")`;

                // Add data to consolidated array with "Team" as "Infrastructure" and "Sub-Team" from customfield_10900
                consolidatedEpics.push([summary, "Infrastructure", subTeam, projectKey, assignee, storyPoints, fixVersion, clickableKey, dueDate, status]);
            });

            // If the total results returned are less than the maxResults, it means we've retrieved all pages
            if (data.issues.length < maxResults) {
                break;
            }

            // Otherwise, update startAt to fetch the next page
            startAt += maxResults;
        }
    });

    // Sort epics alphabetically by Summary (first column in consolidatedEpics)
    consolidatedEpics = consolidatedEpics.slice(1).sort(function(a, b) {
        return a[0].localeCompare(b[0]);  // Sort based on the Summary column (index 0)
    });
    // Add back the header row
    consolidatedEpics.unshift(["Summary", "Team", "Sub-Team", "Project", "Assignee", "Story Point Estimate", "Fix Version", "Epic Key (URL)", "Due Date", "Status"]);

    // Update the "Update" sheet
    updateConsolidatedSheet(consolidatedEpics);
}

function updateConsolidatedSheet(epicData) {
    // Get the active spreadsheet
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "Update"; // Name of the consolidated sheet

    // Check if the "Update" sheet exists; if not, create it
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
        sheet = spreadsheet.insertSheet(sheetName); // Create a new sheet
    } else {
        sheet.clear(); // Clear the sheet if it already exists
    }

    // Write the epic data starting from row 3, column 1
    sheet.getRange(3, 1, epicData.length, epicData[0].length).setValues(epicData);

    // Apply conditional formatting to the Status column (10th column, "Status")
    applyConditionalFormatting(sheet, 10, epicData.length + 2);
}

function applyConditionalFormatting(sheet, statusColumn, totalRows) {
    // Define status colors mapping
    var rules = sheet.getConditionalFormatRules();
    
    var newRules = [
        {
            status: 'In Progress',
            color: '#00FF00'  // Green
        },
        {
            status: 'Backlog',
            color: '#FFFF00'  // Yellow
        },
        {
            status: 'To Do',
            color: '#FFFF00'  // Yellow
        },
        {
            status: 'On Hold',
            color: '#FFA500'  // Orange
        },
        {
            status: 'In Review',
            color: '#00FF00'  // Green
        },
        {
            status: 'Done',
            color: '#0000FF'  // Blue
        }
    ];
    
    newRules.forEach(function(rule) {
        var formatRule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo(rule.status)
            .setBackground(rule.color)
            .setRanges([sheet.getRange(3, statusColumn, totalRows - 2)]) // Skip header rows
            .build();
        rules.push(formatRule);
    });
    
    // Apply the new rules
    sheet.setConditionalFormatRules(rules);
}
