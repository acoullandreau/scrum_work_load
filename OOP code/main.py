from classfile import JiraToolBox
import json

with open('conf.json', encoding='utf-8') as config_file:
    conf_data = json.load(config_file)

sprintmatrix = JiraToolBox(conf_data, 'load')
sprintmatrix.parseJiraIssues('JIRA.csv')
sprintmatrix.processJiraIssues()
sprintmatrix.parseExistingFile('Sprint_load.xlsx')
sprintmatrix.processIssueOwnership()
sprintmatrix.processAssigneesRemainingEstimate()
new_issues = sprintmatrix.processNewIssues()
sprintmatrix.addNewIssues(new_issues)
sprintmatrix.updateExistingIssues()
sprintmatrix.updateTotalsPerAssignee()
sprintmatrix.highlightRemovedIssues()
sprintmatrix.writeUpdatedFile('Sprint_load_updated.xlsx')