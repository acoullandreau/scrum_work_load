=================================
Scrum sprint load estimate script
=================================


---------------------
Purpose of the script
---------------------

This script intends to create or update the base of a sprint load estimate.
The sprint load estimate is divided in two sections:
* the list of issues and the remaining time estimate for each person
* the summary with pre-defined formulaes to estimate the load of each person (percentage of the time spent on the issues listed)

The first time the file is created, it is possible to use as a source file the template files. Once the file has been updated (for example availability of the people, or notes taken in the file), this new file can be used as a source file, and will be updated by the script.


-------------------------------------------
Structure of the script and the input files
-------------------------------------------

The script is written in Python 3. Three libraries are used to parse the configuration, input and output files: csv, json and openpyxl (v. 2.6.0).

The configuration file (a json file) is used to "map" a username in JIRA with the name of the person we want to display in the output file.

There are two input files:
* the 'JIRA.csv' file, exported from JIRA (all the issues of the current sprint)
* the 'Sprint_load.xslx', that can either be the template (first creation) or a version already updated of the load estimate file

No arguments are used to launch the script.


-----------------------------
Structure of the output files
-----------------------------

There is one output file: 'Sprint_load_update.xslx' that contains a copy of the input file and the update from the csv file.

The sprint load estimate contains only "standard issues", i.e no subtasks. The script maps the subtasks to their associated standard issue, and reports all the information of all issues and subtasks for the standard issue.

The script loops through the input 'Sprint_load.xslx' file, and if a key is found the line is updated. The position of the line in the file remains as-is.

If there are some keys in the 'JIRA.csv' file not present in the input matrix, new lines are added at the end of the file.

The total sum of assigned time for each person is computed by the script. The other formulas (percentage in particular) remain unaffected by the script.

In case some issues were removed from the sprint in between two updates of the file, the script highlights the line of the removed story (the removal is not done automatically to allow tracking)


------------
Step by step
------------

1. Verify that the configuration file contains the list of possible assignees included in the sprint matrix
2. Download the sprint report file from JIRA and add it to the same folder from where the script is run (the file must be names 'JIRA.csv')
3. Add either a template (to be renamed 'Sprint_load.xslx') or add an existing file with the name 'Sprint_load.xslx'
4. Execute the script


----------------------
Initial files provided
----------------------

- readme.rst
- run.bat
- main.py
- conf.json
- Sprint_load_template_CMS.xlsx
- Sprint_load_template_FW.xlsx
