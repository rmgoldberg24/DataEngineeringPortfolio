import requests, json, os, datetime
import pandas as pd

# Get Odata from PMO Application
user = ''
password = ''
url = "https://odata.ppmpro.com/InnotasOdataService/Fields('DASHBOARD@INSTANCE')"
response = requests.get(url, auth=(user, password))

# Load JSON data
data = json.loads(response.text)

# Convert to dataframe
df = pd.json_normalize(data, 'value')

# Parse to match BI tool input
Resource = df[df['EntityDescription'] == 'Resource']
Work = df[df['EntityDescription'] == 'Work']
Task = df[df['EntityDescription'] == 'Task']
Time = df[df['EntityDescription'] == 'Time']

Resource = Resource[["ID", "Name\nLast\nFirst", "First\nName", "Middle\nName",
"Last\nName", "Email", "Primary\nRole", "Role\nGroup", "Secondary\nRoles", "Type",
"Immediate\nSupervisor", "(CF)\nITS\nVertical", "Capacity\nStart\nDate",
"Termination\nDate", "Is\nNOT\nTerminated?", "Unit"]]

Resource = Resource.drop_duplicates()

Work = Work[["ID", "Type", "USC\nProject\nID", "Title", "Start\nDate", "Target\nDate",
"Completion\nDate", "Status", "Program"]]

Work = Work.drop_duplicates()

Task = Task[["Start\nDate", "Target\nDate", "Completion\nDate", "Status",
"Program", "Work\nID", "Task\nID", "Task\nTitle"]]

Task = Task.drop_duplicates()

Time = Time[[ "Task\nID", "Task\nTitle", "Week", "Date", "Resource\nID",
"Resource\nName", "Role", "Project\nID", "Project\nTitle", "Total\nHours",
"Time\nType", "Timesheet\nEntry\n(CF)\nWork\nCategory\nor\nAdmin",
"Entry\nApproval\nState", "Entry\nNotes", "Timesheet\nApproval\nState",
"Timesheet\nNote", "Timesheet\nStart\nDate", "Timesheet\nEnd\nDate",
"Timesheet\nTotal\nhours"]]

timestamp = datetime.datetime.now().strftime("%m%d%Y")

# Archive existing files
os.rename('Resource.xlsx','Archive/Resource' + timestamp + '.xlsx')
os.rename('Work.xlsx','Archive/Work' + timestamp + '.xlsx')
os.rename('Task.xlsx','Archive/Task' + timestamp + '.xlsx')
os.rename('Time.xlsx','Archive/Time' + timestamp + '.xlsx')

# Save new files
Resource.to_excel('Resource.xlsx', index=False, sheet_name='Resource')
Work.to_excel('Work.xlsx', index=False, sheet_name='Work')
Task.to_excel('Task.xlsx', index=False, sheet_name='Task')
Time.to_excel('Time.xlsx', index=False, sheet_name='Time')

print('Files completed!')
