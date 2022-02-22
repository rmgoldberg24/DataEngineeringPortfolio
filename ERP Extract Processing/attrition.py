import pandas as pd
import datetime, json

# Read file
df = pd.read_csv('attrition.csv')

# Define functions
def serviceBucket(x):
    if x < 2:
        return 'Less than 2 Years'
    elif x < 5:
        return '2 to 5 Years'
    elif x < 16:
        return '5 to 16 Years'
    else:
        return '16+ Years'

def activeWithinYear(row):
    if pd.isnull(row['TermDateCleaned']):
        return 1
    elif row['TermDateCleaned'] > oneYearAgo:
        return 1
    else:
        return 0

def formatPercentage(x):
    formatted = str(round(x * 100,2))+'%'
    return formatted

def createTabulations(columnToGroup):
    grouped = pd.DataFrame(df.groupby(columnToGroup, dropna=False)[['TermedWithinYear', 'ActiveWithinYear']].sum()).reset_index()
    grouped = grouped[grouped['ActiveWithinYear']>0]
    grouped['TermPercentage'] = grouped['TermedWithinYear']/grouped['ActiveWithinYear']
    grouped['TermPercentage'] = grouped['TermPercentage'].apply(formatPercentage)
    groupedTermType = pd.DataFrame(df.groupby(['TermType', columnToGroup], dropna=False)[['TermedWithinYear', 'ActiveWithinYear']].sum()).reset_index()
    groupedVoluntary = groupedTermType[groupedTermType['TermType'] == 'Voluntary'][[columnToGroup, 'TermedWithinYear']].rename({'TermedWithinYear': 'VoluntaryTerms'}, axis=1)
    final = pd.merge(grouped, groupedVoluntary, on=columnToGroup, how='left')
    final['VoluntaryTerms'] = final['VoluntaryTerms'].fillna(0)
    final['VoluntaryAttritionRate'] = final['VoluntaryTerms']/final['ActiveWithinYear']
    final['VoluntaryAttritionRate'] = final['VoluntaryAttritionRate'].apply(formatPercentage)
    return final

# Get current time
today = datetime.datetime.now()
oneYearAgo = today.replace(year=today.year-1)

# Open Dictionary for Scrubbing
with open('cleanMap.json') as f:
    cleanMap = json.load(f)

# Data Cleaning / Feature Engineering
df['CostCenterCleaned'] = df.apply(lambda x: cleanMap[x['SupervisoryOrg']] if x['SupervisoryOrg'] in cleanMap.keys() else x['CostCenter'], axis=1)
df['Vertical'] = df['CostCenterCleaned'].apply(lambda x: x[9:] if pd.notna(x) else '')
df['Management'] = df['ManagementLevel'].apply(lambda x: 'Staff' if x == '7 Individual Contributor' else 'Management')
df['TermDateCleaned'] = pd.to_datetime(df['TermDateForTermed'])
df['HireDateCleaned'] = pd.to_datetime(df['HireDate'])
df['TermYear'] = df['TermDateCleaned'].apply(lambda x: x.year)
df['TermMonth'] = df['TermDateCleaned'].apply(lambda x: x.month)
df['TermReason'] = df['Primary Termination Reason'].fillna('')
df['TermType'] = df['TermReason'].apply(lambda x: 'Involuntary' if 'Involuntary' in x else 'Voluntary')
df['RaceCleaned'] = df['Race'].apply(lambda x: str(x).split('(')[0])
df['RaceCleaned'] = df['RaceCleaned'].astype(str).replace({'nan': 'None Listed'})
df['LastEvaluation'] = df['LastEvaluation'].fillna('No Evaluation')
df['YearsOfService'] = round(df['MonthsOfService']/12,1)
df['ServiceBucket'] = df['YearsOfService'].apply(serviceBucket)
df['TermedWithinYear'] = df['TermDateCleaned'].apply(lambda x: 1 if x >= oneYearAgo else 0)
df['ActiveWithinYear'] = df.apply(activeWithinYear, axis=1)
df['TerminationReason'] = df['TermReason'].apply(lambda x: x.split('>')[-1])

# Calculate Overall Totals
totalTermed = df['TermedWithinYear'].sum()
totalActive = df['ActiveWithinYear'].sum()
totalTermedByType = df.groupby('TermType')['TermedWithinYear'].sum()
percentTermed = formatPercentage(df['TermedWithinYear'].sum()/df['ActiveWithinYear'].sum())

totalData = [
['Count Attrition Past 12 Months', totalTermed],
['Active Past 12 Months', totalActive],
['Percent Attrition Past 12 Months', percentTermed],
['Voluntary Attrition Rate', formatPercentage(totalTermedByType['Voluntary']/totalActive)]
]

total = pd.DataFrame(totalData)

# Calculate Tabulations
verticalGroup = createTabulations('Vertical')
managementGroup = createTabulations('Management')
staffGroup = createTabulations('EmployeeType')
seniorGroup = createTabulations('SeniorIndividualContributer')
genderGroup = createTabulations('Gender')
raceGroup = createTabulations('RaceCleaned')
evaluationGroup = createTabulations('LastEvaluation')
serviceGroup = createTabulations('ServiceBucket')
terminationReason = createTabulations('TerminationReason')

# Raw Data
rawData = df[df['TermedWithinYear']==1][['Employee ID', 'Full Name', 'Vertical', 'SupervisoryOrg', 'TermDate']]

# Write Calculations to Excel
tables = {'Total ITS Attrition': total,
'Attrition by Vertical': verticalGroup,
'Attrition by Management Level': managementGroup,
'Attrition by Staff Type': staffGroup,
'Attrition by Senior Status':seniorGroup,
'Attrition by Gender': genderGroup,
'Attrition by Race': raceGroup,
'Attrition by Last Evaluation': evaluationGroup,
'Attrition by Years of Service': serviceGroup,
'Attrition by Termination Reason': terminationReason,
'Terminated Workers': rawData}

timestamp = today.strftime('%m%d%Y')
filename = 'ITS_Attrition_{0}.xlsx'.format(timestamp)

writer = pd.ExcelWriter(filename, engine='xlsxwriter')
for sheetname, table in tables.items():  # loop through `dict` of dataframes
    if sheetname == 'Total ITS Attrition':
        table.to_excel(writer, sheet_name=sheetname, index=False, header=False)
    else:
        table.to_excel(writer, sheet_name=sheetname, index=False)  # send df to writer
    worksheet = writer.sheets[sheetname]  # pull worksheet object
    for idx, col in enumerate(table):  # loop through all columns
        series = table[col]
        max_len = max((
            series.astype(str).map(len).max(),  # len of largest item
            len(str(series.name))  # len of column name/header
            )) + 1  # adding a little extra space
        worksheet.set_column(idx, idx, max_len)  # set column width

writer.save()

print('Nice one -> all done! :)')
