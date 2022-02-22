import pandas as pd
import datetime
import re

# Read in Workday Extract
df = pd.read_csv('headcount.csv')

# Define Row Functions
def activeWithinTimeframe(row):
    timeframe = datetime.datetime(2021, 1, 1)
    if pd.isnull(row['TermDate']):
        return 1
    elif row['TermDateCleaned'] >= timeframe:
        return 1
    else:
        return 0

# Convert to datetime and clean (add today if no term date)
df['HireDate'] = pd.to_datetime(df['HireDate'])
df['TermDate'] = pd.to_datetime(df['TermDate'])
df['TermDateCleaned'] = df['TermDate'].fillna(datetime.datetime.today())

# Create active within timeframe column
df['activeWithinTimeframe'] = df.apply(activeWithinTimeframe, axis=1)

# Fitler to only include active personnel in 2021
active = df[df['activeWithinTimeframe'] == 1].copy()

# Get a list of months active between hire date and term date/current date
def getMonths(row):
    datelist = pd.date_range(row['HireDate'].strftime('%Y-%m'), row['TermDateCleaned'].strftime('%Y-%m'), freq='MS').strftime('%Y-%m').tolist()
    monthlist = [x for x in datelist if re.search(r'^2021|^2022', x)]
    return monthlist

# Save month active list to column
active['MonthsActive'] = active.apply(getMonths, axis=1)

# Set start and end date for complete month list and generate
startdate = datetime.datetime(2021, 1, 1)
enddate = datetime.datetime(2022, 12, 31)
completeMonthList = pd.date_range(startdate.strftime('%Y-%m'), enddate.strftime('%Y-%m'), freq='MS').strftime('%Y-%m').tolist()

# Create a column for each month, mark 1 if active in month, else 0
for month in completeMonthList:
    column = 'Active' + str(month)
    active[column] = active['MonthsActive'].apply(lambda x: 1 if month in x else 0)

# Set which columns correspond to each quarter
quarterdict = {'2021Q1': ['Active2021-01', 'Active2021-02', 'Active2021-03'],
 '2021Q2': ['Active2021-04', 'Active2021-05', 'Active2021-06'],
 '2021Q3': ['Active2021-07', 'Active2021-08', 'Active2021-09'],
 '2021Q4': ['Active2021-10', 'Active2021-11', 'Active2021-12'],
 '2022Q1': ['Active2022-01', 'Active2022-02', 'Active2022-03'],
 '2022Q2': ['Active2022-04', 'Active2022-05', 'Active2022-06'],
 '2022Q3': ['Active2022-07', 'Active2022-08', 'Active2022-09'],
 '2022Q4': ['Active2022-10', 'Active2022-11', 'Active2022-12']}

# Create a column for each quarter, mark 1 if active in quarter, else 0
for quarter, month in quarterdict.items():
    columnname = 'Active' + quarter
    active[columnname] = active.apply(lambda row: row[month].max(), axis=1)
    print(quarter, month)

# Filter to only include our calculated "active" columns
columns = list(active.columns)
r = re.compile('^Active')
includecolumns = list(filter(r.match, columns))

# Do active tabulations and save to flatfile
final = active[includecolumns].sum()
final.to_csv('QuarterlyMonthlyHeadcount.csv', header=None)
