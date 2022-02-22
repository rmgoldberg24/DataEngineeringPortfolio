import pandas as pd
import numpy as np

# Read qualtrics extract
df = pd.read_csv('EaseOfUse.csv', skiprows=[1,2])

# Functoin to collapse columns (different conditions have different question items,
# need to consolidate into single column for analysis)
def collapse_columns(row, item1, item2, item3):
    if pd.notnull(row[item1]):
        return row[item1]
    elif pd.notnull(row[item2]):
        return row[item2]
    elif pd.notnull(row[item3]):
        return row[item3]
    else:
        return np.NaN

# Open list that maps which questions should be collapsed into one
# then performs list collapsing
with open('ItemList.csv', 'r', encoding='utf-8-sig') as file:
    for line in file:
        items = line.strip().split(',')
        print('Collapsing ', items, ' . . .')
        df[items[0]] = df.apply(collapse_columns,
            args=(items), axis = 1)

# Remove old columns that were collapsed (no longer needed)
df.drop(df.iloc[:, 55:], inplace = True, axis = 1)

# Make a copy of df to use once tabulations are completed (need to exclude
# 6 values from tabulations, but want to retain in raw data)
df_copy = df.copy()

# Removing 6 for tabulations
df = df.replace(6, np.NaN)

# Calculate different construct means by question number
def content_mean(row):
    group_mean = np.nanmean([row['Q3_1'],
                           row['Q3_2'],
                           row['Q3_3'],
                           row['Q3_4'],
                           row['Q3_5'],
                           row['Q3_6'],
                           row['Q3_7'],
                           row['Q3_8'],
                           row['Q3_9']])
    return group_mean

def exams_mean(row):
    group_mean = np.nanmean([row['Q4_1'],
                           row['Q4_2'],
                           row['Q4_3'],
                           row['Q4_4']])
    return group_mean

def admin_mean(row):
    group_mean = np.nanmean([row['Q5_1'],
                           row['Q5_2'],
                           row['Q5_3'],
                           row['Q5_4'],
                           row['Q5_5']])
    return group_mean

def communication_mean(row):
    group_mean = np.nanmean([row['Q6_1'],
                           row['Q6_2'],
                           row['Q6_3'],
                           row['Q6_4'],
                           row['Q6_5']])
    return group_mean

def grades_mean(row):
    group_mean = np.nanmean([row['Q7_1'],
                           row['Q7_2'],
                           row['Q7_3'],
                           row['Q7_4'],
                           row['Q7_5'],
                           row['Q7_6'],
                           row['Q7_7'],
                           row['Q7_8'],
                           row['Q7_9'],
                           row['Q7_10']])
    return group_mean

# Perform calculations and save as columns
df['Content'] = df.apply(content_mean, axis = 1)
df['Exams'] = df.apply(exams_mean, axis = 1)
df['Administrative'] = df.apply(admin_mean, axis = 1)
df['Communication'] = df.apply(communication_mean, axis = 1)
df['Grades'] = df.apply(grades_mean, axis = 1)

df = df.replace({'Q2': {1: 'Blackboard', 2: 'Canvas', 3: 'D2L Brightspace' }})

# Removed scrubbed columns for raw data and replace with untouched columns
df = df.drop(columns=['Q3_1', 'Q3_2', 'Q3_3',
       'Q3_4', 'Q3_5', 'Q3_6', 'Q3_7', 'Q3_8', 'Q3_9', 'Q4_1', 'Q4_2', 'Q4_3',
       'Q4_4', 'Q5_1', 'Q5_2', 'Q5_3', 'Q5_4', 'Q5_5', 'Q6_1', 'Q6_2', 'Q6_3',
       'Q6_4', 'Q6_5', 'Q7_1', 'Q7_2', 'Q7_3', 'Q7_4', 'Q7_5', 'Q7_6', 'Q7_7',
       'Q7_8', 'Q7_9', 'Q7_10', 'Q8_1', 'Q9_1', 'Q9_2', 'Q10'])

df_columns = df_copy[['Q3_1', 'Q3_2', 'Q3_3',
       'Q3_4', 'Q3_5', 'Q3_6', 'Q3_7', 'Q3_8', 'Q3_9', 'Q4_1', 'Q4_2', 'Q4_3',
       'Q4_4', 'Q5_1', 'Q5_2', 'Q5_3', 'Q5_4', 'Q5_5', 'Q6_1', 'Q6_2', 'Q6_3',
       'Q6_4', 'Q6_5', 'Q7_1', 'Q7_2', 'Q7_3', 'Q7_4', 'Q7_5', 'Q7_6', 'Q7_7',
       'Q7_8', 'Q7_9', 'Q7_10', 'Q8_1', 'Q9_1', 'Q9_2', 'Q10']]


final_df = pd.merge(df, df_columns, left_index=True, right_index=True)

# Save as flatfile
final_df.to_excel('Ease of Use Cleaned.xlsx', index=False)
