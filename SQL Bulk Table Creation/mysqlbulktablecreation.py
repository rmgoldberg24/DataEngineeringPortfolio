import pandas as pd
from sqlalchemy import create_engine
import os

# Get all files that will be turned into tables
dir = os.getcwd() + '/Files'
filelist = []
for file in os.listdir(dir):
    if file.endswith(".csv"):
        filelist.append(file)
        print(file)

filelist.sort()

# Read list of tablenames
with open('TableNameList.txt', 'r') as file:
    tablenamelist = file.read().splitlines()

# Zip table names with files
tablelist = dict(zip(tablenamelist, filelist))

# Connect to MySQL
db_connection_str = 'mysql+pymysql://USER@IP/SCHEMA'
db_connection = create_engine(db_connection_str)

# Create table from each file
for name in tablelist:
    file = 'Files/' + tablelist[name]
    print('Reading file {0} . . . '.format(file))
    df = pd.read_csv(file)
    print('Creating table {0} . . . '.format(name))
    df.to_sql(name, db_connection, if_exists='replace', index = False)
