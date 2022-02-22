import cx_Oracle, os
import pandas as pd
import sys
from sklearn.model_selection import StratifiedShuffleSplit

# Find oracle client and explicity load
lib_dir=''
try:
    cx_Oracle.init_oracle_client(lib_dir=lib_dir)
except Exception as err:
    print("Error connecting: cx_Oracle.init_oracle_client()")
    print(err)
    sys.exit(1)

# Connect to oracle
ip = ''
port = 1521
SID = ''
dsn_tns = cx_Oracle.makedsn(ip, port, SID)

username = ''
password = ''

connection = cx_Oracle.connect(username, password, dsn_tns)

# Take current term as terminal argument
current_term = sys.argv[1]

# Oracle query to get list of currents students
query = """
SELECT
student.FIRST_NAME,
student.LAST_NAME,
email.EADDRESS,
class.ARR_CLASS_C
FROM USC_VIEW.A_STUDENT_00 student
LEFT JOIN
USC_VIEW.A_STUDENT_EMAIL_00 email
ON
student.PERSON_ID = email.PERSON_ID
LEFT JOIN
USC_VIEW.a_arr_00 class
ON
student.PERSON_ID = class.PERSON_ID
WHERE
student.PERSON_ID IN
(
SELECT
PERSON_ID
FROM
usc_view.v_myusc_widget_registration
WHERE TERM_CODE = {0}
)
AND
email.SOURCE_CODE = 'SIS_USC_EMAIL'
AND
student.CONFIDENTIAL_FLAG IS NULL
AND
class.ARR_TERM = {1}
ORDER BY LAST_NAME, FIRST_NAME
""".format(current_term, current_term)

# Send query to oracle
df = pd.read_sql(query, con=connection)

# View distribution of students (to check stratified sample)
# df["ARR_CLASS_C"].value_counts(normalize=True)

# Create split information (only want 4% of enrolled students)
split = StratifiedShuffleSplit(n_splits=1, test_size=.04, random_state=24)

# Perform stratified split
for train_index, test_index in split.split(df, df["ARR_CLASS_C"]):
    extra = df.loc[train_index]
    sample = df.loc[test_index]
# View distribution of students (to check stratified sample)
# sample["ARR_CLASS_C"].value_counts(normalize=True)

# Save to flatfile
filename = '{0}_Current_Student_Sample.xlsx'.format(current_term)
sample.to_excel(filename, index=False)
