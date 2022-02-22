import cx_Oracle, os
import pandas as pd
from datetime import datetime
# need openpyxl

# This will move old files to an existing archive folder, if needed
for file in os.listdir(os.getcwd()):
    extension = os.path.splitext(file)[1][1:]
    if extension == 'xlsx':
        os.rename(file,'Archive/' + file)

# This is for an issue with Mac, you might need to download the Oracle client
# and explicitly point to it
lib_dir='/local/location/of/lib/instantclient_19_8'
try:
    cx_Oracle.init_oracle_client(lib_dir=lib_dir)
except Exception as err:
    print("Error connecting: cx_Oracle.init_oracle_client()")
    print(err)
    sys.exit(1)

# This is for making the connection to the PE
ip = ''
port = 1521
SID = ''
dsn_tns = cx_Oracle.makedsn(ip, port, SID)

# Please enter your username and password for your PE login
username = ''
password = ''
connection = cx_Oracle.connect(username, password, dsn_tns)

# Getting timestamp to append on filenames
timestamp = datetime.now().strftime("%m%d%Y")

# File 1: Orientation Loaded
orientation_loaded = """
SELECT id.person_other_id_value PID, o.* FROM CVENT.ORIENTATION_LOADED o
LEFT JOIN
(
SELECT
  person_ID,
  person_other_id_value
 FROM usc_person.person_other_id
 WHERE PERSON_ID_TYPE = 'PID'
) ID
ON
o.PERSON_ID = id.person_ID
WHERE o.ADMIT_TERM IN ('20212','20213','20221')
"""
df = pd.read_sql(orientation_loaded, con=connection)
filename = 'orientation_loaded_' + timestamp + '.xlsx'
df.to_excel(filename, index=False)

# File 2: Orientation to Load
orientation_to_load = """
SELECT * FROM CVENT.V_ORIENTATION_TO_LOAD
"""
df = pd.read_sql(orientation_to_load, con=connection)
filename = 'orientation_to_load_' + timestamp + '.xlsx'
df.to_excel(filename, index=False)

# File 3: Merge Queue
merge_queue = """
SELECT * FROM usc_view.V_SIS_PE_DUPLICATE_QUEUE
"""
df = pd.read_sql(merge_queue, con=connection)
filename = 'merge_queue_' + timestamp + '.xlsx'
df.to_excel(filename, index=False)

# File 4: Match Queue
match_queue = """
SELECT * FROM esb.match_queue
"""
df = pd.read_sql(match_queue, con=connection)
filename = 'match_queue_' + timestamp + '.xlsx'
df.to_excel(filename, index=False)

# File 5: Advisement Loaded
advisement_loaded = """
SELECT
student.admit_term, advise.*
FROM
advisement.ADVISEMENT_LOADED advise
LEFT JOIN
advisement.STUDENT_LOADED student
ON advise.USCID = student.USCID
"""
df = pd.read_sql(advisement_loaded, con=connection)
filename = 'advisement_loaded_' + timestamp + '.xlsx'
df.to_excel(filename, index=False)

# File 6: Student Loaded
student_loaded = """
SELECT * FROM advisement.STUDENT_LOADED
"""
df = pd.read_sql(student_loaded, con=connection)
filename = 'student_loaded_' + timestamp + '.xlsx'
df.to_excel(filename, index=False)

print('Files completed!')
