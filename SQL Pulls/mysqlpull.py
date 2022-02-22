from sqlalchemy import create_engine
import pandas as pd

# Connect to MySQL
db_connection_str = 'mysql+pymysql://USERNAME:PASSWORD@IP/SCHEMA'
db_connection = create_engine(db_connection_str)

# Labor distribution query
village_query = """
SELECT  l.emplid 'EmployeeID',
e.legal_name 'EmployeeName',
e.employee_last_name 'EmployeeLast Name',
e.employee_first_name 'Employee First Name',
e.employee_middle_name 'Employee Middle Name',
e.usc_id 'USC ID',
ef.actual_hours 'Actual Hours',
CASE WHEN l.fin_coa_cd = 'SC' AND l.fin_balance_typ_cd = 'AC' AND l.trn_debit_crdt_cd = 'C'
                THEN ROUND(((l.trn_ldgr_entr_amt * -1) / pay_rate),2)
WHEN l.fin_coa_cd = 'SC' AND l.fin_balance_typ_cd = 'AC' AND l.trn_debit_crdt_cd = 'D'
                THEN ROUND((l.trn_ldgr_entr_amt / pay_rate),2)
ELSE NULL
END 'Paid Hours',
l.account_nbr 'Account Number',
a.account_nm 'Account Name',
l.fin_object_cd 'Object Code',
o.fin_obj_cd_nm 'Object Name',
e.home_postal_code,
CASE WHEN  e.home_postal_code IN () THEN 'Tier 2' ELSE 'Outside of Radius' END AS 'Local Employee',
CASE WHEN month(l.pay_period_end_dt) IN (1,2,3) THEN 'Q1'
     WHEN month(l.pay_period_end_dt) IN (4,5,6) THEN 'Q2'
     WHEN month(l.pay_period_end_dt) IN (7,8,9) THEN 'Q3'
     WHEN month(l.pay_period_end_dt) IN (10,11,12,13, 14, 15) THEN 'Q4' ELSE NULL END AS 'Quarter',
 year(l.pay_period_end_dt)
FROM LABORLEDGER l
         INNER JOIN EMPLOYEEDIM e
                 ON l.employee_dim_id = e.employee_dim_id
         INNER JOIN ACCOUNTDIM a
                 ON l.account_dim_id = a.account_dim_id
         INNER JOIN EARNINGDIM r
                 ON l.earning_dim_id = r.earning_dim_id
         LEFT JOIN HOMEDEPTDIM h
                 ON l.home_dept_dim_id = h.home_dept_dim_id
         INNER JOIN OBJECTDIM o
                 ON l.object_dim_id = o.object_dim_id
         LEFT JOIN HOMEDEPTDIM p
                ON l.emp_region_cd = p.home_dept_code AND l.trn_post_dt BETWEEN p.created AND p.expired
LEFT JOIN WDINTEGRATION ef
                ON l.workday_id = ef.workday_id
                AND l.fin_balance_typ_cd = ef.fin_balance_typ_cd AND l.fin_balance_typ_cd = 'AC'
                AND l.pay_period_end_dt = ef.pay_period_end_dt
                AND l.cycl_typ = ef.cycl_typ
                AND l.trn_debit_crdt_cd = ef.trn_debit_crdt_cd
WHERE 1=1
AND l.fin_balance_typ_cd IN  ('AC')
AND year(l.pay_period_end_dt) IN (2020, 2019)
AND l.account_nbr IN ()
and e.employee_type <> 'Student'
"""

# Execute query and save raw data to flat file
df = pd.read_sql(village_query , con=db_connection)
df.to_excel('Raw_Village_Data.xlsx', index=False)
