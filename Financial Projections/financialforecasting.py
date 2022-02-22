from sqlalchemy import create_engine
from datetime import datetime
import pandas as pd
# need pymysql, xlsxwriter

# Maps fiscal period to fiscal year to date tabulation
fiscal_period_dict = {
       1 : ['MO1_ACCT_LN_AMT', 'FYTD_AS_OF_12'],
       2 : ['MO2_ACCT_LN_AMT', 'FYTD_AS_OF_01'],
       3 : ['MO3_ACCT_LN_AMT', 'FYTD_AS_OF_02'],
       4 : ['MO4_ACCT_LN_AMT', 'FYTD_AS_OF_03'],
       5 : ['MO5_ACCT_LN_AMT', 'FYTD_AS_OF_04'],
       6 : ['MO6_ACCT_LN_AMT', 'FYTD_AS_OF_05'],
       7 : ['MO7_ACCT_LN_AMT', 'FYTD_AS_OF_06'],
       8 : ['MO8_ACCT_LN_AMT', 'FYTD_AS_OF_07'],
       9 : ['MO9_ACCT_LN_AMT', 'FYTD_AS_OF_08'],
       10:['MO10_ACCT_LN_AMT', 'FYTD_AS_OF_09'],
       11:['MO11_ACCT_LN_AMT', 'FYTD_AS_OF_10'],
       12:['MO12_ACCT_LN_AMT', 'FYTD_AS_OF_11']
}

# Runs MySQL query that gets GL data by fiscal period and maps to translated
# Workday values, returns fiscal year dataframes
def make_df():
    query = """
            SELECT
            gl.*,
            ACLN_ANNL_BAL_AMT - mo15_acct_ln_amt - mo14_acct_ln_amt - mo13_acct_ln_amt AS FYTD_AS_OF_12,
            ACLN_ANNL_BAL_AMT - mo15_acct_ln_amt - mo14_acct_ln_amt - mo13_acct_ln_amt - mo12_acct_ln_amt AS FYTD_AS_OF_11,
            ACLN_ANNL_BAL_AMT - mo15_acct_ln_amt - mo14_acct_ln_amt - mo13_acct_ln_amt - mo12_acct_ln_amt - mo11_acct_ln_amt AS FYTD_AS_OF_10,
            ACLN_ANNL_BAL_AMT - mo15_acct_ln_amt - mo14_acct_ln_amt - mo13_acct_ln_amt - mo12_acct_ln_amt - mo11_acct_ln_amt - mo10_acct_ln_amt AS FYTD_AS_OF_09,
            ACLN_ANNL_BAL_AMT - mo15_acct_ln_amt - mo14_acct_ln_amt - mo13_acct_ln_amt - mo12_acct_ln_amt - mo11_acct_ln_amt - mo10_acct_ln_amt - mo9_acct_ln_amt AS FYTD_AS_OF_08,
            ACLN_ANNL_BAL_AMT - mo15_acct_ln_amt - mo14_acct_ln_amt - mo13_acct_ln_amt - mo12_acct_ln_amt - mo11_acct_ln_amt - mo10_acct_ln_amt - mo9_acct_ln_amt - mo8_acct_ln_amt AS FYTD_AS_OF_07,
            ACLN_ANNL_BAL_AMT - mo15_acct_ln_amt - mo14_acct_ln_amt - mo13_acct_ln_amt - mo12_acct_ln_amt - mo11_acct_ln_amt - mo10_acct_ln_amt - mo9_acct_ln_amt - mo8_acct_ln_amt - mo7_acct_ln_amt AS FYTD_AS_OF_06,
            ACLN_ANNL_BAL_AMT - mo15_acct_ln_amt - mo14_acct_ln_amt - mo13_acct_ln_amt - mo12_acct_ln_amt - mo11_acct_ln_amt - mo10_acct_ln_amt - mo9_acct_ln_amt - mo8_acct_ln_amt - mo7_acct_ln_amt - mo6_acct_ln_amt AS FYTD_AS_OF_05,
            ACLN_ANNL_BAL_AMT - mo15_acct_ln_amt - mo14_acct_ln_amt - mo13_acct_ln_amt - mo12_acct_ln_amt - mo11_acct_ln_amt - mo10_acct_ln_amt - mo9_acct_ln_amt - mo8_acct_ln_amt - mo7_acct_ln_amt - mo6_acct_ln_amt - mo5_acct_ln_amt AS FYTD_AS_OF_04,
            ACLN_ANNL_BAL_AMT - mo15_acct_ln_amt - mo14_acct_ln_amt - mo13_acct_ln_amt - mo12_acct_ln_amt - mo11_acct_ln_amt - mo10_acct_ln_amt - mo9_acct_ln_amt - mo8_acct_ln_amt - mo7_acct_ln_amt - mo6_acct_ln_amt - mo5_acct_ln_amt - mo4_acct_ln_amt AS FYTD_AS_OF_03,
            ACLN_ANNL_BAL_AMT - mo15_acct_ln_amt - mo14_acct_ln_amt - mo13_acct_ln_amt - mo12_acct_ln_amt - mo11_acct_ln_amt - mo10_acct_ln_amt - mo9_acct_ln_amt - mo8_acct_ln_amt - mo7_acct_ln_amt - mo6_acct_ln_amt - mo5_acct_ln_amt - mo4_acct_ln_amt - mo3_acct_ln_amt AS FYTD_AS_OF_02,
            ACLN_ANNL_BAL_AMT - mo15_acct_ln_amt - mo14_acct_ln_amt - mo13_acct_ln_amt - mo12_acct_ln_amt - mo11_acct_ln_amt - mo10_acct_ln_amt - mo9_acct_ln_amt - mo8_acct_ln_amt - mo7_acct_ln_amt - mo6_acct_ln_amt - mo5_acct_ln_amt - mo4_acct_ln_amt - mo3_acct_ln_amt - mo2_acct_ln_amt AS FYTD_AS_OF_01,
            object.fin_obj_cd_nm,
            object.mfs_group_cd,
            object.mfs_group_nm,
            trans.*
            FROM
            BALANCEFACT gl
            LEFT JOIN OBJECTDIM object
            ON
            gl.object_dim_id = object.object_dim_id
            LEFT JOIN
            (
            SELECT distinct
            legacyaccount AS Account,
            ifnull(ifnull(ifnull(workdayproject,workdayprogram),workdaygift),workdaygrant) AS PPGG,
            workdaycostcenter AS CostCenter,
            workdayfund AS Fund,
            workdayfunction AS Fuction,
            legacysubaccount AS ObjectCode,
            legacysubaccountname AS ObjectCodeName,
            ifnull(workdayrevenuecategory, workdayspendcategory) AS SpendRevenueCategory,
            ifnull(workdayrevenuecategoryname, workdayspendcategoryname) AS SpendRevenueCategoryName,
            workdayledgeraccount AS LedgerAccount,
            workdayledgeraccountname AS LedgerAccountName
            FROM
            TRANSLATIONTABLE
            WHERE
            translation_status = 'Correct'
            ) trans
            ON
            gl.account_nbr = trans.Account
            AND
            gl.FIN_OBJECT_CD = trans.ObjectCode
            WHERE
            gl.account_nbr = {0}
            AND
            gl.univ_fiscal_yr IN ("2020", "2021")
            """.format(account)
    df = pd.read_sql(query, con=db_connection)
    df = df.replace({'EX': 'ENC', 'IE': 'ENC'})
    df_2021 = df[df['UNIV_FISCAL_YR']==2021]
    df_2020 = df[df['UNIV_FISCAL_YR']==2020]
    return(df_2021, df_2020)

# Calculate current fiscal period and corresponding column values
def get_fiscal_period():
    month = datetime.today().month
    if month >= 7:
        current_fiscal_period = month - 6
    else:
        current_fiscal_period = month + 6
    current_fp_column, fytd_as_of_fp = fiscal_period_dict[current_fiscal_period]
    previous_fiscal_period = current_fiscal_period - 1
    remaining_months = 12 - previous_fiscal_period
    previous_fp_column = fiscal_period_dict[previous_fiscal_period][0]
    column_list = [fytd_as_of_fp, current_fp_column, previous_fp_column]
    return(column_list, remaining_months)

def rename_columns(df, fy):
    df.columns.values[-4] = 'FYTD_' + fy
    df.columns.values[-3] = 'FYTD_END_OF_PRIOR_FP_' + fy
    df.columns.values[-2] = 'CURRENT_FP_' + fy
    df.columns.values[-1] = 'PREVIOUS_FP_' + fy

# Make tabulations by fiscal year, MFS code, object code, ledger account, and category
def make_tabulations(df, columns, fy):
    df_sum = pd.DataFrame(df.groupby(['FIN_BALANCE_TYP_CD'])[columns].sum()).reset_index()
    rename_columns(df_sum, fy)
    df_mfs_sum = pd.DataFrame(df.groupby(['FIN_BALANCE_TYP_CD', 'mfs_group_cd', 'mfs_group_nm'])[columns].sum()).reset_index()
    rename_columns(df_mfs_sum, fy)
    df_obj_sum = pd.DataFrame(df.groupby(['FIN_BALANCE_TYP_CD', 'FIN_OBJECT_CD', 'fin_obj_cd_nm'])[columns].sum()).reset_index()
    rename_columns(df_obj_sum, fy)
    df_ledger_sum = pd.DataFrame(df.groupby(['FIN_BALANCE_TYP_CD', 'LedgerAccount', 'LedgerAccountName'])[columns].sum()).reset_index()
    rename_columns(df_ledger_sum, fy)
    df_category_sum = pd.DataFrame(df.groupby(['FIN_BALANCE_TYP_CD', 'SpendRevenueCategory', 'SpendRevenueCategoryName'])[columns].sum()).reset_index()
    rename_columns(df_category_sum, fy)
    return(df_sum, df_mfs_sum, df_obj_sum, df_ledger_sum, df_category_sum)

# Combine tabulations for each fiscal year into one
def join_fiscal_years(df_sums, df_mfs_sums, df_obj_sums, df_ledger_sums, df_category_sums):
    df_sum = df_sums[0].join(df_sums[1].set_index('FIN_BALANCE_TYP_CD'), on='FIN_BALANCE_TYP_CD')
    df_mfs_sum = pd.merge(df_mfs_sums[0], df_mfs_sums[1], on=['FIN_BALANCE_TYP_CD','mfs_group_cd', 'mfs_group_nm'])
    df_obj_sum = pd.merge(df_obj_sums[0], df_obj_sums[1], on=['FIN_BALANCE_TYP_CD','FIN_OBJECT_CD', 'fin_obj_cd_nm'])
    df_ledger_sum = pd.merge(df_ledger_sums[0], df_ledger_sums[1], on=['FIN_BALANCE_TYP_CD','LedgerAccount', 'LedgerAccountName'])
    df_category_sum = pd.merge(df_category_sums[0], df_category_sums[1], on=['FIN_BALANCE_TYP_CD','SpendRevenueCategory', 'SpendRevenueCategoryName'])
    return(df_sum, df_mfs_sum, df_obj_sum, df_ledger_sum, df_category_sum)

# Bring budget encumbrance into fiscal year tabulations
def get_current_budget_encumbrance(df, type):
    df_actuals = df[df['FIN_BALANCE_TYP_CD'] =='AC']
    if type == 'SUM':
        df_budget = pd.DataFrame(df[df['FIN_BALANCE_TYP_CD'] =='CB'].reset_index()['FYTD_END_OF_PRIOR_FP_2021'])
        df_budget.columns.values[-1] = 'CURR_BUDGET'
        df_encumbrance = pd.DataFrame(df[df['FIN_BALANCE_TYP_CD'] =='ENC'].reset_index()['FYTD_END_OF_PRIOR_FP_2021'])
        df_encumbrance.columns.values[-1] = 'CURR_ENCUMBRANCE'
        df_sum = df_actuals.join(df_budget)
        df_sum = df_sum.join(df_encumbrance)
    elif type == 'MFS':
        df_budget = pd.DataFrame(df[df['FIN_BALANCE_TYP_CD'] =='CB'][['mfs_group_cd', 'mfs_group_nm', 'FYTD_END_OF_PRIOR_FP_2021']])
        df_budget.columns.values[-1] = 'CURR_BUDGET'
        df_encumbrance = pd.DataFrame(df[df['FIN_BALANCE_TYP_CD'] =='ENC'][['mfs_group_cd', 'mfs_group_nm', 'FYTD_END_OF_PRIOR_FP_2021']])
        df_encumbrance.columns.values[-1] = 'CURR_ENCUMBRANCE'
        df_sum = pd.merge(df_actuals, df_budget, on=['mfs_group_cd', 'mfs_group_nm'], how='outer')
        df_sum = pd.merge(df_sum, df_encumbrance, on=['mfs_group_cd', 'mfs_group_nm'], how='outer')
    elif type == 'OBJ':
        df_budget = pd.DataFrame(df[df['FIN_BALANCE_TYP_CD'] =='CB'][['FIN_OBJECT_CD', 'fin_obj_cd_nm', 'FYTD_END_OF_PRIOR_FP_2021']])
        df_budget.columns.values[-1] = 'CURR_BUDGET'
        df_encumbrance = pd.DataFrame(df[df['FIN_BALANCE_TYP_CD'] =='ENC'][['FIN_OBJECT_CD', 'fin_obj_cd_nm', 'FYTD_END_OF_PRIOR_FP_2021']])
        df_encumbrance.columns.values[-1] = 'CURR_ENCUMBRANCE'
        df_sum = pd.merge(df_actuals, df_budget, on=['FIN_OBJECT_CD', 'fin_obj_cd_nm'], how='outer')
        df_sum = pd.merge(df_sum, df_encumbrance, on=['FIN_OBJECT_CD', 'fin_obj_cd_nm'], how='outer')
    elif type == 'LEDG':
        df_budget = pd.DataFrame(df[df['FIN_BALANCE_TYP_CD'] =='CB'][['LedgerAccount', 'LedgerAccountName', 'FYTD_END_OF_PRIOR_FP_2021']])
        df_budget.columns.values[-1] = 'CURR_BUDGET'
        df_encumbrance = pd.DataFrame(df[df['FIN_BALANCE_TYP_CD'] =='ENC'][['LedgerAccount', 'LedgerAccountName', 'FYTD_END_OF_PRIOR_FP_2021']])
        df_encumbrance.columns.values[-1] = 'CURR_ENCUMBRANCE'
        df_sum = pd.merge(df_actuals, df_budget, on=['LedgerAccount', 'LedgerAccountName'], how='outer')
        df_sum = pd.merge(df_sum, df_encumbrance, on=['LedgerAccount', 'LedgerAccountName'], how='outer')
    elif type == 'CAT':
        df_budget = pd.DataFrame(df[df['FIN_BALANCE_TYP_CD'] =='CB'][['SpendRevenueCategory', 'SpendRevenueCategoryName', 'FYTD_END_OF_PRIOR_FP_2021']])
        df_budget.columns.values[-1] = 'CURR_BUDGET'
        df_encumbrance = pd.DataFrame(df[df['FIN_BALANCE_TYP_CD'] =='ENC'][['SpendRevenueCategory', 'SpendRevenueCategoryName', 'FYTD_END_OF_PRIOR_FP_2021']])
        df_encumbrance.columns.values[-1] = 'CURR_ENCUMBRANCE'
        df_sum = pd.merge(df_actuals, df_budget, on=['SpendRevenueCategory', 'SpendRevenueCategoryName'], how='outer')
        df_sum = pd.merge(df_sum, df_encumbrance, on=['SpendRevenueCategory', 'SpendRevenueCategoryName'], how='outer')
    return(df_sum)

# Calculate final tables
def return_final_tables(columns):
    df_2021, df_2020 = make_df()
    df_sum_2021, df_mfs_sum_2021, df_obj_sum_2021, df_ledger_sum_2021, df_category_sum_2021 = make_tabulations(df_2021, columns, '2021')
    df_sum_2020, df_mfs_sum_2020, df_obj_sum_2020, df_ledger_sum_2020, df_category_sum_2020 = make_tabulations(df_2020, columns, '2020')
    df_sum, df_mfs_sum, df_obj_sum, df_ledger_sum, df_category_sum  =join_fiscal_years([df_sum_2021, df_sum_2020], [df_mfs_sum_2021, df_mfs_sum_2020],
    [df_obj_sum_2021, df_obj_sum_2020], [df_ledger_sum_2021, df_ledger_sum_2020], [df_category_sum_2021, df_category_sum_2020])
    df_sum = get_current_budget_encumbrance(df_sum, "SUM")
    df_mfs_sum = get_current_budget_encumbrance(df_mfs_sum, "MFS")
    df_obj_sum = get_current_budget_encumbrance(df_obj_sum, "OBJ")
    df_ledger_sum = get_current_budget_encumbrance(df_ledger_sum, "LEDG")
    df_category_sum = get_current_budget_encumbrance(df_category_sum, "CAT")
    return(df_sum, df_mfs_sum, df_obj_sum, df_ledger_sum, df_category_sum)

# Define different financial projection methods (as defined within our TM1 instance)
def straight_line(df, remaining_months, current_year):
    # Straight Line Method: Current Month x Remaining Months + Year to Date
    previous_fp = 'PREVIOUS_FP_' + current_year
    fytd_as_of_prior_fp = 'FYTD_END_OF_PRIOR_FP_' + current_year
    df['straight_line'] = (df[previous_fp] * remaining_months) + df[fytd_as_of_prior_fp]

def average(df, remaining_months, current_year):
    # Average Method: Year to Date / Months Elapsed x 12
    fytd_as_of_prior_fp = 'FYTD_END_OF_PRIOR_FP_' + current_year
    df['average'] = (df[fytd_as_of_prior_fp] / (12 - remaining_months)) * 12

def marginal_rate(df, current_year, previous_year):
    # Marginal Rate: Year to Date x (Prior Year Actual / Prior Year to Date)
    fytd_as_of_prior_fp = 'FYTD_END_OF_PRIOR_FP_' + current_year
    previous_fytd_as_of_prior_fp = 'FYTD_END_OF_PRIOR_FP_' + previous_year
    fytd = 'FYTD_' + previous_year
    df['marginal_rate'] = df[fytd_as_of_prior_fp] * (df[fytd] / df[previous_fytd_as_of_prior_fp])

def actual_prior_actual(df, current_year, previous_year):
    fytd_as_of_prior_fp = 'FYTD_END_OF_PRIOR_FP_' + current_year
    fytd = 'FYTD_' + previous_year
    df['actual'] = df[fytd_as_of_prior_fp]
    df['prior_actual'] = df[fytd]

def average_minus_one(df, remaining_months, current_year):
    fytd_as_of_prior_fp = 'FYTD_END_OF_PRIOR_FP_' + current_year
    df['average_minus_one'] = (df[fytd_as_of_prior_fp] / ((12 - remaining_months)-1)) * 12

def prior_year_remaining(df, current_year, previous_year):
    fytd_as_of_prior_fp = 'FYTD_END_OF_PRIOR_FP_' + current_year
    previous_fytd_as_of_prior_fp = 'FYTD_END_OF_PRIOR_FP_' + previous_year
    fytd = 'FYTD_' + previous_year
    df['prior_year_remaining'] = df[fytd_as_of_prior_fp] + (df[fytd] - df[previous_fytd_as_of_prior_fp])

def avg_of_four(df):
    df['avg_of_four'] = (df['straight_line'] + df['average'] + df['marginal_rate'] + df['prior_year_remaining'])/4

def encumbrance(df, current_year):
    fytd_as_of_prior_fp = 'FYTD_END_OF_PRIOR_FP_' + current_year
    df['encumbrance'] = df[fytd_as_of_prior_fp] + df['CURR_ENCUMBRANCE']

def apply_all_methods(df, remaining_months, current_year, previous_year):
    straight_line(df, remaining_months, current_year)
    average(df, remaining_months, current_year)
    marginal_rate(df, current_year, previous_year)
    actual_prior_actual(df, current_year, previous_year)
    average_minus_one(df, remaining_months, current_year)
    prior_year_remaining(df, current_year, previous_year)
    avg_of_four(df)
    encumbrance(df, current_year)

if __name__ == "__main__":
    db_connection_str = 'mysql+pymysql://USERNAME:PASSWORD@IP/SCHEMA'
    db_connection = create_engine(db_connection_str)

    account = input("Please enter your account number: ")

    columns, remaining_months = get_fiscal_period()
    fiscal_year = 'ACLN_ANNL_BAL_AMT'
    columns.insert(0, fiscal_year)
    df_sum, df_mfs_sum, df_obj_sum, df_ledger_sum, df_category_sum = return_final_tables(columns)

    apply_all_methods(df_sum, remaining_months, "2021", "2020")
    apply_all_methods(df_mfs_sum, remaining_months, "2021", "2020")
    apply_all_methods(df_obj_sum, remaining_months, "2021", "2020")
    apply_all_methods(df_ledger_sum, remaining_months, "2021", "2020")
    apply_all_methods(df_category_sum, remaining_months, "2021", "2020")

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('{0}_projection_translated.xlsx'.format(account), engine='xlsxwriter')

    # Write each dataframe to a different worksheet.
    df_sum.to_excel(writer, sheet_name='df_sum')
    df_mfs_sum.to_excel(writer, sheet_name='df_mfs_sum')
    df_obj_sum.to_excel(writer, sheet_name='df_obj_sum')
    df_ledger_sum.to_excel(writer, sheet_name='df_ledger_sum')
    df_category_sum.to_excel(writer, sheet_name='df_category_sum')

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()
