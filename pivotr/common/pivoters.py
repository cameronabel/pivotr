import pandas as pd
import numpy as np
from pivotr.common import helpers


def jh_pivot(filename, head):
    """Returns a pivot table df based on JH assets."""
    source_dict = {'0': 'Profit Sharing',
                   '3': 'PS to Forf',
                   '4': 'Deferrals',
                   '5': 'SH Match',
                   '6': 'Rollover',
                   '7': 'SH Match to Forf',
                   '8': 'SHNEC',
                   '12': 'Roth',
                   '21': 'Rollover',
                   '23': '403(b) Rollover',
                   '25': 'Simple IRA Rollover',
                   '27': 'After-tax Rollover',
                   '29': 'Roth Rollover',
                   '55': 'Matching',
                   '57': 'Match to Forf',
                   '71': 'QACA Match'}
    trans_dict = {'0': 'Beg Bal',
                  '1': 'Contrib',
                  '2': 'Distrib',
                  '11': 'G/L',
                  '5': 'TRF',
                  '7': 'Loan Issue',
                  '8': 'Loan Pay',
                  '99': 'Loan Default',
                  '120': 'Roll In',
                  '999': 'End Bal'}

    worktable = pd.read_fwf(filename,
                            widths=[8, 9, 3, 6, 2, 3, 12, 15, 12, 10],
                            names=['Contract Number', 'SSN', 'Trans Code',
                                   'Period End', 'Source Code', 'Fund', 'Amount',
                                   'Units', 'Loan Interest', 'Loan Charge'])

    worktable['Trans'] = worktable['Trans Code'].apply(lambda x: trans_dict.get(str(x), str(x)))
    pye = str(worktable.at[1, 'Period End']).strip()
    pye = pye[:-2] + '20' + pye[-2:]

    pivoted_data = pd.pivot_table(worktable,
                                  values=['Amount'],
                                  index=['SSN', 'Source Code'],
                                  columns=['Trans'],
                                  aggfunc=np.sum,
                                  fill_value=0)

    pivoted_data.round(2)
    pivoted_data.columns = [s for (_, s) in list(pivoted_data.columns)]

    pivoted_data.reset_index(inplace=True)
    nametable = helpers.namegen(head)

    pivoted_data['SSN'] = pivoted_data['SSN'].astype(str)
    pivoted_data['SSN'] = pivoted_data['SSN'].str.zfill(9)
    cotable = pivoted_data.merge(nametable, on='SSN', how='left')

    trunctable = cotable.loc[:, (cotable != 0).any(axis=0)]

    trunctable['Source'] = trunctable['Source Code'].apply(lambda x: source_dict.get(str(x), str(x)))
    trunctable = trunctable.drop(columns=['Source Code'])
    trunctable['SSN'] = trunctable['SSN'].astype(int)
    columns = list(trunctable)
    for col in columns:
        if col not in ('Beg Bal', 'End Bal'):
            trunctable[col].replace(0, '', inplace=True)

    return trunctable, pye


def voya_pivot(filename):
    """Returns a pivot table df based on Voya assets."""

    worktable = pd.read_excel(filename, sheet_name=3, usecols='B:C,E:F,L:AB,AG')
    plan_name = str(worktable.at[1, 'Plan Name']).strip().replace('(K)', '(k)')
    pye = str(worktable.at[1, 'End Date']).strip().replace('/', '')
    worktable.drop(columns=['Plan Name', 'End Date'], inplace=True)
    worktable['Name'] = worktable['Name'].str.title()
    worktable['G/L'] = worktable['Dividends Earnings'] + worktable['Gain/Loss']
    worktable['TRF'] = worktable['Fund Transfers'] + worktable['Internal Transfers']
    worktable['Fee'] = worktable['Fees'] + worktable['TPA Fees']
    worktable.drop(columns=['Dividends Earnings', 'Gain/Loss',
                            'Fund Transfers', 'Internal Transfers',
                            'Fees', 'TPA Fees'], inplace=True)
    col_name_dict = {'Source Name': 'Source',
                     'Participant Number': 'SSN',
                     'Beginning Balance': 'Beg Bal',
                     'Contributions': 'Contrib',
                     'Takeover Contribution': 'Takeover',
                     'Loan Repayments': 'Loan Pmt',
                     'Loan Repay Principal': 'Principal',
                     'Loan Repay Interest': 'Interest',
                     'Withdrawals': 'Distrib',
                     'Forfeitures': 'Forf',
                     'Ending Balance': 'End Bal'}

    for col in list(worktable):
        worktable.rename(columns={col: col_name_dict.get(col, col)}, inplace=True)
    worktable = worktable[~worktable.Source.str.contains('Loans')]
    # worktable = worktable[~worktable.Name.str.contains('Forfeiture Account')]
    trunctable = worktable.loc[:, (worktable != 0).any(axis=0)]

    trunctable['SSN'] = trunctable['SSN'].str.replace('-', '').astype(int)
    trunctable['Name'] = trunctable['Name'].str.replace(',', ', ').replace(',  ', ', ')
    columns = list(trunctable)
    for col in columns:
        if col not in ('Beg Bal', 'End Bal'):
            trunctable[col].replace(0, '', inplace=True)
    return trunctable, pye, plan_name


def trc_pivot(filename):
    sources = pd.DataFrame([['Profit Sharing', 'Profit Sharing'],
                            ['Deferred Salary', 'Deferrals'],
                            ['Rollover', 'Rollover'],
                            ['Safe Harbor Non-elective', 'SHNEC'],
                            ['Roth Salary Deferral', 'Roth'],
                            ['Company Match', 'Matching']],
                           columns=['Source Name', 'Source'])

    b = open(filename, 'r').readlines()
    # line = b[1]
    pye = b[1][-11:-1].replace('/', '')

    worktable = pd.read_csv(filename, sep='\t',
                            usecols=[4, 5, 6, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23],
                            names=['First', 'Last', 'SSN', 'Source Name', 'Beg Bal', 'Contrib',
                                   'Loan', 'Forf Alloc', 'TRF', 'Distrib', 'Forf', 'Fees', 'Other',
                                   'G/L', 'End Bal'], header=None,
                            skiprows=range(0, 5))
    worktable.drop(worktable.tail(1).index, inplace=True)
    worktable.reset_index(drop=True, inplace=True)

    worktable['Name'] = worktable['Last'] + ', ' + worktable['First']
    worktable['SSN'] = worktable['SSN'].astype(int)

    pivoted_data = pd.pivot_table(worktable,
                                  index=['SSN', 'Source Name'],
                                  aggfunc=np.sum,
                                  fill_value=0)
    pivoted_data.reset_index(inplace=True)

    nametable = worktable[['SSN'] + ['Name']].drop_duplicates(subset='SSN')

    pivoted_data = pivoted_data.merge(nametable, on='SSN', how='left')
    pivoted_data = pivoted_data.merge(sources, on='Source Name', how='left')

    pivoted_data['Source Name'].update(pivoted_data['Source'])
    pivoted_data = pivoted_data.drop(columns=['Source']).rename(columns={'Source Name': 'Source'})

    pivoted_data = pivoted_data[['SSN', 'Name', 'Source', 'Beg Bal', 'Contrib',
                                 'Loan', 'Forf Alloc', 'TRF', 'Distrib', 'Forf', 'Fees', 'Other',
                                 'G/L', 'End Bal']].dropna(how='all')

    trunctable = pivoted_data.sort_values(by=['Source', 'Name'])

    trunctable = trunctable.loc[:, (trunctable != 0).any(axis=0)]
    plan_name = ''
    return trunctable, pye, plan_name


# noinspection PyBroadException
def rkd_pivot(filename):
    """Returns a pivot table df based on American Funds RKDirect assets."""

    worktable = pd.read_csv(filename, sep=',',
                            usecols=[1, 2, 3, 4, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 26],
                            names=['planID', 'SSN', 'fullname', 'MYT', 'Beg Bal', 'Conv', 'Contrib',
                                   'Dividends', 'Gain/Loss', 'Exch', 'Fees', 'Forf', 'Distrib', 'Other',
                                   'TRFIN', 'TRFOUT', 'Loan TRFIN', 'Loan TRFOUT', 'Int', 'Ins', 'End Bal', 'PYE'],
                            header=None,
                            skiprows=1)
    pye = worktable['PYE'][1].replace('/', '')

    planID = worktable['planID'][1]
    worktable['Name'] = worktable['fullname'].apply(helpers.parsename).str.title()
    worktable['G/L'] = worktable['Dividends'] + worktable['Gain/Loss']
    worktable['TRF'] = worktable['Exch'] + worktable['TRFIN'] + worktable['TRFOUT']
    worktable.drop(columns=['fullname', 'Dividends', 'Gain/Loss', 'planID',
                            'Exch', 'TRFIN', 'TRFOUT', 'PYE'], inplace=True)
    pivoted_data = pd.pivot_table(worktable,
                                  index=['Name', 'SSN', 'MYT'],
                                  aggfunc=np.sum,
                                  fill_value=0)
    pivoted_data.reset_index(inplace=True)

    trunctable = pivoted_data.loc[:, (pivoted_data != 0).any(axis=0)]
    trunctable['MYT'] = trunctable['MYT'].astype(str)
    trunctable['Source'] = trunctable['MYT'].apply(helpers.parsemyt)
    trunctable = trunctable.drop(columns=['MYT'])

    trunctable['SSN'] = trunctable['SSN'].astype(int)
    columns = list(trunctable)
    for col in columns:
        if col not in ('Beg Bal', 'End Bal'):
            trunctable[col].replace(0, '', inplace=True)

    try:
        planlist = pd.read_csv('rkdplans.csv')
        planname = planlist.loc[planlist['planID'] == planID]
        planname.reset_index(inplace=True)
        planname = planname.at[0, 'Plan Name']
    except Exception:
        planname = ''

    return trunctable, pye, planname


def emp_pivot(filename, tail):
    """Returns a pivot table df based on Empower assets."""

    worktable = pd.read_excel(filename, sep=',',
                              usecols=[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13],
                              names=['SSN', 'Name', 'Source Code', 'Beg Bal', 'Contrib_Reg',
                                     'Contrib_SGL', 'Loan Pay', 'Cred_Int', 'GL', 'Fees', 'Forf',
                                     'New Loans', 'Distrib', 'End Bal'],
                              header=None,
                              skiprows=1)
    pye = ''
    planname = tail[:-5]
    worktable['Name'] = worktable['Name'].str.title()
    worktable['G/L'] = worktable['Cred_Int'] + worktable['GL']
    worktable['Contribs'] = worktable['Contrib_Reg'] + worktable['Contrib_SGL']
    worktable.drop(columns=['Cred_Int', 'GL',
                            'Contrib_SGL', 'Contrib_Reg'],
                   inplace=True)
    pivoted_data = pd.pivot_table(worktable,
                                  index=['Name', 'SSN', 'Source Code'],
                                  aggfunc=np.sum,
                                  fill_value=0)
    pivoted_data.reset_index(inplace=True)

    trunctable = pivoted_data.loc[:, (pivoted_data != 0).any(axis=0)]
    trunctable['Source Code'] = trunctable['Source Code'].astype(str)
    trunctable['Source'] = trunctable['Source Code'].apply(helpers.parse_emp_src)
    trunctable = trunctable.drop(columns=['Source Code'])

    trunctable['SSN'] = trunctable['SSN'].str.replace('-', '').astype(int)
    columns = list(trunctable)
    for col in columns:
        if col not in ('Beg Bal', 'End Bal'):
            trunctable[col].replace(0, '', inplace=True)

    return trunctable, pye, planname


def prin_pivot(filename):
    """Returns a pivot table based on Principal assets."""
    pye = ''
    worktable = pd.read_csv(filename,
                            usecols=[0, 1, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22],
                            names=['First',
                                   'Last',
                                   'SSN',
                                   'Source',
                                   'Beg Bal',
                                   'End Bal',
                                   'gains',
                                   'Contrib',
                                   'dist',
                                   'ins',
                                   'rmd',
                                   'capgain',
                                   'div',
                                   'Fees',
                                   'Forf',
                                   'Refunds',
                                   'Loan Fees',
                                   'Loan Dist',
                                   'TRF',
                                   'Rollover',
                                   'Loan Prin',
                                   'Loan Int'],
                            header=None,
                            skiprows=1
                            )
    worktable.drop(worktable.tail(1).index, inplace=True)
    worktable.reset_index(drop=True, inplace=True)

    worktable['Name'] = worktable['Last'] + ', ' + worktable['First']
    worktable['SSN'] = worktable['SSN'].astype(int)
    worktable['Distrib'] = worktable['dist'] + worktable['ins'] + worktable['rmd']

    worktable['Name'] = worktable['Name'].str.title()
    worktable['G/L'] = worktable['gains'] + worktable['capgain'] + worktable['div']
    worktable['Distrib'] = worktable['dist'] + worktable['ins'] + worktable['rmd']
    worktable.drop(columns=['gains', 'capgain', 'div',
                            'dist', 'ins', 'rmd'],
                   inplace=True)

    pivoted_data = pd.pivot_table(worktable,
                                  index=['SSN', 'Source'],
                                  aggfunc=np.sum,
                                  fill_value=0)
    pivoted_data.reset_index(inplace=True)

    nametable = worktable[['SSN'] + ['Name']].drop_duplicates(subset='SSN')

    pivoted_data = pivoted_data.merge(nametable, on='SSN', how='left')

    trunctable = pivoted_data.sort_values(by=['Source', 'Name'])

    trunctable = trunctable.loc[:, (trunctable != 0).any(axis=0)]
    plan_name = ''
    return trunctable, pye, plan_name
