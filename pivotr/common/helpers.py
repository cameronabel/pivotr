import os
import pandas as pd
from nameparser import HumanName


def determine_file_type(filename):
    """Determines investment company source and file path."""
    head, tail = os.path.split(filename)
    if is_jh(filename):
        file_type = 'John Hancock'
        valid_file = True
    elif tail[:14] == 'ArchiveService':
        file_type = 'Voya'
        valid_file = True
    elif 'PartcBalance' in tail:
        file_type = 'TRC'
        valid_file = True
    elif is_rkdirect(filename):
        file_type = 'RK Direct'
        valid_file = True
    elif is_emp(filename):
        file_type = 'Empower'
        valid_file = True
    elif is_prin(filename):
        file_type = 'Principal'
        valid_file = True
    elif is_ascensus(filename):
        file_type = 'Ascensus'
        valid_file = True
    else:
        file_type = 'Incompatible file source'
        valid_file = False
    return file_type, valid_file, head, tail


def is_jh(tail):
    if tail[-7:] == 'YTD.TXT':
        return True
    return False


# noinspection PyBroadException
def is_rkdirect(filename):
    try:
        b = open(filename, 'r').readlines()
        # print(b[0])
        # print(b[0][:6])
        if b[0][:6] == 'ICU ID' or b[0][:6] == '"ICU I':
            return True
        else:
            return False
    except Exception:
        return False


# noinspection PyBroadException
def is_emp(filename):
    try:
        test_df = pd.read_excel(filename)
        if 'Contributions, SGL' in test_df.columns:
            return True
        else:
            return False
    except Exception:
        return False


# noinspection PyBroadException
def is_prin(filename):
    try:
        test_df = pd.read_csv(filename)
        if 'Source Text' in test_df.columns:
            return True
        else:
            return False
    except Exception:
        return False

# noinspection PyBroadException
def is_ascensus(filename):
    try:
        test_df = pd.read_excel(filename, skiprows=6)
        if 'Location Name' in test_df.columns:
            return True
        else:
            return False
    except Exception:
        return False

# noinspection PyBroadException
def namegen(head):
    """Returns a df of names and SSNs."""
    try:
        ascname = pd.read_clipboard('\t', header=1)
        ascname.rename(columns={'SocSecNum': 'SSN'}, inplace=True)
        ascname.rename(columns={'Employee Name': 'Name'}, inplace=True)
        ascname['Name'] = ascname['Name'].str.strip()
        ascname['SSN'] = ascname['SSN'].str.replace('-', '')
        ascname['SSN'] = ascname['SSN'].astype(str)
        ascname['SSN'] = ascname['SSN'].str.zfill(9)

    except Exception:
        try:
            ascname = pd.read_csv(head + '\\names.csv', names=['SSN', 'Name'])
            ascname['Name'] = ascname['Name'].str.strip()
            ascname['SSN'] = ascname['SSN'].str.replace('-', '')
            ascname = ascname.drop([0, 1])
            ascname['SSN'] = ascname['SSN'].astype(str)
            ascname['SSN'] = ascname['SSN'].str.zfill(9)

        except Exception:
            ascname = pd.DataFrame([['', '']], columns=['SSN', 'Name'])

    try:
        sub = 'Census_Summary'
        census_report = next((s for s in os.listdir(head) if sub in s), None)
        census_report = head + '\\' + census_report
        nametable = pd.read_csv(census_report,
                                usecols=[2, 3, 4],
                                names=['SSN', 'First', 'Last'])

    except Exception:
        nametable = pd.DataFrame([['', '', '']], columns=['SSN', 'First', 'Last'])

    nametable['Name'] = nametable['Last'] + ', ' + nametable['First']
    nametable['SSN'] = nametable['SSN'].str.replace('-', '')
    nametable = nametable.drop([0]).drop(columns=['First', 'Last'])
    nametable['SSN'] = nametable['SSN'].astype(str)
    nametable['SSN'] = nametable['SSN'].str.zfill(9)
    nametable = pd.merge(nametable, ascname, on='SSN', how='outer')
    nametable['Name_x'].update(nametable['Name_y'])
    nametable = nametable.rename(columns={'Name_x': 'Name'}).drop(columns=['Name_y'])

    return nametable


def parsename(fullname):
    name = HumanName(fullname)
    return name.last + ', ' + name.first


def parsemyt(myt):
    src_dict = {'103': 'Profit Sharing',
                '101': 'Deferrals',
                '137': 'Pension',
                '104': 'Matching',
                '119': 'ESOP Rollover',
                '112': 'Rollover'}
    return src_dict.get(myt, myt)


def parse_emp_src(source_code):
    src_dict = {'ER01': 'Profit Sharing',
                'ER02': 'Prev Wage',
                'BTK1': 'Deferrals',
                '137': 'Pension',
                'ERM1': 'Matching',
                'RTH1': 'Roth',
                'EER1': 'Rollover',
                'SHM1': 'SH Match'}
    return src_dict.get(source_code, source_code)


def end_bal_check(table: pd.DataFrame) -> pd.DataFrame:
    if 'End Bal' not in table.columns:
        table['End Bal'] = 0
    if 'Beg Bal' not in table.columns:
        table['Beg Bal'] = 0
    return table
