"""Produces formatted Excel files containing pivot tables from investment company source files"""
import sys
import os
if (sys.platform == 'win32' and sys.executable.split('\\')[-1] == 'pythonw.exe'):
    sys.stdout = open('log.txt', 'w')
    sys.stderr = open('err.txt', 'w')
from types import SimpleNamespace
import kivy
from kivy.app import App
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.stacklayout import StackLayout
from kivy.core.window import Window
from nameparser import HumanName
from kivy.config import Config
import numpy as np
import pandas as pd

kivy.require("1.9.1")

__author__ = 'Cameron Abel'
__copyright__ = 'Copyright 2019'
__credits__ = ['Cameron Abel']
__license__ = 'MIT'
__maintainer__ = 'Cameron Abel'
__email__ = 'cameronabel@gmail.com'
__status__ = 'Development'

Config.set('graphics', 'width', '400')
Config.set('graphics', 'height', '500')
Config.write()

ActiveTable= SimpleNamespace(
    filename = '',
    file_type = '',
    valid_file = False,
    head = '',
    tail = '')


class DropFile(Button):
    def __init__(self, **kwargs):
        super(DropFile, self).__init__(**kwargs)

        # get app instance to add function from widget
        app = App.get_running_app()

        # add function to the list
        app.drops.append(self.on_dropfile)

    def on_dropfile(self, widget, filename):
        # a function catching a dropped file
        # if it's dropped in the widget's area
        if self.collide_point(*Window.mouse_pos):
            # on_dropfile's filename is bytes (py3)
            ActiveTable.filename = filename.decode('utf-8')
            ActiveTable.file_type, ActiveTable.valid_file, ActiveTable.head, ActiveTable.tail = (
                determine_file_type(ActiveTable.filename))
            self.text = ('Raw asset file:\n' + ActiveTable.tail
                         + '\n\nFile source:\n' + ActiveTable.file_type)


class Pivotr(App):

    def build(self):
        self.title = 'Pivotr'
        self.icon = 'pivotr.png'
        SL = StackLayout(orientation='tb-rl')
        Window.size = (400, 500)
        Window.clearcolor = (.25, .25, .25, 1)
        Window.bind(on_cursor_enter=lambda *__:Window.raise_window())
        # set an empty list that will be later populated
        # with functions from widgets themselves
        self.drops = []

        # bind handling function to 'on_dropfile'
        Window.bind(on_dropfile=self.handledrops)
        drop = DropFile(text='Drop raw asset file here',
                        font_size=20,
                        bold=True,
                        size_hint=(.8, .95),
                        background_normal='',
                        background_color=(0, .475, .42, 1))
        # Creating Multiple Buttons
        btn_hk = Button(text="HK",
                        font_size=20,
                        size_hint=(.2, .186),
                        background_normal='',
                        background_color=(1, 1, 1, .6),
                        color=(.25, .25, .25, 1),
                        bold=True)
        btn_mh = Button(text="MH",
                        font_size=20,
                        size_hint=(.2, .186),
                        background_normal='',
                        background_color=(1, 1, 1, .6),
                        color=(.25, .25, .25, 1),
                        bold=True)
        btn_ms = Button(text="MS",
                        font_size=20,
                        size_hint=(.2, .186),
                        background_normal='',
                        background_color=(1, 1, 1, .6),
                        color=(.25, .25, .25, 1),
                        bold=True)
        btn_sb = Button(text="SB",
                        font_size=20,
                        size_hint=(.2, .186),
                        background_normal='',
                        background_color=(1, 1, 1, .6),
                        color=(.25, .25, .25, 1),
                        bold=True)
        btn_ks = Button(text="KS",
                        font_size=20,
                        size_hint=(.2, .186),
                        background_normal='',
                        background_color=(1, 1, 1, .6),
                        color=(.25, .25, .25, 1),
                        bold=True)
        cred = Label(text='© 2019 by Cameron Abel',
                     size_hint=(.6, .05),
                     halign='right',)
        borda = Label(size_hint=(.005, .95))
        bordb = Label(size_hint=(.2, .005))
        bordc = Label(size_hint=(.2, .005))
        bordd = Label(size_hint=(.2, .005))
        borde = Label(size_hint=(.2, .005))
        bordf = Label(size_hint=(.2, .005))

        # adding widgets
        SL.add_widget(btn_hk)
        btn_hk.bind(on_press=lambda x: hk_boot())
        SL.add_widget(bordb)
        SL.add_widget(btn_mh)
        btn_mh.bind(on_press=lambda x: mh_boot())
        SL.add_widget(bordc)
        SL.add_widget(btn_ms)
        btn_ms.bind(on_press=lambda x: ms_boot())
        SL.add_widget(bordd)
        SL.add_widget(btn_sb)
        SL.add_widget(borde)
        SL.add_widget(btn_ks)
        SL.add_widget(borda)
        SL.add_widget(drop)
        SL.add_widget(cred)

        # returning widgets
        return SL

    def handledrops(self, *args):
        # this will execute each function from list with arguments from
        # Window.on_dropfile
        for func in self.drops:
            func(*args)

 
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
    else:
        file_type = 'Incompatible file source'
        valid_file = False
    return file_type, valid_file, head, tail


def is_jh(tail):
    if tail[-7:] == 'YTD.TXT':
        return True
    return False


def is_rkdirect(filename):
    try:
        b = open(filename, 'r').readlines()
        if b[0][:6] == 'ICU ID':
            return True
        else:
            return False
    except:
        return False







def mh_boot():
    if ActiveTable.file_type == 'John Hancock':
        contractnum = int(ActiveTable.tail[:-7])

        try:
            contractlist = pd.read_csv('contracts.csv')
            contractname = contractlist.loc[contractlist['Contract Number'] == contractnum]
            contractname.reset_index(inplace=True)
            contractname = contractname.at[0, 'Contract Name']
        except:
            contractname = ''
        trunctable, pye = jh_pivot(ActiveTable.filename)
        mh_prep(trunctable, pye, contractname)
    elif ActiveTable.file_type == 'Voya':
        trunctable, pye, plan_name = voya_pivot(ActiveTable.filename)
        mh_prep(trunctable, pye, plan_name)
    elif ActiveTable.file_type == 'TRC':
        trunctable, pye, plan_name = trc_pivot(ActiveTable.filename)
        mh_prep(trunctable, pye, plan_name)
    elif ActiveTable.file_type == 'RK Direct':
        trunctable, pye, plan_name = rkd_pivot(ActiveTable.filename)
        mh_prep(trunctable, pye, plan_name)


def ms_boot():
    if ActiveTable.file_type == 'John Hancock':
        contractnum = int(ActiveTable.tail[:-7])

        try:
            contractlist = pd.read_csv('contracts.csv')
            contractname = contractlist.loc[contractlist['Contract Number'] == contractnum]
            contractname.reset_index(inplace=True)
            contractname = contractname.at[0, 'Contract Name']
        except:
            contractname = ''
        trunctable, pye = jh_pivot(ActiveTable.filename)
        ms_prep(trunctable, pye, contractname)
    elif ActiveTable.file_type == 'Voya':
        trunctable, pye, plan_name = voya_pivot(ActiveTable.filename)
        ms_prep(trunctable, pye, plan_name)
    elif ActiveTable.file_type == 'TRC':
        trunctable, pye, plan_name = trc_pivot(ActiveTable.filename)
        ms_prep(trunctable, pye, plan_name)
    elif ActiveTable.file_type == 'RK Direct':
        trunctable, pye, plan_name = rkd_pivot(ActiveTable.filename)
        ms_prep(trunctable, pye, plan_name)


def hk_boot():
    if ActiveTable.file_type == 'John Hancock':
        contractnum = int(ActiveTable.tail[:-7])

        try:
            contractlist = pd.read_csv('contracts.csv')
            contractname = contractlist.loc[contractlist['Contract Number'] == contractnum]
            contractname.reset_index(inplace=True)
            contractname = contractname.at[0, 'Contract Name']
        except:
            contractname = ''
        trunctable, pye = jh_pivot(ActiveTable.filename)
        hk_prep(trunctable, pye, contractname)
    elif ActiveTable.file_type == 'Voya':
        trunctable, pye, plan_name = voya_pivot(ActiveTable.filename)
        hk_prep(trunctable, pye, plan_name)
    elif ActiveTable.file_type == 'TRC':
        trunctable, pye, plan_name = trc_pivot(ActiveTable.filename)
        hk_prep(trunctable, pye, plan_name)
    elif ActiveTable.file_type == 'RK Direct':
        trunctable, pye, plan_name = rkd_pivot(ActiveTable.filename)
        hk_prep(trunctable, pye, plan_name)


def mh_prep(trunctable, pye, contractname):
    """writes pivot table to excel with MH preferred formatting"""
    trunctable = trunctable[['SSN'] +
                            [col for col in trunctable.columns if col != 'SSN']]
    trunctable = trunctable[['Name'] +
                            [col for col in trunctable.columns if col != 'Name']]
    trunctable = trunctable[['Source'] +
                            [col for col in trunctable.columns if col != 'Source']]

    trunctable = trunctable[[col for col in trunctable.columns if col != 'End Bal'] +
                            ['End Bal']]
    trunctable = trunctable.sort_values(by=['Source', 'Name'])

    outputfile = ActiveTable.head + '\\' + pye[-4:] + ' ' + contractname + ' Data Prep - MH.xlsx'

    writer = pd.ExcelWriter(outputfile, engine='xlsxwriter')
    workbook = writer.book
    worksheet = workbook.add_worksheet('Source Data')
    writer.sheets['Source Data'] = worksheet
    trunctable.to_excel(writer, sheet_name='Source Data', index=False,
                        startrow=4, startcol=0, header=False)

    formatnum = workbook.add_format({'num_format': '#,##0.00_)'})
    formatssn = workbook.add_format({'num_format': '000-00-0000'})
    formattitle = workbook.add_format({'bold': True, 'font_size': 14})
    formatheader = workbook.add_format({'bold': True, 'bottom':0, 'top':0,
                                        'left':0, 'right':0, 'underline':True})
    formattotal = workbook.add_format({'bold': True, 'bottom':1, 'top':1,
                                       'num_format': '#,##0.00_)'})
    formatdouble = workbook.add_format({'top':6})

    # number format and column width
    srclen = trunctable['Source'].map(len).max()
    worksheet.set_column('C:C', 12, formatssn)
    worksheet.set_column('B:B', 26)
    worksheet.set_column('A:A', srclen)
    worksheet.set_column('D:O', 12, formatnum)

    # writes and formats sheet title
    worksheet.write(0, 0, contractname, formattitle)
    worksheet.write(1, 0, ActiveTable.file_type + ' data for PYE ' + pye, formattitle)

    # writes and formats headers
    worksheet.set_row(2, 15, formatheader)
    worksheet.set_row(3, 15, formatheader)
    for col_num, value in enumerate(trunctable.columns.values):
        worksheet.write(3, col_num, value, formatheader)

    # writes and formats totals
    totalrow = trunctable.shape[0]+5
    worksheet.write(totalrow, 0, 'TOTALS', formattotal)
    worksheet.write(totalrow, 1, '', formattotal)
    worksheet.write(totalrow, 2, '', formattotal)

    col_letters = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R']
    colindex = 0
    for _ in trunctable:
        if colindex > 2:
            tindex = col_letters.pop(0)
            tcell = tindex + str(totalrow)
            formula = '=sum(' + tindex + '5:' + tcell + ')'
            worksheet.write(totalrow - 1, colindex, '', formatdouble)
            worksheet.write(totalrow, colindex, formula, formattotal)
        colindex += 1

    worksheet.freeze_panes(4, 0)

    writer.save()


def ms_prep(trunctable, pye, contractname):
    """writes table to excel with MS preferred format"""
    trunctable = trunctable[['SSN'] +
                            [col for col in trunctable.columns if col != 'SSN']]
    trunctable = trunctable[['Name'] +
                            [col for col in trunctable.columns if col != 'Name']]

    trunctable = trunctable[[col for col in trunctable.columns if col != 'End Bal'] +
                            ['End Bal']]
    trunctable = trunctable.sort_values(by=['Source', 'Name'])

    outputfile = ActiveTable.head + '\\' + pye[-4:] + ' ' + contractname + ' Data Prep - MS.xlsx'

    writer = pd.ExcelWriter(outputfile, engine='xlsxwriter')
    workbook = writer.book

    sourcelist = trunctable['Source'].unique().tolist()

    formatnum = workbook.add_format({'num_format': '#,##0.00_)'})
    formatssn = workbook.add_format({'num_format': '000-00-0000'})
    formattitle = workbook.add_format({'bold': True, 'font_size': 14})
    formatheader = workbook.add_format({'bold': True, 'bottom':0, 'top':0,
                                        'left':0, 'right':0, 'underline':True})
    formattotal = workbook.add_format({'bold': True, 'bottom':1, 'top':1,
                                       'num_format': '#,##0.00_)'})
    formatdouble = workbook.add_format({'top':6})

    tab_dict = {'Profit Sharing' : 'PS', 'After-tax Rollover' : 'A-T Roll',
                'Matching' : 'Match', 'Roth Rollover' : 'Roth Roll',
                'Rollover' : 'Roll', 'Simple IRA Rollover' : 'Simp Roll',
                '403(b) Rollover' : '403(b)', 'Deferrals' : 'Def',
                'Employee Pre Tax' : 'Def', 'Employer Matching' : 'Match',
                'Employer Profit Sharing' : 'PS'}

    for tab_name in sourcelist:
        if tab_name in tab_dict:
            tab = tab_dict[tab_name]
        else: tab = tab_name
        temp_table = trunctable.loc[trunctable.Source == tab_name]
        temp_table = temp_table.drop(columns=['Source'])
        temp_table = temp_table.loc[:, (temp_table != 0).any(axis=0)]
        temp_table = temp_table.loc[:, (temp_table != '').any(axis=0)]
        worksheet = workbook.add_worksheet(tab)
        writer.sheets[tab_name] = worksheet
        temp_table.to_excel(writer, sheet_name=tab_name, index=False,
                            startrow=4, startcol=0, header=False)
        worksheet.set_column('B:B', 12, formatssn)
        worksheet.set_column('A:A', 26)
        worksheet.set_column('C:O', 12, formatnum)

        worksheet.write(0, 0, contractname, formattitle)
        worksheet.write(1, 0, ActiveTable.file_type
                        + ' data for PYE ' + pye + ' - ' + tab_name, formattitle)

        # writes and formats headers
        worksheet.set_row(2, 15, formatheader)
        worksheet.set_row(3, 15, formatheader)
        for col_num, value in enumerate(temp_table.columns.values):
            worksheet.write(3, col_num, value, formatheader)

        # writes and formats totals
        totalrow = temp_table.shape[0]+5
        worksheet.write(totalrow, 0, 'TOTALS', formattotal)
        worksheet.write(totalrow, 1, '', formattotal)
        worksheet.write(totalrow, 2, '', formattotal)

        col_letters = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J',
                       'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R']
        colindex = 0
        for _ in temp_table:
            if colindex > 1:
                tindex = col_letters.pop(0)
                tcell = tindex + str(totalrow)
                formula = '=sum(' + tindex + '5:' + tcell + ')'
                worksheet.write(totalrow - 1, colindex, '', formatdouble)
                worksheet.write(totalrow, colindex, formula, formattotal)
            colindex += 1

        worksheet.freeze_panes(4, 0)
    writer.save()


def hk_prep(trunctable, pye, contractname):
    """writes pivot table to excel with HK preferred formatting"""
    trunctable = trunctable[['Source'] +
                            [col for col in trunctable.columns if col != 'Source']]
    trunctable = trunctable[['Name'] +
                            [col for col in trunctable.columns if col != 'Name']]
    trunctable = trunctable[['SSN'] +
                            [col for col in trunctable.columns if col != 'SSN']]
    trunctable = trunctable[[col for col in trunctable.columns if col != 'End Bal'] +
                            ['End Bal']]
    trunctable = trunctable.sort_values(by=['Source', 'Name'])

    outputfile = ActiveTable.head + '\\' + pye[-4:] + ' ' + contractname + ' Data Prep - HK.xlsx'

    writer = pd.ExcelWriter(outputfile, engine='xlsxwriter')
    workbook = writer.book
    worksheet = workbook.add_worksheet('Source Data')
    writer.sheets['Source Data'] = worksheet
    trunctable.to_excel(writer, sheet_name='Source Data', index=False,
                        startrow=4, startcol=0, header=False)

    formatnum = workbook.add_format({'num_format': '#,##0.00_);(#,##0.00)'})
    formatssn = workbook.add_format({'num_format': '000-00-0000'})
    formattitle = workbook.add_format({'bold': True, 'font_size': 14})
    formatheader = workbook.add_format({'bold': True, 'bottom':0, 'top':0,
                                        'left':0, 'right':0, 'underline':True})
    formattotal = workbook.add_format({'bold': True, 'bottom':1, 'top':1,
                                       'num_format': '#,##0.00_)'})
    formatdouble = workbook.add_format({'top':6})

    # number format and column width
    srclen = trunctable['Source'].map(len).max()
    worksheet.set_column('A:A', 12, formatssn)
    worksheet.set_column('B:B', 26)
    worksheet.set_column('C:C', srclen)
    worksheet.set_column('D:O', 12, formatnum)

    # writes and formats sheet title
    worksheet.write(0, 1, contractname, formattitle)
    worksheet.write(1, 1, ActiveTable.file_type + ' data for PYE ' + pye, formattitle)

    # writes and formats headers
    worksheet.set_row(2, 15, formatheader)
    worksheet.set_row(3, 15, formatheader)
    for col_num, value in enumerate(trunctable.columns.values):
        worksheet.write(3, col_num, value, formatheader)

    # writes and formats totals
    totalrow = trunctable.shape[0]+5
    worksheet.write(totalrow, 0, 'TOTALS', formattotal)
    worksheet.write(totalrow, 1, '', formattotal)
    worksheet.write(totalrow, 2, '', formattotal)

    col_letters = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R']
    colindex = 0
    for _ in trunctable:
        if colindex > 2:
            tindex = col_letters.pop(0)
            tcell = tindex + str(totalrow)
            formula = '=sum(' + tindex + '5:' + tcell + ')'
            worksheet.write(totalrow - 1, colindex, '', formatdouble)
            worksheet.write(totalrow, colindex, formula, formattotal)
        colindex += 1

    worksheet.freeze_panes(4, 0)

    writer.save()


def namegen():
    """Returns a df of names and SSNs."""
    try:
        ascname = pd.read_clipboard('\t', header=1)
        ascname.rename(columns={'SocSecNum':'SSN'}, inplace=True)
        ascname.rename(columns={'Employee Name':'Name'}, inplace=True)
        ascname['Name'] = ascname['Name'].str.strip()
        ascname['SSN'] = ascname['SSN'].str.replace('-', '')
        ascname['SSN'] = ascname['SSN'].astype(str)
        ascname['SSN'] = ascname['SSN'].str.zfill(9)
        
    except:
        try:
            ascname = pd.read_csv(ActiveTable.head + '\\names.csv', names=['SSN', 'Name'])
            ascname['Name'] = ascname['Name'].str.strip()
            ascname['SSN'] = ascname['SSN'].str.replace('-', '')
            ascname = ascname.drop([0, 1])
            ascname['SSN'] = ascname['SSN'].astype(str)
            ascname['SSN'] = ascname['SSN'].str.zfill(9)

        except:
            ascname = pd.DataFrame([['', '']], columns=['SSN', 'Name'])

    try:
        sub = 'Census_Summary'
        census_report = next((s for s in os.listdir(ActiveTable.head) if sub in s), None)
        census_report = ActiveTable.head + '\\' + census_report
        nametable = pd.read_csv(census_report,
                                usecols=[2, 3, 4],
                                names=['SSN', 'First', 'Last'])

    except:
        nametable = pd.DataFrame([['', '', '']], columns=['SSN', 'First', 'Last'])

    nametable['Name'] = nametable['Last'] + ', ' + nametable['First']
    nametable['SSN'] = nametable['SSN'].str.replace('-', '')
    nametable = nametable.drop([0]).drop(columns=['First', 'Last'])
    nametable['SSN'] = nametable['SSN'].astype(str)
    nametable['SSN'] = nametable['SSN'].str.zfill(9)
    nametable = pd.merge(nametable, ascname, on='SSN', how='outer')
    nametable['Name_x'].update(nametable['Name_y'])
    nametable = nametable.rename(columns={'Name_x':'Name'}).drop(columns=['Name_y'])
    nametable['Name'] = nametable.Name.str.title()

    return nametable


def parsename(fullname):
    name = HumanName(fullname)
    return name.last + ', ' + name.first


def parsemyt(myt):
    src_dict = {'103' : 'Profit Sharing',
                '101' : 'Deferrals',
                '137' : 'Pension',
                '104' : 'Matching',
                '119' : 'ESOP Rollover',
                '112' : 'Rollover'}
    return src_dict.get(myt, myt)


def jh_pivot(filename):
    """Returns a pivot table df based on JH assets."""
    sources = pd.DataFrame([[0, 'Profit Sharing'],
                            [3, 'PS to Forf'],
                            [4, 'Deferrals'],
                            [5, 'SH Match'],
                            [6, 'Rollover'],
                            [7, 'SH Match to Forf'],
                            [8, 'SHNEC'],
                            [12, 'Roth'],
                            [21, 'Rollover'],
                            [23, '403(b) Rollover'],
                            [25, 'Simple IRA Rollover'],
                            [27, 'After-tax Rollover'],
                            [29, 'Roth Rollover'],
                            [55, 'Matching'],
                            [71, 'QACA Match']],
                           columns=['Source Code', 'Source'])

    transactions = pd.DataFrame([[0, 'Beg Bal'],
                                 [1, 'Contrib'],
                                 [2, 'Distrib'],
                                 [11, 'G/L'],
                                 [5, 'TRF'],
                                 [7, 'Loan Issue'],
                                 [8, 'Loan Pay'],
                                 [120, 'Roll In'],
                                 [999, 'End Bal']],
                                columns=['Trans Code', 'Trans'])

    worktable = pd.read_fwf(filename,
                            widths=[8, 9, 3, 6, 2, 3, 12, 15, 12, 10],
                            names=['Contract Number', 'SSN', 'Trans Code',
                                   'Period End', 'Source Code', 'Fund', 'Amount',
                                   'Units', 'Loan Interest', 'Loan Charge'])
    worktable = worktable.merge(transactions, on='Trans Code', how='left')
    pye = str(worktable.at[1, 'Period End']).strip()
    pye = pye[:-2] + '20' + pye[-2:]
    pye = pye[:-6] + '/' + pye[-6:-4] + '/' + pye[-4:]

    pivoted_data = pd.pivot_table(worktable,
                                  values=['Amount'],
                                  index=['SSN', 'Source Code'],
                                  columns=['Trans'],
                                  aggfunc=np.sum,
                                  fill_value=0)

    pivoted_data.round(2)
    pivoted_data.columns = [s2 for (s1, s2) in pivoted_data.columns.tolist()]
    pivoted_data.reset_index(inplace=True)

    nametable = namegen()

    pivoted_data['SSN'] = pivoted_data['SSN'].astype(str)
    pivoted_data['SSN'] = pivoted_data['SSN'].str.zfill(9)
    cotable = pivoted_data.merge(nametable, on='SSN', how='left')

    trunctable = cotable.loc[:, (cotable != 0).any(axis=0)]
    trunctable = trunctable.merge(sources, on='Source Code', how='left')
    trunctable = trunctable.drop(columns=['Source Code'])

    trunctable['SSN'] = trunctable['SSN'].astype(int)
    columns = list(trunctable)
    for col in columns:
        if col not in ('Beg Bal', 'End Bal'):
            trunctable[col].replace(0, '', inplace=True)

    return trunctable, pye


def voya_pivot(filename):
    """Returns a pivot table df based on Voya assets."""

    worktable = pd.read_excel(filename, sheet_name=3, usecols=('B:C,E:F,L:AB,AG'))
    plan_name = str(worktable.at[1, 'Plan Name']).strip().replace('(K)', '(k)')
    pye = str(worktable.at[1, 'End Date']).strip()
    worktable.drop(columns=['Plan Name', 'End Date'], inplace=True)
    worktable['Name'] = worktable['Name'].str.title()
    worktable['G/L'] = worktable['Dividends Earnings'] + worktable['Gain/Loss']
    worktable['TRF'] = worktable['Fund Transfers'] + worktable['Internal Transfers']
    worktable['Fee'] = worktable['Fees'] + worktable['TPA Fees']
    worktable.drop(columns=['Dividends Earnings', 'Gain/Loss',
                            'Fund Transfers', 'Internal Transfers',
                            'Fees', 'TPA Fees'], inplace=True)
    col_name_dict = {'Source Name' : 'Source',
                     'Participant Number' : 'SSN',
                     'Beginning Balance' : 'Beg Bal',
                     'Contributions' : 'Contrib',
                     'Takeover Contribution' : 'Takeover',
                     'Loan Repayments' : 'Loan Pmt',
                     'Loan Repay Principal' : 'Principal',
                     'Loan Repay Interest' : 'Interest',
                     'Withdrawals' : 'Distrib',
                     'Forfeitures' : 'Forf',
                     'Ending Balance' : 'End Bal'}

    for col in list(worktable):
        worktable.rename(columns={col:col_name_dict.get(col, col)}, inplace=True)
    worktable = worktable[~worktable.Source.str.contains('Loans')]
    worktable = worktable[~worktable.Name.str.contains('Forfeiture Account')]
    trunctable = worktable.loc[:, (worktable != 0).any(axis=0)]

    trunctable['SSN'] = trunctable['SSN'].str.replace('-', '').astype(int)
    trunctable['Name'] = trunctable['Name'].str.replace(',', ', ').replace(',  ', ', ')
    columns = list(trunctable)
    for col in columns:
        if col not in ('Beg Bal', 'End Bal'):
            trunctable[col].replace(0, '', inplace=True)

    return (trunctable, pye, plan_name)


def trc_pivot(filename):
    sources = pd.DataFrame([['Profit Sharing', 'Profit Sharing'],
                            ['Deferred Salary', 'Deferrals'],
                            ['Rollover', 'Rollover'],
                            ['Safe Harbor Non-elective', 'SHNEC'],
                            ['Roth Salary Deferral', 'Roth'],
                            ['Company Match', 'Matching']],
                            columns=['Source Name','Source'])
    
    b = open(filename, 'r').readlines()
    #line = b[1]
    pye = b[1][-11:-1]
    
    worktable = pd.read_csv(filename, sep='\t',
                            usecols=[4,5,6,12,13,14,15,16,17,18,19,20,21,22,23],
                            names=['First', 'Last', 'SSN', 'Source Name', 'Beg Bal', 'Contrib',
                            'Loan', 'Forf Alloc', 'TRF', 'Distrib', 'Forf', 'Fees', 'Other',
                            'G/L', 'End Bal'], header=None,
                            skiprows=range(0,5))
    worktable.drop(worktable.tail(1).index, inplace=True)
    worktable.reset_index(drop=True, inplace=True)

    worktable['Name'] = worktable['Last'] + ', ' + worktable['First']
    worktable['SSN'] = worktable['SSN'].astype(int)

    pivoted_data = pd.pivot_table(worktable,
                                  index = ['SSN','Source Name'],
                                  aggfunc=np.sum,
                                  fill_value=0)
    pivoted_data.reset_index(inplace=True)
    
    nametable = worktable[['SSN'] + ['Name']].drop_duplicates(subset='SSN')

    pivoted_data = pivoted_data.merge(nametable, on='SSN', how='left')
    pivoted_data = pivoted_data.merge(sources, on='Source Name', how='left')

    pivoted_data['Source Name'].update(pivoted_data['Source'])
    pivoted_data = pivoted_data.drop(columns=['Source']).rename(columns={'Source Name':'Source'})

    pivoted_data = pivoted_data[['SSN', 'Name', 'Source', 'Beg Bal', 'Contrib', 
    'Loan', 'Forf Alloc', 'TRF', 'Distrib', 'Forf', 'Fees', 'Other', 
    'G/L', 'End Bal']].dropna(how='all')

    trunctable = pivoted_data.sort_values(by=['Source', 'Name'])
    
    trunctable = trunctable.loc[:, (trunctable != 0).any(axis=0)]
    plan_name = ''
    return trunctable, pye, plan_name


def rkd_pivot(filename):
    """Returns a pivot table df based on American Funds RKDirect assets."""

    worktable = pd.read_csv(filename, sep=',',
                            usecols=[1, 2, 3, 4, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 26],
                            names=['planID', 'SSN', 'fullname', 'MYT', 'Beg Bal', 'Conv', 'Contrib',
                            'Dividends', 'Gain/Loss', 'Exch', 'Fees', 'Forf', 'Distrib', 'Other',
                            'TRFIN', 'TRFOUT', 'Loan TRFIN', 'Loan TRFOUT', 'Int', 'Ins', 'End Bal', 'PYE'], header=None,
                            skiprows=1)
    pye = worktable['PYE'][1]
    planID = worktable['planID'][1]
    worktable['Name'] = worktable['fullname'].apply(parsename).str.title()
    worktable['G/L'] = worktable['Dividends'] + worktable['Gain/Loss']
    worktable['TRF'] = worktable['Exch'] + worktable['TRFIN'] + worktable['TRFOUT']
    worktable.drop(columns=['fullname', 'Dividends', 'Gain/Loss', 'planID',
                            'Exch', 'TRFIN', 'TRFOUT', 'PYE'], inplace=True)
    pivoted_data = pd.pivot_table(worktable,
                                  index = ['Name', 'SSN', 'MYT'],
                                  aggfunc=np.sum,
                                  fill_value=0)
    pivoted_data.reset_index(inplace=True)

    trunctable = pivoted_data.loc[:, (pivoted_data != 0).any(axis=0)]
    trunctable['MYT'] = trunctable['MYT'].astype(str)
    trunctable['Source'] = trunctable['MYT'].apply(parsemyt)
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
    except:
        planname = ''
    
    
    
    return trunctable, pye, planname



if __name__ == '__main__':
    Pivotr().run()
