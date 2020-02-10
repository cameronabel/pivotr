"""Produces formatted Excel files containing pivot tables from investment company source files"""

from types import SimpleNamespace

import kivy
import pandas as pd
from kivy.app import App
from kivy.config import Config
from kivy.core.window import Window
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.stacklayout import StackLayout

from pivotr.common import helpers as hp
from pivotr.common import pivoters as pv

kivy.require("1.9.1")

__author__ = 'Cameron Abel'
__copyright__ = 'Copyright 2019 - 2020'
__credits__ = ['Cameron Abel', 'Marissa Hartwig', 'Natalie Fultz']
__license__ = 'MIT'
__maintainer__ = 'Cameron Abel'
__email__ = 'cameronabel@gmail.com'
__status__ = 'Production'

Config.set('graphics', 'width', '400')
Config.set('graphics', 'height', '400')
Config.write()

ActiveTable: SimpleNamespace = SimpleNamespace(
    filename='',
    file_type='',
    valid_file=False,
    head='',
    tail='',
    tail_disp='')


class DropFile(Button):
    def __init__(self, **kwargs):
        super(DropFile, self).__init__(**kwargs)

        # get app instance to add function from widget
        app = App.get_running_app()

        # add function to the list
        app.drops.append(self.on_dropfile)

    def on_dropfile(self, _, filename):

        if self.collide_point(*Window.mouse_pos):
            ActiveTable.filename = filename.decode('utf-8')
            ActiveTable.file_type, ActiveTable.valid_file, ActiveTable.head, ActiveTable.tail = (
                hp.determine_file_type(ActiveTable.filename))
            if len(ActiveTable.tail) > 18:
                ActiveTable.tail_disp = ActiveTable.tail[:8] + '...' + ActiveTable.tail[-7:]
            else:
                ActiveTable.tail_disp = ActiveTable.tail
            self.text = ('Raw asset file:\n' + ActiveTable.tail_disp
                         + '\n\nFile source:\n' + ActiveTable.file_type)


class Pivotr(App):

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.drops = []
        Window.bind(mouse_pos=self.on_mouse_pos)

    def build(self):
        self.title = 'Pivotr'
        self.icon = 'img/pivotr.png'
        SL = StackLayout(orientation='tb-rl')
        Window.size = (400, 400)
        Window.clearcolor = (.25, .25, .25, 1)
        Window.borderless = False

        Window.bind(on_cursor_enter=lambda *__: Window.raise_window())
        # set an empty list that will be later populated
        # with functions from widgets themselves

        # bind handling function to 'on_dropfile'
        Window.bind(on_dropfile=self.handledrops)
        drop = DropFile(text='Drop raw asset file here',
                        font_size=20,
                        bold=True,
                        size_hint=(.6, .95),
                        background_normal='',
                        background_color=(0, .475, .42, 1),
                        background_down='img/drop.png')
        # Creating Multiple Buttons
        btn_stacked = Button(text="Stacked\n\n",
                             font_size=20,
                             size_hint=(.4, .311666),
                             background_normal='img/stk.png',
                             background_color=(1, 1, 1, 1),
                             color=(.2, .2, .2, 1),
                             bold=True,
                             background_down='img/stkd.png')
        btn_tabbed = Button(text="Tabbed",
                            font_size=20,
                            size_hint=(.4, .311666),
                            background_normal='img/tabbed2.png',
                            background_color=(1, 1, 1, 1),
                            color=(.2, .2, .2, 1),
                            bold=True,
                            background_down='img/tabbed2d.png')
        btn_wide = Button(text="Wide\ncoming\nsoon",
                          font_size=20,
                          size_hint=(.4, .311666),
                          background_normal='',
                          background_color=(1, 1, 1, 0),
                          color=(.4, .4, .4, 1),
                          bold=True,
                          background_down='img/clear.png')
        cred = Label(text='© 2019 - 2020 by Cameron Abel',
                     size_hint=(.2, .05),
                     halign='right', )
        borda = Label(size_hint=(.005, .95))
        bordb = Label(size_hint=(.2, .005))
        bordc = Label(size_hint=(.2, .005))
        bordd = Label(size_hint=(.2, .005))

        # adding widgets
        SL.add_widget(bordb)
        SL.add_widget(btn_stacked)
        btn_stacked.bind(on_press=lambda x: stacked_boot())
        SL.add_widget(bordc)
        SL.add_widget(btn_tabbed)
        btn_tabbed.bind(on_press=lambda x: tabbed_boot())
        SL.add_widget(bordd)
        SL.add_widget(btn_wide)
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

    def on_mouse_pos(self, _, pos):
        if pos[0] > 240 and pos[1] > 146:
            Window.set_system_cursor('hand')
        else:
            Window.set_system_cursor('arrow')


# noinspection PyBroadException
def stacked_boot():
    if ActiveTable.file_type == 'John Hancock':
        contract_num = int(ActiveTable.tail[:-7])

        try:
            contract_list = pd.read_csv('contracts.csv')
            contract_name = contract_list.loc[contract_list['Contract Number'] == contract_num]
            contract_name.reset_index(inplace=True)
            contract_name = contract_name.at[0, 'Contract Name']
        except Exception:
            contract_name = ''
        trunctable, pye = pv.jh_pivot(ActiveTable.filename, ActiveTable.head)
        stacked_prep(trunctable, pye, contract_name)

    elif ActiveTable.file_type == 'Voya':
        trunctable, pye, plan_name = pv.voya_pivot(ActiveTable.filename)
        stacked_prep(trunctable, pye, plan_name)
    elif ActiveTable.file_type == 'TRC':
        trunctable, pye, plan_name = pv.trc_pivot(ActiveTable.filename)
        stacked_prep(trunctable, pye, plan_name)
    elif ActiveTable.file_type == 'RK Direct\nAviator':
        trunctable, pye, plan_name = pv.rkd_pivot(ActiveTable.filename)
        stacked_prep(trunctable, pye, plan_name)
    elif ActiveTable.file_type == 'Empower':
        trunctable, pye, plan_name = pv.emp_pivot(ActiveTable.filename, ActiveTable.tail)
        stacked_prep(trunctable, pye, plan_name)
    elif ActiveTable.file_type == 'Principal':
        trunctable, pye, plan_name = pv.prin_pivot(ActiveTable.filename)
        stacked_prep(trunctable, pye, plan_name)


# noinspection PyBroadException
def tabbed_boot():
    if ActiveTable.file_type == 'John Hancock':
        contract_num = int(ActiveTable.tail[:-7])

        try:
            contract_list = pd.read_csv('contracts.csv')
            contract_name = contract_list.loc[contract_list['Contract Number'] == contract_num]
            contract_name.reset_index(inplace=True)
            contract_name = contract_name.at[0, 'Contract Name']
        except Exception:
            contract_name = ''
        trunctable, pye = pv.jh_pivot(ActiveTable.filename, ActiveTable.head)
        tabbed_prep(trunctable, pye, contract_name)

    elif ActiveTable.file_type == 'Voya':
        trunctable, pye, plan_name = pv.voya_pivot(ActiveTable.filename)
        tabbed_prep(trunctable, pye, plan_name)
    elif ActiveTable.file_type == 'TRC':
        trunctable, pye, plan_name = pv.trc_pivot(ActiveTable.filename)
        tabbed_prep(trunctable, pye, plan_name)
    elif ActiveTable.file_type == 'RK Direct':
        trunctable, pye, plan_name = pv.rkd_pivot(ActiveTable.filename)
        tabbed_prep(trunctable, pye, plan_name)
    elif ActiveTable.file_type == 'Empower':
        trunctable, pye, plan_name = pv.emp_pivot(ActiveTable.filename, ActiveTable.tail)
        tabbed_prep(trunctable, pye, plan_name)
    elif ActiveTable.file_type == 'Principal':
        trunctable, pye, plan_name = pv.prin_pivot(ActiveTable.filename)
        tabbed_prep(trunctable, pye, plan_name)


# noinspection PyBroadException
def stacked_prep(trunctable, pye, contract_name):
    """Writes pivot table to excel with stacked formatting"""
    trunctable = hp.end_bal_check(trunctable)
    trunctable = trunctable[['SSN'] +
                            [col for col in trunctable.columns if col != 'SSN']]
    trunctable = trunctable[['Name'] +
                            [col for col in trunctable.columns if col != 'Name']]
    trunctable = trunctable[['Source'] +
                            [col for col in trunctable.columns if col != 'Source']]

    try:
        trunctable = trunctable[[col for col in trunctable.columns if col != 'End Bal'] +
                                ['End Bal']]
    except Exception:
        pass
    trunctable = trunctable.sort_values(by=['Source', 'Name'])
    if pye == '':
        pye_space = ''
    else:
        pye_space = pye + ' '
    outputfile = ActiveTable.head + '\\' + pye_space + contract_name + ' Data Prep.xlsx'

    writer = pd.ExcelWriter(outputfile, engine='xlsxwriter')
    workbook = writer.book
    worksheet = workbook.add_worksheet('Source Data')
    writer.sheets['Source Data'] = worksheet
    trunctable.to_excel(writer, sheet_name='Source Data', index=False,
                        startrow=4, startcol=0, header=False)

    formatnum = workbook.add_format({'num_format': '#,##0.00_)'})
    formatssn = workbook.add_format({'num_format': '000-00-0000'})
    formattitle = workbook.add_format({'bold': True, 'font_size': 14})
    formatheader = workbook.add_format({'bold': True, 'bottom': 0, 'top': 0,
                                        'left': 0, 'right': 0, 'underline': True})
    formattotal = workbook.add_format({'bold': True, 'bottom': 1, 'top': 1,
                                       'num_format': '#,##0.00_)'})
    formatdouble = workbook.add_format({'top': 6})

    # number format and column width
    srclen = trunctable['Source'].map(len).max()
    worksheet.set_column('C:C', 12, formatssn)
    worksheet.set_column('B:B', 26)
    worksheet.set_column('A:A', srclen)
    worksheet.set_column('D:O', 12, formatnum)

    # writes and formats sheet title
    worksheet.write(0, 0, contract_name, formattitle)
    worksheet.write(1, 0, ActiveTable.file_type + ' data for PYE ' + pye, formattitle)

    # writes and formats headers
    worksheet.set_row(2, 15, formatheader)
    worksheet.set_row(3, 15, formatheader)
    for col_num, value in enumerate(trunctable.columns.values):
        worksheet.write(3, col_num, value, formatheader)

    # writes and formats totals
    tot_row = trunctable.shape[0] + 5
    worksheet.write(tot_row, 0, 'TOTALS', formattotal)
    worksheet.write(tot_row, 1, '', formattotal)
    worksheet.write(tot_row, 2, '', formattotal)

    col_letters = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R']
    col_index = 0
    for _ in trunctable:
        if col_index > 2:
            tot_index = col_letters.pop(0)
            tot_cell = tot_index + str(tot_row)
            formula = '=sum(' + tot_index + '5:' + tot_cell + ')'
            worksheet.write(tot_row - 1, col_index, '', formatdouble)
            worksheet.write(tot_row, col_index, formula, formattotal)
        col_index += 1

    worksheet.freeze_panes(4, 0)

    writer.save()


# noinspection PyBroadException
def tabbed_prep(trunctable, pye, contract_name):
    """Writes table to excel with tabbed format"""
    trunctable = hp.end_bal_check(trunctable)
    trunctable = trunctable[['SSN'] +
                            [col for col in trunctable.columns if col != 'SSN']]
    trunctable = trunctable[['Name'] +
                            [col for col in trunctable.columns if col != 'Name']]

    try:
        trunctable = trunctable[[col for col in trunctable.columns if col != 'End Bal'] +
                                ['End Bal']]
    except Exception:
        pass
    trunctable = trunctable.sort_values(by=['Source', 'Name'])
    if pye == '':
        pye_space = ''
    else:
        pye_space = pye + ' '
    outputfile = ActiveTable.head + '\\' + pye_space + contract_name + ' Data Prep.xlsx'

    writer = pd.ExcelWriter(outputfile, engine='xlsxwriter')
    workbook = writer.book

    sourcelist = [str(x) for x in trunctable['Source'].unique().tolist()]

    formatnum = workbook.add_format({'num_format': '#,##0.00_)'})
    formatssn = workbook.add_format({'num_format': '000-00-0000'})
    formattitle = workbook.add_format({'bold': True, 'font_size': 14})
    formatheader = workbook.add_format({'bold': True, 'bottom': 0, 'top': 0,
                                        'left': 0, 'right': 0, 'underline': True})
    formattotal = workbook.add_format({'bold': True, 'bottom': 1, 'top': 1,
                                       'num_format': '#,##0.00_)'})
    formatdouble = workbook.add_format({'top': 6})

    tab_dict = {'Profit Sharing': 'PS', 'After-tax Rollover': 'A-T Roll',
                'Matching': 'Match', 'Roth Rollover': 'Roth Roll',
                'Rollover': 'Roll', 'Simple IRA Rollover': 'Simp Roll',
                '403(b) Rollover': '403(b)', 'Deferrals': 'Def',
                'Employee Pre Tax': 'Def', 'Employer Matching': 'Match',
                'Employer Profit Sharing': 'PS'}

    for tab_name in sourcelist:
        if tab_name in tab_dict:
            tab = tab_dict[tab_name]
        else:
            tab = tab_name
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

        worksheet.write(0, 0, contract_name, formattitle)
        worksheet.write(1, 0, ActiveTable.file_type
                        + ' data for PYE ' + pye + ' - ' + tab_name, formattitle)

        # writes and formats headers
        worksheet.set_row(2, 15, formatheader)
        worksheet.set_row(3, 15, formatheader)
        for col_num, value in enumerate(temp_table.columns.values):
            worksheet.write(3, col_num, value, formatheader)

        # writes and formats totals
        tot_row = temp_table.shape[0] + 5
        worksheet.write(tot_row, 0, 'TOTALS', formattotal)
        worksheet.write(tot_row, 1, '', formattotal)
        worksheet.write(tot_row, 2, '', formattotal)

        col_letters = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J',
                       'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R']
        col_index = 0
        for _ in temp_table:
            if col_index > 1:
                tot_index = col_letters.pop(0)
                tot_cell = tot_index + str(tot_row)
                formula = '=sum(' + tot_index + '5:' + tot_cell + ')'
                worksheet.write(tot_row - 1, col_index, '', formatdouble)
                worksheet.write(tot_row, col_index, formula, formattotal)
            col_index += 1

        worksheet.freeze_panes(4, 0)
    writer.save()


if __name__ == '__main__':
    Pivotr().run()
