from zipfile import ZipFile
import re
import pandas as pd
import sys
from PyQt5.QtWidgets import QApplication, QFileDialog
import xlsxwriter
from dataclasses import dataclass
import numpy as np


@dataclass
class formats():
    """Class holding workbook formats"""
    wb: xlsxwriter.Workbook

    @property
    def header(self):
        return wb.add_format({
            'bg_color': '#538DD5',
            'bold': True,
            'bottom': 1
        })

    @property
    def matching(self):
        return wb.add_format({
            'bg_color': '#FFE699'
        })

    @property
    def matching_diffs(self):
        return wb.add_format({
            'bg_color': '#FFE699',
            'font_color': '#ff0000'
        })

    @property
    def remove(self):
        return wb.add_format({
            'bg_color': '#FF7C80'
        })

    @property
    def add(self):
        return wb.add_format({
            'bg_color': '#A9D08E'
        })


def read_config_rpt(filepath: str) -> dict:
    """Reads an entire CORD Configuration report into a dictionary
    of dataframes with key as csv name without the '_rpt_0000000000.csv'"""
    zip_file = ZipFile(filepath)
    dfs = {re.sub('(.*?)_rpt_\d*\.csv', r'\1', text_file.filename): pd.read_csv(zip_file.open(
        text_file.filename), encoding='unicode_escape', skiprows=3) for text_file in zip_file.infolist() if text_file.filename.endswith('.csv') and text_file.filename != "zzhashtotals.csv"}
    return dfs


def convert_to_date_created(df):
    # deep copy to get rid of chained assignment warning
    df = df[['Date Created']].copy(deep=True)
    df['id'] = df.index
    df = df.set_index('Date Created', drop=True)
    return df


def find_matching(pre_df: pd.DataFrame, post_df: pd.DataFrame) -> pd.DataFrame:
    """Returns list of tuples of (pre_index, post_index) for calcs that have
    the same Date Created between pre and post."""
    pre = convert_to_date_created(pre_df)
    post = convert_to_date_created(post_df)
    match_list = list(set(pre.index) & set(post.index))
    return list(zip(pre.loc[match_list, 'id'].tolist(), post.loc[match_list, 'id'].tolist()))


def find_removed(pre_df: pd.DataFrame, post_df: pd.DataFrame) -> pd.DataFrame:
    """Returns list of matches calcs between pre and post based on Date Created"""
    pre = convert_to_date_created(pre_df)
    post = convert_to_date_created(post_df)
    return [pre.loc[i, 'id'] for i in pre.index if i not in post.index]


def find_added(pre_df: pd.DataFrame, post_df: pd.DataFrame) -> pd.DataFrame:
    """Returns list of matches calcs between pre and post based on Date Created"""
    pre = convert_to_date_created(pre_df)
    post = convert_to_date_created(post_df)
    return [post.loc[i, 'id'] for i in post.index if i not in pre.index]


def compare(before: list, after: list) -> pd.Series:
    """Compares each cell of a pre and post series and returns
    a boolean series indicating where there are differences"""
    return [False if cell == after[i] else True for i, cell in enumerate(before)]


def reformat_tasks(orig: pd.DataFrame, changed: list):
    """Tasks report doesn't include the task lines, reformat it so that
    Tasks changes include changes to task lines."""
    orig = orig.reindex(columns=orig.columns.tolist() +
                        ['Task Lines', 'Object Type', 'Parallel Marker?'])
    df = pd.DataFrame(columns=orig.columns)
    changed_rows, changed_cols = zip(*changed)
    changed_rows = np.array(changed_rows)
    changed_cols = np.array(changed_cols)
    configs = {'change from': pre_config,
               'change to': post_config,
               'remove': pre_config,
               'add': post_config}

    def get_task_lines(action: str, i: int) -> pd.DataFrame:
        """Fetches the config dataframe based on action and row"""
        if action == '':
            return pd.DataFrame()
        return configs[action]['TskLn-'+orig.loc[i, 'Name'].replace(
            ' ', '_').replace('.', '_')]

    new_changed = []
    groups = []

    for i in orig.index:
        start = len(df)
        df.loc[len(df)] = orig.loc[i]
        task_lines = get_task_lines(orig.loc[i, 'Action'], i)
        # store the old task liness for comparison to new
        if orig.loc[i, 'Action'] == 'change from':
            old_task = task_lines
        # add task lines to the dataframe
        if orig.loc[i, 'Action'] in ['change from', 'change to', 'add']:
            for line in task_lines.index:
                if line == 0:
                    idx = len(df)-1
                else:
                    idx = len(df)
                df.loc[idx, 'Action'] = orig.loc[i, 'Action']
                df.loc[idx, 'Task Lines'] = task_lines.loc[line, 'Name']
                df.loc[idx, 'Object Type'] = task_lines.loc[line, 'Type']
                df.loc[idx, 'Parallel Marker?'] = task_lines.loc[line,
                                                                 'Is Line Parallel']

        # convert old changed to new
        for change in np.flatnonzero(changed_rows == i):
            new_changed.append((start, changed_cols[change]))
        # add new task lines to changed if they differ from old
        if orig.loc[i, 'Action'] == 'change to' and not task_lines.equals(old_task):
            lines_rows = range(start, len(df))
            new_changed += list(zip(lines_rows, [4]*len(lines_rows)))
            new_changed += list(zip(lines_rows, [5]*len(lines_rows)))
            new_changed += list(zip(lines_rows, [6]*len(lines_rows)))
        # always add task lines to new tasks
        if orig.loc[i, 'Action'] == 'add':
            lines_rows = range(start, len(df))
            new_changed += list(zip(lines_rows, [4]*len(lines_rows)))
            new_changed += list(zip(lines_rows, [5]*len(lines_rows)))
            new_changed += list(zip(lines_rows, [6]*len(lines_rows)))
        if orig.loc[i, 'Action'] != '':
            groups.append((start+1, len(df)+1))
    return df.fillna(''), new_changed, groups


def write_sheet(wb: xlsxwriter.workbook, sheet: str, df: pd.DataFrame, changed: list):
    """"""
    f = formats(wb)
    ws = wb.add_worksheet(sheet)
    ws.outline_settings(True, False, True, True)
    groups = False
    if sheet == "Tasks":
        df, changed, groups = reformat_tasks(df, changed)
    for c, header in enumerate(df.columns):
        ws.write(0, c, header, f.header)
    for r in df.index:
        act = df.loc[r, 'Action']
        for c, val in enumerate(df.loc[r]):
            if act == 'change from':
                ws.write(r+1, c, val, f.matching)
            if act == 'change to':
                form = f.matching_diffs if (r, c) in changed else f.matching
                ws.write(r+1, c, val, form)
            if act == 'remove':
                ws.write(r+1, c, val, f.remove)
            if act == 'add':
                ws.write(r+1, c, val, f.add)
    if groups:
        collapsed = {'change from': True,
                     'change to': False,
                     'remove': True,
                     'add': False,
                     '': False}
        for (start, end) in groups:
            for g in range(start+1, end):
                ws.set_row(g, None, None, {
                           'level': 1, 'hidden': collapsed[df.loc[start, 'Action']]})
            ws.set_row(start, None, None, {
                       'collapsed': collapsed[df.loc[start, 'Action']]})


def create_sheet(rpt: str) -> pd.DataFrame:
    """Creates and returns the calculations specification page as a pandas DataFrame,
    along with a tuple index list of cells with differences"""
    pre = pre_config[rpt]
    post = post_config[rpt]

    matching = find_matching(pre, post)
    removed = find_removed(pre, post)
    added = find_added(pre, post)
    # drop date columns as we don't need them for spec
    pre = pre.drop(['Date Created', 'Date Last Used'], axis=1).fillna('')
    post = post.drop(['Date Created', 'Date Last Used'], axis=1).fillna('')
    df = pd.DataFrame(columns=['Action']+post.columns.tolist())
    # changed is a list of (row, column) indices marking format should highlight changes
    changed = []
    # add matching to the df
    for before_idx, after_idx in matching:
        before = pre.loc[before_idx].tolist()
        after = post.loc[after_idx].tolist()
        compared = compare(before, after)
        if any(compared):
            df.loc[len(df)] = ['change from']+before
            diff = [i+1 for i, d in enumerate(compared) if d]
            changed += list(zip([len(df)]*len(diff), diff))
            df.loc[len(df)] = ['change to']+after
            df.loc[len(df)] = ['']*(len(post.columns)+1)
    # add removed to the df
    for rem in removed:
        df.loc[len(df)] = ['remove']+pre.loc[rem].tolist()
        df.loc[len(df)] = ['']*(len(post.columns)+1)
    # add added to the df
    for a in added:
        df.loc[len(df)] = ['add']+post.loc[a].tolist()
        df.loc[len(df)] = ['']*(len(post.columns)+1)
    return df, changed


if __name__ == "__main__":
    app = QApplication(sys.argv)
    global pre_config, post_config
    # pre_config, _ = QFileDialog.getOpenFileName(
    #     None, 'Select config report taken before changes...', '', '*.zip')
    # if not pre_config:
    #     raise Exception('Must select Pre-Change config report')
    # post_config, _ = QFileDialog.getOpenFileName(
    #     None, 'Select config report taken after changes...', '', '*.zip')
    # if not pre_config:
    #     raise Exception('Must select Post-Change config report')
    pre_config_path = 'Config Reports\Inventories Local_Config_Report_BlueBook_06_Dec_2022.zip'
    post_config_path = 'Config Reports\Inventories Local_Config_Report_BAT_06_Dec_2022.zip'
    pre_config = read_config_rpt(pre_config_path)
    post_config = read_config_rpt(post_config_path)
    with xlsxwriter.Workbook('Outputs/test.xlsx') as wb:
        for rpt in ['Calculations', 'Tasks']:
            df, changed = create_sheet(rpt)
            write_sheet(wb, rpt, df, changed)
