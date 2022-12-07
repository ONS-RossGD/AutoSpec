from zipfile import ZipFile
import re
import pandas as pd
import sys
from PyQt5.QtWidgets import QApplication, QFileDialog
import xlsxwriter
from dataclasses import dataclass


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


def write_sheet(wb: xlsxwriter.workbook, sheet: str, df: pd.DataFrame, changed: dict):
    f = formats(wb)
    ws = wb.add_worksheet(sheet)
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


def create_sheet(pre: pd.DataFrame, post: pd.DataFrame) -> pd.DataFrame:
    """Creates and returns the calculations specification page as a pandas DataFrame,
    along with a dictionary for changed cells with differences"""
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
    print(changed)
    return df, changed


if __name__ == "__main__":
    app = QApplication(sys.argv)
    # pre_config, _ = QFileDialog.getOpenFileName(
    #     None, 'Select config report taken before changes...', '', '*.zip')
    # if not pre_config:
    #     raise Exception('Must select Pre-Change config report')
    # post_config, _ = QFileDialog.getOpenFileName(
    #     None, 'Select config report taken after changes...', '', '*.zip')
    # if not pre_config:
    #     raise Exception('Must select Post-Change config report')
    pre_config = 'Config Reports\Inventories Local_Config_Report_BlueBook_06_Dec_2022.zip'
    post_config = 'Config Reports\Inventories Local_Config_Report_BAT_06_Dec_2022.zip'
    pre = read_config_rpt(pre_config)
    post = read_config_rpt(post_config)
    with xlsxwriter.Workbook('Outputs/test.xlsx') as wb:
        for rpt in ['Calculations', 'Tasks']:
            df, changed = create_sheet(pre[rpt], post[rpt])
            write_sheet(wb, rpt, df, changed)
