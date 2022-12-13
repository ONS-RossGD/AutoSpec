from typing import Tuple
from zipfile import ZipFile
import re
import pandas as pd
import sys
from PyQt5.QtWidgets import QApplication, QFileDialog
import xlsxwriter
from dataclasses import dataclass
import numpy as np


GENERIC_SHEETS = ['Dataset_Definitions', 'Calculations',
                  'Consistency_Checks', 'Visualisations', 'Import_Definitions', 'Parameters']


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
    def delete(self):
        return wb.add_format({
            'bg_color': '#FF7C80'
        })

    @property
    def create(self):
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


def read_group_names(filepath: str) -> dict:
    """Returns a dictionary where group name as written in file name is the key,
    and group name as written in CORD is the value"""
    zip_file = ZipFile(filepath)
    groups_as_df = {re.sub('(.*?)_rpt_\d*\.csv', r'\1', text_file.filename): pd.read_csv(zip_file.open(
        text_file.filename), encoding='unicode_escape', skiprows=1, nrows=1, header=None) for text_file in zip_file.infolist() if text_file.filename.endswith('.csv') and re.match('(Grp\d-).*', text_file.filename)}
    groups = dict()
    for name, line in groups_as_df.items():
        groups[name] = line.iloc[0, 0].split(' Group: ')[-1].strip()
    return groups


def read_task_names(filepath: str) -> dict:
    """Returns a dictionary where task name as written in file name is the key,
    and task name as written in CORD is the value"""
    zip_file = ZipFile(filepath)
    tasks_as_df = {re.sub('(.*?)_rpt_\d*\.csv', r'\1', text_file.filename): pd.read_csv(zip_file.open(
        text_file.filename), encoding='unicode_escape', skiprows=1, nrows=1, header=None) for text_file in zip_file.infolist() if text_file.filename.endswith('.csv') and re.match('(TskLn-).*', text_file.filename)}
    tasks = dict()
    for name, line in tasks_as_df.items():
        tasks[re.sub(
            '.* Task: (.*) \(Not Parallel\)|.* Task: (.*) \(Parallel\)', r'\1\2', line.iloc[0, 0])] = name
    return tasks


def convert_to_date_created(df):
    # deep copy to get rid of chained assignment warning
    df = df[['Date Created', 'Name']].copy(deep=True)
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


def find_deleted(pre_df: pd.DataFrame, post_df: pd.DataFrame) -> pd.DataFrame:
    """Returns list of matches calcs between pre and post based on Date Created"""
    pre = convert_to_date_created(pre_df).sort_values(
        by='Name', key=lambda col: col.str.lower())
    post = convert_to_date_created(post_df)
    return [pre.loc[i, 'id'] for i in pre.index if i not in post.index]


def find_created(pre_df: pd.DataFrame, post_df: pd.DataFrame) -> pd.DataFrame:
    """Returns list of matches calcs between pre and post based on Date Created"""
    pre = convert_to_date_created(pre_df)
    post = convert_to_date_created(post_df).sort_values(
        by='Name', key=lambda col: col.str.lower())
    return [post.loc[i, 'id'] for i in post.index if i not in pre.index]


def compare(before: list, after: list) -> pd.Series:
    """Compares each cell of a pre and post series and returns
    a boolean series indicating where there are differences"""
    return [False if cell == after[i] else True for i, cell in enumerate(before)]


def write_generic_sheet(wb: xlsxwriter.workbook, sheet: str, df: pd.DataFrame, changed: list):
    """"""
    if df.empty:
        return
    f = formats(wb)
    ws = wb.add_worksheet(sheet.replace('_', ' '))
    # change outline settings to make groups collapse from the top
    ws.outline_settings(True, False, True, True)
    for c, header in enumerate(df.columns):
        ws.write(0, c, header, f.header)
    # write each line of the dataframe to the worksheet with proper format
    for r in df.index:
        act = df.loc[r, 'Action']
        for c, val in enumerate(df.loc[r]):
            if act == 'change from':
                ws.write(r+1, c, val, f.matching)
            if act == 'change to':
                form = f.matching_diffs if (r, c) in changed else f.matching
                ws.write(r+1, c, val, form)
            if act == 'delete':
                ws.write(r+1, c, val, f.delete)
            if act == 'create':
                ws.write(r+1, c, val, f.create)


def create_classifications_sheet() -> Tuple[pd.DataFrame, list]:
    """"""


def write_tasks_sheet(wb: xlsxwriter.Workbook, master_matching: pd.DataFrame) -> Tuple[pd.DataFrame, list]:
    """Writes the Tasks sheet to the specification file tasks come from multiple config files"""

    df = pd.DataFrame(columns=['Action', 'Task Name', 'Task Description',
                      'Parallel Execution?', 'Task Lines', 'Object Type', 'Parallel Marker?'])
    pre_names, post_names = read_task_names(
        pre_config_path), read_task_names(post_config_path)
    pre_tasks, post_tasks = dict(), dict()
    for name, filename in pre_names.items():
        pre_tasks[name] = pre_config[filename].fillna('n/a').drop(
            'Order No.', axis=1)
    for name, filename in post_names.items():
        post_tasks[name] = post_config[filename].fillna('n/a').drop(
            'Order No.', axis=1)
    matching = find_matching(pre_config['Tasks'], post_config['Tasks'])
    deleted = find_deleted(pre_config['Tasks'], post_config['Tasks'])
    created = find_created(pre_config['Tasks'], post_config['Tasks'])
    groups = []
    changed = []

    def add_task(act: str, start: int, i: int, df: pd.DataFrame, orig: pd.DataFrame):
        """Adds a task and it's task lines to the dataframe"""
        df.loc[start, 'Action'] = act
        df.loc[start, 'Task Name'] = orig.loc[i, 'Name']
        df.loc[start, 'Task Description'] = orig.loc[i, 'Description']
        df.loc[start, 'Parallel Execution?'] = orig.loc[i, 'Is Parallel']
        # add in task lines
        if act in ['Change From', 'Delete']:
            tasks = pre_tasks
        if act in ['Change To', 'Create']:
            tasks = post_tasks
        lines_df = tasks[df.loc[start, 'Task Name']]
        lines_df = lines_df.rename(columns={
                                   'Name': 'Task Lines', 'Type': 'Object Type', 'Is Line Parallel': 'Parallel Marker?'})
        lines_df.index = range(start, start+len(lines_df))
        # fill in columns with empty string ready for concat
        for col in df.columns:
            if col not in lines_df.columns:
                lines_df[col] = ['']*len(lines_df)
        # return joined task lines dataframe to main df
        return pd.concat([df, lines_df]).groupby(level=0).sum()

    def get_changed(pre: pd.DataFrame, post: pd.DataFrame, conv: dict) -> list:
        """Checks if there have been changes to the task"""
        mini_changed = []
        # add new tasks to conversion dict
        for task in post['Task Lines']:
            if task not in conv.keys():
                conv[task] = ''
        # mark changes in basic task details
        for col in range(1, 3):
            if pre.iloc[0, col] != post.iloc[0, col]:
                mini_changed.append((post.loc[0, 'index'], col))
        # mark changes in task lines
        if pre['Task Lines'].tolist() != [conv[post] for post in post['Task Lines'].tolist()]:
            for i in post['index']:
                mini_changed.append((i, 4))
                mini_changed.append((i, 5))
                mini_changed.append((i, 6))
        else:  # mark changes in type or parallel markers
            for i in range(len(post)):
                for c in [5, 6]:
                    if post.iloc[i, c] != pre.iloc[i, c]:
                        mini_changed.append((post.loc[0, 'index'], c))
        return mini_changed

    # start sheet by specifying tasks to delete
    for i in deleted:
        start = len(df)
        df = add_task('Delete', start, i, df, pre_config['Tasks'])
        df.loc[start:, 'Action'] = 'Delete'
        # add grouping details
        groups.append((start+1, len(df)+1))
        # add in blank line
        df.loc[len(df)] = ['']*len(df.columns)
    # then move to matching so long as they don't contain tasks
    pre_i, post_i = zip(*matching)
    # create a new matching dataframe for easy sorting
    matching_df = pd.DataFrame({'pre': pre_i, 'post': post_i, 'pre_name': [
                               pre_config['Tasks'].loc[i, 'Name'] for i in pre_i], 'post_name': [post_config['Tasks'].loc[i, 'Name'] for i in post_i]})
    # create a boolean column to identify tasks containing tasks
    matching_df = matching_df.assign(tasks=[
        'Y' if 'TASK' in post_tasks[name]['Type'].tolist() else 'N' for name in matching_df['post_name']]).sort_values(by='tasks')
    # sort the tasks alphabetically by pre name and whether task contains tasks
    matching_df = matching_df.sort_values(
        by=['tasks', 'pre_name'], key=lambda col: col.str.lower()).reset_index(drop=True)
    # add matching to master matching and create conversion dict
    master_matching = pd.concat([master_matching, matching_df.drop(
        'tasks', axis=1).assign(rpt=['task']*len(matching_df))])
    conversion = dict(
        zip(master_matching['post_name'].tolist(), master_matching['pre_name'].tolist()))
    # add changes without tasks to main df
    for i in matching_df[matching_df['tasks'] == 'N'].index:
        pre_i, post_i = matching_df.loc[i, 'pre'], matching_df.loc[i, 'post']
        pre_start = len(df)
        df = add_task('Change From', pre_start, pre_i, df, pre_config['Tasks'])
        df.loc[pre_start:, 'Action'] = 'Change From'
        pre_change = df.loc[pre_start:].copy().drop(
            'Action', axis=1).reset_index()
        # store end of pre for grouping later
        end_pre = len(df)+1
        post_start = len(df)
        df = add_task('Change To', post_start, post_i,
                      df, post_config['Tasks'])
        df.loc[post_start:, 'Action'] = 'Change To'
        post_change = df.loc[post_start:].copy().drop(
            'Action', axis=1).reset_index()
        # find what's changed between pre and post
        sub_changed = get_changed(pre_change, post_change, conversion)
        # if there are changes add them to the changed formatting and add groups
        if sub_changed:
            changed += sub_changed
            # add grouping details for change to
            groups.append((pre_start+1, end_pre))
            groups.append((post_start+1, len(df)+1))
        # otherwise if nothings changed we don't need it on the sheet
        else:
            df = df.loc[:pre_start-2]
        # add in blank line
        df.loc[len(df)] = ['']*len(df.columns)
    # add creations to main df
    for i in created:
        start = len(df)
        df = add_task('Create', start, i, df, post_config['Tasks'])
        df.loc[start:, 'Action'] = 'Create'
        # add grouping details
        groups.append((start+1, len(df)+1))
        # add in blank line
        df.loc[len(df)] = ['']*len(df.columns)
    # add changes with tasks to main df
    for i in matching_df[matching_df['tasks'] == 'Y'].index:
        pre_i, post_i = matching_df.loc[i, 'pre'], matching_df.loc[i, 'post']
        pre_start = len(df)
        df = add_task('Change From', pre_start, pre_i, df, pre_config['Tasks'])
        df.loc[pre_start:, 'Action'] = 'Change From'
        pre_change = df.loc[pre_start:].copy().drop(
            'Action', axis=1).reset_index()
        # store end of pre for grouping later
        end_pre = len(df)+1
        post_start = len(df)
        df = add_task('Change To', post_start, post_i,
                      df, post_config['Tasks'])
        df.loc[post_start:, 'Action'] = 'Change To'
        post_change = df.loc[post_start:].copy().drop(
            'Action', axis=1).reset_index()
        # find what's changed between pre and post
        sub_changed = get_changed(pre_change, post_change, conversion)
        # if there are changes add them to the changed formatting and add groups
        if sub_changed:
            changed += sub_changed
            # add grouping details for change to
            groups.append((pre_start+1, end_pre))
            groups.append((post_start+1, len(df)+1))
        # otherwise if nothings changed we don't need it on the sheet
        else:
            df = df.loc[:pre_start-2]
        # add in blank line
        df.loc[len(df)] = ['']*len(df.columns)

    # If no task changes, no need to create a tasks sheet
    if df.empty:
        return

    f = formats(wb)
    ws = wb.add_worksheet('Tasks')
    # change outline settings to make groups collapse from the top
    ws.outline_settings(True, False, True, True)
    for c, header in enumerate(df.columns):
        ws.write(0, c, header, f.header)
    # write each line of the dataframe to the worksheet with proper format
    for r in df.index:
        act = df.loc[r, 'Action']
        for c, val in enumerate(df.loc[r]):
            if act == 'Change From':
                ws.write(r+1, c, val, f.matching)
            if act == 'Change To':
                form = f.matching_diffs if (r, c) in changed else f.matching
                ws.write(r+1, c, val, form)
            if act == 'Delete':
                ws.write(r+1, c, val, f.delete)
            if act == 'Create':
                ws.write(r+1, c, val, f.create)
    # add grouping to lines to better summarise changes
    collapsed = {'Change From': True,
                 'Change To': False,
                 'Delete': True,
                 'Create': False,
                 '': False}
    for (start, end) in groups:
        for g in range(start+1, end):
            ws.set_row(g, None, None, {
                'level': 1, 'hidden': collapsed[df.loc[start, 'Action']]})
        ws.set_row(start, None, None, {
            'collapsed': collapsed[df.loc[start, 'Action']]})

    ws.set_column('A:A', width=13)
    ws.set_column('B:B', width=40)
    ws.set_column('C:C', width=100)
    ws.set_column('D:D', width=17)
    ws.set_column('E:E', width=40)
    ws.set_column('F:F', width=30)
    ws.set_column('G:G', width=15)
    ws.freeze_panes(1, 0)


def create_generic_sheet(rpt: str) -> Tuple[pd.DataFrame, list]:
    """Creates and returns the a specification page as a pandas DataFrame,
    along with a tuple index list of cells with differences and a dataframe of 
    matching objects between pre and post"""
    none_pre = False
    none_post = False
    try:
        pre = pre_config[rpt]
    except KeyError:
        none_pre = True
    try:
        post = post_config[rpt]
    except KeyError:
        none_post = True

    # if object doesn't exist in either return empty
    if none_pre and none_post:
        return pd.DataFrame(), []

    # if exists in one, create the missing dataframe using columns from the other
    if none_pre:
        pre = pd.DataFrame(columns=post_config[rpt].columns)
    if none_post:
        post = pd.DataFrame(columns=pre_config[rpt].columns)

    matching = find_matching(pre, post)
    deleted = find_deleted(pre, post)
    created = find_created(pre, post)
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
    # add deleted to the df
    for rem in deleted:
        df.loc[len(df)] = ['delete']+pre.loc[rem].tolist()
        df.loc[len(df)] = ['']*(len(post.columns)+1)
    # add created to the df
    for a in created:
        df.loc[len(df)] = ['create']+post.loc[a].tolist()
        df.loc[len(df)] = ['']*(len(post.columns)+1)
    # create the matching df
    pre_i, post_i = [pre for (pre, _) in matching], [
        post for (_, post) in matching]
    matching_df = pd.DataFrame({'rpt': [rpt]*len(pre_i), 'pre': pre_i, 'post':  post_i, 'pre_name': [
                               pre_config[rpt].loc[i, 'Name'] for i in pre_i], 'post_name': [post_config[rpt].loc[i, 'Name'] for i in post_i]})
    return df, changed, matching_df


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
        master_matching = pd.DataFrame()
        for rpt in GENERIC_SHEETS:
            df, changed, matching = create_generic_sheet(rpt)
            master_matching = pd.concat([master_matching, matching])
            write_generic_sheet(wb, rpt, df, changed)
        write_tasks_sheet(wb, master_matching)
        # df, changed = create_classifications_sheet()
