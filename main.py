from typing import Tuple
from zipfile import ZipFile
import re
import pandas as pd
import sys
from PyQt5.QtWidgets import QApplication, QFileDialog
import xlsxwriter
from dataclasses import dataclass
import numpy as np
from natsort import natsort_keygen


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
        text_file.filename), encoding='unicode_escape', skiprows=3, dtype=str, na_values="") for text_file in zip_file.infolist() if text_file.filename.endswith('.csv') and text_file.filename != "zzhashtotals.csv"}
    return dfs


def read_group_names(filepath: str) -> dict:
    """Returns a dictionary where group name as written in file name is the key,
    and group name as written in CORD is the value"""
    zip_file = ZipFile(filepath)
    groups_as_df = {re.sub('(.*?)_rpt_\d*\.csv', r'\1', text_file.filename): pd.read_csv(zip_file.open(
        text_file.filename), encoding='unicode_escape', skiprows=1, nrows=1, header=None, dtype=str, na_values="") for text_file in zip_file.infolist() if text_file.filename.endswith('.csv') and re.match('(Grp\d-).*', text_file.filename)}
    groups = dict()
    for name, line in groups_as_df.items():
        groups[name] = line.iloc[0, 0].split(' Group: ')[-1].strip()
    return groups


def read_class_items_names(filepath: str) -> dict:
    """Returns a dictionary where classification file name is the value,
    and classification name as written in CORD is the key"""
    zip_file = ZipFile(filepath)
    items_as_df = {re.sub('(.*?)_rpt_\d*\.csv', r'\1', text_file.filename): pd.read_csv(zip_file.open(
        text_file.filename), encoding='unicode_escape', skiprows=1, nrows=1, header=None, dtype=str, na_values="") for text_file in zip_file.infolist() if text_file.filename.endswith('.csv') and re.match('(Itms-).*', text_file.filename)}
    items = dict()
    for name, line in items_as_df.items():
        items[re.sub(
            '.* Classification: (.*)', r'\1', line.iloc[0, 0])] = name
    return items


def read_class_groups(filepath: str) -> dict:
    """Returns a dictionary where classification file name is the value,
    and classification name as written in CORD is the key"""
    zip_file = ZipFile(filepath)
    groups_as_df = {re.sub('(.*?)_rpt_\d*\.csv', r'\1', text_file.filename): pd.read_csv(zip_file.open(
        text_file.filename), encoding='unicode_escape', skiprows=1, nrows=1, header=None, dtype=str, na_values="") for text_file in zip_file.infolist() if text_file.filename.endswith('.csv') and re.match('(Grp\d-).*', text_file.filename)}
    groups = dict()
    for name, line in groups_as_df.items():
        groups[re.sub(
            '.* Classification: (.*?) *Group: (.*)', r'\1#\2', line.iloc[0, 0])] = name
    return groups


def read_class_links(filepath: str) -> dict:
    """Returns a dictionary where classification pc link file name is the value,
    and a dataframe of parent child links is the key"""
    zip_file = ZipFile(filepath)
    link_classes = [re.sub(
        '.* Classification: (.*?)', r'\1', pd.read_csv(zip_file.open(
            text_file.filename), encoding='unicode_escape', skiprows=1, nrows=1, header=None, dtype=str, na_values="").iloc[0, 0]) for text_file in zip_file.infolist() if text_file.filename.endswith('.csv') and re.match('(PCLnk-).*', text_file.filename)]
    link_dfs = [pd.read_csv(zip_file.open(
        text_file.filename), encoding='unicode_escape', skiprows=3, dtype=str, na_values="")[['Parent Code', 'Child Code']] for text_file in zip_file.infolist() if text_file.filename.endswith('.csv') and re.match('(PCLnk-).*', text_file.filename)]
    return dict(zip(link_classes, link_dfs))


def read_grp_descs(filepath: str) -> dict:
    """Returns a dictionary where classification group filename (without _rpt_...) is the key,
    and group description is the value"""
    zip_file = ZipFile(filepath)
    groups = {re.sub('(.*?)_rpt_\d*\.csv', r'\1', text_file.filename): pd.read_csv(zip_file.open(
        text_file.filename), encoding='unicode_escape', skiprows=3, dtype=str, na_values="") for text_file in zip_file.infolist() if text_file.filename.endswith('.csv') and re.match('(Grps-).*', text_file.filename)}
    grps_master = dict()
    for class_filename, grp_df in groups.items():
        for r, grp in enumerate(grp_df['Name']):
            grps_master[('#').join([class_filename, grp])
                        ] = grp_df.loc[r, 'Description']
    return grps_master


def read_task_names(filepath: str) -> dict:
    """Returns a dictionary where task name as written in file name is the value,
    and task name as written in CORD is the key"""
    zip_file = ZipFile(filepath)
    tasks_as_df = {re.sub('(.*?)_rpt_\d*\.csv', r'\1', text_file.filename): pd.read_csv(zip_file.open(
        text_file.filename), encoding='unicode_escape', skiprows=1, nrows=1, header=None, dtype=str, na_values="") for text_file in zip_file.infolist() if text_file.filename.endswith('.csv') and re.match('(TskLn-).*', text_file.filename)}
    tasks = dict()
    for name, line in tasks_as_df.items():
        tasks[re.sub(
            '.* Task: (.*) \(Not Parallel\)|.* Task: (.*) \(Parallel\)', r'\1\2', line.iloc[0, 0])] = name
    return tasks


def convert_to_date_created(df: pd.DataFrame):
    # deep copy to get rid of chained assignment warning
    df = df[['Date Created', 'Name']].copy(deep=True)
    # sort values by name to solve for cases where multiple objects
    # have been created in the same minute
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


def convert_to_filename(name: str) -> str:
    """Converts a CORD object name to what it would be saved as for CSV report"""
    return name.replace('.', '_').replace(' ', '_')


def create_generic_sheet(rpt: str) -> Tuple[pd.DataFrame, list]:
    """Creates and returns the a specification page as a pandas DataFrame,
    along with a tuple index list of cells with differences and a dataframe of
    matching objects between pre and post"""
    none_pre = False
    none_post = False
    # using natsorted sorting on name for when multiple objects have same creation time
    # this holds so long as CORD objects have been prefixed with ordered code and so long as task order
    # hasn't changed.
    # otherwise no information is lost but spec say to edit calc of one kind to be another
    try:
        pre = pre_config[rpt].sort_values(
            by='Name', key=natsort_keygen()).reset_index(drop=True)
    except KeyError:
        none_pre = True
    try:
        post = post_config[rpt].sort_values(
            by='Name', key=natsort_keygen()).reset_index(drop=True)
    except KeyError:
        none_post = True
    # if object doesn't exist in either return empty
    if none_pre and none_post:
        return pd.DataFrame(), [], pd.DataFrame()

    # if exists in one, create the missing dataframe using columns from the other
    if none_pre:
        pre = pd.DataFrame(columns=post_config[rpt].columns)
    if none_post:
        post = pd.DataFrame(columns=pre_config[rpt].columns)

    # get indices and create their relevant dataframes
    matching = find_matching(pre, post)
    pre_i, post_i = [pre for (pre, _) in matching], [
        post for (_, post) in matching]
    matching_df = pd.DataFrame({'rpt': [rpt]*len(pre_i), 'pre': pre_i, 'post':  post_i, 'pre_name': [
                                pre.loc[i, 'Name'] for i in pre_i], 'post_name': [post.loc[i, 'Name'] for i in post_i]})
    if not matching_df.empty:
        matching_df = matching_df.sort_values(
            'pre_name', key=lambda col: col.str.lower()).reset_index(drop=True)
    deleted = find_deleted(pre, post)
    created = find_created(pre, post)
    # drop date columns as we don't need them for spec
    pre = pre.drop(['Date Created', 'Date Last Used'], axis=1).fillna('')
    post = post.drop(['Date Created', 'Date Last Used'], axis=1).fillna('')
    df = pd.DataFrame(columns=['Action']+post.columns.tolist())
    # changed is a list of (row, column) indices marking format should highlight changes
    changed = []
    # add deleted to the df
    for d in deleted:
        df.loc[len(df)] = ['Delete']+pre.loc[d].tolist()
        df.loc[len(df)] = ['']*(len(post.columns)+1)
    # add matching to the df
    for r in matching_df.index:
        before = pre.loc[matching_df.loc[r, 'pre']].tolist()
        after = post.loc[matching_df.loc[r, 'post']].tolist()
        compared = compare(before, after)
        if any(compared):
            df.loc[len(df)] = ['Change From']+before
            diff = [i+1 for i, d in enumerate(compared) if d]
            changed += list(zip([len(df)]*len(diff), diff))
            df.loc[len(df)] = ['Change To']+after
            df.loc[len(df)] = ['']*(len(post.columns)+1)
    # add created to the df
    for c in created:
        df.loc[len(df)] = ['Create']+post.loc[c].tolist()
        df.loc[len(df)] = ['']*(len(post.columns)+1)
    return df, changed, matching_df


def tidy_format(ws: xlsxwriter.Workbook.worksheet_class, sheet: str):
    """Function to add formatting to worksheet columns based on sheet name"""
    if sheet == 'Dataset_Definitions':
        ws.set_column('A:A', width=13)
        ws.set_column('B:C', width=30)
        ws.set_column('D:N', width=11)
    elif sheet == 'Calculations':
        ws.set_column('A:A', width=13)
        ws.set_column('B:C', width=40)
        ws.set_column('D:E', width=24)
        ws.set_column('F:F', width=10)
        ws.set_column('G:G', width=12)
        ws.set_column('H:I', width=10)
        ws.set_column('J:L', width=25)
    elif sheet == 'Consistency_Checks':
        ws.set_column('A:A', width=13)
        ws.set_column('B:C', width=40)
        ws.set_column('D:D', width=15)
        ws.set_column('E:E', width=30)
        ws.set_column('F:F', width=10)
        ws.set_column('G:G', width=12)
        ws.set_column('H:J', width=10)
        ws.set_column('K:M', width=24)
    elif sheet == 'Visualisations':
        ws.set_column('A:A', width=13)
        ws.set_column('B:C', width=40)
        ws.set_column('D:D', width=8)
        ws.set_column('E:E', width=13)
        ws.set_column('F:F', width=30)
        ws.set_column('G:G', width=20)
        ws.set_column('H:I', width=10)
        ws.set_column('J:L', width=25)
    elif sheet == 'Import_Definitions':
        ws.set_column('A:A', width=13)
        ws.set_column('B:B', width=40)
        ws.set_column('C:C', width=50)
        ws.set_column('D:D', width=13)
        ws.set_column('E:E', width=30)
        ws.set_column('F:F', width=15)
        ws.set_column('G:I', width=36)
    elif sheet == 'Parameters':
        ws.set_column('A:A', width=13)
        ws.set_column('B:B', width=40)
        ws.set_column('C:C', width=100)
        ws.set_column('D:E', width=15)
    elif sheet == 'Tasks':
        ws.set_column('A:A', width=13)
        ws.set_column('B:B', width=40)
        ws.set_column('C:C', width=100)
        ws.set_column('D:D', width=17)
        ws.set_column('E:E', width=40)
        ws.set_column('F:F', width=30)
        ws.set_column('G:G', width=15)
    elif sheet == 'Classifications':
        ws.set_column('A:A', width=13)
        ws.set_column('B:B', width=40)
        ws.set_column('C:C', width=100)
        ws.set_column('D:D', width=17)
        ws.set_column('E:E', width=57)
        ws.set_column('F:F', width=30)
    elif sheet == 'Groups':
        ws.set_column('A:A', width=13)
        ws.set_column('B:C', width=40)
        ws.set_column('D:D', width=100)
        ws.set_column('E:E', width=40)
    elif sheet == 'Hierarchy':
        ws.set_column('A:A', width=13)
        ws.set_column('B:D', width=40)
    else:
        raise Exception('Unrecognised sheet name')

    ws.freeze_panes(1, 0)


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
            if act == 'Change From':
                ws.write(r+1, c, val, f.matching)
            if act == 'Change To':
                form = f.matching_diffs if (r, c) in changed else f.matching
                ws.write(r+1, c, val, form)
            if act == 'Delete':
                ws.write(r+1, c, val, f.delete)
            if act == 'Create':
                ws.write(r+1, c, val, f.create)
    tidy_format(ws, sheet)


def write_classifications_sheet(wb: xlsxwriter.Workbook, master_matching: pd.DataFrame):
    """Writes the Classifications sheet to the specification file since
    classifications come from multiple config files"""
    df = pd.DataFrame(columns=['Action', 'Name', 'Description',
                      'Item Code', 'Item Description', 'Table Header'])
    pre_names, post_names = read_class_items_names(
        pre_config_path), read_class_items_names(post_config_path)
    pre_items, post_items = dict(), dict()
    for name, filename in pre_names.items():
        pre_items[name] = pre_config[filename].fillna('n/a').drop(
            'Level', axis=1).apply(lambda x: x.str.strip())
    for name, filename in post_names.items():
        post_items[name] = post_config[filename].fillna('n/a').drop(
            'Level', axis=1).apply(lambda x: x.str.strip())
    matching = find_matching(
        pre_config['Classifications'], post_config['Classifications'])
    deleted = find_deleted(
        pre_config['Classifications'], post_config['Classifications'])
    created = find_created(
        pre_config['Classifications'], post_config['Classifications'])
    groups = []
    changed = []

    def add_classification(act: str, start: int, i: int, df: pd.DataFrame, orig: pd.DataFrame):
        """Adds a task and it's task lines to the dataframe"""
        df.loc[start, 'Action'] = act
        df.loc[start, 'Name'] = orig.loc[i, 'Name']
        df.loc[start, 'Description'] = orig.loc[i, 'Description']
        df.loc[start, ['Item Code', 'Item Description', 'Table Header']] = ['n/a']*3
        # add in classification items
        if act in ['Change From', 'Delete']:
            items = pre_items
        if act in ['Change To', 'Create']:
            items = post_items
        lines_df = items[df.loc[start, 'Name']]
        lines_df = lines_df.rename(columns={
                                   'Code': 'Item Code', 'Description': 'Item Description'})
        lines_df.index = range(start, start+len(lines_df))
        # fill in columns with empty string ready for concat
        for col in df.columns:
            if col not in lines_df.columns:
                lines_df[col] = ['']*len(lines_df)
        # return joined task lines dataframe to main df
        return pd.concat([df, lines_df]).reset_index(drop=True)

    def handle_items(pre: pd.DataFrame, post: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, list]:
        """"Checks if there have been changes to the task"""
        mini_changed = []
        # mark changes in name / description
        for col in range(1, 2):
            if pre.iloc[0, col] != post.iloc[0, col]:
                mini_changed.append((post.loc[0, 'index'], col))
        # find deleted, created and matching for items
        deleted = set(pre['Item Code']) - set(post['Item Code'])
        created = set(post['Item Code']) - set(pre['Item Code'])
        matching = set(pre['Item Code']) & set(post['Item Code'])
        # add deleted items into the post dataframe
        del_slice = pd.DataFrame(columns=post.columns,
                                 data=pre.loc[pre['Item Code'].isin(deleted)].assign(Action='Delete'))
        if not del_slice.empty:
            post = pd.concat(
                [post.iloc[0:1], del_slice, post.iloc[1:]]).reset_index(drop=True)
        # change Action to Create for created items
        post.loc[post[post['Item Code'].isin(
            created)].index, 'Action'] = 'Create'
        # add changed descriptions and headers to changed
        pre_itms = pre.loc[pre[pre['Item Code'].isin(matching)].index, [
            'Item Code', 'Item Description', 'Table Header']]
        post_itms = post.loc[post[post['Item Code'].isin(matching)].index, [
            'Item Code', 'Item Description', 'Table Header']]
        for i in post_itms.index:
            code = post_itms.loc[i, 'Item Code']
            chngd = ~post_itms.loc[i].eq(
                pre_itms[pre_itms['Item Code'] == code])
            if chngd.iloc[0].any():
                post.loc[i, 'Action'] = 'Change To'
                chng_cols = post_itms.columns[(chngd > 0).all()].tolist()
                for col in chng_cols:
                    mini_changed.append((i, post.columns.get_loc(col)))
        return pd.concat([pre, post]), mini_changed

    # start sheet by specifying classifications to delete
    for i in deleted:
        start = len(df)
        df = add_classification('Delete', start, i, df,
                                pre_config['Classifications'])
        df.loc[start:, 'Action'] = 'Delete'
        # add grouping details
        groups.append((start+1, len(df)+1))
        # add in blank line
        df.loc[len(df)] = ['']*len(df.columns)
    # then move to matching
    pre_i, post_i = zip(*matching)
    # create a new matching dataframe for easy sorting
    matching_df = pd.DataFrame({'pre': pre_i, 'post': post_i, 'pre_name': [
        pre_config['Classifications'].loc[i, 'Name'] for i in pre_i], 'post_name': [post_config['Classifications'].loc[i, 'Name'] for i in post_i]})
    # sort the classifications alphabetically by pre name
    matching_df = matching_df.sort_values(
        by='pre_name', key=lambda col: col.str.lower()).reset_index(drop=True)
    # add matching to master matching
    master_matching = pd.concat([master_matching, matching_df.assign(
        rpt=['Classifications']*len(matching_df))])
    # add classifications with changes to main df
    for i in matching_df.index:
        pre_i, post_i = matching_df.loc[i, 'pre'], matching_df.loc[i, 'post']
        pre_start = len(df)
        df = add_classification('Change From', pre_start,
                                pre_i, df, pre_config['Classifications'])
        df.loc[pre_start:, 'Action'] = 'Change From'
        pre_change = df.loc[pre_start:].copy().reset_index(drop=True)
        # store end of pre for grouping later
        end_pre = len(df)+1
        post_start = len(df)
        df = add_classification('Change To', post_start, post_i,
                                df, post_config['Classifications'])
        # only changing to for name and description of classification itself
        df.loc[post_start:, 'Action'] = 'Change To'
        post_change = df.loc[post_start:].copy().reset_index(drop=True)
        # find what's changed between pre and post
        pre_post, sub_changed = handle_items(pre_change, post_change)
        df = pd.concat([df.loc[:pre_start-1], pre_post]).reset_index(drop=True)
        # if there are changes add them to the changed formatting and add groups
        if sub_changed:
            for r, c in sub_changed:
                changed.append((r+post_start, c))
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
        df = add_classification('Create', start, i, df,
                                post_config['Classifications'])
        df.loc[start:, 'Action'] = 'Create'
        # add grouping details
        groups.append((start+1, len(df)+1))
        # add in blank line
        df.loc[len(df)] = ['']*len(df.columns)

    # If no task changes, no need to create a tasks sheet
    if df.replace('', np.nan).dropna(how='all').empty:
        return master_matching

    f = formats(wb)
    ws = wb.add_worksheet('Classifications')
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
            if act == '':
                ws.write(r+1, c, val)
    # add grouping to lines to better summarise changes
    collapsed = {'Change From': True,
                 'Change To': False,
                 'Delete': False,
                 'Create': False,
                 '': False}
    for (start, end) in groups:
        for g in range(start+1, end):
            ws.set_row(g, None, None, {
                'level': 1, 'hidden': collapsed[df.loc[start, 'Action']]})
        ws.set_row(start, None, None, {
            'collapsed': collapsed[df.loc[start, 'Action']]})
    # add tidy formatting
    tidy_format(ws, 'Classifications')
    return master_matching


def write_classifications_groups_sheet(wb: xlsxwriter.Workbook, master_matching: pd.DataFrame):
    """Writes the Classifications group changes to the specification file since
    classifications come from multiple config files"""
    df = pd.DataFrame(columns=['Action', 'Classification',
                      'Group Name', 'Group Description', 'Group Items'])
    pre_groups, post_groups = read_class_groups(
        pre_config_path), read_class_groups(post_config_path)
    pre_descs, post_descs = read_grp_descs(
        pre_config_path), read_grp_descs(post_config_path)
    pre_group_items, post_group_items = dict(), dict()
    for name, filename in pre_groups.items():
        pre_group_items[name] = pre_config[filename]['Code'].tolist()
    for name, filename in post_groups.items():
        post_group_items[name] = post_config[filename]['Code'].tolist()
    matching_class = master_matching[master_matching['rpt']
                                     == 'Classifications']
    # convert group item keys to updated post classfications
    conversion = dict(
        zip(matching_class['post_name'].tolist(), matching_class['pre_name'].tolist()))
    for name in list(pre_group_items.keys()):
        [class_name, grp] = name.split('#')
        class_name = conversion[class_name]
        pre_group_items[('#').join([class_name, grp])
                        ] = pre_group_items.pop(name)
    # identify deleted, matching and created
    deleted = set(set(list(pre_group_items.keys())) -
                  set(list(post_group_items.keys())))
    matching = set(set(list(post_group_items.keys()))
                   & set(list(pre_group_items.keys())))
    created = set(set(list(post_group_items.keys())) -
                  set(list(pre_group_items.keys())))
    groups = []
    changed = []

    # start sheet by specifying groups to delete
    for grp in deleted:
        [class_name, g] = grp.split('#')
        start = len(df)
        df.loc[start, 'Classification'] = class_name
        df.loc[start, 'Group Name'] = g
        df.loc[start, 'Group Description'] = pre_descs['Grps-'+class_name.replace(
            '.', '_').replace(' ', '_')+'#'+g]
        df.loc[start, 'Group Items'] = ''
        for item in pre_group_items[grp]:
            df.loc[len(df), 'Group Items'] = item
        df.loc[start:, 'Action'] = 'Delete'
        # add grouping details
        groups.append((start+1, len(df)+1))
        # add in blank line
        df.loc[len(df)] = ['']*len(df.columns)
    # move on to matching
    for grp in matching:
        [class_name, g] = grp.split('#')
        pre_start = len(df)
        # have to unconvert pre name to find old description and classfication
        old_class_name = matching_class.loc[matching_class[matching_class['post_name'] == class_name].index.tolist()[
            0], 'pre_name']
        old_desc = pre_descs['Grps-' +
                             convert_to_filename(old_class_name)+'#'+g]
        df.loc[pre_start, 'Classification'] = old_class_name
        df.loc[pre_start, 'Group Name'] = g
        df.loc[pre_start, 'Group Description'] = old_desc
        df.loc[pre_start, 'Group Items'] = ''
        for item in pre_group_items[grp]:
            df.loc[len(df), 'Group Items'] = item
        df.loc[pre_start:, 'Action'] = 'Change From'

        post_start = len(df)
        new_desc = post_descs['Grps-'+class_name.replace(
            '.', '_').replace(' ', '_')+'#'+g]
        df.loc[post_start, 'Action'] = 'Change To'
        df.loc[post_start, 'Classification'] = class_name
        df.loc[post_start, 'Group Name'] = g
        df.loc[post_start, 'Group Description'] = new_desc
        df.loc[post_start, 'Group Items'] = ''
        created_items = list(
            set(post_group_items[grp]) - set(pre_group_items[grp]))
        deleted_items = list(
            set(pre_group_items[grp]) - set(post_group_items[grp]))
        if deleted_items:
            for item in deleted_items:
                r = len(df)
                df.loc[r, 'Group Items'] = item
                df.loc[r, 'Action'] = 'Remove'
        if created_items:
            for item in created_items:
                r = len(df)
                df.loc[r, 'Group Items'] = item
                df.loc[r, 'Action'] = 'Add'
        # if nothing has changed remove it from the dataframe
        if not deleted_items and not created_items and old_desc == new_desc:
            df = df.loc[:pre_start-2]  # TODO -1?
        else:  # add grouping details
            groups.append((pre_start+1, post_start+1))
            groups.append((post_start+1, len(df)+1))
        if old_desc != new_desc:
            changed.append((post_start, 3))
        # add in blank line
        df.loc[len(df)] = ['']*len(df.columns)
    # finally specify groups to create
    for grp in created:
        [class_name, g] = grp.split('#')
        start = len(df)
        df.loc[start, 'Classification'] = class_name
        df.loc[start, 'Group Name'] = g
        df.loc[start, 'Group Description'] = post_descs['Grps-'+class_name.replace(
            '.', '_').replace(' ', '_')+'#'+g]
        df.loc[start, 'Group Items'] = ''
        for item in post_group_items[grp]:
            df.loc[len(df), 'Group Items'] = item
        df.loc[start:, 'Action'] = 'Create'
        # add grouping details
        groups.append((start+1, len(df)+1))
        # add in blank line
        df.loc[len(df)] = ['']*len(df.columns)

    # If no task changes, no need to create a tasks sheet
    if df.replace('', np.nan).dropna(how='all').empty:
        return

    df = df.fillna('')

    f = formats(wb)
    ws = wb.add_worksheet('Classification Groups')
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
            if act in ['Delete', 'Remove']:
                ws.write(r+1, c, val, f.delete)
            if act in ['Create', 'Add']:
                ws.write(r+1, c, val, f.create)
    # add grouping to lines to better summarise changes
    collapsed = {'Change From': True,
                 'Change To': False,
                 'Delete': True,
                 'Remove': False,
                 'Create': False,
                 'Add': False,
                 '': False}
    for (start, end) in groups:
        for g in range(start+1, end):
            ws.set_row(g, None, None, {
                'level': 1, 'hidden': collapsed[df.loc[start, 'Action']]})
        ws.set_row(start, None, None, {
            'collapsed': collapsed[df.loc[start, 'Action']]})
    # add tidy formatting
    tidy_format(ws, 'Groups')


def write_classifications_hierarchy_sheet(wb: xlsxwriter.Workbook, master_matching: pd.DataFrame):
    """Writes the Classifications parent-child link changes to the specification
    file since they come from multiple config files"""
    df = pd.DataFrame(
        columns=['Action', 'Classification', 'Parent Code', 'Child Code'])
    pre_link_items, post_link_items = read_class_links(
        pre_config_path), read_class_links(post_config_path)
    matching_class = master_matching[master_matching['rpt']
                                     == 'Classifications']
    # convert group item keys to updated post classfications
    conversion = dict(
        zip(matching_class['pre_name'].tolist(), matching_class['post_name'].tolist()))
    for class_name in list(pre_link_items.keys()):
        pre_link_items[conversion[class_name]] = pre_link_items.pop(class_name)
    # identify matching and created
    matching = set(set(list(post_link_items.keys()))
                   & set(list(pre_link_items.keys())))
    created = set(set(list(post_link_items.keys())) -
                  set(list(pre_link_items.keys())))
    groups = []
    changed = []

    def find_link_differences(pre: pd.DataFrame, post: pd.DataFrame) -> Tuple[list, list]:
        """Returns 2 dataframes for links added and removed"""
        pre['comb'] = pre['Parent Code'] + ['#'] + pre['Child Code']
        post['comb'] = post['Parent Code'] + ['#'] + post['Child Code']
        added = list(set(post['comb']) - set(pre['comb']))
        removed = list(set(pre['comb']) - set(post['comb']))
        return post[post['comb'].isin(added)].assign(Action='Add').drop('comb', axis=1), pre[pre['comb'].isin(removed)].assign(Action='Remove').drop('comb', axis=1)

    # start on matching
    for class_name in matching:
        pre_start = len(df)
        old_class_name = list(conversion.keys())[list(
            conversion.values()).index(class_name)]
        df.loc[pre_start, 'Classification'] = old_class_name
        df = pd.concat([df, pre_link_items[class_name]]).reset_index(drop=True)
        df.loc[pre_start:, 'Action'] = 'Change From'

        post_start = len(df)
        df.loc[post_start, 'Action'] = 'Change To'
        df.loc[post_start, 'Classification'] = class_name
        added, removed = find_link_differences(
            pre_link_items[class_name], post_link_items[class_name])
        if not removed.empty:
            df = pd.concat([df, removed]).reset_index(drop=True)
        if not added.empty:
            df = pd.concat([df, added]).reset_index(drop=True)
        if added.empty and removed.empty:
            df = df.loc[:pre_start-1]
        else:
            # add in blank line
            df.loc[len(df)] = ['']*len(df.columns)
            # add grouping details
            groups.append((pre_start+1, post_start+1))
            groups.append((post_start+1, len(df)))
    # finally specify links for new classifications
    for class_name in created:
        start = len(df)
        df.loc[start, 'Classification'] = class_name
        df = pd.concat([df, post_link_items[class_name]]
                       ).reset_index(drop=True)
        df.loc[start:, 'Action'] = 'Create'
        # add grouping details
        groups.append((start+1, len(df)+1))
        # add in blank line
        df.loc[len(df)] = ['']*len(df.columns)

    # If no task changes, no need to create a tasks sheet
    if df.replace('', np.nan).dropna(how='all').empty:
        return

    df = df.fillna('')

    f = formats(wb)
    ws = wb.add_worksheet('Classification Hierarchy')
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
            if act in ['Delete', 'Remove']:
                ws.write(r+1, c, val, f.delete)
            if act in ['Create', 'Add']:
                ws.write(r+1, c, val, f.create)
    # add grouping to lines to better summarise changes
    collapsed = {'Change From': True,
                 'Change To': False,
                 'Delete': True,
                 'Remove': False,
                 'Create': False,
                 'Add': False,
                 '': False}
    for (start, end) in groups:
        for g in range(start+1, end):
            ws.set_row(g, None, None, {
                'level': 1, 'hidden': collapsed[df.loc[start, 'Action']]})
        ws.set_row(start, None, None, {
            'collapsed': collapsed[df.loc[start, 'Action']]})
    # add tidy formatting
    tidy_format(ws, 'Hierarchy')


def write_tasks_sheet(wb: xlsxwriter.Workbook, master_matching: pd.DataFrame):
    """Writes the Tasks sheet to the specification file since tasks come from multiple config files"""

    df = pd.DataFrame(columns=['Action', 'Name', 'Description',
                               'Parallel Execution?', 'Task Lines', 'Object Type', 'Parallel Marker?'])
    pre_names, post_names = read_task_names(
        pre_config_path), read_task_names(post_config_path)
    pre_tasks, post_tasks = dict(), dict()
    for name, filename in pre_names.items():
        pre_tasks[name] = pre_config[filename].sort_values(
            by='Name', key=natsort_keygen()).reset_index(drop=True).fillna('n/a').drop(
            'Order No.', axis=1)
    for name, filename in post_names.items():
        post_tasks[name] = post_config[filename].sort_values(
            by='Name', key=natsort_keygen()).reset_index(drop=True).fillna('n/a').drop(
            'Order No.', axis=1)
    matching = find_matching(pre_config['Tasks'], post_config['Tasks'])
    deleted = find_deleted(pre_config['Tasks'], post_config['Tasks'])
    created = find_created(pre_config['Tasks'], post_config['Tasks'])
    groups = []
    changed = []

    def add_task(act: str, start: int, i: int, df: pd.DataFrame, orig: pd.DataFrame):
        """Adds a task and it's task lines to the dataframe"""
        df.loc[start, 'Action'] = act
        df.loc[start, 'Name'] = orig.loc[i, 'Name']
        df.loc[start, 'Description'] = orig.loc[i, 'Description']
        df.loc[start, 'Parallel Execution?'] = orig.loc[i, 'Is Parallel']
        # add in task lines
        if act in ['Change From', 'Delete']:
            tasks = pre_tasks
        if act in ['Change To', 'Create']:
            tasks = post_tasks
        lines_df = tasks[df.loc[start, 'Name']]
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
    if df.replace('', np.nan).dropna(how='all').empty:
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
    # add tidy formatting
    tidy_format(ws, 'Tasks')


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
    pre_config_path = r'Config Reports\Unbalanced Inventories_Config_Report_BlueBook_06_Dec_2022.zip'
    post_config_path = r'Config Reports\Unbalanced Inventories_Config_Report_BAT_06_Dec_2022.zip'
    pre_config = read_config_rpt(pre_config_path)
    post_config = read_config_rpt(post_config_path)
    with xlsxwriter.Workbook('Outputs/test.xlsx') as wb:
        master_matching = pd.DataFrame(
            columns=['rpt', 'pre', 'post', 'pre_name', 'post_name'])
        master_matching = write_classifications_sheet(wb, master_matching)
        write_classifications_groups_sheet(wb, master_matching)
        write_classifications_hierarchy_sheet(wb, master_matching)
        for rpt in GENERIC_SHEETS:
            df, changed, matching = create_generic_sheet(rpt)
            master_matching = pd.concat([master_matching, matching])
            write_generic_sheet(wb, rpt, df, changed)
        write_tasks_sheet(wb, master_matching)
