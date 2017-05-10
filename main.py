# -*- coding: utf-8 -*-

import pandas as pd
import datetime

sourcefile_path = 'Auditing Export 2017-05-04.csv'

csv = pd.read_csv(sourcefile_path)
# csv.to_excel('1.xlsx', encoding='utf-8')

df = csv.loc[:, ['Date', 'Event category', 'Change summary', 'Changed object', 'Operation details', 'Associated items']]
for i, a in enumerate(df['Date']):
    df['Date'][i] = datetime.datetime.strptime(a, '%d/%b/%y %I:%M %p')

df = df.sort_index(ascending=False)
df = df[df['Event category'] == 'projects']
df.drop('Event category', axis=1, inplace=True)
df['index'] = range(len(df))
df.set_index(['index'], inplace=True)
# df.to_excel('2.xlsx', encoding='utf-8')


def od_parse(od, label):
    old = ''
    new = ''
    pos = od.find(label)
    if pos > -1:
        f_pos = pos + od[pos:].find('from [')
        if f_pos > -1:
            old = od[f_pos + 6: f_pos + od[f_pos:].find(']')]
        t_pos = pos + od[pos:].find('to [')
        new = od[t_pos + 4: t_pos + od[t_pos:].find(']')]
    return old, new

project = pd.DataFrame(columns=['Project', 'Version', 'Start', 'End', 'Create', 'Release', 'Description',
                                'Update time', 'Deleted'])

for row in df.iterrows():
    update_time = row[1]['Date']
    if row[1]['Change summary'] == 'Project created':
        project = project.append([{'Project': row[1]['Changed object'], 'Update time': update_time,
                                   'Create':update_time.date()}], ignore_index=True)
    elif row[1]['Change summary'] == 'Project updated':
        pn_old, pn_new = od_parse(row[1]['Operation details'], 'Name')
        if pn_old == '':
            pass
        else:
            p_index = project[project['Project'] == pn_old].index.values
            for index in p_index:
                project.loc[index, 'Project'] = pn_new
                project.loc[index, 'Update time'] = update_time
    elif row[1]['Change summary'] == 'Project deleted':
        p_index = project[project['Project'] == row[1]['Changed object']].index.values
        for index in p_index:
            project.loc[index, 'Deleted'] = 'True'
            project.loc[index, 'Update time'] = update_time
    elif row[1]['Change summary'] == 'Project version created':
        modified = False
        p_name = row[1]['Associated items'][row[1]['Associated items'].find('[')+1:row[1]['Associated items'].find(']')]
        p_version = row[1]['Changed object']
        _, p_start = od_parse(row[1]['Operation details'], 'Start date')
        _, p_end = od_parse(row[1]['Operation details'], 'Release date')
        _, p_des = od_parse(row[1]['Operation details'], 'Description')
        p_create = update_time.date()
        p_index = project[project['Project'] == p_name].index.values
        for index in p_index:
            if type(project.loc[index]['Version']) == float:
                project.loc[index, 'Version'] = p_version
                project.loc[index, 'Start'] = p_start
                project.loc[index, 'End'] = p_end
                project.loc[index, 'Create'] = p_create
                project.loc[index, 'Description'] = p_des
                project.loc[index, 'Update time'] = update_time
                modified = True
        if not modified:
            project = project.append([{'Project': p_name, 'Version': p_version, 'Update time': update_time,
                                       'Start': p_start, 'End': p_end, 'Create': p_create,
                                       'Description': p_des}], ignore_index=True)
    elif row[1]['Change summary'] == 'Project version updated':
        p_name = row[1]['Associated items'][row[1]['Associated items'].find('[')+1:row[1]['Associated items'].find(']')]
        p_version = row[1]['Changed object']
        p_ver_old, p_ver_new = od_parse(row[1]['Operation details'], 'Name')
        _, p_start = od_parse(row[1]['Operation details'], 'Start date')
        _, p_end = od_parse(row[1]['Operation details'], 'Release date')
        _, p_des = od_parse(row[1]['Operation details'], 'Description')
        if p_version == p_ver_new:
            p_index = project[(project['Project'] == p_name) & (project['Version'] == p_ver_old)].index.values
        else:
            p_index = project[(project['Project'] == p_name) & (project['Version'] == p_version)].index.values
        for index in p_index:
            if p_ver_new != '':
                project.loc[index, 'Version'] = p_ver_new
            if p_start != '':
                project.loc[index, 'Start'] = p_start
            if p_end != '':
                project.loc[index, 'End'] = p_end
            if p_des != '':
                project.loc[index, 'Description'] = p_des
            project.loc[index, 'Update time'] = update_time
    elif row[1]['Change summary'] == 'Project version deleted':
        p_name = row[1]['Associated items'][row[1]['Associated items'].find('[')+1:row[1]['Associated items'].find(']')]
        p_index = project[(project['Project'] == p_name) &
                          (project['Version'] == row[1]['Changed object'])].index.values
        for index in p_index:
            project.loc[index, 'Deleted'] = 'True'
            project.loc[index, 'Update time'] = update_time
    elif row[1]['Change summary'] == 'Project version released':
        p_name = row[1]['Associated items'][row[1]['Associated items'].find('[')+1:row[1]['Associated items'].find(']')]
        p_index = project[(project['Project'] == p_name) &
                          (project['Version'] == row[1]['Changed object'])].index.values
        for index in p_index:
            project.loc[index, 'Release'] = update_time.date()
            project.loc[index, 'Update time'] = update_time

project.to_excel('project.xlsx', encoding='utf-8')

