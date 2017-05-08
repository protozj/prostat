# -*- coding: utf-8 -*-

import numpy as np
import pandas as pd
import datetime

sourcefile_path = 'Auditing Export 2017-05-04.csv'
start_date = datetime.datetime(2016, 4, 1)
end_date = datetime.datetime(2016, 4, 28)

csv = pd.read_csv(sourcefile_path)
# csv.to_excel('1.xlsx', encoding='utf-8')

df = csv.loc[:, ['Date', 'Event category', 'Change summary', 'Author name', 'Changed object', 'Associated items']]
for i, a in enumerate(df['Date']):
    df['Date'][i] = datetime.datetime.strptime(a, '%d/%b/%y %I:%M %p')

# df = df[(df['Date'] >= start_date) & (df['Date'] < end_date)]
df = df[df['Event category'] == 'projects']
df.drop('Event category', axis=1, inplace=True)
df['index'] = range(len(df))
df.set_index(['index'], inplace=True)

stop_projects = []
for i, a in enumerate(df['Change summary']):
    if a == 'Project deleted':
        stop_projects.append(df['Changed object'][i])

pvc = df[df['Change summary'] == 'Project version created']
pvc = pvc.drop(['Change summary'], axis=1)
pvc.set_index(pd.Series(range(len(pvc))), inplace=True)
# pvc.to_excel('pvc.xlsx', encoding='utf-8')


def pvc_apply(row):
    row['Date'] = row['Date'].date()
    row['Associated items'] =  \
        row['Associated items'][row['Associated items'].find('[')+1:row['Associated items'].find(']')]
    if (row['Associated items'] not in stop_projects) and (row['Changed object'].find('.') != -1):
        return row
    else:
        row['Date'] = np.nan
        return row

pvc.apply(pvc_apply, axis=1)
pvc = pvc.dropna(subset=['Date'])
move = pvc.pop('Changed object')
pvc.insert(0, 'version', move)
move = pvc.pop('Associated items')
pvc.insert(0, 'project', move)
pvc.rename(columns={'Date': 'Start date'}, inplace=True)
pvc = pvc.sort_values(by=['project', 'version'], ascending=False)
# pvc.to_excel('pvc_apply.xlsx', encoding='utf-8')

pvr = df[df['Change summary'] == 'Project version released']
pvr = pvr.drop(['Change summary'], axis=1)
pvr.set_index(pd.Series(range(len(pvr))), inplace=True)
# pvr.to_excel('pvr.xlsx', encoding='utf-8')


def pvr_apply(row):
    row['Date'] = row['Date'].date()
    row['Associated items'] =  \
        row['Associated items'][row['Associated items'].find('[')+1:row['Associated items'].find(']')]
    if (row['Associated items'] not in stop_projects) and (row['Changed object'].find('.') != -1):
        return row
    else:
        row['Date'] = np.nan
        return row

pvr.apply(pvr_apply, axis=1)
pvr = pvr.dropna(subset=['Date'])
move = pvr.pop('Changed object')
pvr.insert(0, 'version', move)
move = pvr.pop('Associated items')
pvr.insert(0, 'project', move)
pvr.rename(columns={'Date': 'Release date'}, inplace=True)
pvr = pvr.sort_values(by=['project', 'version'], ascending=False)
# pvr.to_excel('pvr_apply.xlsx', encoding='utf-8')

project = pd.merge(pvc, pvr, how='outer', on=['project', 'version'])
project = project.sort_values(by=['project', 'version'], ascending=False)
# project.to_excel('project.xlsx', encoding='utf-8')

prostat = project[(project['Release date'] >= end_date.date())
                  | ((project['Release date'].isnull()) & (project['Start date'] >= end_date.date()))]
prostat = prostat.drop(['Author name_x', 'Author name_y'], axis=1)
prostat.to_excel('prostat.xlsx', encoding='utf-8')
