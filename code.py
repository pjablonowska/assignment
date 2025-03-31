import pandas as pd
import openpyxl
import datetime
import calendar

#files were initially renamed to jan.xlsx and feb.xlsx
#skipping 2 first rows with filter values for the report
df1=pd.read_excel('jan.xlsx', sheet_name = 'Sheet1', header = 2) 
df2=pd.read_excel('feb.xlsx', sheet_name = 'Sheet1', header = 2)

#assuming the report is created on the last day of the month, variable for column names
#cmonth = current month, pmonth = previous month

#today = datetime.date.today() - would be useful
#since we have data for Feb and Jan we have to report on the 28.02.2025
today = datetime.date(2025, 2, 20)

cmonth = calendar.month_name[today.month]

if today.month ==1:
    pmonth = calendar.month_name[12]
else:
    pmonth = calendar.month_name[today.month - 1]
    
#adding prefix to df1 before merging
prefix = pmonth[:3].lower() + '.'
df11=df1.add_prefix(prefix)

#copy of df2 in order to create merged df
df_all=df2
#merging df2 and df1 on first 4 columns
df_all = df_all.merge(df11, how='outer', left_on = ['Application Name -> Process -> Activity -> Metric',
 'Process',
 'Activity',
 'Metric'], right_on=[f'{prefix}Application Name -> Process -> Activity -> Metric',
 f'{prefix}Process',
 f'{prefix}Activity',
 f'{prefix}Metric'])

#depends how the code is executed, additional check
#df_all.columns.tolist()
#df_all.isnull().sum()

#deleting last six columns
df_all = df_all[df_all.columns[:-6]]

#calculating difference between adherence, checking if the change occured
df_all.loc[:, 'dif_adherence'] = df_all['Adherence'] - df_all[f'{prefix}Adherence']
df_all.loc[:,'change'] = (df_all['dif_adherence'] != 0).astype(str).str.upper()

#creating column with the change type
df_all.loc[:, 'change_type'] = df_all['dif_adherence'].replace({
   -1: 'Adherent -> Non - Adherent',
    1: 'Non - Adherent -> Adherent', 
    0: 'No change'})

#aggregate by app
df_all_app = df_all.groupby('Application Name -> Process -> Activity -> Metric').agg({
    f'{prefix}Adherence': 'mean',
    'Adherence': 'mean'
})
df_all_app.columns = [pmonth, cmonth]
df_all_app['Delta'] = df_all_app[cmonth] - df_all_app[pmonth]
df_all_app = df_all_app.apply(lambda x: round(x, 4))
df_all_app = df_all_app.reset_index()

#aggregate by process
df_all_proc = df_all.groupby('Process').agg({
    f'{prefix}Adherence': 'mean',
    'Adherence': 'mean'
})
df_all_proc.columns = [pmonth, cmonth]
df_all_proc['Delta'] = df_all_proc[cmonth] - df_all_proc[pmonth]
df_all_proc = df_all_proc.apply(lambda x: round(x, 4))
df_all_proc = df_all_proc.reset_index()

#aggregate by process
df_all_proc = df_all.groupby('Process').agg({
    f'{prefix}Adherence': 'mean',
    'Adherence': 'mean'
})
df_all_proc.columns = [pmonth, cmonth]
df_all_proc['Delta'] = df_all_proc[cmonth] - df_all_proc[pmonth]
df_all_proc = df_all_proc.apply(lambda x: round(x, 4))
df_all_proc = df_all_proc.reset_index()

#aggregate by activity
df_all_act = df_all.groupby(['Process', 'Activity']).agg({
    f'{prefix}Adherence': 'mean',
    'Adherence': 'mean'
})
df_all_act.columns = [pmonth, cmonth]
df_all_act['Delta'] = df_all_act[cmonth] - df_all_act[pmonth]
df_all_act = df_all_act.apply(lambda x: round(x, 4))
df_all_act = df_all_act.reset_index()

#aggregate by metric
df_all_met = df_all.groupby(['Process', 'Activity', 'Metric']).agg({
    f'{prefix}Adherence': 'mean',
    'Adherence': 'mean'
})
df_all_met.columns = [pmonth, cmonth]
df_all_met['Delta'] = df_all_met[cmonth] - df_all_met[pmonth]
df_all_met = df_all_met.apply(lambda x: round(x, 4))
df_all_met = df_all_met.reset_index()

#aggregate by adherence state
df_all_ad = df_all.groupby('change_type').agg({
    'Metric': 'count',
})
df_all_ad = df_all_ad.reset_index()

with pd.ExcelWriter('output.xlsx', engine='openpyxl') as writer:
    df1.to_excel(writer, sheet_name=pmonth, index=False)
    df2.to_excel(writer, sheet_name=cmonth, index=False)
    df_all.to_excel(writer, sheet_name='combined', index=False)
    df_all_app.to_excel(writer, sheet_name = 'app', index = False)
    df_all_ad.to_excel(writer, sheet_name = 'change', index = False)
    df_all_proc.to_excel(writer, sheet_name = 'process', index = False)
    df_all_act.to_excel(writer, sheet_name = 'activity', index = False)
    df_all_met.to_excel(writer, sheet_name = 'metric', index = False)
    
import xlwings as xw
from xlwings import constants
from openpyxl import load_workbook

#file opening
app = xw.App(visible=True)
wb=app.books.open('output.xlsx')

#formatting combined data in order to create a pivot
cws=wb.sheets['combined']
data_range = cws.range('A1').expand('table')
table=cws.tables.add(source=data_range, name = 'data', has_headers=True)
table.table_style='TableStyleMedium3'

#adding sheet for pivot
pws = wb.sheets.add('pivot apps')

#inserting pivot - completely new skill for me
pcache = wb.api.PivotCaches().Create(
    SourceType=constants.PivotTableSourceType.xlDatabase,
    SourceData=data_range.api)
ptable = pcache.CreatePivotTable(
    TableDestination=pws.range('B2').api,
    TableName = 'AppAdherence')

#pivot fields
ptable.PivotFields("Adherence").Orientation = constants.PivotFieldOrientation.xlRowField
ptable.PivotFields(f"{prefix}Adherence").Orientation = constants.PivotFieldOrientation.xlColumnField
ptable.PivotFields("Application Name -> Process -> Activity -> Metric").Orientation = constants.PivotFieldOrientation.xlDataField

pws.range('A4').value = f'{cmonth} values'
pws.range('A4').api.Font.Bold = True
pws.range('A4:A5').color = (1, 150, 32)
pws.range('C1').value = f'{pmonth} values'
pws.range('C1').api.Font.Bold = True
pws.range('C1:D1').color = (1, 150, 32)
pws.autofit()

#formatting rest of the sheets
ws_app=wb.sheets['app']
data_range = ws_app.range('A1').expand('table')
table=ws_app.tables.add(source=data_range, name = 'app', has_headers=True)
table.table_style='TableStyleMedium2'
ws_app.range('B2').expand('table').number_format = '0.00%'
ws_app.autofit()

ws_proc=wb.sheets['process']
data_range = ws_proc.range('A1').expand('table')
table=ws_proc.tables.add(source=data_range, name = 'process', has_headers=True)
table.table_style='TableStyleMedium1'
ws_proc.range('B2:D12').number_format = '0.00%'
ws_proc.autofit()

ws_act=wb.sheets['activity']
data_range = ws_act.range('A1').expand('table')
table=ws_act.tables.add(source=data_range, name = 'activity', has_headers=True)
table.table_style='TableStyleMedium4'
ws_act.range('C2').expand('table').number_format = '0.00%'
ws_act.autofit()

ws_met=wb.sheets['metric']
data_range = ws_met.range('A1').expand('table')
table=ws_met.tables.add(source=data_range, name = 'metric', has_headers=True)
table.table_style='TableStyleMedium5'
ws_met.range('C2').expand('table').number_format = '0.00%'
ws_met.autofit()

ws_ad=wb.sheets['change']
data_range = ws_ad.range('A1').expand('table')
table=ws_ad.tables.add(source=data_range, name = 'change', has_headers=True)
table.table_style='TableStyleMedium5'
ws_ad.autofit()

wb.save()
#wb.close()