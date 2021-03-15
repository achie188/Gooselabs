import pandas as pd
import datetime as dt
from datetime import timedelta, datetime

dateparse = lambda x: dt.datetime.strptime(x, '%Y-%m-%d')

#get control sheet
url = r'C:\Users\achie\Github\Gooselabs\Control.xlsx'
Control = pd.read_excel(url, sheet_name = 0, header=0, index_col=0)
Venue_Summary = pd.read_excel(url, sheet_name = 1, header=0, index_col=0)
Round_Summary = pd.read_excel(url, sheet_name = 2, header=0, index_col=0)
Quiz_Difficulty = pd.read_excel(url, sheet_name = 3, header=0, index_col=0)
Round_Difficulty = pd.read_excel(url, sheet_name = 4, header=0, index_col=0)


Q_Dates = Control.loc['Easy', 'Recent date range']
Q_Count = Control.loc['Easy', 'Recent count range'] - 1

print("Importing Data...")
#get data
location = r'C:\Users\achie\Github\Gooselabs\2.Databases\All_Data.csv'

A_Db = pd.read_csv(location, parse_dates=['Date'], date_parser=dateparse)


print("Analysing questions...")
#get all time question average
All_time_avg = A_Db.groupby(['QRef'], as_index=False).agg({'Correct': ['mean', 'min', 'max', 'count']})
All_time_avg.columns = ['QRef', 'Mean', 'Min', 'Max', 'Count']

All_time_avg['Avg'] = All_time_avg['Mean'] / All_time_avg['Max']

All_time_avg.to_csv(r'C:\Users\achie\Github\Gooselabs\2.Databases\Q_all_time_average.csv')


#question difficulty analysis
All_time_avg = All_time_avg.sort_values(by=['Avg'], ascending=False) 
All_time_avg = All_time_avg.reset_index()
All_time_avg['Index'] = All_time_avg.index

NQs = len(All_time_avg)
E_Threshold = Control.loc['Easy', 'Threshold']
M_Threshold = Control.loc['Medium', 'Threshold']
H_Threshold = Control.loc['Hard', 'Threshold']

MediumQs = round(NQs * M_Threshold)
HardQs = round(NQs * H_Threshold)
EasyQs = NQs - MediumQs - HardQs

E_Per_Threshold = All_time_avg.loc[NQs - EasyQs, 'Avg']
M_Per_Threshold = All_time_avg.loc[NQs - MediumQs, 'Avg']
H_Per_Threshold = All_time_avg.loc[NQs - HardQs, 'Avg']

Round_Difficulty.loc['Easy', 'Overall'] = E_Per_Threshold
Round_Difficulty.loc['Medium', 'Overall'] = M_Per_Threshold
Round_Difficulty.loc['Hard', 'Overall'] = H_Per_Threshold



#get recent question average - last X days
Date_threshold = datetime.now().replace(microsecond=0).isoformat(' ')
Date_threshold = datetime.strptime(Date_threshold, '%Y-%m-%d %H:%M:%S')
Date_threshold = Date_threshold - timedelta(Q_Dates)

#A_Db['Date'] = pd.to_datetime(A_Db['Date'])
Recent_Db = A_Db[(A_Db['Date'] > Date_threshold)]
Recent_date_avg = Recent_Db.groupby(['QRef'], as_index=False).agg({'Correct': ['mean', 'min', 'max', 'count']})
Recent_date_avg.columns = ['QRef', 'Mean', 'Min', 'Max', 'Count']

#get last X - average
A_Db = A_Db.sort_values(by=['Date'], ascending=False) 
A_Db['xAsked'] = A_Db.groupby('QRef').cumcount()

Recent_count_Db = A_Db[(A_Db['xAsked'] < Q_Count)]
Recent_count_avg = Recent_count_Db.groupby(['QRef'], as_index=False).agg({'Correct': ['mean', 'min', 'max', 'count']})
Recent_count_avg.columns = ['QRef', 'Mean', 'Min', 'Max', 'Count']

Recent_date_avg.to_csv(r'C:\Users\achie\Github\Gooselabs\2.Databases\Q_Recent_date_avg.csv')
Recent_count_avg.to_csv(r'C:\Users\achie\Github\Gooselabs\2.Databases\Q_Recent_count_avg.csv')


url = r'C:\Users\achie\Github\Gooselabs\Control.xlsx'
writer = pd.ExcelWriter(url , engine='xlsxwriter')
Control.to_excel(writer, sheet_name='Parameters')
Venue_Summary.to_excel(writer, sheet_name='Venues')
Round_Summary.to_excel(writer, sheet_name='Rounds')
Quiz_Difficulty.to_excel(writer, sheet_name='Quiz_Diff_Structure')
Round_Difficulty.to_excel(writer, sheet_name='Round_Diff')
writer.save()