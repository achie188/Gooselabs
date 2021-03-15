import pandas as pd
import numpy as np
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

V_Dates = Control.loc['Medium', 'Recent date range']
V_Count = Control.loc['Medium', 'Recent count range'] - 1

print("Importing Data...")
#get data
location = r'C:\Users\achie\Github\Gooselabs\2.Databases\All_Data.csv'

A_Db = pd.read_csv(location, parse_dates=['Date'], date_parser=dateparse)


print("Analysing venues...")
#all time venue averages
Venue_avgs = A_Db.groupby(['Location', 'Round'], as_index=False).agg({'Correct': ['mean', 'count']})
Venue_avgs.columns = ['Location', 'Round', 'Mean', 'Count']
Venue_avgs.to_csv(r'C:\Users\achie\Github\Gooselabs\2.Databases\Venue_all_time_average.csv')


#get recent venue average - last X days
Date_threshold = datetime.now().replace(microsecond=0).isoformat(' ')
Date_threshold = datetime.strptime(Date_threshold, '%Y-%m-%d %H:%M:%S')
Date_threshold = Date_threshold - timedelta(V_Dates)

#A_Db['Date'] = pd.to_datetime(A_Db['Date'])
Recent_Db = A_Db[(A_Db['Date'] > Date_threshold)]
Recent_date_avg = Recent_Db.groupby(['Location', 'Round'], as_index=False).agg({'Correct': ['mean', 'count']})
Recent_date_avg.columns = ['Location', 'Round', 'Mean', 'Count']

#get last X - average
A_Db = A_Db.sort_values(by=['Date'], ascending=False) 
A_Db['xAsked'] = A_Db.groupby('QRef').cumcount()

Recent_count_Db = A_Db[(A_Db['xAsked'] < V_Count)]
Recent_count_avg = Recent_count_Db.groupby(['Location', 'Round'], as_index=False).agg({'Correct': ['mean', 'count']})
Recent_count_avg.columns = ['Location', 'Round', 'Mean', 'Count']

Recent_date_avg.to_csv(r'C:\Users\achie\Github\Gooselabs\2.Databases\Venue_Recent_date_avg.csv')
Recent_count_avg.to_csv(r'C:\Users\achie\Github\Gooselabs\2.Databases\Venue_Recent_count_avg.csv')



#overall round averages
Round_avgs = A_Db.groupby(['Round'], as_index=False).agg({'Correct': ['mean', 'count']})
Round_avgs.columns = ['Round', 'Mean', 'Count']

Round_count = A_Db

Round_count['Duplicated'] = Round_count.duplicated(subset = {'QRef'}, keep = 'last')
Round_count['Remove'] = np.where((Round_count['Duplicated']==True), 'Yes', 'No')
Round_count = Round_count[Round_count['Remove'].str.contains("Yes") == False]
Round_count = Round_count.drop(['Duplicated', 'Remove'], axis=1) 

Round_count = Round_count.groupby(['Round'], as_index=False).agg({'Ref': ['count']})
Round_count.columns = ['Round', 'Count']

Round_avgs = Round_avgs.merge(Round_count, on="Round")
Round_avgs.rename(columns = {'Count_x':'xAsked'}, inplace = True) 
Round_avgs.rename(columns = {'Count_y':'QCount'}, inplace = True) 

Round_avgs.to_csv(r'C:\Users\achie\Github\Gooselabs\2.Databases\Round_avg.csv')




#Round Difficulty Range
All_time_avg = A_Db.groupby(['QRef', 'Round'], as_index=False).agg({'Correct': ['mean', 'min', 'max', 'count']})
All_time_avg.columns = ['QRef', 'Round', 'Mean', 'Min', 'Max', 'Count']
All_time_avg['Avg'] = All_time_avg['Mean'] / All_time_avg['Max']

All_time_avg = All_time_avg.sort_values(by=['Avg'], ascending=False) 
All_time_avg = All_time_avg.reset_index()
All_time_avg['Index'] = All_time_avg.index

R_Diffs = Round_avgs.Round.unique()


x = len(R_Diffs)

for y in range(0, x):
#    y=0
    Round = R_Diffs[y]

    Round_Db = All_time_avg[All_time_avg['Round'] == Round]  

    Round_Db = Round_Db.sort_values(by=['Avg'], ascending=False) 
    Round_Db = Round_Db.reset_index()
    Round_Db['Index'] = Round_Db.index
    
    NQs = len(Round_Db)
    E_Threshold = Control.loc['Easy', 'Threshold']
    M_Threshold = Control.loc['Medium', 'Threshold']
    H_Threshold = Control.loc['Hard', 'Threshold']
    
    MediumQs = round(NQs * M_Threshold)
    HardQs = round(NQs * H_Threshold)
    EasyQs = NQs - MediumQs - HardQs
    
    E_Per_Threshold = Round_Db.loc[NQs - EasyQs, 'Avg']
    M_Per_Threshold = Round_Db.loc[NQs - MediumQs, 'Avg']
    H_Per_Threshold = Round_Db.loc[NQs - HardQs, 'Avg']
    
    Round_Difficulty.loc['Easy', Round] = E_Per_Threshold
    Round_Difficulty.loc['Medium', Round] = M_Per_Threshold
    Round_Difficulty.loc['Hard', Round] = H_Per_Threshold



#last asked in venue 
Last_Ask_Db = A_Db[['QRef', 'Location', 'Date']]

Last_Ask_Db['LocQRef'] = Last_Ask_Db['Location'].str.cat(Last_Ask_Db['QRef'])
Last_Ask_Db = Last_Ask_Db.sort_values(by=['Date'])


Last_Ask_Db['Duplicated'] = Last_Ask_Db.duplicated(subset = {'LocQRef'}, keep = 'last')
Last_Ask_Db['Remove'] = np.where((Last_Ask_Db['Duplicated']==True), 'Yes', 'No')
Last_Ask_Db = Last_Ask_Db[Last_Ask_Db['Remove'].str.contains("Yes") == False]
Last_Ask_Db = Last_Ask_Db.drop(['Duplicated', 'Remove'], axis=1) 

Last_Ask_Db.to_csv(r'C:\Users\achie\Github\Gooselabs\2.Databases\Last_asked.csv')




#combine average dataframes
Recent_count_avg['VenueRound'] = Recent_count_avg['Location'].str.cat(Recent_count_avg['Round'])
Recent_date_avg['VenueRound'] = Recent_date_avg['Location'].str.cat(Recent_date_avg['Round'])
Venue_avgs['VenueRound'] = Venue_avgs['Location'].str.cat(Venue_avgs['Round'])

Overall_avg = pd.merge(Venue_avgs, Recent_date_avg, on='VenueRound')
Overall_avg = pd.merge(Overall_avg, Recent_count_avg, on='VenueRound')


Overall_avg = Overall_avg.drop(columns={"Location_x", "Round_x", "Location_y", "Round_y"})
Overall_avg = Overall_avg.rename(columns={"Mean_x":"Overall Avg", "Count_x":"Overall Count", "Mean_y":"Date Avg", "Count_y":"Date Count", "Mean":"Count Avg", "Count":"Count Count"})

Overall_avg = Overall_avg[['Location', 'Round', 'Overall Avg', 'Overall Count', 'Date Avg', 'Date Count', 'Count Avg', 'Count Count']]


url = r'C:\Users\achie\Github\Gooselabs\Control.xlsx'
writer = pd.ExcelWriter(url , engine='xlsxwriter')
Control.to_excel(writer, sheet_name='Parameters')
Overall_avg.to_excel(writer, sheet_name='Venues')
Round_avgs.to_excel(writer, sheet_name='Rounds')
Quiz_Difficulty.to_excel(writer, sheet_name='Quiz_Diff_Structure')
Round_Difficulty.to_excel(writer, sheet_name='Round_Diff')
writer.save()
