import pandas as pd
import datetime as dt
from datetime import timedelta, datetime
import random
import sys
import numpy as np

dateparse = lambda x: dt.datetime.strptime(x, '%Y-%m-%d')

print("Importing Data...")
#get control sheet
url = r'C:\Users\achie\Github\Gooselabs\Control.xlsx'
Control = pd.read_excel(url, sheet_name = 0, header=0, index_col=0)
Venue_Summary = pd.read_excel(url, sheet_name = 1, header=0, index_col=0)
Round_Summary = pd.read_excel(url, sheet_name = 2, header=0, index_col=0)
Quiz_Difficulty = pd.read_excel(url, sheet_name = 3, header=0, index_col=0)
Round_Difficulty = pd.read_excel(url, sheet_name = 4, header=0, index_col=0)

Exclude_Date = Control.loc['Easy', 'Exclude Venue Q Dates']

#get quiz requests
location = r'C:\Users\achie\Github\Gooselabs\3.Quiz_Requests\Example_Quiz_Request.csv'

Quiz_Requests = pd.read_csv(location)

#get venue database
location = r'C:\Users\achie\Github\Gooselabs\2.Databases\Venue_all_time_average.csv'

V_Db = pd.read_csv(location)
V_Db = V_Db.drop(columns={"Unnamed: 0"})

#get quiz database
location = r'C:\Users\achie\Github\Gooselabs\2.Databases\Q_all_time_average.csv'

Q_Db_Mas = pd.read_csv(location)
Q_Db_Mas = Q_Db_Mas.drop(columns={"Unnamed: 0"})

Q_Db_Mas['Round'] = Q_Db_Mas['QRef'].str.slice(3,6)

Q_Db_Mas['Mean'] = pd.to_numeric(Q_Db_Mas['Mean'], errors='coerce')

Q_Db_Mas['Mean'] = Q_Db_Mas['Mean'].astype('float')

#get round database
location = r'C:\Users\achie\Github\Gooselabs\2.Databases\Round_avg.csv'

R_Db = pd.read_csv(location)
R_Db = R_Db.drop(columns={"Unnamed: 0"})

#get last asked database
location = r'C:\Users\achie\Github\Gooselabs\2.Databases\Last_asked.csv'

LA_Db = pd.read_csv(location, parse_dates=['Date'], date_parser=dateparse)
LA_Db = LA_Db.drop(columns={"Unnamed: 0"})

QCount = len(Quiz_Requests)

All_Quiz_Qs = pd.DataFrame()

for x in range(0, QCount):
#    x = 0

    i = "Building Quizzes: " + "{0:.1f}%".format(((x+1)/QCount)*100)
    sys.stdout.write("\r" + i)
    #get venue round difficulty ranges
    Q_Loc = Quiz_Requests.loc[x, 'Location']
    Q_Round = Quiz_Requests.loc[x, 'Round']
    QPR = Quiz_Requests.loc[x, 'QPR']
    Bespoke = Quiz_Requests.loc[x, 'Bespoke']
    
    Easy_Q_No = Quiz_Difficulty.loc['Easy', QPR]
    Med_Q_No = Quiz_Difficulty.loc['Medium', QPR]
    Hard_Q_No = Quiz_Difficulty.loc['Hard', QPR]   
    
    Q_low = Round_Difficulty.loc['Easy', Q_Round]
    Q_high = Round_Difficulty.loc['Medium', Q_Round] 

    V_Diff = V_Db[V_Db['Location'] == Q_Loc]  
    V_Diff = V_Diff[V_Diff['Round'] == Q_Round]
    
    

    #if venue has round difficulty data adjust for them:
    if len(V_Diff) > 0 and Bespoke == "Y":
           
        V_Diff = V_Diff.reset_index()
        V_Round_Diff = float(V_Diff.loc[0, 'Mean'])
        
        #get average round difficulty
        R_Diff = R_Db[R_Db['Round'] == Q_Round]
        R_Diff = R_Diff.reset_index()
        Round_Diff = float(R_Diff.loc[0, 'Mean'])
        
        V_Delta = V_Round_Diff - Round_Diff
        
        Q_low = Q_low - V_Delta
        Q_high = Q_high - V_Delta





    #get appropriate ranges   
    Q_Db = Q_Db_Mas[Q_Db_Mas['Round'] == Q_Round]
    
    Easy_Db = Q_Db[Q_Db['Avg'] > Q_low]
    Easy_Db = Easy_Db.loc[Easy_Db['Avg'] < 1 - V_Delta]
    
    Med_Db = Q_Db[Q_Db['Avg'] > Q_high]
    Med_Db = Med_Db.loc[Med_Db['Avg'] < Q_low]
    
    Diff_Db = Q_Db[Q_Db['Avg'] > 0 - V_Delta]
    Diff_Db = Diff_Db.loc[Diff_Db['Avg'] < Q_high]



    #exclude question with a year
    Date_threshold = datetime.now().replace(microsecond=0).isoformat(' ')
    Date_threshold = datetime.strptime(Date_threshold, '%Y-%m-%d %H:%M:%S')
    Date_threshold = Date_threshold - timedelta(Exclude_Date)
    RQ_Db = LA_Db[(LA_Db['Date'] > Date_threshold)]
    
    Recent_Qs = RQ_Db['QRef']
    
    Easy_Db = Easy_Db[~Easy_Db['QRef'].isin(Recent_Qs)]
    Med_Db = Med_Db[~Med_Db['QRef'].isin(Recent_Qs)]
    Diff_Db = Diff_Db[~Diff_Db['QRef'].isin(Recent_Qs)]


    #cleanse
    Easy_Db = Easy_Db.reset_index()
    Easy_Db['index'] = Easy_Db.index
    Med_Db = Med_Db.reset_index()
    Med_Db['index'] = Med_Db.index
    Diff_Db = Diff_Db.reset_index()
    Diff_Db['index'] = Diff_Db.index

    
    
    
    
    
    
    #select questions of right difficulty
    Easy_len = len(Easy_Db)
    if Easy_len > Easy_Q_No:
        Easy_Qs = random.sample(range(0, Easy_len), Easy_Q_No)
        Easy_Db = Easy_Db[Easy_Db['index'].isin(Easy_Qs)]
        Easy_Db['Difficulty'] = 'Easy'
        Easy_Db['Total_Pop'] = Easy_len
    else:
        Easy_Db['Difficulty'] = 'Easy'
        Easy_Db['Total_Pop'] = 'ERROR - NOT ENOUGH QUESTIONS'
        
    Med_len = len(Med_Db)
    if Med_len > Med_Q_No:
        Med_Qs = random.sample(range(0, Med_len), Med_Q_No)
        Med_Db = Med_Db[Med_Db['index'].isin(Med_Qs)]
        Med_Db['Difficulty'] = 'Medium'
        Med_Db['Total_Pop'] = Med_len
    else:
        Med_Db['Difficulty'] = 'Medium'
        Med_Db['Total_Pop'] = 'ERROR - NOT ENOUGH QUESTIONS'

    Diff_len = len(Diff_Db)
    if Diff_len > Hard_Q_No:
        Diff_Qs = random.sample(range(0, Diff_len), Hard_Q_No)
        Diff_Db = Diff_Db[Diff_Db['index'].isin(Diff_Qs)]
        Diff_Db['Difficulty'] = 'Hard'
        Diff_Db['Total_Pop'] = Diff_len
    else:
        Diff_Db['Difficulty'] = 'Hard'
        Diff_Db['Total_Pop'] = 'ERROR - NOT ENOUGH QUESTIONS'

    
    
    Round = Easy_Db.append(Med_Db)
    Round = Round.append(Diff_Db)
    
    Round = Round.drop(columns={"index", "Min"})
    Round['Bespoke'] = Bespoke
    Round['Location'] = Q_Loc
    
   
    
    Round = Round[['Location', 'Round', 'Bespoke', 'QRef', 'Difficulty', 'Avg', 'Max', 'Count', 'Total_Pop']]
    
    if len(Round) == 0:
        Round.loc[0, 'Location'] = Q_Loc
        Round.loc[0, 'Round'] = Q_Round
        Round.loc[0, 'Total_Pop'] = 'ERROR - NOT ENOUGH QUESTIONS'
    
    All_Quiz_Qs = All_Quiz_Qs.append(Round)


All_Quiz_Qs['Length'] = All_Quiz_Qs['QRef'].apply(len)
All_Quiz_Qs['QRef'] = np.where(All_Quiz_Qs['Length']==8, All_Quiz_Qs['QRef'].str.slice(0,7) + "0" + All_Quiz_Qs['QRef'].str.slice(7,8),All_Quiz_Qs['QRef'])
    
All_Quiz_Qs = All_Quiz_Qs[['Location', 'Round', 'Bespoke', 'QRef', 'Difficulty', 'Avg', 'Max', 'Count', 'Total_Pop']]
    
All_Quiz_Qs.to_csv(r'C:\Users\achie\Github\Gooselabs\4.Quiz_Outputs\Todays_Quizzes.csv')


url = r'C:\Users\achie\Github\Gooselabs\Control.xlsx'
writer = pd.ExcelWriter(url , engine='xlsxwriter')
Control.to_excel(writer, sheet_name='Parameters')
Venue_Summary.to_excel(writer, sheet_name='Venues')
Round_Summary.to_excel(writer, sheet_name='Rounds')
Quiz_Difficulty.to_excel(writer, sheet_name='Quiz_Diff_Structure')
Round_Difficulty.to_excel(writer, sheet_name='Round_Diff')
writer.save()
