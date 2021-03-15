import pandas as pd


location = r'C:\Users\achie\Github\Gooselabs\4.Quiz_Outputs\Todays_Quizzes.csv'

All_Quiz_Qs = pd.read_csv(location)

url = r'C:\Users\achie\Github\Gooselabs\1.Input\Questions Database V2.xlsm'
Questions = pd.read_excel(url, sheet_name = 0, header=0, index_col=0)

Questions['QRef'] = Questions.index

All_Quiz_WQs = All_Quiz_Qs.merge(Questions, on="QRef", how="left")

All_Quiz_WQs.to_csv(r'C:\Users\achie\Github\Gooselabs\4.Quiz_Outputs\Todays_Quizzes_with Qs.csv')
