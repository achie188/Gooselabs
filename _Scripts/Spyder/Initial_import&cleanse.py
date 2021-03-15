import importlib.util
import datetime as dt
from datetime import date, timedelta, datetime
import pandas as pd
pd.set_option('mode.chained_assignment', None)

location2 = r'C:\Users\achie\Github\Gooselabs'

Importer = location2 +r'\Initial_import&cleanse.py'
Q_Analysis = location2 +r'\Questions_analysis.py'
V_Analysis = location2 +r'\Venue_analysis.py'

Quiz_Builder = location2 +r'\Quiz_builder.py'


Import = 0
Questions_Analysis = 1
Venue_Analysis = 1

Build_Quiz = 0



if Import == 1:
    print("Importing new data...")
    
    spec = importlib.util.spec_from_file_location("Initial_import&cleanse", Importer)
    Imptr = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(Imptr)

    print("Done")
    print("-------------------------------------------------------------------------")
    print(" ")   
    
  
if Questions_Analysis == 1:
    
    print("QUESTION ANALYSER")  
    
    spec = importlib.util.spec_from_file_location("Questions_analysis", Q_Analysis)
    Q_A = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(Q_A)

    print("Done")
    print("-------------------------------------------------------------------------")
    print(" ")   
    
    
if Venue_Analysis == 1:
    
    print("VENUE ANALYSER")  
    
    spec = importlib.util.spec_from_file_location("Venue_analysis", V_Analysis)
    V_A = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(V_A)

    print("Done")
    print("-------------------------------------------------------------------------")
    print(" ")   
       
    
if Build_Quiz == 1:
    
    print("QUIZ BUILDER")   
    
    spec = importlib.util.spec_from_file_location("Quiz_builder", Build_Quiz)
    B_Q = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(B_Q)
    
    print("Done")
    print("-------------------------------------------------------------------------")
    print(" ")  