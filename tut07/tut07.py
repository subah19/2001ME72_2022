from datetime import datetime
start_time = datetime.now()

#Help
def octant_analysis(mod=5000):
	pass


###Code
#importing required libraries
import os
import pandas as pd
import numpy as np
import math
import openpyxl as xl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font, fills


# Mod rank count function
def octant_range_names(mod, filename):
    octant_name_id_mapping = {"1":"Internal outward interaction", "-1":"External outward interaction", "2":"External Ejection", "-2":"Internal Ejection", "3":"External inward interaction", "-3":"Internal inward interaction", "4":"Internal sweep", "-4":"External sweep"}
    
    # Reading Excel File
    
    df = pd.read_excel(f'input\\{filename}')
    rows = df.shape[0]
    
    
    try:
    
        u_avg = df['U'].mean()
        v_avg = df['V'].mean()
        w_avg = df['W'].mean()

        
        df.insert(4, column="U Avg", value="")
        df.insert(5, column="V Avg", value="")
        df.insert(6, column="W Avg", value="")

    
        df.insert(7, column="U'=U - U avg", value="")
        df.insert(8, column="V'=V - V avg", value="")
        df.insert(9, column="W'=W - W avg", value="")

        df["U'=U - U avg"] = round(df['U'] - u_avg, 3)
        df["V'=V - V avg"] = round(df['V'] - v_avg, 3)
        df["W'=W - W avg"] = round(df['W'] - w_avg , 3)
        
        df.at[0, 'U Avg'] = round(u_avg, 3)
        df.at[0, 'V Avg'] = round(v_avg, 3)
        df.at[0, 'W Avg'] = round(w_avg, 3)
        
        df['U'] = round(df['U'], 3)
        df['V'] = round(df['V'], 3)
        df['W'] = round(df['W'], 3)
        df['T'] = round(df['T'], 3)
    except:
        print("Error in calculating average!")
        exit()
        
    try:

        df.insert(10, column="Octant", value="")
        df.insert(11, column=" ", value="")
        df.insert(12, column="", value="")
        df.insert(13, column="Octant ID", value="")
        df.insert(14, column="1", value="")
        df.insert(15, column="-1", value="")
        df.insert(16, column="2", value="")
        df.insert(17, column="-2", value="")
        df.insert(18, column="3", value="")
        df.insert(19, column="-3", value="")
        df.insert(20, column="4", value="")
        df.insert(21, column="-4", value="")

        df.iloc[1, 12] = "Mod "+ str(mod)
        df.at[0, 'Octant ID'] = "Overall Count"
        l=[]
    except:
        print("Error in inserting columns.")
        exit()

    try:
        # Calculating the octant values
        for i in range(0, rows):
            if df.at[i,"U'=U - U avg"] >= 0 and  df.at[i,"V'=V - V avg"] >= 0:
                if df.at[i,"W'=W - W avg"] >= 0:
                  df.at[i, 'Octant'] = 1
                else:
                  df.at[i, 'Octant'] = -1
            elif df.at[i,"U'=U - U avg"] < 0 and  df.at[i,"V'=V - V avg"] >= 0:
                if df.at[i,"W'=W - W avg"] >= 0:
                  df.at[i, 'Octant'] = 2
                else:
                  df.at[i, 'Octant'] = -2
            elif df.at[i,"U'=U - U avg"] < 0 and  df.at[i,"V'=V - V avg"] < 0:
                if df.at[i,"W'=W - W avg"] >= 0:
                  df.at[i, 'Octant'] = 3
                else:
                  df.at[i, 'Octant'] = -3
            elif df.at[i,"U'=U - U avg"] >= 0 and  df.at[i,"V'=V - V avg"] < 0:
                if df.at[i,"W'=W - W avg"] >= 0:
                  df.at[i, 'Octant'] = 4
                else:
                  df.at[i, 'Octant'] = -4
            l.append(df.at[i, 'Octant'])

        df.at[0, "1"] = l.count(1)
        df.at[0, "-1"] = l.count(-1)
        df.at[0 ,"2"] = l.count(2)
        df.at[0 ,"-2"] = l.count(-2)
        df.at[0 ,"3"] = l.count(3)
        df.at[0 ,"-3"] = l.count(-3)
        df.at[0 ,"4"] = l.count(4)
        df.at[0 ,"-4"] = l.count(-4)
    except:
        print("Error in counting octant values.")
        exit()
    