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
        # Calculating Average Values
        u_avg = df['U'].mean()
        v_avg = df['V'].mean()
        w_avg = df['W'].mean()

        # Calculating Average Value of U, V, W
        df.insert(4, column="U Avg", value="")
        df.insert(5, column="V Avg", value="")
        df.insert(6, column="W Avg", value="")

        # Calculating U', V', W' 
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
        #Inserting new column for Octant
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
    
    try:
        # Splitting list into ranges and finding the count of octant values
        start = 0
        end = len(l)
        step = int(mod)
        idx=1
        rows_tot_mod = math.ceil(rows/step)
        rows_tot = rows_tot_mod
        for i in range(start, end, step):
            x = i
            sub_list = l[x:x+step]
            y = x+step-1
            if y>rows:
                y=rows-1
            df.at[idx ,'Octant ID'] = str(x)+"-"+str(y)
            df.at[idx, '1'] = sub_list.count(1)
            df.at[idx, '-1'] = sub_list.count(-1)
            df.at[idx, '2'] = sub_list.count(2)
            df.at[idx, '-2'] = sub_list.count(-2)
            df.at[idx, '3'] = sub_list.count(3)
            df.at[idx, '-3'] = sub_list.count(-3)
            df.at[idx, '4'] = sub_list.count(4)
            df.at[idx, '-4'] = sub_list.count(-4)
            idx+=1
    except:
        print("Error in counting octant values for ranges!")
        exit()
    
    try:
        # Inserting Rank Columns 
        col_num = 22
        for i in range(1,5):
            header = "Rank Octant "+str(i)
            df.insert(col_num, column=header, value="")
            col_num+=1
            header = "Rank Octant "+str(-1*i)
            df.insert(col_num, column=header, value="")
            col_num+=1
        df.insert(col_num, column="Rank 1 Octant ID", value="")
        col_num+=1
        df.insert(col_num, column="Rank 1 Octant Name", value="")
        col_num+=1
        
        # Calculating rank for Overall Octant Count
        dict={}
        l=[]
        for i in range(1,5):
            oct_cnt = df.at[0, str(i)]
            dict[oct_cnt] = str(i)
            l.append(oct_cnt)
            oct_cnt = df.at[0, str(-1*i)]
            dict[oct_cnt] = str(-1*i)
            l.append(oct_cnt)
        
        l.sort(reverse=True)
        rank = 1
        df.at[0, "Rank 1 Octant ID"] = dict[l[0]]
        df.at[0, "Rank 1 Octant Name"] = octant_name_id_mapping[dict[l[0]]]
        
        for i in l:
            oct_id = "Rank Octant "+dict[i]
            df.at[0, oct_id] = rank
            rank+=1
        
        # Calculating rank for Mod Octant Count
        rank1=[]
        for idx in range(1, rows_tot_mod+1): 
            dict={}
            l=[]
            for i in range(1,5):
                oct_cnt = df.at[idx, str(i)]
                dict[oct_cnt] = str(i)
                l.append(oct_cnt)
                oct_cnt = df.at[idx, str(-1*i)]
                dict[oct_cnt] = str(-1*i)
                l.append(oct_cnt)

            l.sort(reverse=True)
            df.at[idx, "Rank 1 Octant ID"] = dict[l[0]]
            rank1.append(dict[l[0]])
            df.at[idx, "Rank 1 Octant Name"] = octant_name_id_mapping[dict[l[0]]]
            
            rank = 1
            for i in l:
                oct_id = "Rank Octant "+dict[i]
                df.at[idx, oct_id] = rank
                rank+=1
        
        # Count of Rank 1 Mod Values
        idx = rows_tot_mod+5
        df.at[idx, 'Rank Octant 4'] = "Octant ID"
        df.at[idx, 'Rank Octant -4'] = "Octant Name"
        df.at[idx, 'Rank 1 Octant ID'] = "Count of Rank 1 Mod Values"
        idx+=1
        for i in range(1,5):
            oct_id = str(i)
            cnt = rank1.count(oct_id)
            df.at[idx, 'Rank Octant 4'] = oct_id
            df.at[idx, 'Rank Octant -4'] = octant_name_id_mapping[oct_id]
            df.at[idx, 'Rank 1 Octant ID'] = cnt
            idx+=1
            
            oct_id = str(-1*i)
            cnt = rank1.count(oct_id)
            df.at[idx, 'Rank Octant 4'] = oct_id
            df.at[idx, 'Rank Octant -4'] = octant_name_id_mapping[oct_id]
            df.at[idx, 'Rank 1 Octant ID'] = cnt
            idx+=1
            
    except Exception as e:
        print("Error in calculating rank.", e)
            
    try:
        return df
    except:
        print("Error in exporting to CSV.")
        exit()
    
# Octant Transition Count function
def octant_transition_count(mod, df):
    try:
        # Reading Excel File
        rows = df.shape[0]
        step = mod
        cols = df.shape[1]
        df.insert(cols, column="                     ", value="")
        cols+=1
    except:
        print("Error in reading Excel file!")
        exit()
    
    try:
        # Overall Transition Count 
        for l in range(2,12):
            blank = ""
            for i in range(1,l+1):
                blank+=" "
            df.insert(cols, column=blank, value="")
            cols+=1
        blank_dict={}
        bl_len = 4        
        for i in range(1,5):
            blank=""
            for idx in range(0, bl_len):
                blank += " "
            blank_dict[str(i)] = blank
            bl_len+=1
            blank=""
            for idx in range(0, bl_len):
                blank += " "
            blank_dict[str(-1*i)] = blank
            bl_len+=1
        blank_dict['f'] = '  '
        blank_dict['s'] = '   '
        
        idx=0
        df.at[idx, blank_dict['1']] = 'To'
        idx+=1
        df.at[idx, blank_dict['s']] = 'Count'
        for k in range(-4,5):
            if k==0:
                continue
            df.at[idx, blank_dict[str(k)]] =  k
        idx+=1
        df.at[idx, blank_dict['f']] = "From"

    
        data=[]
        df1 = pd.DataFrame(data, index=['1','-1','2','-2','3','-3','4','-4'],
                        columns=['1','-1','2','-2','3','-3','4','-4'])

        df1 = df1.fillna(0)  

        
        for i in range(0,rows-1):
            first = str(df.at[i,'Octant'])
            second = str(df.at[i+1, 'Octant'])
            df1.at[first, second] += 1

    
        for k in range (1,5):
            df.at[idx, blank_dict['s']] = str(k)
            for l in range (-4,5):
                if l==0:
                    continue
                df.at[idx, blank_dict[str(l)]] = df1.at[str(k), str(l)]
            idx+=1
            df.at[idx, blank_dict['s']] = str(-1*k)
            for l in range (-4,5):
                if l==0:
                    continue
                df.at[idx, blank_dict[str(l)]] = df1.at[str(-1*k), str(l)]
            idx+=1
    except Exception as e:
        print("Error in calculating Overall Transition Count!", e)
        exit()

    try:
        # Mod Transition Count
        for i in range(0, rows, step):
            lim = i+step
            if lim>=rows:
                lim = rows
            idx+=2
            df.at[idx, blank_dict['s']] = 'Mod Transition Count'
            idx+=1
            df.at[idx, blank_dict['s']] = str(i)+'-'+str(lim-1)
            df.at[idx, blank_dict['1']] = 'To'
            idx+=1
            df.at[idx, blank_dict['s']] = 'Octant #'
            for k in range(-4,5):
                if k==0:
                    continue
                df.at[idx, blank_dict[str(k)]] =  k
            idx+=1
            df.at[idx, blank_dict['f']] = "From"

            data=[]
            df1 = pd.DataFrame(data, index=['1','-1','2','-2','3','-3','4','-4'],
                            columns=['1','-1','2','-2','3','-3','4','-4'])
            df1 = df1.fillna(0)

            if lim==rows:
                lim-=1
            for j in range(i,lim):
                first = str(df.at[j,'Octant'])
                second = str(df.at[j+1, 'Octant'])
                df1.at[first, second] += 1

    
            for k in range (1,5):
                df.at[idx, blank_dict['s']] = str(k)
                for l in range (-4,5):
                    if l==0:
                        continue
                    df.at[idx, blank_dict[str(l)]] = df1.at[str(k), str(l)]
                idx+=1
                df.at[idx, blank_dict['s']] = str(-1*k)
                for l in range (-4,5):
                    if l==0:
                        continue
                    df.at[idx, blank_dict[str(l)]] = df1.at[str(-1*k), str(l)]
                idx+=1
    except:
        print("Error in calculating Mod Transition Count!")
        exit()
    