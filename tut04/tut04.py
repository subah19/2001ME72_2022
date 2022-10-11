import pandas as pd
import numpy as np
import openpyxl as opxl

df = pd.read_excel('C:/Users/hp/Documents/GitHub/CS384_2022/tut04/input_octant_longest_subsequence_with_range.xlsx', 'Sheet1')

df['Uavg'] = df['U'].mean()
df['Vavg'] = df['V'].mean()
df['Wavg'] = df['W'].mean()
df['U-Uavg'] = df['U'] - df['Uavg']
df['V-Vavg'] = df['V'] - df['Vavg']
df['W-Wavg'] = df['W'] - df['Wavg']

def identify_octant(x, y, z):
    if(x>0 and y>0 and z>0):
        octant = 1
    if(x>0 and y>0 and z<0):
        octant = -1
    if(x<0 and y>0 and z>0):
        octant = 2
    if(x<0 and y>0 and z<0):
        octant = -2
    if(x<0 and y<0 and z>0):
        octant = 3
    if(x<0 and y<0 and z<0):
        octant = -3
    if(x>0 and y<0 and z>0):
        octant = 4
    if(x>0 and y<0 and z<0):
        octant = -4
    return octant

n = len(df)

def octant_identification_count(df):
    octants = []
    octant_count = {1:0, -1:0, 2:0, -2:0, 3:0, -3:0, 4:0, -4:0}
    for i in (-1, 1, 2, -2, 3, -3, 4, -4):
        octant_count[i] = 0
    for i in range(n):
        x = df.loc[i, "U-Uavg"]
        y = df.loc[i, "V-Vavg"]
        z = df.loc[i, "W-Wavg"]
        octants.append(identify_octant(x, y, z))
        octant_count[identify_octant(x, y, z)] = octant_count[identify_octant(x, y, z)]+1
    # print(octants, '\n')
    # print(octant_count, '\n')
    return (octants, octant_count)

octants_overall = octant_identification_count(df)[0]
df['Octants'] = octants_overall

len_count_matrix = np.zeros((8, 2), int)
from_to = ([],[],[],[],[],[],[],[])

for i in range(n):
    octant_index = {1:0, -1:1, 2:2, -2:3, 3:4, -3:5, 4:6, -4:7}
    count = 1
    max = len_count_matrix[octant_index[df['Octants'][i]]][0]
    while i<n-1 and df['Octants'][i]==df['Octants'][i+1]:
        count = count + 1
        i = i + 1
    if count>max:
        len_count_matrix[octant_index[df['Octants'][i]]][0] = count

for i in range(n):
    octant_index = {1:0, -1:1, 2:2, -2:3, 3:4, -3:5, 4:6, -4:7}
    count = 1
    max = len_count_matrix[octant_index[df['Octants'][i]]][0]
    while i<n-1 and df['Octants'][i]==df['Octants'][i+1]:
        count = count + 1
        i = i + 1
    if count==max:
        len_count_matrix[octant_index[df['Octants'][i]]][1] = len_count_matrix[octant_index[df['Octants'][i]]][1] + 1
        from_to[octant_index[df['Octants'][i]]].append((df['Time'][i],df['Time'][i-max+1]))

df['Octant ID'] = ''
df['Longest Subsequence Length'] = ''
df['Count'] = ''

possible_octant_values = [1, -1, 2, -2, 3, -3, 4, -4]

for i in range(8):
    df['Octant ID'][i] = possible_octant_values[i]
    df['Longest Subsequence Length'][i] = len_count_matrix[i][0]
    df['Count'][i] = len_count_matrix[i][1]

df['Octant_ID'] = ''
df['Longest_Subsequence_Length'] = ''
df['Count_of_Longest_Subsequnces'] = ''

row_number = 0

for i in range(8):
    df['Octant_ID'][row_number] = possible_octant_values[i]
    df['Longest_Subsequence_Length'][row_number] = len_count_matrix[i][0]
    df['Count_of_Longest_Subsequnces'][row_number] = len_count_matrix[i][1]
    row_number = row_number+1
    df['Octant_ID'][row_number] = 'Time'
    df['Longest_Subsequence_Length'][row_number] = 'From'
    df['Count_of_Longest_Subsequnces'][row_number] = 'To'
    row_number = row_number+1
    for j in range(len(from_to[i])):
        df['Longest_Subsequence_Length'][row_number] = from_to[i][j][0]
        df['Count_of_Longest_Subsequnces'][row_number] = from_to[i][j][1]
        row_number = row_number+1

df.to_excel('octant_output_4.xlsx')