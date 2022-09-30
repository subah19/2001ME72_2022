import pandas as pd

from platform import python_version
ver = python_version()

if ver == "3.8.10":
    print("Correct Version Installed")
else:
    print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")
mod=5000

df = pd.read_csv("octant_input.csv")

#-----------------------------------------------------

# Average of U
U_avg = df['U'].mean()

# Average of V
V_avg = df['V'].mean()

# Average of W
W_avg = df['W'].mean()

#-----------------------------------------------------

# defining columns
df["U_avg"] = U_avg
# this extra line is there to print U_avg once, else it would have filled all the rows with the same value(U_avg)
df["U_avg"] = df["U_avg"].head(1)

df["V_avg"] = V_avg
df["V_avg"] = df["V_avg"].head(1)

df["W_avg"] = W_avg
df["W_avg"] = df["W_avg"].head(1)

#-----------------------------------------------------

U_minus_U_avg = (df['U'] - U_avg)
V_minus_V_avg = (df['V'] - V_avg)
W_minus_W_avg = (df['W'] - W_avg)

df["U' = U - U_avg"] = U_minus_U_avg
df["V' = V - V_avg"] = V_minus_V_avg
df["W' = W - W_avg"] = W_minus_W_avg

#-----------------------------------------------------

# defining an array that will store the octant values of the points
octant = []

for i in range(len(df)):

    # if U is +ve
    if df.loc[i, "U' = U - U_avg"]>0:
        #if V is +ve
        if df.loc[i, "V' = V - V_avg"]>0:
            # perform for W
            if df.loc[i, "W' = W - W_avg"]>0:
                octant.append(int(1))
            else:
                octant.append(int(-1))

        # if V is -ve
        else:
            if df.loc[i, "W' = W - W_avg"]>0:
                octant.append(int(4))
            else:
                octant.append(int(-4))

    #if U is -ve
    else:
        #if V is +ve
        if df.loc[i, "V' = V - V_avg"]>0:
            # perform for W
            if df.loc[i, "W' = W - W_avg"]>0:
                octant.append(int(2))
            else:
                octant.append(int(-2))

        # if V is -ve
        else:
            if df.loc[i, "W' = W - W_avg"]>0:
                octant.append(int(3))
            else:
                octant.append(int(-3))

# defining new column Octant and assigning it values of octant array
df["Octant"] = octant

#--------------------------------------------------

# Printing and tabulating overall octant data
top_row = ["", "Octant ID", "1", "-1", "2", "-2", "3", "-3", "4", "-4"]
for i in range(len(top_row)):
    df.insert(i+11, top_row[i], value="")

df.iloc[0, 12] = "Overall count"

for i in range(8):
    df.iloc[0, i+13] = octant.count(int(top_row[2+i]))

#--------------------------------------------------

df.iloc[1, 11] = "User input"
df.iloc[1, 12] = "Mod " + str(mod)

i=0
k=3
df.iloc[2,12] = f"{i}-{i+mod-1}"
i+=mod

while i<len(df):
    # printing ranges
    df.iloc[k, 12] = f"{i}-{min(i+mod-1, len(df))}"
    # move to next row k+=1
    k+=1
    i+=mod


# defining chunk size to split the octant array
chunk_size = mod
chunked_list = []

# chunked_list is and array of arrays, whose length is chunk_size(mod)
for i in range(0, len(octant), chunk_size):
    chunked_list.append(octant[i:i+chunk_size])

# printing the counts at their positions
for m in range(len(chunked_list)):
    for j in range(8):
        df.iloc[m+2, j+13] = chunked_list[m].count(int(top_row[2+j]))

# finally writing output to file
df.to_csv('octant_output.csv', index=False)
