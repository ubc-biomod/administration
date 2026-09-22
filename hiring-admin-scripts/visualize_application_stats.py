import pandas as pd

df = pd.read_excel("./BIOMOD+Hiring+2026_27+Application_September+21,+2026_15.12.xlsx", skiprows=1)

print(df.shape)

df1 = df[df['First and Last Name'].str.len() > 1]
print(df1.shape) 


ser1 = df1['Which subteam are you applying for? (Select multiple if you are interested in applying for multiple)\n\nTo see the specific responsibilities, our hiring package link is here: https://docs.google.com/document/d/1rfUi9SrG0KYOu2ElkSTnx5-fK_Zpv2s6OSDZHXQ0UGo/edit?usp=sharing']
ser2 = ser1.str.split(",")

flat = ser2.explode(ignore_index=True)
print(flat.shape)

print(flat.value_counts())

print(df1['Faculty'].value_counts())