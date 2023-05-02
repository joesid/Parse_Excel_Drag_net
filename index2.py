import pandas as pd

# Read in the first sheet
df1 = pd.read_excel('1st_Sheet.xlsx')

# Read in the second sheet
df2 = pd.read_excel('2nd_Sheet.xlsx')

print(df1.columns)
print(df2.columns)

#error as a result of merge object and float64, this is to change conent to sttring
df2['Email'] = df2['Email'].astype(str)
df2['FirstName'] = df2['FirstName'].astype(str)
df2['MiddleName'] = df2['MiddleName'].astype(str)
df2['LastName'] = df2['LastName'].astype(str)


#Merge the two dataframes based on matching first, middle, and last names
merged_df = pd.merge(df1, df2, on=['FirstName', 'MiddleName', 'LastName', 'Email'])

# Create a new dataframe with the desired columns
new_df = merged_df[['FirstName', 'MiddleName', 'LastName', 'Email']]

#Write the new dataframe to a new excel file
new_df.to_excel('new_sheet.xlsx', index=False)
