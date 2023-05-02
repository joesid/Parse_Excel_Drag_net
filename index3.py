import pandas as pd

# Load the first sheet
df1 = pd.read_excel('1st_Sheet.xlsx')

# Load the second sheet
df2 = pd.read_excel('2nd_Sheet.xlsx')

# Make all the string values lower case for case insensitive matching
#df1 = df1.applymap(lambda x: x.lower() if isinstance(x, str) else x)
#df2 = df2.applymap(lambda x: x.lower() if isinstance(x, str) else x)

# Create an empty DataFrame to store the matching rows
matched_df = pd.DataFrame(columns=['FirstName', 'MiddleName', 'LastName', 'Email'])


for index1, row1 in df1.iterrows():
    for index2, row2 in df2.iterrows():
        if row1.str.lower().equals(row2.str.lower()):
            matched_df = pd.concat([matched_df, row1[['FirstName', 'MiddleName', 'LastName', 'Email']]], ignore_index=True)

# Iterate through the rows in the first sheet
#for _, row1 in df1.iterrows():
    # Iterate through the rows in the second sheet
   # for _, row2 in df2.iterrows():
        # Check if the FirstName, MiddleName, and LastName columns match
        #if row1['FirstName'] == row2['FirstName'] and row1['MiddleName'] == row2['MiddleName'] and row1['LastName'] == row2['LastName']:
            # If there's a match, add the row from the first sheet to the matched_df
           # matched_df = matched_df.append(row1[['FirstName', 'MiddleName', 'LastName', 'Email']])

# Save the matched rows to a new Excel file
matched_df.to_excel('matched_rows.xlsx', index=False)

