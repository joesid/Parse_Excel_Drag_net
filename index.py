import pandas as pd

# Load the Excel sheet into a pandas dataframe
df = pd.read_excel('input_file.xlsx')

# Create an empty dataframe for the parsed names
parsed_df = pd.DataFrame(columns=['First Name', 'Last Name'])

# Loop through each row in the Full Name column
for name in df['Full Name']:
    # Split the name into separate strings
    name_parts = name.split()
    if len(name_parts) == 2:
        # If there are only two strings, use them as first and last name
        first_name = name_parts[0]
        last_name = name_parts[1]
    elif len(name_parts) > 2:
        # If there are more than two strings, use the first two as first name and the last as last name
        first_name = name_parts[0] + ' ' + name_parts[1]
        last_name = name_parts[-1]
    else:
        # If there is only one string, use it as the last name
        first_name = ''
        last_name = name_parts[0]
    
    # Add the parsed name to the new dataframe
    #DEPRECATED IN PANDA VERSION 2.0  ->    parsed_df = parsed_df.append({'First Name': first_name, 'Last Name': last_name}, ignore_index=True)

    parsed_df = pd.concat([parsed_df, pd.DataFrame({'First Name': first_name, 'Last Name':last_name}, index=[0])], ignore_index=True)


# Save the new dataframe to a new Excel sheet
parsed_df.to_excel('output_file.xlsx', index=False)
