import pandas as pd

# Load the existing Excel file
df = pd.read_excel('master_table.xlsx')

# Initialize variables to track the last seen budget type
last_budget_type = ''

# Iterate through the DataFrame row by row
for index, row in df.iterrows():
    # Check if column D (Funding) is filled and column C (Budget Type) is empty
    if pd.notna(row['Funding']) and row['Funding'] != '' and (pd.isna(row['Budget Type']) or row['Budget Type'] == ''):
        # Fill column C with the last seen budget type
        df.at[index, 'Budget Type'] = last_budget_type
    # Update the last seen budget type if column C is filled
    elif pd.notna(row['Budget Type']) and row['Budget Type'] != '':
        last_budget_type = row['Budget Type']

# Save the updated data back to the Excel file
df.to_excel('master_table.xlsx', index=False)

print("Filled Budget Type column with the most recent value from above where Funding is filled.") 