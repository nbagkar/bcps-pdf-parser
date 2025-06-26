import pandas as pd

# Load the existing Excel file
df = pd.read_excel('master_table.xlsx')

# Remove values in column C (Budget Type) if column D (Funding) is empty
df.loc[df['Funding'].isna() | (df['Funding'] == ''), 'Budget Type'] = ''

# Save the cleaned data back to the Excel file
df.to_excel('master_table.xlsx', index=False)

print("Cleaned the Excel file: removed Budget Type values where Funding is empty.") 