import pandas as pd

# Read the existing Excel file
xls_file_path = 'example.xls'
original_xls_df = pd.read_excel(xls_file_path)

# Create a copy of the original DataFrame
xls_df = original_xls_df.copy()

# Define the ranges for each new sheet
ranges = [(1, 30), (31, 43), (44, 56), (57, 69), (70, 90)]

# Insert new rows into xls_df at specific locations
new_row_values = [
    ['New Value 1A', 'New Value 1B', 'New Value 1C'],
    ['New Value 2A', 'New Value 2B', 'New Value 2C'],
    ['New Value 3A', 'New Value 3B', 'New Value 3C'],
    ['New Value 4A', 'New Value 4B', 'New Value 4C']
]
insert_locations = [5, 10, 35, 54, 78]

for loc, values in zip(insert_locations, new_row_values):
    # Create a new DataFrame for the row to insert
    print(f"Inserting at location: {loc}")
    print(f"New row values: {values}")
    
    new_row = pd.DataFrame([values[:len(xls_df.columns)]], columns=xls_df.columns)
    print("New row:")
    print(new_row)
    print("\n")

    # Append the new row to the DataFrame
    xls_df = pd.concat([xls_df.iloc[:loc], new_row, xls_df.iloc[loc:]]).reset_index(drop=True)

# Create a new Excel writer object with the 'openpyxl' engine
with pd.ExcelWriter(xls_file_path, engine='openpyxl') as writer:
    # Write the original data to the original sheet
    original_xls_df.to_excel(writer, sheet_name='Original', index=False)

    # Write data to separate sheets based on specified ranges
    for sheet_number, (start_row, end_row) in enumerate(ranges):
        if sheet_number == 0:
            sheet_name = 'Planilha 1'
        elif sheet_number == 1:
            sheet_name = 'Planilha 2'
        elif sheet_number == 2:
            sheet_name = 'Planilha 3'
        elif sheet_number == 3:
            sheet_name = 'Planilha 4'
        elif sheet_number == 4:
            sheet_name = 'Planilha 5'      

        # Extract the subset of data for the current sheet
        subset_df = xls_df.iloc[start_row - 1:end_row]

        # Write the subset to the new sheet
        subset_df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Data copied into new sheets based on specified ranges. Original data preserved in 'Original' sheet.")
