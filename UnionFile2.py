import pandas as pd
import os

input_loc = r"\\IRGRAFTP\DashboardData\Folder Moving Project\Python Input"
output_loc = r"\\IRGRAFTP\DashboardData\Folder Moving Project"
fileList = os.listdir(input_loc)
data = []

for file in fileList:
    if file.endswith(".xlsx"):
        file_path = os.path.join(input_loc, file)
        sheet1 = pd.read_excel(file_path, sheet_name="Summary")

        # Find the row where the word "Location" appears in Column B (second column)
        location_row = sheet1[sheet1.iloc[:, 1].astype(str).str.contains("Location", na=False)].index

        if len(location_row) > 0:
            # Get the first occurrence of "Location" and select all rows after that
            location_index = location_row[0]
            sheet1 = sheet1.iloc[location_index + 1:].reset_index(drop=True)

        # Reset index to avoid blank rows
        sheet1.reset_index(drop=True, inplace=True)

        # Drop the last 2 rows
        sheet1.drop(sheet1.tail(2).index, inplace=True)

        # Append the cleaned sheet to the list
        data.append(sheet1)

# Concatenate all the DataFrames
finalDf = pd.concat(data, ignore_index=True)

# Drop the first column if it's blank (all NaN values)
if finalDf.iloc[:, 0].isna().all():
    finalDf.drop(finalDf.columns[0], axis=1, inplace=True)

# Drop any completely blank rows
finalDf.dropna(how='all', inplace=True)

# Save the final DataFrame to Excel without header and index
output_file = os.path.join(output_loc, "Test.xlsx")
finalDf.to_excel(output_file, header=False, index=False)
