import pandas as pd
import openpyxl

def reorder_csv_to_bulk_format(csv_file_path, output_excel_path):
    # Load the CSV file
    csv_df = pd.read_csv(csv_file_path)

    # Desired order of columns (final columns to keep)
    final_order = [
        'Name',
        'inscor__Product__r.Name',
        'inscor__Condition_Code__r.Name',
        'inscor__Quantity_Available__c',
        'inscor__UOM__c',  # This column will be added with a default value
        'SSP_Updated__c',
        'inscor__Keyword__c'
    ]

    # Retain only the columns that are in the final_order
    csv_df = csv_df[[col for col in final_order if col in csv_df.columns]]

    # Add the missing 'inscor__UOM__c' column with default value "EA"
    if 'inscor__UOM__c' not in csv_df.columns:
        csv_df.loc[:, 'inscor__UOM__c'] = "EA"  # Using .loc to avoid SettingWithCopyWarning

    # Reorder columns as per the final order, including the new 'inscor__UOM__c' column
    csv_reordered = csv_df.reindex(columns=final_order)

    # Save the reordered CSV as an Excel file
    csv_reordered.to_excel(output_excel_path, index=False)

# Example usage
csv_file_path = r'F:\CAMERON\temp\aero x new on common bulk\List Coder08_19_2024-07_47_02.csv'
output_excel_path = r'F:\CAMERON\temp\aero x new on common bulk\out1.xlsx'
reorder_csv_to_bulk_format(csv_file_path, output_excel_path)
