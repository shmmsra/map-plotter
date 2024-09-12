import os
import pandas as pd

# Create sample data for multiple countries (worksheets)
data_usa = {
    'employee_id': [101, 102, 103],
    'employee_name': ['Alice', 'Bob', 'Charlie'],
    'home_latitude': [37.7749, 34.0522, 40.7128],  # San Francisco, Los Angeles, New York
    'home_longitude': [-122.4194, -118.2437, -74.0060],
    'office_latitude': [37.7749, 40.7128, 34.0522],  # Office locations in the same cities
    'office_longitude': [-122.4194, -74.0060, -118.2437]
}

data_india = {
    'employee_id': [201, 202, 203],
    'employee_name': ['Deepak', 'Priya', 'Karthik'],
    'home_latitude': [28.7041, 19.0760, 13.0827],  # Delhi, Mumbai, Chennai
    'home_longitude': [77.1025, 72.8777, 80.2707],
    'office_latitude': [28.7041, 19.0760, 13.0827],  # Office locations in the same cities
    'office_longitude': [77.1025, 72.8777, 80.2707]
}

# Create a dictionary of dataframes representing each country's worksheet
excel_data = {
    'USA': pd.DataFrame(data_usa),
    'India': pd.DataFrame(data_india)
}

# Write the data to an Excel file with multiple sheets
data_directory = os.path.join(os.path.dirname(__file__), '../data')
if not os.path.exists(data_directory):
    os.makedirs(data_directory)

file_path = os.path.join(data_directory, 'sample_employee_data.xlsx')
with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
    for sheet_name, df in excel_data.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

file_path  # Returning the file path for download
