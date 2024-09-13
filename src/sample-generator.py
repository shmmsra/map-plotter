import os
import pandas as pd
import numpy as np

# Function to generate random latitude and longitude for offices and employees
def generate_random_lat_lon(num_entries, lat_range, lon_range):
    lats = np.random.uniform(lat_range[0], lat_range[1], num_entries)
    lons = np.random.uniform(lon_range[0], lon_range[1], num_entries)
    return lats, lons

# Generate data for a single country
def generate_country_data(country_name, num_offices, num_employees_per_office, lat_range, lon_range):
    # Generate office locations
    office_lats, office_lons = generate_random_lat_lon(num_offices, lat_range, lon_range)
    
    employee_data = {
        'employee_id': [],
        'employee_name': [],
        'home_latitude': [],
        'home_longitude': [],
        'office_latitude': [],
        'office_longitude': []
    }
    
    employee_id = 1
    # Generate employees for each office
    for office_lat, office_lon in zip(office_lats, office_lons):
        for _ in range(num_employees_per_office):
            employee_data['employee_id'].append(employee_id)
            employee_data['employee_name'].append(f'Employee_{employee_id}')
            # Generate employee home location near the office (within 0.5 degree radius)
            home_lat = np.random.uniform(office_lat - 0.5, office_lat + 0.5)
            home_lon = np.random.uniform(office_lon - 0.5, office_lon + 0.5)
            employee_data['home_latitude'].append(home_lat)
            employee_data['home_longitude'].append(home_lon)
            employee_data['office_latitude'].append(office_lat)
            employee_data['office_longitude'].append(office_lon)
            employee_id += 1
            
    return pd.DataFrame(employee_data)

# Generate data for multiple countries
def generate_sample_data():
    # Define data ranges for different countries
    countries = {
        'USA': {'lat_range': [25, 49], 'lon_range': [-125, -66], 'num_offices': 3},
        'India': {'lat_range': [8, 37], 'lon_range': [68, 97], 'num_offices': 4}
    }
    
    # Create a dictionary of dataframes for each country
    excel_data = {}
    
    for country, info in countries.items():
        num_employees_per_office = np.random.randint(5, 10)  # 100 to 150 employees per office
        excel_data[country] = generate_country_data(country, info['num_offices'], num_employees_per_office, 
                                                    info['lat_range'], info['lon_range'])
    
    # Write the data to an Excel file with multiple sheets
    data_directory = os.path.join(os.path.dirname(__file__), '../data')
    if not os.path.exists(data_directory):
        os.makedirs(data_directory)

    file_path = os.path.join(data_directory, 'sample_employee_data.xlsx')
    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        for sheet_name, df in excel_data.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    return file_path

# Generate the sample data
generate_sample_data()
