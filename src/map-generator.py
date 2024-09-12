import os
import pandas as pd
import googlemaps
import matplotlib.pyplot as plt
from matplotlib.patches import Circle
import numpy as np

# Function to read the Excel file with multiple worksheets
def read_employee_data(file_path):
    return pd.read_excel(file_path, sheet_name=None)  # Reading all sheets into a dictionary

# Function to find the nearest office for each employee using Google Maps Distance Matrix API
def find_nearest_office(gmaps, employee_lat, employee_lon, office_locations):
    min_distance = float('inf')
    nearest_office = None

    for office in office_locations:
        office_lat, office_lon = office
        # Calculate distance using Google Maps API
        distance_result = gmaps.distance_matrix(origins=[(employee_lat, employee_lon)],
                                                destinations=[(office_lat, office_lon)],
                                                mode="driving")  # Driving distance

        distance = distance_result['rows'][0]['elements'][0]['distance']['value']  # Distance in meters

        if distance < min_distance:
            min_distance = distance
            nearest_office = (office_lat, office_lon)

    return nearest_office, min_distance

# Function to plot employees and offices on a graph
def plot_employee_office_map(employee_data, office_locations, country):
    fig, ax = plt.subplots()

    # Plot offices as colored translucent circles
    colors = plt.cm.rainbow(np.linspace(0, 1, len(office_locations)))

    for office, color in zip(office_locations, colors):
        office_lat, office_lon = office
        circle = Circle((office_lon, office_lat), 0.1, color=color, alpha=0.5)
        ax.add_patch(circle)

    # Plot employee home locations
    for _, row in employee_data.iterrows():
        employee_lat = row['home_latitude']
        employee_lon = row['home_longitude']
        ax.plot(employee_lon, employee_lat, 'bo')  # Plot as blue dots

    ax.set_title(f"Employee and Office Locations in {country}")
    plt.xlabel("Longitude")
    plt.ylabel("Latitude")
    plt.grid(True)
    plt.show()

# Main function to process employee data and plot maps
def process_employee_data(file_path, gmaps_api_key):
    # Initialize Google Maps client
    gmaps = googlemaps.Client(key=gmaps_api_key)

    # Read data from Excel file
    employee_sheets = read_employee_data(file_path)

    for country, employee_data in employee_sheets.items():
        # Extract office locations from the employee data
        office_locations = list(set(zip(employee_data['office_latitude'], employee_data['office_longitude'])))
    
        # For each employee, find the nearest office
        employee_data['nearest_office'] = employee_data.apply(
            lambda row: find_nearest_office(gmaps, row['home_latitude'], row['home_longitude'], office_locations),
            axis=1
        )

        # Plot the data
        plot_employee_office_map(employee_data, office_locations, country)


# Example usage:
data_directory = os.path.join(os.path.dirname(__file__), '../data')
if os.path.exists(data_directory):
    file_path = os.path.join(data_directory, 'sample_employee_data.xlsx')
    gmaps_api_key = 'AIzaSyBVQjWfN67G9MwBFmk3BeKHMb5Bh1CCwOw'
    process_employee_data(file_path, gmaps_api_key)
