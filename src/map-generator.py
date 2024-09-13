import os
import pandas as pd
import googlemaps
import folium

data_directory = os.path.join(os.path.dirname(__file__), '../data')

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

        element = distance_result['rows'][0]['elements'][0]
        if element['status'] != 'OK':
            continue
        else:
            print(f"Distance: {element['distance']['text']}, Duration: {element['duration']['text']}")

        distance = distance_result['rows'][0]['elements'][0]['distance']['value']  # Distance in meters

        if distance < min_distance:
            min_distance = distance
            nearest_office = (office_lat, office_lon)

    return nearest_office, min_distance

# Function to plot employees and offices on a Google Maps terrain view
def plot_on_google_map(employee_data, office_locations, country):
    # Create a folium map centered on the country's first office location
    first_office_lat, first_office_lon = office_locations[0]
    map_ = folium.Map(
        location=[first_office_lat, first_office_lon],
        zoom_start=5
    )

    # Add office locations with colored translucent circles
    for office in office_locations:
        office_lat, office_lon = office
        folium.Circle(
            location=[office_lat, office_lon],
            radius=50000,  # Radius in meters
            color="blue",
            fill=True,
            fill_opacity=0.4
        ).add_to(map_)

    # Add employee home locations
    for _, row in employee_data.iterrows():
        employee_lat = row['home_latitude']
        employee_lon = row['home_longitude']
        folium.Marker(
            location=[employee_lat, employee_lon],
            popup=f"Employee: {row['employee_name']}",
            icon=folium.Icon(color="red", icon="home")
        ).add_to(map_)

    # Save map as an HTML file
    map_file_path = os.path.join(data_directory, f"{country}_employee_office_map.html")
    map_.save(map_file_path)
    print(f"Map saved as {map_file_path}")

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

        # Plot the data on a Google Map terrain view
        plot_on_google_map(employee_data, office_locations, country)

# Example usage:
if os.path.exists(data_directory):
    file_path = os.path.join(data_directory, 'sample_employee_data.xlsx')
    gmaps_api_key = 'GOOGLE_MAPS_API_KEY'
    process_employee_data(file_path, gmaps_api_key)
