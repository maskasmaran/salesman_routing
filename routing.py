import googlemaps
import folium
import numpy as np
import pandas as pd
from datetime import datetime
import colorsys
import random
import os
import sys

# Your Google Maps API key
api_key = 'Google API'

# Create a Google Maps client
gmaps = googlemaps.Client(key=api_key)

# Load locations from Excel file
excel_file_path = 'file.xlsx'
df = pd.read_excel(excel_file_path, sheet_name='Cluster_5')

# Convert dataframe to dictionary
locations = {row['NAMA OUTLET']: (float(row['Latitude']), float(row['Longitude'])) for _, row in df.iterrows()}

def generate_colors(num_colors):
    colors = []
    for i in range(num_colors):
        hue = i / num_colors
        rgb = colorsys.hsv_to_rgb(hue, 1.0, 1.0)
        colors.append(tuple(int(c * 255) for c in rgb))
    return ['#%02x%02x%02x' % color for color in colors]

def nearest_neighbor_algorithm(starting_point, unvisited, max_outlets):
    current_point = starting_point
    route = [starting_point]
    while unvisited and len(route) <= max_outlets:
        distances_to_unvisited = {loc: np.linalg.norm(np.array(locations[current_point]) - np.array(locations[loc]))
                                  for loc in unvisited}
        nearest = min(distances_to_unvisited, key=distances_to_unvisited.get)
        route.append(nearest)
        current_point = nearest
        unvisited.remove(nearest)
    route.append(starting_point)  # Ensure the route goes back to the starting point
    return route

def distribute_outlets(routes, max_outlets):
    outlets_per_route = max_outlets // len(routes)
    remaining_outlets = max_outlets % len(routes)
    distributed_outlets = []

    for i, route in enumerate(routes):
        start = i * outlets_per_route + min(i, remaining_outlets)
        end = start + outlets_per_route + (1 if i < remaining_outlets else 0)
        distributed_outlets.extend(route[start:end])

    return distributed_outlets[:max_outlets]  # Ensure the total number of outlets does not exceed max_outlets

# Function to get directions between two locations using Google Maps API
def get_distance_and_directions(origin, destination):
    try:
        # Make the directions request
        directions_result = gmaps.directions(origin, destination, mode="driving", departure_time=datetime.now())
        if directions_result:
            steps = directions_result[0]['legs'][0]['steps']
            distances = [step['distance']['value'] for step in steps]
            return distances, steps
    except googlemaps.exceptions.ApiError as e:
        print(f"Error getting directions for: {origin} -> {destination}")
        print(e)
    return [], []

def get_location(locations, location_name):
    # Case-insensitive lookup without stripping
    for key, value in locations.items():
        if key.lower() == location_name.lower():
            return value
    return None

# Function to calculate total distance and travel time of a route using Google Maps API
def calculate_distance_and_time(route, locations):
    total_distance = 0
    total_time_seconds = 0

    for i in range(len(route) - 1):
        origin, destination = get_location(locations, route[i]), get_location(locations, route[i + 1])
        if origin is not None and destination is not None:
            distances, steps = get_distance_and_directions(origin, destination)
            total_distance += sum(distances)

            # Calculate total travel time in seconds
            total_time_seconds += sum(step['duration']['value'] for step in steps)

        else:
            print(f"Location not found in dictionary: {route[i]} or {route[i + 1]}")

    # Convert total travel time to hours and minutes
    total_time_hours, remainder = divmod(total_time_seconds, 3600)
    total_time_minutes = remainder // 60

    return total_distance, total_time_hours, total_time_minutes

# Function to plot Folium Map and export to HTML
def plot_folium_map(route, locations, directory_path, route_num):
    if not locations:
        print(f"Warning: Empty locations dictionary. Skipping route {route_num}.")
        return

    map_center = np.mean(list(locations.values()), axis=0)

    folium_map = folium.Map(location=map_center, zoom_start=13, tiles='https://mt1.google.com/vt/lyrs=m&x={x}&y={y}&z={z}',
                            attr='Google')

    route_coordinates = [locations[loc] for loc in route]

    # Get distances, directions, and travel time for the route using Google Maps API
    total_distance, total_time_hours, total_time_minutes = calculate_distance_and_time(route, locations)

    directions = []
    for j in range(len(route) - 1):
        start_location = get_location(locations, route[j])
        end_location = get_location(locations, route[j + 1])
        _, steps = get_distance_and_directions(start_location, end_location)
        directions += steps

    # Extract coordinates and distances from directions
    coordinates = [(step['end_location']['lat'], step['end_location']['lng']) for step in directions]

    try:
        # Plot the route with random color
        route_color = "#{:06x}".format(random.randint(0, 0xFFFFFF))
        folium.PolyLine(coordinates, color=route_color, weight=5, opacity=1, tooltip=f"Route {route_num}").add_to(folium_map)
    except Exception as e:
        print(f"Error plotting route {filename}: {e}")
        print(f"Locations: {locations}")
        print(f"Route: {route}")
        print(f"Coordinates: {coordinates}")
        raise

    # Add markers for each location with actual distances, travel time, and route numbers
    for j in range(len(route) - 1):
        loc1, loc2 = route[j], route[j + 1]
        coord1, coord2 = locations[loc1], locations[loc2]
        distance = np.linalg.norm(np.array(coord1) - np.array(coord2))
        popup_text = f"{loc1} to {loc2}: {distance:.2f} km\nTravel Time: {total_time_hours:.0f} hours {total_time_minutes:.0f} minutes"
        
        # Add outlet number next to the pin (bold, black, larger, and thicker)
        folium.Marker(location=coord1, icon=folium.DivIcon(html=f"<div style='font-size: 14px; color: black; font-weight: bold; line-height: 1.5;'>{j + 1}</div>")).add_to(folium_map)

        folium.Marker(coord1, popup=popup_text, icon=folium.Icon(color=route_color, icon_color=route_color, prefix='fa', icon='circle')).add_to(folium_map)

    # Extract distances of each outlet in the route
    outlet_distances_and_time = []
    for j in range(len(route) - 1):
        loc1, loc2 = route[j], route[j + 1]
        distances, steps = get_distance_and_directions(locations[loc1], locations[loc2])
        distance = sum(distances)

        # Slow down travel time (increase duration for each step)
        for step in steps:
            step['duration']['value'] *= 2  # Adjust the factor as needed to slow down the travel time

        travel_time_seconds = sum(step['duration']['value'] for step in steps)

        # Convert travel time to hours and minutes
        travel_time_hours, remainder = divmod(travel_time_seconds, 3600)
        travel_time_minutes = remainder // 60

        outlet_distances_and_time.append((f"{loc1} to {loc2}", distance, travel_time_hours, travel_time_minutes))

    # Save the map
    try:
        folium_map.save(os.path.join(directory_path, f"{filename}.html"))
        print(f"Folium Map for Route saved as '{filename}.html'")
    except Exception as e:
        print(f"Error saving Folium Map for route {filename}: {e}")

    # Export distances and travel time to Excel
    try:
        with pd.ExcelWriter(os.path.join(directory_path, f"{filename}.xlsx")) as writer:
            df_distances = pd.DataFrame(outlet_distances_and_time, columns=['Outlet Pair', 'Distance (m)', 'Travel Time Hours', 'Travel Time Minutes'])
            df_distances.to_excel(writer, sheet_name='Distances', index=False)

        print(f"Distances and Travel Time data for Route exported to '{filename}.xlsx'")
    except Exception as e:
        print(f"Error exporting distances and travel time to Excel for route {filename}: {e}")

# Apply Nearest Neighbor Algorithm for each driver
if len(sys.argv) > 1:
    starting_point = sys.argv[1]
    max_outlets_per_route = int(sys.argv[2])
    total_routes = int(sys.argv[3])
else:
    # Function to get user input for starting point, max outlets per route, and total routes
    def get_user_input():
        starting_point = input("Input DEPO: ")
        max_outlets_per_route = int(input("Input max outlets per route: "))
        total_routes = int(input("Input jumlah salesman: "))
        
        return starting_point, max_outlets_per_route, total_routes

    starting_point, max_outlets_per_route, total_routes = get_user_input()

routes = []

unvisited_all = list(locations.keys())

for route_num in range(total_routes):
    unvisited = unvisited_all.copy()

    # Check if starting_point is in the unvisited list before removing
    if starting_point in unvisited:
        unvisited.remove(starting_point)
    else:
        print(f"Warning: Starting point '{starting_point}' not in the unvisited list.")

    current_starting_point = starting_point

    route = []
    while unvisited and len(route) < max_outlets_per_route:
        current_unvisited = list(set(unvisited) - set(route))
        current_route = nearest_neighbor_algorithm(current_starting_point, current_unvisited, max_outlets_per_route)
        route += current_route[:-1]
        current_starting_point = current_route[-2]

    route.append(starting_point)  # Ensure the route goes back to the starting point
    routes.append(route)

    # Update the list of unvisited outlets
    unvisited_all = list(set(unvisited_all) - set(route))

    # Create unique filenames, route color, and route numbers for each route
    filename = f"filename{route_num + 1}"
    route_color = generate_colors(1)[0]  # Generate a unique color for each route
    directory_path = r'filepath'

    # Plot and save Folium Map and Excel file for each route
    plot_folium_map(route, locations, directory_path, route_num + 1)
