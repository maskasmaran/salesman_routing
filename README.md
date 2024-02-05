Route Optimization Readme
This repository contains Python code for optimizing routes using the Nearest Neighbor Algorithm. The script utilizes the Google Maps API for obtaining distances, directions, and travel times between locations. The resulting routes are visualized on interactive Folium maps and exported to Excel files for further analysis.

Prerequisites
Python 3.x
pandas
numpy
folium
googlemaps
Install the required libraries using the following command:

bash
Copy code
pip install pandas numpy folium googlemaps
Usage
Clone the Repository:

bash
Copy code
git clone https://github.com/your-username/your-repository.git
cd your-repository
Install Dependencies:
Make sure to install the necessary Python libraries using the provided requirements.txt file.

bash
Copy code
pip install -r requirements.txt
Set Google Maps API Key:
Replace the api_key variable in the script with your Google Maps API key.

Prepare Data:
Replace the ALL_NEW_CLUSTER_BANJARBARU.xlsx file in the repository with your own Excel file containing clustered outlet data.

Run the Script:
Execute the route_optimization.py script to optimize routes based on user input.

bash
Copy code
python route_optimization.py
User Input:

Input the starting point (DEPO).
Input the maximum number of outlets per route.
Input the total number of routes (salesmen).
Output:

The script will apply the Nearest Neighbor Algorithm for each driver, optimizing routes.
Interactive Folium maps for each route will be saved in the specified output directory.
Excel files containing distances and travel time data for each route will also be exported.
Directory Structure
route_optimization.py: Main Python script for route optimization.
ALL_NEW_CLUSTER_BANJARBARU.xlsx: Sample Excel file containing clustered outlet data (replace with your own data).
requirements.txt: List of required Python libraries.
C:\Users\fatah\OneDrive\文件\PMA\KOORDINAT\RUTE BANJARBARU: Default output directory for maps and Excel files.
Feel free to customize the script and adapt it to your specific use case.

Note: Ensure that you have the required permissions to access and modify the specified output directory.

For any issues or questions, please open an issue.

Optimize those routes efficiently!
