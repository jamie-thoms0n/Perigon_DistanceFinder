"""
Python script: journey distance calculations
The following python script uses Open Street Maps API to calculate the distance between two points. 

"""

import pandas as pd
from geopy.geocoders import Nominatim # Open Street Maps
from geopy.distance import geodesic
# Pre-Processing like removing dups, add only net news
def get_distance(city1, city2):
    geolocator = Nominatim(user_agent="distance_calculator")
    location1 = geolocator.geocode(city1)
    location2 = geolocator.geocode(city2)
    if location1 and location2:
        coord1 = (location1.latitude, location1.longitude)
        coord2 = (location2.latitude, location2.longitude)
        distance = geodesic(coord1, coord2).kilometers
        return distance
    else:
        return None
# Read both worksheets from the Excel file
excel_file = r'C:\Users\RohanKhoja\Downloads\Annual Travel Data 2023.xlsx' # Excel file location
sheet1 = pd.read_excel(excel_file, sheet_name=0)
sheet2 = pd.read_excel(excel_file, sheet_name=1)
# Process each worksheet
for sheet_name, df in [("Sheet1", sheet1), ("Sheet2", sheet2)]:
    df['Distance'] = df.apply(lambda row: get_distance(row['Starting'], row['Destination']), axis=1) 
    # Save the updated dataframe back to Excel
    with pd.ExcelWriter(excel_file, mode='a', if_sheet_exists='replace',engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
print("Distance calculation completed for both worksheets.")
# Lookup against the original data - add distance
