# -*- coding: utf-8 -*-
"""
Created on Thu Nov  5 11:31:08 2020

@author: Hagen

Turn addresses to Longitude and Latitudes
Based off https://towardsdatascience.com/pythons-geocoding-convert-a-list-of-addresses-into-a-map-f522ef513fd6
"""

#%% Setting up
import pandas as pd
import googlemaps
import folium
from datetime import datetime
#from folium.plugins import MarkerCluster
now = datetime.now()
dt_string = now.strftime("%Y.%m.%d - %H.%M")

    # Extracts the output of an array/dictionary and drill down the layers of the first result "0", looking for
    # the key "geometry", and then lat and location respectively nested within that 
def lnglat_results(x):
    lnglat = gmaps.geocode(x)[0]["geometry"]["location"]
    return lnglat

#Generating base url and key https://developers.google.com/maps/documentation/geocoding/intro
gmaps = googlemaps.Client(key = {API KEY HERE})

#Set up Pandas df
target = '{PATH TO FILE HERE}'
df = pd.read_excel(target)

#%% Apply geolocator, calling GMaps API
# Only calls API if lng/lat are empty
df["lng_lat"] = df["Location"].apply(lambda x: x if x == dict else lnglat_results(x))

#%% Sorting and splitting into lng and lat
df["Longitude"] = df['lng_lat'].apply(lambda x: x.get('lng'))
df["Latitude"] = df['lng_lat'].apply(lambda x: x.get('lat'))
# df = df.drop(columns=["lng_lat"])
# Export
with pd.ExcelWriter('{PATH TO FILE HERE}', engine="openpyxl", mode='a') as writer:
    df.to_excel(writer, sheet_name= "Updated")

#%% Generating the Map
# center to the mean of all points
m = folium.Map(location=df[["Latitude","Longitude"]].mean().to_list(), zoom_start=17)
# if the points are too close to each other, cluster them, create a cluster overlay with MarkerCluster
# marker_cluster = MarkerCluster().add_to(m)
# draw the markers and assign popup and hover texts
# add the markers the the cluster layers so that they are automatically clustered
for i,r in df.iterrows():
    location = (r["Latitude"], r["Longitude"])
    if r['Price'] == 'POI':
        folium.Marker(location=location,
                      popup=r['Price'],
                      tooltip=r['Name'],
                      icon=folium.Icon(color='red'))\
            .add_to(m)
    else:
        folium.Marker(location=location,
                      popup=r['Price'],
                      tooltip=r['Name'])\
            .add_to(m)
        
                
#   .add_to(marker_cluster) #if wanting clustered 
# display the map
# m
# save to a file
m.save(f"Office_List Map ver {dt_string}.html")
print(f"Office_List Map ver {dt_string}.html created.")
