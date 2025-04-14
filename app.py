import pandas as pd
from geopy.distance import geodesic
from geopy.geocoders import Nominatim
import streamlit as st
import folium
from streamlit_folium import st_folium
import folium.plugins as plugins

# Streamlit setup
st.set_page_config(page_title="Closest Centres Map", layout="wide")
st.title("📍 Find 8 Closest Centres")

input_address = st.text_input("Enter an address:")

if input_address:
    try:
        # Geocode the input address with a longer timeout
        geolocator = Nominatim(user_agent="centre_map_app", timeout=10)  # Timeout set to 10 seconds
        location = geolocator.geocode(input_address)

        if location is None:
            st.error("❌ Address not found. Try again.")
        else:
            input_coords = (location.latitude, location.longitude)

            # Load and process data
            file_path = "Database IC.xlsx"  # Ensure you have this file in your repo or use a URL
            sheets = ["Comps", "Active Centre", "Centre Opened"]
            all_data = []

            with st.spinner("Loading and processing data..."):
                for sheet in sheets:
                    df = pd.read_excel(file_path, sheet_name=sheet, engine="openpyxl")
                    df["Source Sheet"] = sheet
                    all_data.append(df)

                data = pd.concat(all_data)
                data = data.dropna(subset=["Latitude", "Longitude"])

                # Calculate distances
                data["Distance (miles)"] = data.apply(
                    lambda row: geodesic(input_coords, (row["Latitude"], row["Longitude"])).miles, axis=1
                )

                # Find 8 closest
                closest = data.nsmallest(8, "Distance (miles)")

            # Create map centered on the input address
            m = folium.Map(location=input_coords, zoom_start=12)

            # Add marker for input address
            folium.Marker(
                location=input_coords,
                popup=f"Your Address: {input_address}",
                icon=folium.Icon(color="green")
            ).add_to(m)

            # Prepare text data to display the distances below the map
            distance_text = f"Your Address: {input_address} - Coordinates: {input_coords[0]}, {input_coords[1]}\n"
            distance_text += "\nClosest Centres (Distances in miles):\n"

            # Draw lines and add markers for the closest centres
            for _, row in closest.iterrows():
                dest_coords = (row["Latitude"], row["Longitude"])

                # Draw a line from input address to the closest centre
                folium.PolyLine([input_coords, dest_coords], color="blue", weight=2.5, opacity=1).add_to(m)

                # Add marker for the closest centre
                folium.Marker(
                    location=dest_coords,
                    popup=f"Centre #{row['Centre Number']}<br>Address: {row['Addresses']}<br>Type: {row['Type of Centre']}<br>Transaction Milestone: {row['Transaction Milestone Status']}<br>Distance: {row['Distance (miles)']:.2f} miles",
                    icon=folium.Icon(color="blue")
                ).add_to(m)

                # Add distance to text output
                distance_text += f"Centre #{row['Centre Number']} - {row['Addresses']} - Type: {row['Type of Centre']} - Milestone: {row['Transaction Milestone Status']} - {row['Distance (miles)']:.2f} miles\n"

            # Display the map with the lines and markers
            st_folium(m, width=800, height=600)

            # Display the distances as text below the map
            st.subheader("Distances from Your Address to the Closest Centres:")
            st.text(distance_text)

    except Exception as e:
        st.error(f"Error: {e}")
