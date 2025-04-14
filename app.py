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

            # Calculate the bounding box to fit all markers and determine dynamic zoom level
            lats = [input_coords[0]] + closest["Latitude"].tolist()
            lngs = [input_coords[1]] + closest["Longitude"].tolist()
            lat_min, lat_max = min(lats), max(lats)
            lng_min, lng_max = min(lngs), max(lngs)

            # Dynamic zoom based on bounding box
            lat_diff = lat_max - lat_min
            lng_diff = lng_max - lng_min
            max_diff = max(lat_diff, lng_diff)

            # Adjust zoom level for tighter view on the markers
            zoom_level = 14 - (max_diff * 3)  # Increase zoom level for a closer view
            zoom_level = max(zoom_level, 14)  # Minimum zoom level

            # Create map centered on the input address
            m = folium.Map(location=input_coords, zoom_start=int(zoom_level))

            # Add marker for input address
            folium.Marker(
                location=input_coords,
                popup=f"Your Address: {input_address}",
                icon=folium.Icon(color="green")
            ).add_to(m)

            # Prepare text data to display the distances below the map
            distance_text = f"Your Address: {input_address} - Coordinates: {input_coords[0]}, {input_coords[1]}\n"
            distance_text += "\nClosest Centres (Distances in miles):\n"

            # For staggering the labels vertically
            stagger_offsets = [-0.002, 0.002, -0.0015, 0.0015, -0.001, 0.001, -0.0005, 0.0005]

            # Draw lines and add markers for the closest centres
            for i, (_, row) in enumerate(closest.iterrows()):
                dest_coords = (row["Latitude"], row["Longitude"])

                # Draw a line from input address to the closest centre
                folium.PolyLine([input_coords, dest_coords], color="blue", weight=2.5, opacity=1).add_to(m)

                # Add marker for the closest centre
                folium.Marker(
                    location=dest_coords,
                    popup=f"Centre #{row['Centre Number']}<br>Address: {row['Addresses']}<br>Format: {row['Format - Type of Centre']}<br>Transaction Milestone: {row['Transaction Milestone Status']}<br>Distance: {row['Distance (miles)']:.2f} miles",
                    icon=folium.Icon(color="blue")
                ).add_to(m)

                # Add distance to text output
                distance_text += f"Centre #{row['Centre Number']} - {row['Addresses']} - Format: {row['Format - Type of Centre']} - Milestone: {row['Transaction Milestone Status']} - {row['Distance (miles)']:.2f} miles\n"

                # Floating label box that appears automatically
                label_text = f"#{row['Centre Number']} - {row['Addresses']} ({row['Distance (miles)']:.2f} mi)"
                offset_lat = stagger_offsets[i % len(stagger_offsets)]

                folium.Marker(
                    location=(row["Latitude"] + offset_lat, row["Longitude"]),
                    icon=folium.DivIcon(
                        icon_size=(250, 40),
                        icon_anchor=(0, 0),
                        html=f"""
                            <div style="
                                background-color: white;
                                color: black;
                                padding: 6px 10px;
                                border: 1px solid black;
                                border-radius: 6px;
                                font-size: 13px;
                                font-family: Arial, sans-serif;
                                white-space: nowrap;
                                box-shadow: 1px 1px 3px rgba(0,0,0,0.2);
                            ">
                                {label_text}
                            </div>
                        """
                    )
                ).add_to(m)

            # Display the map with the lines and markers
            st_folium(m, width=950, height=650)

            # Display the distances as text below the map
            st.subheader("Distances from Your Address to the Closest Centres:")
            st.text(distance_text)

    except Exception as e:
        st.error(f"Error: {e}")
