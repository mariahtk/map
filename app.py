import pandas as pd
from geopy.distance import geodesic
from geopy.geocoders import Nominatim
import streamlit as st
import folium
from streamlit_folium import st_folium

# Streamlit setup
st.set_page_config(page_title="Closest Centres Map", layout="wide")
st.title("📍 Find 8 Closest Centres")

input_address = st.text_input("Enter an address:")

if input_address:
    try:
        geolocator = Nominatim(user_agent="centre_map_app", timeout=10)
        location = geolocator.geocode(input_address)

        if location is None:
            st.error("❌ Address not found. Try again.")
        else:
            input_coords = (location.latitude, location.longitude)

            # Load and process data
            file_path = "Database IC.xlsx"
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

                closest = data.nsmallest(8, "Distance (miles)")

            # Create map centered on input address
            m = folium.Map(location=input_coords, zoom_start=14)  # Starting zoom level

            # Add marker for input address
            folium.Marker(
                location=input_coords,
                popup=f"Your Address: {input_address}",
                icon=folium.Icon(color="green")
            ).add_to(m)

            # Coordinates list for all centers
            center_coords = [input_coords]  # Add input address as first point

            # Add markers with text box that appears automatically
            for i, (_, row) in enumerate(closest.iterrows()):
                dest_coords = (row["Latitude"], row["Longitude"])
                center_coords.append(dest_coords)

                # Draw line
                folium.PolyLine([input_coords, dest_coords], color="blue", weight=2.5, opacity=1).add_to(m)

                # Create a label as popup content
                label_text = f"""
                    <div style="background-color: white; color: black; padding: 10px 15px; border: 1px solid black; border-radius: 6px; font-size: 13px; font-family: Arial, sans-serif; box-shadow: 2px 2px 5px rgba(0,0,0,0.2);">
                        <strong>Centre #{row['Centre Number']}</strong><br>
                        <strong>Address:</strong> {row['Addresses']}<br>
                        <strong>Distance:</strong> {row['Distance (miles)']:.2f} miles
                    </div>
                """

                # Popup that will show the label text automatically
                folium.Marker(
                    location=dest_coords,
                    popup=folium.Popup(label_text, max_width=300),
                    icon=folium.Icon(color="blue")
                ).add_to(m)

            # Calculate bounds (min/max latitudes and longitudes)
            latitudes = [coord[0] for coord in center_coords]
            longitudes = [coord[1] for coord in center_coords]

            # Get bounding box
            min_lat, max_lat = min(latitudes), max(latitudes)
            min_lon, max_lon = min(longitudes), max(longitudes)

            # Adjust the zoom dynamically based on the bounding box
            m.fit_bounds([[min_lat, min_lon], [max_lat, max_lon]])

            # Show the map
            st_folium(m, width=950, height=650)

    except Exception as e:
        st.error(f"Error: {e}")
