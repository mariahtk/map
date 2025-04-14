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
            m = folium.Map(location=input_coords, zoom_start=14)  # Increased zoom level

            # Add marker for input address
            folium.Marker(
                location=input_coords,
                popup=f"Your Address: {input_address}",
                icon=folium.Icon(color="green")
            ).add_to(m)

            # Add markers and solid white text boxes for closest centres
            for i, (_, row) in enumerate(closest.iterrows()):
                dest_coords = (row["Latitude"], row["Longitude"])

                # Draw line
                folium.PolyLine([input_coords, dest_coords], color="blue", weight=2.5, opacity=1).add_to(m)

                # Marker at centre
                folium.Marker(
                    location=dest_coords,
                    popup=f"Centre #{row['Centre Number']}<br>Address: {row['Addresses']}<br>Format: {row['Format - Type of Centre']}<br>Transaction Milestone: {row['Transaction Milestone Status']}<br>Distance: {row['Distance (miles)']:.2f} miles",
                    icon=folium.Icon(color="blue")
                ).add_to(m)

                # Solid white text box that shows automatically
                label_text = f"Centre #{row['Centre Number']} - {row['Addresses']} ({row['Distance (miles)']:.2f} mi)"
                
                # Offsetting the text box slightly for better visibility
                offset_lat = 0.0008 * (i + 1)  # Slight vertical offset for each label
                offset_lon = 0.0008 * (i + 1)  # Slight horizontal offset for each label

                folium.Marker(
                    location=(dest_coords[0] + offset_lat, dest_coords[1] + offset_lon),
                    icon=folium.DivIcon(
                        icon_size=(250, 40),
                        icon_anchor=(0, 0),
                        html=f"""
                            <div style="
                                background-color: white;
                                color: black;
                                padding: 10px 15px;
                                border: 1px solid black;
                                border-radius: 6px;
                                font-size: 13px;
                                font-family: Arial, sans-serif;
                                box-shadow: 2px 2px 5px rgba(0,0,0,0.2);
                                width: auto;
                                white-space: nowrap;
                            ">
                                {label_text}
                            </div>
                        """
                    )
                ).add_to(m)

            # Show the map
            st_folium(m, width=950, height=650)

    except Exception as e:
        st.error(f"Error: {e}")
