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
        geolocator = Nominatim(user_agent="centre_map_app", timeout=10)
        location = geolocator.geocode(input_address)

        if location is None:
            st.error("❌ Address not found. Try again.")
        else:
            input_coords = (location.latitude, location.longitude)

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

                data["Distance (miles)"] = data.apply(
                    lambda row: geodesic(input_coords, (row["Latitude"], row["Longitude"])).miles, axis=1
                )

                closest = data.nsmallest(8, "Distance (miles)")

            m = folium.Map(location=input_coords, zoom_start=12)

            folium.Marker(
                location=input_coords,
                popup=f"Your Address: {input_address}",
                icon=folium.Icon(color="green")
            ).add_to(m)

            distance_text = f"Your Address: {input_address} - Coordinates: {input_coords[0]}, {input_coords[1]}\n"
            distance_text += "\nClosest Centres (Distances in miles):\n"

            for _, row in closest.iterrows():
                dest_coords = (row["Latitude"], row["Longitude"])

                # Draw line
                folium.PolyLine([input_coords, dest_coords], color="blue", weight=2.5, opacity=1).add_to(m)

                # Add marker for centre with popup
                popup_text = (
                    f"Centre #{row['Centre Number']}<br>"
                    f"Address: {row['Addresses']}<br>"
                    f"Format: {row['Format - Type of Centre']}<br>"
                    f"Transaction Milestone: {row['Transaction Milestone Status']}<br>"
                    f"Distance: {row['Distance (miles)']:.2f} miles"
                )
                folium.Marker(
                    location=dest_coords,
                    popup=popup_text,
                    icon=folium.Icon(color="blue")
                ).add_to(m)

                # Add floating text box above the line (at midpoint)
                midpoint_lat = (input_coords[0] + dest_coords[0]) / 2
                midpoint_lon = (input_coords[1] + dest_coords[1]) / 2
                floating_label = (
                    f"{row['Addresses']}<br>{row['Distance (miles)']:.2f} miles"
                )
                folium.map.Marker(
                    [midpoint_lat, midpoint_lon],
                    icon=folium.DivIcon(
                        html=f"""
                            <div style="
                                font-size: 11px;
                                background-color: white;
                                padding: 4px;
                                border: 1px solid gray;
                                border-radius: 4px;
                                box-shadow: 2px 2px 5px rgba(0,0,0,0.2);
                                max-width: 250px;
                            ">
                                {floating_label}
                            </div>
                        """
                    )
                ).add_to(m)

                # Text output
                distance_text += f"Centre #{row['Centre Number']} - {row['Addresses']} - Format: {row['Format - Type of Centre']} - Milestone: {row['Transaction Milestone Status']} - {row['Distance (miles)']:.2f} miles\n"

            # Display results
            st_folium(m, width=800, height=600)
            st.subheader("Distances from Your Address to the Closest Centres:")
            st.text(distance_text)

    except Exception as e:
        st.error(f"Error: {e}")
