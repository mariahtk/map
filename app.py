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
            m = folium.Map(location=input_coords, zoom_start=12)

            # Add marker for input address
            folium.Marker(
                location=input_coords,
                popup=f"Your Address: {input_address}",
                icon=folium.Icon(color="green")
            ).add_to(m)

            # Prepare text output
            distance_text = f"Your Address: {input_address} - Coordinates: {input_coords[0]}, {input_coords[1]}\n"
            distance_text += "\nClosest Centres (Distances in miles):\n"

            # Used to stagger floating boxes to avoid overlap
            stagger_offsets = [-0.002, 0.002, -0.0015, 0.0015, -0.001, 0.001, -0.0005, 0.0005]

            for i, (_, row) in enumerate(closest.iterrows()):
                dest_coords = (row["Latitude"], row["Longitude"])

                # Draw line to centre
                folium.PolyLine([input_coords, dest_coords], color="blue", weight=2.5, opacity=1).add_to(m)

                # Add standard blue marker
                folium.Marker(
                    location=dest_coords,
                    popup=f"Centre #{row['Centre Number']}<br>Address: {row['Addresses']}<br>Format: {row['Format - Type of Centre']}<br>Transaction Milestone: {row['Transaction Milestone Status']}<br>Distance: {row['Distance (miles)']:.2f} miles",
                    icon=folium.Icon(color="blue")
                ).add_to(m)

                # Add floating white box with black text (clean design)
                label_text = f"{row['Addresses']} ({row['Distance (miles)']:.2f} mi)"
                offset_lat = stagger_offsets[i % len(stagger_offsets)]  # stagger vertically
                folium.Marker(
                    location=(row["Latitude"] + offset_lat, row["Longitude"]),
                    icon=folium.DivIcon(
                        icon_size=(300, 40),
                        icon_anchor=(0, 0),
                        html=f"""
                            <div style="
                                background-color: white;
                                color: black;
                                padding: 6px 10px;
                                border: 1px solid #333;
                                border-radius: 6px;
                                font-size: 13px;
                                font-family: Arial, sans-serif;
                                box-shadow: 2px 2px 5px rgba(0,0,0,0.2);
                                white-space: nowrap;
                            ">
                                {label_text}
                            </div>
                        """
                    )
                ).add_to(m)

                # Add to list text
                distance_text += f"{label_text}\n"

            # Display the map
            st_folium(m, width=900, height=600)

            # Display distances as text
            st.subheader("Distances from Your Address to the Closest Centres:")
            st.text(distance_text)

    except Exception as e:
        st.error(f"Error: {e}")

