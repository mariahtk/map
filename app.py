import pandas as pd
from geopy.distance import geodesic
from geopy.geocoders import Nominatim
import streamlit as st
import folium
from streamlit_folium import st_folium

st.set_page_config(page_title="Closest Centres Map", layout="wide")

st.title("📍 Find 8 Closest Centres")
input_address = st.text_input("Enter an address:")

if input_address:
    try:
        # Geocode input address
        geolocator = Nominatim(user_agent="centre_map_app")
        location = geolocator.geocode(input_address)

        if location is None:
            st.error("❌ Address not found. Try again.")
        else:
            input_coords = (location.latitude, location.longitude)

            # Load all relevant sheets
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

                # Error handling for missing columns
                required_columns = ["Latitude", "Longitude", "Centre Number", "Addresses", "City", "Transaction Milestone Status"]
                if not all(col in data.columns for col in required_columns):
                    st.error(f"❌ Missing required columns in the dataset!")
                    st.stop()

                # Calculate distances
                data["Distance (miles)"] = data.apply(
                    lambda row: geodesic(input_coords, (row["Latitude"], row["Longitude"])).miles, axis=1
                )

                # Find 8 closest
                closest = data.nsmallest(8, "Distance (miles)")

            # Display results
            st.markdown(f"**Results for:** `{input_address}`")

            # Adjust map size based on user input
            map_width = 800
            map_height = 600

            m = folium.Map(location=input_coords, zoom_start=10)

            # Add input address marker
            folium.Marker(
                location=input_coords,
                popup=f"Input Address: {input_address}",
                icon=folium.Icon(color="red", icon="info-sign")
            ).add_to(m)

            # Add markers for the closest centres
            for _, row in closest.iterrows():
                dest_coords = (row["Latitude"], row["Longitude"])
                popup_text = f"""
                <b>Centre #{row['Centre Number']}</b><br>
                {row['Addresses']}<br>
                City: {row['City']}<br>
                Status: {row['Transaction Milestone Status']}<br>
                Distance: {row['Distance (miles)']:.2f} mi
                """
                folium.Marker(
                    location=dest_coords,
                    popup=folium.Popup(popup_text, max_width=300),
                    icon=folium.Icon(color="blue")
                ).add_to(m)
                folium.PolyLine(locations=[input_coords, dest_coords], color="blue").add_to(m)

            # Display the map
            st_folium(m, width=map_width, height=map_height)

    except Exception as e:
        st.error(f"Error: {e}")
