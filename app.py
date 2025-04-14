import pandas as pd
from geopy.distance import geodesic
from geopy.geocoders import Nominatim
import streamlit as st
import folium
from streamlit_folium import st_folium
import folium.plugins as plugins
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
from pptx import Presentation
from pptx.util import Inches
import os

# Streamlit setup
st.set_page_config(page_title="Closest Centres Map", layout="wide")
st.title("📍 Find 8 Closest Centres")

input_address = st.text_input("Enter an address:")

def capture_map_screenshot(m, map_filename="map.html"):
    # Save the folium map to an HTML file
    m.save(map_filename)

    # Use selenium to take a screenshot of the saved HTML
    options = webdriver.ChromeOptions()
    options.headless = True
    options.add_argument("--window-size=1280x1024")

    # Set up selenium webdriver
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get(f"file://{os.path.abspath(map_filename)}")
    time.sleep(3)  # Allow the map to render
    screenshot_filename = "map_screenshot.png"
    driver.save_screenshot(screenshot_filename)
    driver.quit()

    return screenshot_filename

def create_powerpoint(data, screenshot_filename="map_screenshot.png", pptx_filename="Closest_Centres_Presentation.pptx"):
    # Create PowerPoint Presentation
    prs = Presentation()

    # Slide 1: Map Screenshot
    slide_1 = prs.slides.add_slide(prs.slide_layouts[5])  # 5 is a blank layout
    slide_1.shapes.add_picture(screenshot_filename, Inches(0), Inches(0), width=Inches(10), height=Inches(7.5))

    # Slide 2: Distance Data
    slide_2 = prs.slides.add_slide(prs.slide_layouts[1])  # 1 is the title and content layout
    slide_2.shapes.title.text = "Closest Centres Distance Data"

    # Prepare data to add in the second slide
    distance_data = "Your Address: " + input_address + "\n\nClosest Centres (Distances in miles):\n"
    for _, row in data.iterrows():
        distance_data += f"Centre #{row['Centre Number']} - {row['Addresses']} - Format: {row['Format - Type of Centre']} - Milestone: {row['Transaction Milestone Status']} - {row['Distance (miles)']:.2f} miles\n"
    
    # Add data to slide
    slide_2.shapes.placeholders[1].text = distance_data

    # Save the PowerPoint presentation
    prs.save(pptx_filename)

    return pptx_filename

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

                # Adjust label placement if too close to the edges of the map
                label_lat = row["Latitude"] + offset_lat
                label_lon = row["Longitude"]
                if label_lat > lat_max:
                    label_lat = lat_max - 0.0005  # Ensure it's within bounds
                if label_lat < lat_min:
                    label_lat = lat_min + 0.0005  # Ensure it's within bounds
                if label_lon > lng_max:
                    label_lon = lng_max - 0.0005  # Ensure it's within bounds
                if label_lon < lng_min:
                    label_lon = lng_min + 0.0005  # Ensure it's within bounds

                # Add the adjusted label with proper offset
                folium.Marker(
                    location=(label_lat, label_lon),
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

            # Capture the screenshot
            screenshot_filename = capture_map_screenshot(m)

            # Create the PowerPoint and allow for download
            pptx_filename = create_powerpoint(closest, screenshot_filename)

            # Provide download link for the PowerPoint
            st.download_button(
                label="Download PowerPoint Presentation",
                data=open(pptx_filename, "rb").read(),
                file_name=pptx_filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

    except Exception as e:
        st.error(f"Error: {e}")
