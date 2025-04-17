import pandas as pd
from geopy.distance import geodesic
import streamlit as st
import folium
from streamlit_folium import st_folium
import folium.plugins as plugins
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt
import requests
import urllib.parse
import streamlit as st

# --- LOGIN SYSTEM ---
def login():
    st.image("IWG Logo.jpg", width=150)
    st.title("Internal Map Login")

    email = st.text_input("Email")
    password = st.text_input("Password", type="password")

    if st.button("Login"):
        if password == "IWG123" and email.endswith("@iwgplc.com"):
            st.session_state["authenticated"] = True
            st.session_state["user_email"] = email
            st.success("Login successful!")
            st.experimental_rerun()
        else:
            st.error("Invalid email or password.")

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    login()
    st.stop()

# --- REST OF THE APP ---

# Streamlit setup
st.set_page_config(page_title="Closest Centres Map", layout="wide")
st.title("📍 Find 8 Closest Centres")

# Your API key
api_key = "edd4cb8a639240daa178b4c6321a60e6"

input_address = st.text_input("Enter an address:")

if input_address:
    try:
        # URL encode the input address to handle special characters and spaces
        encoded_address = urllib.parse.quote(input_address)

        # Use OpenCage API to get location data
        url = f"https://api.opencagedata.com/geocode/v1/json?q={encoded_address}&key={api_key}"
        response = requests.get(url)
        data = response.json()

        # Check the status code and results
        if response.status_code != 200:
            st.error(f"❌ API Error: {response.status_code}. Try again.")
        elif not data.get('results'):
            st.error("❌ Address not found. Try again.")
        else:
            # Geocode the input address
            location = data['results'][0]
            input_coords = (location['geometry']['lat'], location['geometry']['lng'])

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
            for i, (index, row) in enumerate(closest.iterrows()):
                dest_coords = (row["Latitude"], row["Longitude"])

                # Draw a line from input address to the closest centre
                folium.PolyLine([input_coords, dest_coords], color="blue", weight=2.5, opacity=1).add_to(m)

                # Add marker for the closest centre
                folium.Marker(
                    location=dest_coords,
                    popup=f"Centre #{int(row['Centre Number'])}<br>Address: {row['Addresses']}<br>Format: {row['Format - Type of Centre']}<br>Transaction Milestone: {row['Transaction Milestone Status']}<br>Distance: {row['Distance (miles)']:.2f} miles",
                    icon=folium.Icon(color="blue")
                ).add_to(m)

                # Add distance to text output
                distance_text += f"Centre #{int(row['Centre Number'])} - {row['Addresses']} - Format: {row['Format - Type of Centre']} - Milestone: {row['Transaction Milestone Status']} - {row['Distance (miles)']:.2f} miles\n"

                # Floating label box that appears automatically
                label_text = f"#{int(row['Centre Number'])} - {row['Addresses']} ({row['Distance (miles)']:.2f} mi)"
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
                        icon_size=(150, 40),  # Adjust size to accommodate text
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
                                display: inline-block;
                                white-space: nowrap;  /* Ensure text stays in one line */
                                text-overflow: ellipsis; /* Add ellipsis if text overflows */
                                box-shadow: 1px 1px 3px rgba(0,0,0,0.2);
                            ">
                                {label_text}
                            </div>
                        """
                    )
                ).add_to(m)

            # Display the map with the lines and markers
            folium_map_path = "closest_centres_map.html"
            m.save(folium_map_path)
            st_folium(m, width=950, height=650)

            # Display the distances as text below the map
            st.subheader("Distances from Your Address to the Closest Centres:")
            st.text(distance_text)

            # Save PowerPoint presentation
            prs = Presentation()

            # Title Slide
            slide = prs.slides.add_slide(prs.slide_layouts[0])
            title = slide.shapes.title
            subtitle = slide.placeholders[1]
            title.text = "Closest Centres Presentation"
            subtitle.text = f"Closest Centres to: {input_address}"

            # Add slide with placeholder for the map image
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            title = slide.shapes.title
            title.text = "Closest Centres Map"
            # Add the placeholder text in the slide
            slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(8), Inches(4)).text = "Insert screenshot here."

            # Add slide with table of closest centres
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            title = slide.shapes.title
            title.text = "Distances to Closest Centres"
            table = slide.shapes.add_table(rows=len(closest) + 1, cols=5, left=Inches(0.5), top=Inches(1.5), width=Inches(9), height=Inches(4.5))
            table.table.cell(0, 0).text = "Centre #"
            table.table.cell(0, 1).text = "Address"
            table.table.cell(0, 2).text = "Type of Centre"
            table.table.cell(0, 3).text = "Transaction Milestone"
            table.table.cell(0, 4).text = "Distance (miles)"

            for i, row in enumerate(closest.iterrows()):
                table.table.cell(i + 1, 0).text = str(int(row[1]['Centre Number']))
                table.table.cell(i + 1, 1).text = str(row[1]["Addresses"])
                table.table.cell(i + 1, 2).text = str(row[1]["Format - Type of Centre"])
                table.table.cell(i + 1, 3).text = str(row[1]["Transaction Milestone Status"])
                table.table.cell(i + 1, 4).text = f"{row[1]['Distance (miles)']:.2f}"

            # Save the PowerPoint presentation to a file
            presentation_file = BytesIO()
            prs.save(presentation_file)
            presentation_file.seek(0)
            st.download_button("Download Presentation", presentation_file, "closest_centres_presentation.pptx")

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")

