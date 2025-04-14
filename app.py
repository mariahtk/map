import pandas as pd
from geopy.distance import geodesic
from geopy.geocoders import Nominatim
import streamlit as st
import folium
from streamlit_folium import st_folium
import folium.plugins as plugins
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt

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

            # Adjusted table dimensions to fit within the slide
            rows = len(closest) + 1  # Include header row
            cols = 5  # Centre Number, Address, Format, Milestone, Distance

            table_width = Inches(9)  # Set table width to 9 inches (leaving margins)
            table_height = Inches(5)  # Set table height to 5 inches

            # Set the table position (left, top, width, height)
            left = Inches(0.5)  # Center the table with some margin
            top = Inches(1)  # Starting from the top of the slide

            table = slide.shapes.add_table(rows, cols, left, top, table_width, table_height).table

            # Set the header row
            table.cell(0, 0).text = "Centre #"
            table.cell(0, 1).text = "Address"
            table.cell(0, 2).text = "Format"
            table.cell(0, 3).text = "Milestone Status"
            table.cell(0, 4).text = "Distance (miles)"

            # Fill the table with the closest centres data
            for i, (_, row) in enumerate(closest.iterrows()):
                table.cell(i + 1, 0).text = str(row["Centre Number"])
                table.cell(i + 1, 1).text = str(row["Addresses"])
                table.cell(i + 1, 2).text = str(row["Format - Type of Centre"])
                table.cell(i + 1, 3).text = str(row["Transaction Milestone Status"])
                table.cell(i + 1, 4).text = f"{row['Distance (miles)']:.2f}"

            # Adjust font size for table text
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(8)

            # Save PowerPoint file
            pptx_file = "closest_centres_presentation.pptx"
            prs.save(pptx_file)

            # Offer PowerPoint file for download
            st.download_button(
                label="Download PowerPoint Presentation",
                data=open(pptx_file, "rb").read(),
                file_name=pptx_file,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
    except Exception as e:
        st.error(f"An error occurred: {e}")
