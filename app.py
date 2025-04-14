import pandas as pd
from geopy.distance import geodesic
from geopy.geocoders import Nominatim
import streamlit as st
import folium
from streamlit_folium import st_folium
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# Streamlit setup
st.set_page_config(page_title="Closest Centres Map", layout="wide")
st.title("📍 Find 8 Closest IWG Centres")

input_address = st.text_input("Enter an address:")

if input_address:
    try:
        # Geocode the input address
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

                # Find 8 closest
                closest = data.nsmallest(8, "Distance (miles)")

            # Map
            m = folium.Map(location=input_coords, zoom_start=12)
            folium.Marker(
                location=input_coords,
                popup=f"Your Address: {input_address}",
                icon=folium.Icon(color="green")
            ).add_to(m)

            for _, row in closest.iterrows():
                dest_coords = (row["Latitude"], row["Longitude"])
                folium.PolyLine([input_coords, dest_coords], color="blue", weight=2.5, opacity=1).add_to(m)
                folium.Marker(
                    location=dest_coords,
                    popup=f"Centre #{row['Centre Number']}<br>{row['Addresses']}<br>{row['Format - Type of Centre']}<br>{row['Transaction Milestone Status']}<br>{row['Distance (miles)']:.2f} miles",
                    icon=folium.Icon(color="blue")
                ).add_to(m)

            m.save("closest_centres_map.html")
            st_folium(m, width=950, height=650)

            # --- PowerPoint Slides ---
            prs = Presentation()

            # Slide 1 - Map placeholder
            slide1 = prs.slides.add_slide(prs.slide_layouts[1])
            slide1.shapes.title.text = "Map Overview"
            content1 = slide1.placeholders[1].text_frame
            content1.text = "Insert map screenshot here."

            # Slide 2 - Table of closest centres
            slide2 = prs.slides.add_slide(prs.slide_layouts[5])
            slide2.shapes.title.text = "Closest Centres Data"

            rows = len(closest) + 1
            cols = 5
            left = Inches(0.5)
            top = Inches(1.5)
            width = Inches(9)
            height = Inches(0.8)

            table = slide2.shapes.add_table(rows, cols, left, top, width, height).table

            # Set column headings
            col_names = ["Centre #", "Address", "Format", "Milestone", "Distance (mi)"]
            for col_idx, name in enumerate(col_names):
                cell = table.cell(0, col_idx)
                cell.text = name
                cell.text_frame.paragraphs[0].font.bold = True
                cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(200, 200, 200)

            # Populate rows
            for i, (_, row) in enumerate(closest.iterrows(), start=1):
                table.cell(i, 0).text = str(row["Centre Number"])
                table.cell(i, 1).text = row["Addresses"]
                table.cell(i, 2).text = str(row["Format - Type of Centre"])
                table.cell(i, 3).text = str(row["Transaction Milestone Status"])
                table.cell(i, 4).text = f"{row['Distance (miles)']:.2f}"

                for j in range(cols):
                    table.cell(i, j).text_frame.paragraphs[0].font.size = Pt(10)
                    table.cell(i, j).text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT

            # Adjust column widths
            table.columns[0].width = Inches(1)
            table.columns[1].width = Inches(3.5)
            table.columns[2].width = Inches(1.6)
            table.columns[3].width = Inches(2.3)
            table.columns[4].width = Inches(1)

            # Save PowerPoint
            pptx_file = "closest_centres_presentation.pptx"
            prs.save(pptx_file)

            st.download_button(
                label="Download PowerPoint",
                data=open(pptx_file, "rb").read(),
                file_name=pptx_file,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

    except Exception as e:
        st.error(f"Error: {e}")
