import pandas as pd
from geopy.distance import geodesic
from geopy.geocoders import Nominatim
import streamlit as st
import folium
from streamlit_folium import st_folium
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

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

                data["Distance (miles)"] = data.apply(
                    lambda row: geodesic(input_coords, (row["Latitude"], row["Longitude"])).miles, axis=1
                )

                closest = data.nsmallest(8, "Distance (miles)")

            # Generate map
            m = folium.Map(location=input_coords, zoom_start=14)
            folium.Marker(location=input_coords, popup=f"Your Address: {input_address}",
                          icon=folium.Icon(color="green")).add_to(m)

            for i, (_, row) in enumerate(closest.iterrows()):
                dest_coords = (row["Latitude"], row["Longitude"])
                folium.Marker(
                    location=dest_coords,
                    popup=f"Centre #{row['Centre Number']}<br>Address: {row['Addresses']}<br>Format: {row['Format - Type of Centre']}<br>Milestone: {row['Transaction Milestone Status']}<br>Distance: {row['Distance (miles)']:.2f} mi",
                    icon=folium.Icon(color="blue")
                ).add_to(m)
                folium.PolyLine([input_coords, dest_coords], color="blue", weight=2.5).add_to(m)

            m.save("closest_centres_map.html")
            st_folium(m, width=950, height=650)

            # PowerPoint creation
            prs = Presentation()

            # Slide 1 - Title + Map Placeholder
            slide1 = prs.slides.add_slide(prs.slide_layouts[1])
            slide1.shapes.title.text = "Closest Centres Map"
            content = slide1.placeholders[1]
            content.text = "Insert screenshot of the map here from the Streamlit app."

            # Slide 2 - Table with 8 closest centres
            slide2 = prs.slides.add_slide(prs.slide_layouts[5])
            shapes = slide2.shapes
            slide2.shapes.title.text = "8 Closest Centres"

            rows, cols = 9, 5
            left = Inches(0.5)
            top = Inches(1.5)
            width = Inches(9)
            height = Inches(4.5)

            table = shapes.add_table(rows, cols, left, top, width, height).table

            col_names = ["Centre #", "Address", "Format", "Milestone", "Distance (mi)"]
            for col_idx, name in enumerate(col_names):
                cell = table.cell(0, col_idx)
                cell.text = name
                cell.text_frame.paragraphs[0].font.bold = True

            for row_idx, (_, row) in enumerate(closest.iterrows(), start=1):
                table.cell(row_idx, 0).text = str(row['Centre Number'])
                table.cell(row_idx, 1).text = str(row['Addresses'])
                table.cell(row_idx, 2).text = str(row['Format - Type of Centre'])
                table.cell(row_idx, 3).text = str(row['Transaction Milestone Status'])
                table.cell(row_idx, 4).text = f"{row['Distance (miles)']:.2f}"

            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.text_frame.paragraphs:
                        paragraph.font.size = Pt(12)

            # Save and download
            pptx_file = "closest_centres_presentation.pptx"
            prs.save(pptx_file)

            st.download_button(
                label="📥 Download PowerPoint",
                data=open(pptx_file, "rb").read(),
                file_name=pptx_file,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

    except Exception as e:
        st.error(f"Error: {e}")
