import pandas as pd
from geopy.distance import geodesic
from geopy.geocoders import Nominatim
import streamlit as st
import folium
from streamlit_folium import st_folium
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO

# Streamlit setup
st.set_page_config(page_title="Closest Centres Map", layout="wide")
st.title("📍 Find 8 Closest Centres")

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

            # Load data
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

            # Create map
            m = folium.Map(location=input_coords, zoom_start=12)

            # Marker for input address
            folium.Marker(
                location=input_coords,
                popup=f"Your Address: {input_address}",
                icon=folium.Icon(color="green")
            ).add_to(m)

            bullet_points = []

            for _, row in closest.iterrows():
                dest_coords = (row["Latitude"], row["Longitude"])

                # Draw line
                folium.PolyLine([input_coords, dest_coords], color="blue", weight=2.5, opacity=1).add_to(m)

                # One-line readable label
                label = (
                    f"Centre #{row['Centre Number']} | {row['Addresses']} | "
                    f"{row['Distance (miles)']:.2f} miles"
                )
                folium.Marker(
                    location=dest_coords,
                    icon=folium.DivIcon(html=f'<div style="font-size: 10px; color: black; background-color: white;">{label}</div>')
                ).add_to(m)

                bullet_points.append(label)

            # Display map
            st_data = st_folium(m, width=800, height=600)

            # PowerPoint Export Section
            def generate_pptx(address, bullets):
                prs = Presentation()
                slide_layout = prs.slide_layouts[5]  # Title only
                slide = prs.slides.add_slide(slide_layout)

                title = slide.shapes.title
                title.text = "8 Closest Centres"

                # Add bullet list
                left = Inches(0.5)
                top = Inches(1.2)
                width = Inches(8)
                height = Inches(4)

                textbox = slide.shapes.add_textbox(left, top, width, height)
                tf = textbox.text_frame
                tf.word_wrap = True

                for point in bullets:
                    p = tf.add_paragraph()
                    p.text = point
                    p.level = 0

                # Add placeholder for screenshot
                slide.shapes.add_textbox(
                    Inches(0.5), Inches(5.5), Inches(8), Inches(1)
                ).text = "[Paste screenshot of map here]"

                # Return binary PPTX
                pptx_io = BytesIO()
                prs.save(pptx_io)
                pptx_io.seek(0)
                return pptx_io

            pptx_bytes = generate_pptx(input_address, bullet_points)

            # Download button
            st.download_button(
                label="📥 Download PowerPoint Summary",
                data=pptx_bytes,
                file_name="Closest_Centres_Summary.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

    except Exception as e:
        st.error(f"Error: {e}")
