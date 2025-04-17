import pandas as pd
from geopy.distance import geodesic
import streamlit as st
import folium
from streamlit_folium import st_folium
import folium.plugins as plugins
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches
import requests
import urllib.parse

# Streamlit setup
st.set_page_config(page_title="Closest Centres Map", layout="wide")
st.title("📍 Find 8 Closest Centres")

# Your API key
api_key = "edd4cb8a639240daa178b4c6321a60e6"

input_address = st.text_input("Enter an address:")

if input_address:
    try:
        encoded_address = urllib.parse.quote(input_address)
        url = f"https://api.opencagedata.com/geocode/v1/json?q={encoded_address}&key={api_key}"
        response = requests.get(url)
        data = response.json()

        if response.status_code != 200:
            st.error(f"❌ API Error: {response.status_code}. Try again.")
        elif not data.get('results'):
            st.error("❌ Address not found. Try again.")
        else:
            location = data['results'][0]
            input_coords = (location['geometry']['lat'], location['geometry']['lng'])

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

            # Map boundaries
            lats = [input_coords[0]] + closest["Latitude"].tolist()
            lngs = [input_coords[1]] + closest["Longitude"].tolist()
            lat_min, lat_max = min(lats), max(lats)
            lng_min, lng_max = min(lngs), max(lngs)

            # Dynamic zoom approximation
            m = folium.Map(location=input_coords, zoom_start=13)

            # Input address marker
            folium.Marker(
                location=input_coords,
                popup=f"Your Address: {input_address}",
                icon=folium.Icon(color="green")
            ).add_to(m)

            distance_text = f"Your Address: {input_address} - Coordinates: {input_coords[0]}, {input_coords[1]}\n"
            distance_text += "\nClosest Centres (Distances in miles):\n"

            stagger_offsets = [-0.002, 0.002, -0.0015, 0.0015, -0.001, 0.001, -0.0005, 0.0005]

            for i, (index, row) in enumerate(closest.iterrows()):
                dest_coords = (row["Latitude"], row["Longitude"])
                folium.PolyLine([input_coords, dest_coords], color="blue", weight=2.5, opacity=1).add_to(m)

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

                label_text = f"#{row['Centre Number']} - {row['Addresses']} ({row['Distance (miles)']:.2f} mi)"
                offset_lat = stagger_offsets[i % len(stagger_offsets)]
                label_lat = row["Latitude"] + offset_lat
                label_lon = row["Longitude"]

                label_lat = min(max(label_lat, lat_min + 0.0005), lat_max - 0.0005)
                label_lon = min(max(label_lon, lng_min + 0.0005), lng_max - 0.0005)

                folium.Marker(
                    location=(label_lat, label_lon),
                    icon=folium.DivIcon(
                        icon_size=(150, 40),
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
                                text-overflow: ellipsis;
                                box-shadow: 1px 1px 3px rgba(0,0,0,0.2);
                            ">
                                {label_text}
                            </div>
                        """
                    )
                ).add_to(m)

                distance_text += f"Centre #{row['Centre Number']} - {row['Addresses']} - Format: {row['Format - Type of Centre']} - Milestone: {row['Transaction Milestone Status']} - {row['Distance (miles)']:.2f} miles\n"

            st_folium(m, width=950, height=650)

            st.subheader("Distances from Your Address to the Closest Centres:")
            st.text(distance_text)

            prs = Presentation()

            slide = prs.slides.add_slide(prs.slide_layouts[0])
            slide.shapes.title.text = "Closest Centres Presentation"
            slide.placeholders[1].text = f"Closest Centres to: {input_address}"

            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = "Closest Centres Map"
            slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(8), Inches(4)).text = "Insert screenshot here."

            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = "Distances to Closest Centres"
            table = slide.shapes.add_table(rows=len(closest) + 1, cols=5, left=Inches(0.5), top=Inches(1.5), width=Inches(9), height=Inches(4.5))
            table_data = table.table

            headers = ["Centre #", "Address", "Type of Centre", "Transaction Milestone", "Distance (miles)"]
            for col, header in enumerate(headers):
                table_data.cell(0, col).text = header

            for i, (_, row) in enumerate(closest.iterrows()):
                table_data.cell(i + 1, 0).text = str(row["Centre Number"])
                table_data.cell(i + 1, 1).text = str(row["Addresses"])
                table_data.cell(i + 1, 2).text = str(row["Format - Type of Centre"])
                table_data.cell(i + 1, 3).text = str(row["Transaction Milestone Status"])
                table_data.cell(i + 1, 4).text = f"{row['Distance (miles)']:.2f}"

            presentation_file = BytesIO()
            prs.save(presentation_file)
            presentation_file.seek(0)
            st.download_button("Download Presentation", presentation_file, "closest_centres_presentation.pptx")

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
