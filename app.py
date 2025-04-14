import pandas as pd
from geopy.distance import geodesic
from geopy.geocoders import Nominatim
import streamlit as st
import folium
from streamlit_folium import st_folium
import folium.plugins as plugins
from io import BytesIO
from pptx import Presentation
from pptx.util import Pt

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

            # Calculate bounds
            lats = [input_coords[0]] + closest["Latitude"].tolist()
            lngs = [input_coords[1]] + closest["Longitude"].tolist()
            lat_min, lat_max = min(lats), max(lats)
            lng_min, lng_max = min(lngs), max(lngs)
            max_diff = max(lat_max - lat_min, lng_max - lng_min)
            zoom_level = max(14 - (max_diff * 3), 14)

            # Map
            m = folium.Map(location=input_coords, zoom_start=int(zoom_level))
            folium.Marker(
                location=input_coords,
                popup=f"Your Address: {input_address}",
                icon=folium.Icon(color="green")
            ).add_to(m)

            distance_text = f"Your Address: {input_address} - Coordinates: {input_coords[0]}, {input_coords[1]}\n"
            distance_text += "\nClosest Centres (Distances in miles):\n"

            stagger_offsets = [-0.002, 0.002, -0.0015, 0.0015, -0.001, 0.001, -0.0005, 0.0005]

            for i, (_, row) in enumerate(closest.iterrows()):
                dest_coords = (row["Latitude"], row["Longitude"])
                folium.PolyLine([input_coords, dest_coords], color="blue", weight=2.5, opacity=1).add_to(m)

                folium.Marker(
                    location=dest_coords,
                    popup=f"Centre #{row['Centre Number']}<br>Address: {row['Addresses']}<br>Format: {row['Format - Type of Centre']}<br>Transaction Milestone: {row['Transaction Milestone Status']}<br>Distance: {row['Distance (miles)']:.2f} miles",
                    icon=folium.Icon(color="blue")
                ).add_to(m)

                distance_text += (
                    f"Centre #{row['Centre Number']} - {row['Addresses']} - "
                    f"Format: {row['Format - Type of Centre']} - "
                    f"Milestone: {row['Transaction Milestone Status']} - "
                    f"{row['Distance (miles)']:.2f} miles\n"
                )

                label_text = f"#{row['Centre Number']} - {row['Addresses']} ({row['Distance (miles)']:.2f} mi)"
                offset_lat = stagger_offsets[i % len(stagger_offsets)]
                label_lat = max(min(row["Latitude"] + offset_lat, lat_max - 0.0005), lat_min + 0.0005)
                label_lon = max(min(row["Longitude"], lng_max - 0.0005), lng_min + 0.0005)

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

            # Display map and distances
            m.save("closest_centres_map.html")
            st_folium(m, width=950, height=650)
            st.subheader("Distances from Your Address to the Closest Centres:")
            st.text(distance_text)

            # --- PowerPoint Slides ---
            prs = Presentation()

            # Slide 1 - Title Placeholder
            slide = prs.slides.add_slide(prs.slide_layouts[0])
            title = slide.shapes.title
            subtitle = slide.placeholders[1]
            title.text = "Closest Centres Presentation"
            subtitle.text = f"Closest Centres to: {input_address}"

            # Slide 2 - Distance data
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            title = slide.shapes.title
            title.text = "Distances to Closest Centres"
            textbox = slide.shapes.placeholders[1].text_frame
            textbox.clear()
            textbox.text = f"Your Address: {input_address} - Coordinates: {input_coords[0]:.6f}, {input_coords[1]:.6f}"

            for i, (_, row) in enumerate(closest.iterrows(), start=1):
                paragraph = textbox.add_paragraph()
                paragraph.text = (
                    f"\n{i}. Centre #{row['Centre Number']} - {row['Addresses']} - "
                    f"Format: {row['Format - Type of Centre']} - "
                    f"Milestone: {row['Transaction Milestone Status']} - "
                    f"Distance: {row['Distance (miles)']:.2f} miles"
                )
                paragraph.space_after = Pt(10)

            # Save the PowerPoint presentation
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
