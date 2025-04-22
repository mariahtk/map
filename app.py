import pandas as pd
from geopy.distance import geodesic
import streamlit as st
import folium
from streamlit_folium import st_folium
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches
import requests
import urllib.parse

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

# --- APP START ---
st.set_page_config(page_title="Closest Centres Map", layout="wide")
st.title("📍 Find 8 Closest Centres")

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

            # Map bounds
            lats = [input_coords[0]] + closest["Latitude"].tolist()
            lngs = [input_coords[1]] + closest["Longitude"].tolist()
            lat_min, lat_max = min(lats), max(lats)
            lng_min, lng_max = min(lngs), max(lngs)
            max_diff = max(lat_max - lat_min, lng_max - lng_min)
            zoom_level = max(14 - (max_diff * 3), 14)

            m = folium.Map(location=input_coords, zoom_start=int(zoom_level))

            folium.Marker(
                location=input_coords,
                popup=f"Your Address: {input_address}",
                icon=folium.Icon(color="green")
            ).add_to(m)

            distance_text = f"Your Address: {input_address} - Coordinates: {input_coords[0]}, {input_coords[1]}\n"
            distance_text += "\nClosest Centres (Distances in miles):\n"

            def get_marker_color(format_type):
                colors = {
                    "Regus": "blue",
                    "HQ": "darkblue",
                    "Signature": "purple",
                    "Spaces": "black",
                    "Non-Standard Brand": "gold",
                    "": "red"
                }
                return colors.get(format_type, "gray")

            # Smart label and distance placement
            used_label_positions = set()
            used_distances = set()

            def find_non_overlapping_offset(base_lat, base_lon, used_positions, step=0.0003):
                offset_pairs = [
                    (step * i, step * j)
                    for i in range(-5, 6)
                    for j in range(-5, 6)
                    if not (i == 0 and j == 0)
                ]
                for offset_lat, offset_lon in offset_pairs:
                    new_lat = base_lat + offset_lat
                    new_lon = base_lon + offset_lon
                    key = (round(new_lat, 6), round(new_lon, 6))
                    if key not in used_positions:
                        used_positions.add(key)
                        return new_lat, new_lon
                return base_lat, base_lon

            for i, (index, row) in enumerate(closest.iterrows()):
                dest_coords = (row["Latitude"], row["Longitude"])
                distance_miles = round(row['Distance (miles)'], 2)

                while distance_miles in used_distances:
                    distance_miles = round(distance_miles + 0.001, 3)
                used_distances.add(distance_miles)

                # Line from input to centre
                folium.PolyLine([input_coords, dest_coords], color="blue", weight=2.5, opacity=1).add_to(m)

                marker_color = get_marker_color(row["Format - Type of Centre"])

                # Marker
                folium.Marker(
                    location=dest_coords,
                    popup=f"Centre #{int(row['Centre Number'])}<br>Address: {row['Addresses']}<br>Format: {row['Format - Type of Centre']}<br>Transaction Milestone: {row['Transaction Milestone Status']}<br>Distance: {distance_miles:.3f} miles",
                    icon=folium.Icon(color=marker_color)
                ).add_to(m)

                distance_text += f"Centre #{int(row['Centre Number'])} - {row['Addresses']} - Format: {row['Format - Type of Centre']} - Milestone: {row['Transaction Milestone Status']} - {distance_miles:.3f} miles\n"

                label_text = f"#{int(row['Centre Number'])} - {row['Addresses']} ({distance_miles:.3f} mi)"
                label_lat, label_lon = find_non_overlapping_offset(row["Latitude"], row["Longitude"], used_label_positions)

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

            # Save map and display
            m.save("closest_centres_map.html")
            col1, col2 = st.columns([4, 1])

            with col1:
                st_folium(m, width=950, height=650)

            with col2:
                st.markdown("""
                    <div style="background-color: white; padding: 10px; border: 2px solid grey; border-radius: 10px; width: 100%; margin-top: 20px;">
                        <b>Centre Type Legend</b><br>
                        <i style="background-color: blue; padding: 5px;">&#9724;</i> Regus<br>
                        <i style="background-color: darkblue; padding: 5px;">&#9724;</i> HQ<br>
                        <i style="background-color: purple; padding: 5px;">&#9724;</i> Signature<br>
                        <i style="background-color: black; padding: 5px;">&#9724;</i> Spaces<br>
                        <i style="background-color: red; padding: 5px;">&#9724;</i> Mature<br>
                        <i style="background-color: gold; padding: 5px;">&#9724;</i> Non-Standard Brand
                    </div>
                """, unsafe_allow_html=True)

            st.subheader("Distances from Your Address to the Closest Centres:")
            st.text(distance_text)

            # PowerPoint section
            uploaded_map = st.file_uploader("Upload a screenshot or image of the map for PowerPoint", type=["png", "jpg", "jpeg"])
            prs = Presentation()

            # Title slide
            slide = prs.slides.add_slide(prs.slide_layouts[0])
            slide.shapes.title.text = "Closest Centres Presentation"
            slide.placeholders[1].text = f"Closest Centres to: {input_address}"

            # Map slide
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = "Closest Centres Map"
            if uploaded_map:
                img_stream = BytesIO(uploaded_map.read())
                slide.shapes.add_picture(img_stream, Inches(1), Inches(1.5), width=Inches(8), height=Inches(4.5))
            else:
                slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(8), Inches(4)).text = "Insert map screenshot here."

            def add_table_slide(prs, title_text, centres_chunk):
                slide = prs.slides.add_slide(prs.slide_layouts[5])
                slide.shapes.title.text = title_text
                table = slide.shapes.add_table(
                    rows=len(centres_chunk)+1, cols=5,
                    left=Inches(0.5), top=Inches(1.5),
                    width=Inches(8), height=Inches(5)
                ).table

                headers = ["Centre #", "Address", "Format - Type of Centre", "Transaction Milestone", "Distance (miles)"]
                for col, header in enumerate(headers):
                    table.cell(0, col).text = header

                for i, (_, row) in enumerate(centres_chunk.iterrows(), start=1):
                    table.cell(i, 0).text = str(int(row['Centre Number'])) if pd.notna(row['Centre Number']) else "N/A"
                    table.cell(i, 1).text = row['Addresses'] if pd.notna(row['Addresses']) else "N/A"
                    table.cell(i, 2).text = row['Format - Type of Centre'] if pd.notna(row['Format - Type of Centre']) else "N/A"
                    table.cell(i, 3).text = row['Transaction Milestone Status'] if pd.notna(row['Transaction Milestone Status']) else "N/A"
                    table.cell(i, 4).text = f"{row['Distance (miles)']:.2f}" if pd.notna(row['Distance (miles)']) else "N/A"

            add_table_slide(prs, "Distances to Closest Centres (1–4)", closest.iloc[:4])
            add_table_slide(prs, "Distances to Closest Centres (5–8)", closest.iloc[4:])

            pptx_path = "closest_centres_presentation.pptx"
            prs.save(pptx_path)
            st.download_button("Download PowerPoint Presentation", data=open(pptx_path, "rb"), file_name=pptx_path, mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

    except Exception as e:
        st.error(f"❌ An error occurred: {e}")
