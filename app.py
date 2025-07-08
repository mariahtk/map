import pandas as pd
from geopy.distance import geodesic
import streamlit as st
import folium
from streamlit_folium import st_folium
from folium.plugins import MarkerCluster
from pptx import Presentation
from pptx.util import Inches, Pt
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

# --- MAIN APP ---
st.set_page_config(page_title="Closest Centres Map", layout="wide")
st.title("📍 Find 5 Closest Centres")

api_key = "edd4cb8a639240daa178b4c6321a60e6"

input_address = st.text_input("Enter an address:")

if input_address:
    try:
        # --- GEOCODING ---
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

            # --- LOAD AND PROCESS EXCEL ---
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
                data = data.drop_duplicates(subset=["Centre Number"])

                for col in ["City", "State", "Zipcode"]:
                    if col not in data.columns:
                        data[col] = ""

                data["Distance (miles)"] = data.apply(
                    lambda row: geodesic(input_coords, (row["Latitude"], row["Longitude"])).miles, axis=1
                )

                data_sorted = data.sort_values("Distance (miles)").reset_index(drop=True)

                # --- SELECT 5 CLOSEST ---
                selected_centres = []
                seen_distances = []
                seen_centre_numbers = set()

                for _, row in data_sorted.iterrows():
                    centre_number = row["Centre Number"]
                    current_distance = row["Distance (miles)"]
                    if centre_number not in seen_centre_numbers and all(abs(current_distance - d) >= 0.005 for d in seen_distances):
                        selected_centres.append(row)
                        seen_centre_numbers.add(centre_number)
                        seen_distances.append(current_distance)
                    if len(selected_centres) == 5:
                        break

                closest = pd.DataFrame(selected_centres)

            # --- FOLIUM MAP ---
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

            marker_cluster = MarkerCluster().add_to(m)

            def get_marker_color(format_type):
                return {
                    "Regus": "blue",
                    "HQ": "darkblue",
                    "Signature": "purple",
                    "Spaces": "black",
                    "Non-Standard Brand": "gold"
                }.get(format_type, "red" if pd.isna(format_type) or format_type == "" else "gray")

            distance_text = f"Your Address: {input_address} - Coordinates: {input_coords[0]}, {input_coords[1]}\n\n"
            distance_text += "Closest Centres (Distances in miles):\n"

            for _, row in closest.iterrows():
                dest_coords = (row["Latitude"], row["Longitude"])
                folium.PolyLine([input_coords, dest_coords], color="blue", weight=2.5).add_to(m)
                marker_color = get_marker_color(row["Format - Type of Centre"])
                label_text = f"#{int(row['Centre Number'])} - ({row['Distance (miles)']:.2f} mi)"
                folium.Marker(
                    location=dest_coords,
                    popup=(f"#{int(row['Centre Number'])} - {row['Addresses']} | "
                           f"{row.get('City', 'N/A')}, {row.get('State', 'N/A')} {row.get('Zipcode', 'N/A')} | "
                           f"{row['Format - Type of Centre']} | {row['Transaction Milestone Status']} | "
                           f"{row['Distance (miles)']:.2f} mi"),
                    tooltip=folium.Tooltip(label_text, permanent=True, direction='right'),
                    icon=folium.Icon(color=marker_color)
                ).add_to(marker_cluster)

                distance_text += (
                    f"Centre #{int(row['Centre Number'])} - {row['Addresses']}, "
                    f"{row.get('City', 'N/A')}, {row.get('State', 'N/A')} {row.get('Zipcode', 'N/A')} - "
                    f"Format: {row['Format - Type of Centre']} - Milestone: {row['Transaction Milestone Status']} - "
                    f"{row['Distance (miles)']:.2f} miles\n"
                )

            # --- DRAW RADIUS & LEGEND ---
            folium.Circle(location=input_coords, radius=8046.72, color="green", fill=True, fill_opacity=0.2).add_to(m)
            legend_html = """<div style="position: fixed; bottom: 50px; left: 50px; width: 200px; height: 150px;
                                border:2px solid grey; z-index:9999; font-size:14px;
                                background-color:white; opacity: 0.85;">
                                &nbsp; <b>Legend</b> <br>
                                &nbsp; Your Address &nbsp; <i class="fa fa-map-marker fa-2x" style="color:green"></i><br>
                                &nbsp; Centre &nbsp; <i class="fa fa-map-marker fa-2x" style="color:blue"></i><br>
                                &nbsp; 5-mile Radius &nbsp; <i class="fa fa-circle" style="color:green"></i><br>
                            </div>"""
            m.get_root().html.add_child(folium.Element(legend_html))

            # --- DISPLAY ---
            col1, col2, col3 = st.columns([4, 1.2, 1])
            with col1:
                st_folium(m, width=950, height=650)
            with col2:
                st.markdown("""<div style="background-color: white; padding: 10px; border: 2px solid grey;
                                border-radius: 10px; width: 100%; margin-top: 20px;">
                                <b>Centre Type Legend</b><br>
                                <i style="background-color: lightgreen; padding: 5px;">&#9724;</i> Proposed Address<br>
                                <i style="background-color: lightblue; padding: 5px;">&#9724;</i> Regus<br>
                                <i style="background-color: darkblue; padding: 5px;">&#9724;</i> HQ<br>
                                <i style="background-color: purple; padding: 5px;">&#9724;</i> Signature<br>
                                <i style="background-color: black; padding: 5px;">&#9724;</i> Spaces<br>
                                <i style="background-color: gold; padding: 5px;">&#9724;</i> Non-Standard Brand
                            </div>""", unsafe_allow_html=True)
            with col3:
                st.markdown("""<div style="background-color: white; padding: 10px; border: 2px solid grey;
                                border-radius: 10px; width: 100%; margin-top: 20px;">
                                <b>Radius Legend</b><br>
                                <i style="background-color: green; padding: 5px;">&#9679;</i> 5-mile Radius
                            </div>""", unsafe_allow_html=True)

            st.subheader("Distances from Your Address to the Closest Centres:")
            st.text(distance_text)

            # --- POWERPOINT GENERATION ---
            st.subheader("Upload Map Screenshot for PowerPoint (Optional)")
            uploaded_image = st.file_uploader("Upload an image (e.g., screenshot of map)", type=["png", "jpg", "jpeg"])

            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[0])
            slide.shapes.title.text = "Closest Centres Presentation"
            slide.placeholders[1].text = f"Closest Centres to: {input_address}"

            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = "Closest Centres Map"

            if uploaded_image:
                slide.shapes.add_picture(uploaded_image, Inches(1), Inches(1.5), width=Inches(6))
            else:
                slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(8), Inches(4)).text = "Insert screenshot here."

            def add_distance_slide(prs, title_text, data):
                rows = len(data) + 1
                cols = 7
                slide = prs.slides.add_slide(prs.slide_layouts[5])
                slide.shapes.title.text = title_text
                table = slide.shapes.add_table(rows=rows, cols=cols, left=Inches(0.5), top=Inches(1.5), width=Inches(9), height=Inches(5)).table
                headers = ["Centre #", "Address", "City", "State", "Zip", "Distance (miles)", "Transaction Milestone"]
                for i, h in enumerate(headers):
                    cell = table.cell(0, i)
                    cell.text = h
                    for paragraph in cell.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.bold = True
                            run.font.size = Pt(12)
                for i, (_, row) in enumerate(data.iterrows(), start=1):
                    table.cell(i, 0).text = str(int(row['Centre Number'])) if pd.notna(row['Centre Number']) else "N/A"
                    table.cell(i, 1).text = row['Addresses'] or "N/A"
                    table.cell(i, 2).text = row.get("City", "") or "N/A"
                    table.cell(i, 3).text = row.get("State", "") or "N/A"
                    table.cell(i, 4).text = str(row.get("Zipcode", "")) or "N/A"
                    table.cell(i, 5).text = f"{row['Distance (miles)']:.2f}" if pd.notna(row['Distance (miles)']) else "N/A"
                    table.cell(i, 6).text = row.get("Transaction Milestone Status", "") or "N/A"

            half = (len(closest) + 1) // 2
            add_distance_slide(prs, "Distances to Closest Centres (1–3)", closest.iloc[:half])
            add_distance_slide(prs, "Distances to Closest Centres (4–5)", closest.iloc[half:])

            pptx_path = "closest_centres_presentation.pptx"
            prs.save(pptx_path)
            st.download_button("Download PowerPoint Presentation", data=open(pptx_path, "rb"), file_name=pptx_path,
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

    except Exception as e:
        st.error(f"Error: {e}")

else:
    st.info("Please enter an address above to get started.")
