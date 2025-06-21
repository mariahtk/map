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
from folium.plugins import MarkerCluster

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

st.set_page_config(page_title="Closest Centres Map", layout="wide")
st.title("📍 Find 5 Closest Centres")

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
                data = data.drop_duplicates(subset=["Centre Number"])

                for col in ["City", "State", "Zipcode"]:
                    if col not in data.columns:
                        data[col] = ""

                data["Distance (miles)"] = data.apply(
                    lambda row: geodesic(input_coords, (row["Latitude"], row["Longitude"])).miles, axis=1
                )

                data_sorted = data.sort_values("Distance (miles)").reset_index(drop=True)

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

            lats = [input_coords[0]] + closest["Latitude"].tolist()
            lngs = [input_coords[1]] + closest["Longitude"].tolist()
            lat_min, lat_max = min(lats), max(lats)
            lng_min, lng_max = min(lngs), max(lngs)
            lat_diff = lat_max - lat_min
            lng_diff = lng_max - lng_min
            max_diff = max(lat_diff, lng_diff)
            zoom_level = 14 - (max_diff * 3)
            zoom_level = max(zoom_level, 14)

            m = folium.Map(location=input_coords, zoom_start=int(zoom_level))

            folium.Marker(
                location=input_coords,
                popup=f"Your Address: {input_address}",
                icon=folium.Icon(color="green")
            ).add_to(m)

            distance_text = f"Your Address: {input_address} - Coordinates: {input_coords[0]}, {input_coords[1]}\n"
            distance_text += "\nClosest Centres (Distances in miles):\n"

            def get_marker_color(format_type):
                if format_type == "Regus":
                    return "blue"
                elif format_type == "HQ":
                    return "darkblue"
                elif format_type == "Signature":
                    return "purple"
                elif format_type == "Spaces":
                    return "black"
                elif format_type == "Non-Standard Brand":
                    return "gold"
                elif pd.isna(format_type) or format_type == "":
                    return "red"
                return "gray"

            marker_cluster = MarkerCluster().add_to(m)

            for i, (index, row) in enumerate(closest.iterrows()):
                dest_coords = (row["Latitude"], row["Longitude"])
                folium.PolyLine([input_coords, dest_coords], color="blue", weight=2.5, opacity=1).add_to(m)

                marker_color = get_marker_color(row["Format - Type of Centre"])

                label_text = f"#{int(row['Centre Number'])} - {row['Addresses']} ({row['Distance (miles)']:.2f} mi)"

                # Add marker with popup and permanent tooltip (text box) that moves with the marker
                folium.Marker(
                    location=dest_coords,
                    popup=(
                        f"#{int(row['Centre Number'])} - {row['Addresses']} | "
                        f"{row.get('City', 'N/A')}, {row.get('State', 'N/A')} {row.get('Zipcode', 'N/A')} | "
                        f"{row['Format - Type of Centre']} | {row['Transaction Milestone Status']} | "
                        f"{row['Distance (miles)']:.2f} mi"
                    ),
                    tooltip=folium.Tooltip(label_text, permanent=True, sticky=False, direction='right'),
                    icon=folium.Icon(color=marker_color)
                ).add_to(marker_cluster)

                distance_text += (
                    f"Centre #{int(row['Centre Number'])} - {row['Addresses']}, "
                    f"{row.get('City', 'N/A')}, {row.get('State', 'N/A')} {row.get('Zipcode', 'N/A')} - "
                    f"Format: {row['Format - Type of Centre']} - Milestone: {row['Transaction Milestone Status']} - "
                    f"{row['Distance (miles)']:.2f} miles\n"
                )

            folium.Circle(
                location=input_coords,
                radius=8046.72,
                color="green",
                fill=True,
                fill_opacity=0.2,
                fill_color="green"
            ).add_to(m)

            legend_html = """
            <div style="
                position: fixed;
                bottom: 50px; left: 50px; width: 200px; height: 150px;
                border:2px solid grey; z-index:9999; font-size:14px;
                background-color:white; opacity: 0.85;">
                &nbsp; <b>Legend</b> <br>
                &nbsp; Your Address &nbsp; <i class="fa fa-map-marker fa-2x" style="color:green"></i><br>
                &nbsp; Centre &nbsp; <i class="fa fa-map-marker fa-2x" style="color:blue"></i><br>
                &nbsp; 5-mile Radius &nbsp; <i class="fa fa-circle" style="color:green"></i><br>
            </div>
            """
            m.get_root().html.add_child(folium.Element(legend_html))

            col1, col2, col3 = st.columns([4, 1.2, 1])

            with col1:
                st_folium(m, width=950, height=650)

            with col2:
                st.markdown(""" 
                    <div style="background-color: white; padding: 10px; border: 2px solid grey; border-radius: 10px; width: 100%; margin-top: 20px;">
                        <b>Centre Type Legend</b><br>
                        <i style="background-color: lightgreen; padding: 5px;">&#9724;</i> Proposed Address<br>
                        <i style="background-color: lightblue; padding: 5px;">&#9724;</i> Regus<br>
                        <i style="background-color: darkblue; padding: 5px;">&#9724;</i> HQ<br>
                        <i style="background-color: purple; padding: 5px;">&#9724;</i> Signature<br>
                        <i style="background-color: black; padding: 5px;">&#9724;</i> Spaces<br>
                        <i style="background-color: gold; padding: 5px;">&#9724;</i> Non-Standard Brand
                    </div>
                """, unsafe_allow_html=True)

            with col3:
                st.markdown(""" 
                    <div style="background-color: white; padding: 10px; border: 2px solid grey; border-radius: 10px; width: 100%; margin-top: 20px;">
                        <b>Radius Legend</b><br>
                        <i style="background-color: green; padding: 5px;">&#9679;</i> 5-mile Radius
                    </div>
                """, unsafe_allow_html=True)

            st.subheader("Distances from Your Address to the Closest Centres:")
            st.text(distance_text)

            st.subheader("Upload Map Screenshot for PowerPoint (Optional)")
            uploaded_image = st.file_uploader("Upload an image (e.g., screenshot of map)", type=["png", "jpg", "jpeg"])

            prs = Presentation()
            slide_layout = prs.slide_layouts[5]
            slide = prs.slides.add_slide(slide_layout)
            title_shape = slide.shapes.title
            title_shape.text = "Closest Centres Presentation"

            left = Inches(0.5)
            top = Inches(1.5)
            width = Inches(9)
            height = Inches(5)

            if uploaded_image is not None:
                image_stream = BytesIO(uploaded_image.read())
                slide.shapes.add_picture(image_stream, left, top, width=width, height=height)

            txBox = slide.shapes.add_textbox(left, Inches(6.7), width, Inches(1.5))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = distance_text
            p.font.size = Pt(14)

            pptx_bytes = BytesIO()
            prs.save(pptx_bytes)
            pptx_bytes.seek(0)

            st.download_button(
                label="Download PowerPoint",
                data=pptx_bytes,
                file_name="Closest_Centres_Presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

    except Exception as e:
        st.error(f"Error: {e}")

else:
    st.info("Please enter an address above to get started.")
