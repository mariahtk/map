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
                data = data.drop_duplicates(subset=["Centre Number"])  # Remove duplicate centre numbers
                data["Distance (miles)"] = data.apply(
                    lambda row: geodesic(input_coords, (row["Latitude"], row["Longitude"])).miles, axis=1
                )

                data_sorted = data.sort_values("Distance (miles)").reset_index(drop=True)
                selected_centres = []
                seen_distances = []

                for _, row in data_sorted.iterrows():
                    current_distance = row["Distance (miles)"]
                    if all(abs(current_distance - d) >= 0.01 for d in seen_distances):
                        selected_centres.append(row)
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

            stagger_offsets = [-0.002, 0.002, -0.0015, 0.0015, -0.001, 0.001, -0.0005, 0.0005]

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

            for i, (index, row) in enumerate(closest.iterrows()):
                dest_coords = (row["Latitude"], row["Longitude"])
                folium.PolyLine([input_coords, dest_coords], color="blue", weight=2.5, opacity=1).add_to(m)

                marker_color = get_marker_color(row["Format - Type of Centre"])

                folium.Marker(
                    location=dest_coords,
                    popup=f"Centre #{int(row['Centre Number'])}<br>Address: {row['Addresses']}<br>Format: {row['Format - Type of Centre']}<br>Transaction Milestone: {row['Transaction Milestone Status']}<br>Distance: {row['Distance (miles)']:.2f} miles",
                    icon=folium.Icon(color=marker_color)
                ).add_to(m)

                distance_text += f"Centre #{int(row['Centre Number'])} - {row['Addresses']} - Format: {row['Format - Type of Centre']} - Milestone: {row['Transaction Milestone Status']} - {row['Distance (miles)']:.2f} miles\n"

                label_text = f"#{int(row['Centre Number'])} - {row['Addresses']} ({row['Distance (miles)']:.2f} mi)"
                offset_lat = stagger_offsets[i % len(stagger_offsets)]

                label_lat = row["Latitude"] + offset_lat
                label_lon = row["Longitude"]
                if label_lat > lat_max:
                    label_lat = lat_max - 0.0005
                if label_lat < lat_min:
                    label_lat = lat_min + 0.0005
                if label_lon > lng_max:
                    label_lon = lng_max - 0.0005
                if label_lon < lng_min:
                    label_lon = lng_min + 0.0005

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
                                display: inline-block;
                                white-space: nowrap;
                                text-overflow: ellipsis;
                                box-shadow: 1px 1px 3px rgba(0,0,0,0.2);
                            ">
                                {label_text}
                            </div>
                        """
                    )
                ).add_to(m)

            folium_map_path = "closest_centres_map.html"
            m.save(folium_map_path)

            col1, col2 = st.columns([4, 1])

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

            st.subheader("Distances from Your Address to the Closest Centres:")
            st.text(distance_text)

            # --- POWERPOINT GENERATION ---
            st.subheader("Upload Map Screenshot for PowerPoint (Optional)")
            uploaded_image = st.file_uploader("Upload an image (e.g., screenshot of map)", type=["png", "jpg", "jpeg"])

            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[0])
            title = slide.shapes.title
            subtitle = slide.placeholders[1]
            title.text = "Closest Centres Presentation"
            subtitle.text = f"Closest Centres to: {input_address}"

            slide = prs.slides.add_slide(prs.slide_layouts[5])
            title = slide.shapes.title
            title.text = "Closest Centres Map"

            if uploaded_image:
                slide.shapes.add_picture(uploaded_image, Inches(1), Inches(1.5), width=Inches(6))
            else:
                slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(8), Inches(4)).text = "Insert screenshot here."

            group1 = closest.iloc[:4]
            group2 = closest.iloc[4:]

            def add_distance_slide(prs, title_text, data):
                slide = prs.slides.add_slide(prs.slide_layouts[5])
                title = slide.shapes.title
                title.text = title_text
                table = slide.shapes.add_table(rows=5, cols=5, left=Inches(0.5), top=Inches(1.5), width=Inches(8), height=Inches(5)).table

                table.cell(0, 0).text = "Centre #"
                table.cell(0, 1).text = "Address"
                table.cell(0, 2).text = "Format - Type of Centre"
                table.cell(0, 3).text = "Transaction Milestone"
                table.cell(0, 4).text = "Distance (miles)"

                for i, (_, row) in enumerate(data.iterrows(), start=1):
                    table.cell(i, 0).text = str(int(row['Centre Number'])) if pd.notna(row['Centre Number']) else "N/A"
                    table.cell(i, 1).text = row['Addresses'] if pd.notna(row['Addresses']) else "N/A"
                    table.cell(i, 2).text = row['Format - Type of Centre'] if pd.notna(row['Format - Type of Centre']) else "N/A"
                    table.cell(i, 3).text = row['Transaction Milestone Status'] if pd.notna(row['Transaction Milestone Status']) else "N/A"
                    table.cell(i, 4).text = f"{row['Distance (miles)']:.2f}" if pd.notna(row['Distance (miles)']) else "N/A"

            add_distance_slide(prs, "Distances to Closest Centres (1–4)", group1)
            add_distance_slide(prs, "Distances to Closest Centres (5–8)", group2)

            pptx_path = "closest_centres_presentation.pptx"
            prs.save(pptx_path)
            st.download_button("Download PowerPoint Presentation", data=open(pptx_path, "rb"), file_name=pptx_path, mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

    except Exception as e:
        st.error(f"❌ An error occurred: {e}")
