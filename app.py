import pandas as pd
from geopy.distance import geodesic
import streamlit as st
import folium
from streamlit_folium import st_folium
from pptx import Presentation
from pptx.util import Inches, Pt
import requests
import urllib.parse
import traceback
from branca.element import Template, MacroElement
import os
import tempfile

# MUST BE FIRST Streamlit call
st.set_page_config(page_title="Closest Centres Map", layout="wide")

# Stronger CSS to hide Streamlit share/github/feedback buttons + menu/footer
st.markdown("""
    <style>
    #MainMenu {visibility: hidden !important; display: none !important;}
    footer {visibility: hidden !important; display: none !important;}
    .stShareWidget, button[title="Share"], button[title="Feedback"] {
        display: none !important;
    }
    a[href*="github.com"] {
        display: none !important;
    }
    </style>

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

# --- Area Type Inference ---
def infer_area_type(location):
    components = location.get("components", {})
    formatted_str = location.get("formatted", "").lower()
    big_cities_keywords = [
        "new york", "los angeles", "chicago", "houston", "phoenix", "philadelphia",
        "san antonio", "san diego", "dallas", "san jose", "austin", "jacksonville",
        "fort worth", "columbus", "charlotte", "san francisco", "indianapolis",
        "seattle", "denver", "washington", "boston", "el paso", "detroit",
        "nashville", "memphis", "portland", "oklahoma city", "las vegas", "louisville",
        "baltimore", "milwaukee", "albuquerque", "tucson", "fresno", "sacramento",
        "mesa", "kansas city", "atlanta", "long beach", "colorado springs", "raleigh",
        "miami", "virginia beach", "oakland", "minneapolis", "tulsa", "arlington",
        "new orleans", "wichita", "cleveland", "tampa", "bakersfield", "aurora",
        "honolulu", "anaheim", "santa ana", "corpus christi", "riverside", "lexington",
        "stockton", "henderson", "saint paul", "st. louis", "cincinnati", "pittsburgh",
        "greensboro", "anchorage", "plano", "lincoln", "orlando", "irvine",
        "toledo", "jersey city", "chula vista", "durham", "fort wayne", "st. petersburg",
        "laredo", "buffalo", "madison", "lubbock", "chandler", "scottsdale",
        "glendale", "reno", "norfolk", "winston–salem", "north las vegas", "irving",
        "chesapeake", "gilbert", "hialeah", "garland", "fremont", "richmond",
        "boise", "baton rouge",
        # Canada major cities
        "toronto", "montreal", "vancouver", "calgary", "ottawa", "edmonton",
        "mississauga", "winnipeg", "queens", "hamilton", "kitchener", "london",
        "victoria", "halifax", "oshawa", "windsor", "saskatoon", "regina", "st. john's",
        # Mexico major cities
        "mexico city", "guadalajara", "monterrey", "puebla", "tijuana", "leon",
        "mexicali", "culiacan", "queretaro", "san luis potosi", "toluca", "morelia",
        # Central and South America major cities
        "buenos aires", "rio de janeiro", "sao paulo", "bogota", "lima", "santiago",
        "caracas", "quito", "montevideo", "asuncion", "guayaquil", "cali",
    ]
    if any(city in formatted_str for city in big_cities_keywords):
        return "CBD"
    if "suburb" in components:
        return "Suburb"
    if "city" in components or "city_district" in components:
        return "CBD"
    if any(key in components for key in ["village", "hamlet", "town"]):
        return "Rural"
    return "Suburb"

# --- Minimize Padding ---
st.markdown("""
    <style>
    div.block-container {padding-top: 1rem; padding-bottom: 1rem;}
    .distance-text {margin-top: -35px !important;}
    </style>
    """, unsafe_allow_html=True)

# --- MAIN APP ---
st.title("\U0001F4CD Find 5 Closest Centres")
api_key = "edd4cb8a639240daa178b4c6321a60e6"
input_address = st.text_input("Enter an address:")

if input_address:
    try:
        encoded_address = urllib.parse.quote(input_address)
        url = f"https://api.opencagedata.com/geocode/v1/json?q={encoded_address}&key={api_key}"
        response = requests.get(url)
        data = response.json()

        if response.status_code != 200:
            st.error(f"\u274C API Error: {response.status_code}. Try again.")
        elif not data.get('results'):
            st.error("\u274C Address not found. Try again.")
        else:
            location = data['results'][0]
            input_coords = (location['geometry']['lat'], location['geometry']['lng'])
            area_type = infer_area_type(location)
            st.write(f"Area type detected: **{area_type}**")

            file_path = "Database IC.xlsx"
            sheets = ["Comps", "Active Centre", "Centre Opened"]
            all_data = []
            for sheet in sheets:
                df = pd.read_excel(file_path, sheet_name=sheet, engine="openpyxl")
                df["Source Sheet"] = sheet
                all_data.append(df)
            data = pd.concat(all_data).dropna(subset=["Latitude", "Longitude"]).drop_duplicates(subset=["Centre Number"])
            for col in ["City", "State", "Zipcode"]:
                if col not in data.columns:
                    data[col] = ""

            data["Distance (miles)"] = data.apply(
                lambda row: geodesic(input_coords, (row["Latitude"], row["Longitude"])).miles, axis=1)
            data_sorted = data.sort_values("Distance (miles)").reset_index(drop=True)

            selected_centres = []
            seen_distances, seen_centre_numbers = [], set()
            for _, row in data_sorted.iterrows():
                d = row["Distance (miles)"]
                if row["Centre Number"] not in seen_centre_numbers and all(abs(d - x) >= 0.005 for x in seen_distances):
                    selected_centres.append(row)
                    seen_centre_numbers.add(row["Centre Number"])
                    seen_distances.append(d)
                if len(selected_centres) == 5:
                    break
            closest = pd.DataFrame(selected_centres)

            # Folium map with zoom controls and scale bar
            m = folium.Map(
                location=input_coords,
                zoom_start=14,
                zoom_control=True,
                control_scale=True,
                scrollWheelZoom=True
            )
            folium.Marker(location=input_coords, popup=f"Your Address: {input_address}", icon=folium.Icon(color="green")).add_to(m)

            def get_marker_color(ftype):
                return {"Regus": "blue", "HQ": "darkblue", "Signature": "purple", "Spaces": "black", "Non-Standard Brand": "gold"}.get(ftype, "red")

            distance_text = ""
            for _, row in closest.iterrows():
                dest_coords = (row["Latitude"], row["Longitude"])
                folium.PolyLine([input_coords, dest_coords], color="blue", weight=2.5).add_to(m)
                color = get_marker_color(row["Format - Type of Centre"])
                label = f"#{int(row['Centre Number'])} - ({row['Distance (miles)']:.2f} mi)"
                folium.Marker(location=dest_coords,
                              popup=(f"#{int(row['Centre Number'])} - {row['Addresses']} | {row.get('City', '')}, {row.get('State', '')} {row.get('Zipcode', '')} | {row['Format - Type of Centre']} | {row['Transaction Milestone Status']} | {row['Distance (miles)']:.2f} mi"),
                              tooltip=folium.Tooltip(f"<div style='font-size:16px;font-weight:bold'>{label}</div>", permanent=True, direction='right'),
                              icon=folium.Icon(color=color)).add_to(m)
                distance_text += f"Centre #{int(row['Centre Number'])} - {row['Addresses']}, {row.get('City', '')}, {row.get('State', '')} {row.get('Zipcode', '')} - Format: {row['Format - Type of Centre']} - Milestone: {row['Transaction Milestone Status']} - {row['Distance (miles)']:.2f} miles\n"

            radius_miles = {"CBD": 1, "Suburb": 5, "Rural": 10}
            radius_meters = radius_miles.get(area_type, 5) * 1609.34
            folium.Circle(location=input_coords, radius=radius_meters, color="green", fill=True, fill_opacity=0.2).add_to(m)

            legend_template = f"""
                {{% macro html(this, kwargs) %}}
                <div style='position: absolute; top: 10px; left: 10px; width: 170px; z-index: 9999;
                            background-color: white; padding: 10px; border: 2px solid gray;
                            border-radius: 5px; font-size: 14px;'>
                    <b>Radius</b><br>
                    <span style='color:green;'>&#x25CF;</span> {radius_miles.get(area_type, 5)}-mile Zone
                </div>
                {{% endmacro %}}
            """
            legend = MacroElement()
            legend._template = Template(legend_template)
            m.get_root().add_child(legend)

            col1, col2 = st.columns([5, 2])
            with col1:
                st_folium(m, width=950, height=650)
                styled_text = f"""
                <div class='distance-text' style='font-size:16px; line-height:1.4; padding: 0; margin-top: -35px;'>
                  <b>{distance_text.replace(chr(10), '<br>')}</b>
                </div>
                """
                st.markdown(styled_text, unsafe_allow_html=True)

            with col2:
                st.markdown(f"""<div style="background-color: white; padding: 10px; border: 2px solid grey;
                                    border-radius: 10px; width: 100%; margin-top: 20px;">
                                    <b>Centre Type Legend</b><br>
                                    <i style="background-color: lightgreen; padding: 5px;">&#9724;</i> Proposed Address<br>
                                    <i style="background-color: lightblue; padding: 5px;">&#9724;</i> Regus<br>
                                    <i style="background-color: darkblue; padding: 5px;">&#9724;</i> HQ<br>
                                    <i style="background-color: purple; padding: 5px;">&#9724;</i> Signature<br>
                                    <i style="background-color: black; padding: 5px;">&#9724;</i> Spaces<br>
                                    <i style="background-color: gold; padding: 5px;">&#9724;</i> Non-Standard Brand
                                </div>""", unsafe_allow_html=True)

            # Upload image moved below map and data
            uploaded_image = st.file_uploader("\U0001F5BC\ufe0f Optional: Upload Map Screenshot for PowerPoint", type=["png", "jpg", "jpeg"])

            # --- PowerPoint Export ---
            if st.button("\U0001F4E4 Export to PowerPoint"):
                try:
                    prs = Presentation()
                    slide_layout = prs.slide_layouts[5]

                    slide = prs.slides.add_slide(slide_layout)
                    title = slide.shapes.title
                    title.text = f"5 Closest Centres to:\n{input_address}"

                    if uploaded_image:
                        image_path = os.path.join(tempfile.gettempdir(), uploaded_image.name)
                        with open(image_path, "wb") as img_file:
                            img_file.write(uploaded_image.read())
                        slide.shapes.add_picture(image_path, Inches(0.5), Inches(1.5), height=Inches(3.5))
                        top_text = Inches(5.2)
                    else:
                        top_text = Inches(1.5)

                    def add_centres_to_slide_table(centres_subset, title_text=None):
                        slide = prs.slides.add_slide(slide_layout)
                        if title_text:
                            slide.shapes.title.text = title_text

                        rows = len(centres_subset) + 1  # header + data rows
                        cols = 6
                        left = Inches(0.5)
                        top = Inches(1)
                        width = Inches(9)
                        height = Inches(0.8 + 0.4 * rows)  # adapt height

                        table = slide.shapes.add_table(rows, cols, left, top, width, height).table

                        # Set column widths (adjust as needed)
                        table.columns[0].width = Inches(1)     # Centre #
                        table.columns[1].width = Inches(3)     # Address
                        table.columns[2].width = Inches(2)     # City, State, Zip
                        table.columns[3].width = Inches(1.5)   # Format
                        table.columns[4].width = Inches(1.5)   # Milestone
                        table.columns[5].width = Inches(1)     # Distance

                        # Set header row
                        headers = ["Centre #", "Address", "City, State, Zip", "Format", "Milestone", "Distance (miles)"]
                        for col_idx, header_text in enumerate(headers):
                            cell = table.cell(0, col_idx)
                            cell.text = header_text
                            for paragraph in cell.text_frame.paragraphs:
                                paragraph.font.bold = True
                                paragraph.font.size = Pt(14)

                        # Fill data rows
                        for i, row in enumerate(centres_subset, start=1):
                            table.cell(i, 0).text = str(int(row["Centre Number"]))
                            table.cell(i, 1).text = row["Addresses"]
                            city_state_zip = f"{row.get('City', '')}, {row.get('State', '')} {row.get('Zipcode', '')}".strip(", ")
                            table.cell(i, 2).text = city_state_zip
                            table.cell(i, 3).text = row["Format - Type of Centre"]
                            table.cell(i, 4).text = row["Transaction Milestone Status"]
                            table.cell(i, 5).text = f"{row['Distance (miles)']:.2f}"

                            # Set font size for all cells
                            for col_idx in range(cols):
                                cell = table.cell(i, col_idx)
                                for paragraph in cell.text_frame.paragraphs:
                                    paragraph.font.size = Pt(12)

                    centre_rows = closest.to_dict(orient="records")
                    for i in range(0, len(centre_rows), 4):
                        group = centre_rows[i:i+4]
                        if i == 0 and not uploaded_image:
                            add_centres_to_slide_table(group, title_text=f"5 Closest Centres to:\n{input_address}")
                        else:
                            add_centres_to_slide_table(group)

                    pptx_path = os.path.join(tempfile.gettempdir(), "ClosestCentres.pptx")
                    prs.save(pptx_path)

                    with open(pptx_path, "rb") as f:
                        st.download_button("\u2B07\uFE0F Download PowerPoint", f, file_name="ClosestCentres.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

                except Exception as pptx_error:
                    st.error("\u274C PowerPoint export failed.")
                    st.text(str(pptx_error))

    except Exception as e:
        st.error("An error occurred:")
        st.text(str(e))
        st.text(traceback.format_exc())
else:
    st.info("Please enter an address above to get started.")
