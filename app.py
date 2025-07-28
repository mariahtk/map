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
import streamlit.components.v1 as components

# MUST BE FIRST Streamlit call
st.set_page_config(page_title="Closest Centres Map", layout="wide")

# --- Hide Streamlit UI Chrome & Branding ---
st.markdown("""
    <style>
    #MainMenu {visibility: hidden !important;}
    footer {visibility: hidden !important;}
    header {visibility: hidden !important;}
    [data-testid="stStatusWidget"] {display: none !important;}
    .stDeployButton {display: none !important;}
    iframe[src*="streamlit.io"] {display: none !important;}
    .st-emotion-cache-13ln4jf,
    .st-emotion-cache-zq5wmm,
    .st-emotion-cache-1v0mbdj,
    .st-emotion-cache-1dp5vir {display: none !important;}
    div.block-container {
        padding-top: 1rem !important;
        padding-bottom: 1rem !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- JavaScript to remove "Manage App" or other unknown floating elements ---
components.html("""
<script>
const killFloaters = () => {
    const floaters = document.querySelectorAll('div[aria-label*="Manage"], div[role="complementary"], a[href*="streamlit.app"]');
    floaters.forEach(el => { el.style.display = "none"; });
};
const interval = setInterval(() => {
    killFloaters();
    if (document.readyState === "complete") {
        clearInterval(interval);
        killFloaters();
    }
}, 500);
</script>
""", height=0)

# --- Custom IWG Support Link ---
components.html("""
<div style="position: fixed; bottom: 12px; right: 16px; z-index: 10000;
            background-color: white; padding: 8px 14px; border-radius: 8px;
            border: 1px solid #ccc; font-size: 14px; font-family: sans-serif;
            box-shadow: 0 2px 6px rgba(0,0,0,0.1);">
  💬 <a href="mailto:support@iwgplc.com" style="text-decoration: none; color: #004d99;" target="_blank">
    Contact IWG Support
  </a>
</div>
""", height=0)

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
        "new york", "manhattan", "brooklyn", "queens", "bronx", "staten island",
        "los angeles", "chicago", "houston", "phoenix", "philadelphia",
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
        "greensboro", "anchorage", "plano", "lincoln", "orlando", "irvine", "toledo",
        "jersey city", "chula vista", "durham", "fort wayne", "st. petersburg", "laredo",
        "buffalo", "madison", "lubbock", "chandler", "scottsdale", "glendale", "reno",
        "norfolk", "winston-salem", "north las vegas", "irving", "chesapeake", "gilbert",
        "hialeah", "garland", "fremont", "richmond", "boise", "baton rouge",
        "toronto", "scarborough", "etobicoke", "north york", "montreal", "vancouver",
        "calgary", "ottawa", "edmonton", "mississauga", "winnipeg", "quebec city",
        "hamilton", "kitchener", "london", "victoria", "halifax", "oshawa", "windsor",
        "saskatoon", "regina", "st. john's", "mexico city", "guadalajara", "monterrey",
        "puebla", "tijuana", "leon", "mexicali", "culiacan", "queretaro", "san luis potosi",
        "toluca", "morelia", "buenos aires", "rio de janeiro", "sao paulo", "bogota",
        "lima", "santiago", "caracas", "quito", "montevideo", "asuncion", "guayaquil", "cali"
    ]
    if any(city in formatted_str for city in big_cities_keywords):
        return "CBD"
    if any(key in components for key in ["village", "hamlet", "town"]):
        return "Rural"
    return "Suburb"

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
            combined_data = pd.concat(all_data)

            # Clean Centre Number and drop rows missing important data
            combined_data["Centre Number"] = combined_data["Centre Number"].astype(str).str.strip()
            combined_data = combined_data.dropna(subset=["Latitude", "Longitude", "Centre Number"])

            address_col = "Addresses"

            def has_valid_address(val):
                if pd.isna(val):
                    return False
                if isinstance(val, str) and val.strip() == "":
                    return False
                return True

            # --- FIX LEADING ZEROS ---
            combined_data["Centre Number"] = combined_data["Centre Number"].str.lstrip("0")

            # --- REMOVE DUPLICATES BUT KEEP "Comps" OVER OTHERS ---
            dupe_centre_nums = combined_data["Centre Number"][combined_data["Centre Number"].duplicated(keep=False)].unique()
            combined_data = combined_data[~(
                combined_data["Centre Number"].isin(dupe_centre_nums) &
                (combined_data["Source Sheet"].isin(["Active Centre", "Centre Opened"])) &
                (combined_data["Centre Number"].isin(
                    combined_data.loc[combined_data["Source Sheet"] == "Comps", "Centre Number"]
                ))
            )]

            # Remove rows with invalid addresses
            combined_data = combined_data[combined_data[address_col].apply(has_valid_address)]

            # Assign priority to sheets
            priority_order = {"Comps": 0, "Active Centre": 1, "Centre Opened": 2}
            combined_data["Sheet Priority"] = combined_data["Source Sheet"].map(priority_order)

            data = (
                combined_data
                .sort_values(by="Sheet Priority")
                .drop_duplicates(subset=["Centre Number"], keep="first")
                .drop(columns=["Sheet Priority"])
            )

            # --- Override Transaction Milestone Status with Active Centre values ---
            active_centre_df = pd.read_excel(file_path, sheet_name="Active Centre", engine="openpyxl")
            active_centre_df["Centre Number"] = active_centre_df["Centre Number"].astype(str).str.strip().str.lstrip("0")
            active_status_map = active_centre_df.dropna(subset=["Centre Number", "Transaction Milestone Status"]) \
                                               .set_index("Centre Number")["Transaction Milestone Status"].to_dict()
            def replace_transaction_status(row):
                cn = row["Centre Number"]
                return active_status_map.get(cn, row["Transaction Milestone Status"])
            data["Transaction Milestone Status"] = data.apply(replace_transaction_status, axis=1)

            # Ensure City, State, Zipcode exist
            for col in ["City", "State", "Zipcode"]:
                if col not in data.columns:
                    data[col] = ""

            # Calculate distances
            data["Distance (miles)"] = data.apply(
                lambda row: geodesic(input_coords, (row["Latitude"], row["Longitude"])).miles, axis=1)
            data_sorted = data.sort_values("Distance (miles)").reset_index(drop=True)

            selected_centres = []
            seen_distances, seen_centre_numbers = [], set()
            for _, row in data_sorted.iterrows():
                d = row["Distance (miles)"]
                centre_num = row["Centre Number"]
                if centre_num in seen_centre_numbers:
                    continue
                if all(abs(d - x) >= 0.005 for x in seen_distances):
                    selected_centres.append(row)
                    seen_centre_numbers.add(centre_num)
                    seen_distances.append(d)
                if len(selected_centres) == 5:
                    break
            closest = pd.DataFrame(selected_centres)

            # Folium map creation
            m = folium.Map(location=input_coords, zoom_start=14, zoom_control=True, control_scale=True)
            folium.Marker(location=input_coords, popup=f"Your Address: {input_address}", icon=folium.Icon(color="green")).add_to(m)

            def get_marker_color(ftype):
                return {
                    "Regus": "blue", "HQ": "darkblue", "Signature": "purple",
                    "Spaces": "black", "Non-Standard Brand": "gold"
                }.get(ftype, "red")

            distance_text = ""
            for _, row in closest.iterrows():
                dest_coords = (row["Latitude"], row["Longitude"])
                folium.PolyLine([input_coords, dest_coords], color="blue", weight=2.5).add_to(m)
                color = get_marker_color(row["Format - Type of Centre"])
                label = f"#{row['Centre Number']} - ({row['Distance (miles)']:.2f} mi)"
                folium.Marker(
                    location=dest_coords,
                    popup=(f"#{row['Centre Number']} - {row['Addresses']} | {row.get('City', '')}, {row.get('State', '')} {row.get('Zipcode', '')} | {row['Format - Type of Centre']} | {row['Transaction Milestone Status']} | {row['Distance (miles)']:.2f} mi"),
                    tooltip=folium.Tooltip(f"<div style='font-size:16px;font-weight:bold'>{label}</div>", permanent=True, direction='right'),
                    icon=folium.Icon(color=color)
                ).add_to(m)
                distance_text += f"Centre #{row['Centre Number']} - {row['Addresses']}, {row.get('City', '')}, {row.get('State', '')} {row.get('Zipcode', '')} - Format: {row['Format - Type of Centre']} - Milestone: {row['Transaction Milestone Status']} - {row['Distance (miles)']:.2f} miles\n"

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
                <div class='distance-text' style='font-size:20px; line-height:1.6; padding: 10px 0; margin-top: -20px; font-weight: bold;'>
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

            uploaded_image = st.file_uploader("\U0001F5BC\ufe0f Optional: Upload Map Screenshot for PowerPoint", type=["png", "jpg", "jpeg"])
            if st.button("\U0001F4E4 Export to PowerPoint"):
                try:
                    prs = Presentation()
                    slide_layout = prs.slide_layouts[5]
                    slide = prs.slides.add_slide(slide_layout)
                    slide.shapes.title.text = f"5 Closest Centres to:\n{input_address}"

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
                        rows = len(centres_subset) + 1
                        cols = 6
                        table = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(1), Inches(9), Inches(0.8 + 0.4 * rows)).table
                        headers = ["Centre #", "Address", "City, State, Zip", "Format", "Milestone", "Distance (miles)"]
                        for col_idx, header_text in enumerate(headers):
                            cell = table.cell(0, col_idx)
                            cell.text = header_text
                            for p in cell.text_frame.paragraphs:
                                p.font.bold = True
                                p.font.size = Pt(14)
                        for i, row in enumerate(centres_subset, start=1):
                            table.cell(i, 0).text = str(row["Centre Number"])
                            table.cell(i, 1).text = str(row["Addresses"])
                            table.cell(i, 2).text = f"{row.get('City', '')}, {row.get('State', '')} {row.get('Zipcode', '')}"
                            table.cell(i, 3).text = str(row["Format - Type of Centre"])
                            table.cell(i, 4).text = str(row["Transaction Milestone Status"])
                            table.cell(i, 5).text = f"{row['Distance (miles)']:.2f}"

                    add_centres_to_slide_table(closest.to_dict(orient="records"), title_text="Details of 5 Closest Centres")
                    pptx_path = os.path.join(tempfile.gettempdir(), "closest_centres.pptx")
                    prs.save(pptx_path)
                    with open(pptx_path, "rb") as pptx_file:
                        st.download_button("Download PowerPoint", pptx_file, file_name="closest_centres.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
                except Exception as e:
                    st.error(f"PowerPoint generation failed: {e}")
                    st.text(traceback.format_exc())
    except Exception as e:
        st.error(f"\u274C Error: {e}")
        st.text(traceback.format_exc())
