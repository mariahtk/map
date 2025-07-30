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
from branca.element import Template, MacroElement, Element
import os
import tempfile
import streamlit.components.v1 as components
from folium import plugins

# MUST BE FIRST Streamlit call
st.set_page_config(page_title="Closest Centres Map", layout="wide")

# --- Hide Streamlit UI Chrome, Branding, "Made by Streamlit" Badge & GitHub Icon ---
st.markdown("""
    <style>
    #MainMenu {visibility: hidden !important;}
    footer {visibility: hidden !important;}
    header {visibility: hidden !important;}
    .viewerBadge_container__1QSob,
    .stAppViewerBadge,
    .st-emotion-cache-1wbqy5l,
    .st-emotion-cache-12fmjuu,
    .st-emotion-cache-1gulkj5,
    .stActionButton,
    a[href*="github.com"] {
        display: none !important;
        visibility: hidden !important;
        opacity: 0 !important;
        pointer-events: none !important;
        height: 0px !important;
        width: 0px !important;
    }
    div.block-container {
        padding-top: 1rem !important;
        padding-bottom: 1rem !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- JavaScript backup removal for dynamically injected badges & GitHub icon ---
components.html("""
<script>
const hideBadges = () => {
    const selectors = [
        'div.viewerBadge_container__1QSob',
        'div.stAppViewerBadge',
        'div.st-emotion-cache-1wbqy5l',
        'div.st-emotion-cache-12fmjuu',
        'div.st-emotion-cache-1gulkj5',
        'a[href*="streamlit.io"]',
        'a[href*="github.com"]'
    ];
    selectors.forEach(sel => {
        document.querySelectorAll(sel).forEach(el => el.style.display = "none");
    });
};
setInterval(hideBadges, 500);
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
    cbd_keywords = [
        "new york","manhattan","brooklyn","queens","bronx","staten island",
        "chicago","los angeles","san francisco","boston","washington",
        "philadelphia","houston","seattle","miami","atlanta","dallas",
        "phoenix","detroit","san diego","minneapolis","denver","austin",
        "portland","nashville","new orleans","las vegas",
        "toronto","montreal","vancouver","ottawa","calgary","edmonton",
        "winnipeg","halifax","victoria","quebec city",
        "mexico city","guadalajara","monterrey","tijuana"
    ]
    suburb_keywords = [
        "westmount","laval","longueuil","brossard","côte-saint-luc","ndg",
        "saint-laurent","west island","mississauga","brampton","markham",
        "vaughan","richmond hill","pickering","ajax","oshawa","milton",
        "oakville","burlington","burnaby","surrey","richmond bc","coquitlam",
        "delta","langley","maple ridge","north vancouver","west vancouver",
        "okotoks","airdrie","sherwood park","st. albert","gatineau","kanata",
        "orleans","cambridge","brookline","somerville","newton","quincy",
        "jersey city","hoboken","newark","yonkers","staten island","flushing",
        "long island city","bronxville","white plains","oakland","berkeley",
        "san mateo","redwood city","palo alto","pasadena","burbank",
        "santa monica","long beach","anaheim","evanston","oak park",
        "naperville","schaumburg","coral gables","hialeah","kendall","aventura",
        "zapopan","tlajomulco","santa catarina","san nicolas de los garza"
    ]
    if any(city in formatted_str for city in cbd_keywords):
        return "CBD"
    elif any(nhood in formatted_str for nhood in suburb_keywords):
        return "Suburb"
    elif any(key in components for key in ["village", "hamlet"]):
        return "Rural"
    else:
        return "Suburb"

def normalize_centre_number(val):
    if pd.isna(val): return ""
    val_str = str(val).strip()
    return val_str.lstrip("0") or "0"

def normalize_address(addr):
    if pd.isna(addr): return ""
    return addr.strip().lower()

def filter_duplicates(df):
    preferred_statuses = {
        "Under Construction","Contract Signed","IC Approved",
        "Not Paid But Contract Signed","Centre Open"
    }
    df["Normalized Address"] = df["Addresses"].apply(normalize_address)
    grouped = df.groupby(["Centre Number","Normalized Address"])
    def select_preferred(group):
        preferred = group[group["Transaction Milestone Status"].isin(preferred_statuses)]
        return preferred if not preferred.empty else group
    filtered_df = grouped.apply(select_preferred).reset_index(drop=True)
    return filtered_df.drop(columns=["Normalized Address"])

# --- MAIN APP ---
st.title("\U0001F4CD Find 5 Closest Centres")
api_key = "edd4cb8a639240daa178b4c6321a60e6"
input_address = st.text_input("Enter an address:")

if input_address:
    try:
        with st.spinner("Loading, please wait..."):
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
                    df["Centre Number"] = df["Centre Number"].apply(normalize_centre_number)
                    if sheet == "Active Centre" or sheet == "Centre Opened":
                        df["Addresses"] = df["Address Line 1"]
                    else:
                        if "Addresses" not in df.columns and "Address Line 1" in df.columns:
                            df["Addresses"] = df["Address Line 1"]
                    df["Source Sheet"] = sheet
                    all_data.append(df)

                combined_data = pd.concat(all_data, ignore_index=True)
                combined_data = combined_data.dropna(subset=["Latitude","Longitude","Centre Number"])

                data = filter_duplicates(combined_data)
                data["Distance (miles)"] = data.apply(lambda row: geodesic(input_coords, (row["Latitude"], row["Longitude"])).miles, axis=1)
                data_sorted = data.sort_values("Distance (miles)").reset_index(drop=True)
                closest = data_sorted.head(5)

                # --- Map (Zoom control enabled and moved top-right) ---
                m = folium.Map(location=input_coords, zoom_start=14, zoom_control=True, control_scale=True)
                folium.Marker(location=input_coords, popup=f"Your Address: {input_address}", icon=folium.Icon(color="green")).add_to(m)
                for _, row in closest.iterrows():
                    dest_coords = (row["Latitude"],row["Longitude"])
                    folium.PolyLine([input_coords,dest_coords], color="blue", weight=2.5).add_to(m)
                    folium.Marker(dest_coords,
                        popup=(f"#{int(row['Centre Number'])} - {row['Addresses']} | {row.get('City','')}, "
                               f"{row.get('State','')} {row.get('Zipcode','')} | "
                               f"{row['Format - Type of Centre']} | {row['Transaction Milestone Status']} | "
                               f"{row['Distance (miles)']:.2f} mi"),
                        tooltip=folium.Tooltip(f"<b>#{int(row['Centre Number'])}</b>", permanent=True),
                        icon=folium.Icon(color="blue")).add_to(m)

                # --- Move zoom buttons top-right ---
                move_zoom_js = """
                <script>
                document.addEventListener("DOMContentLoaded", function() {
                    var zc = document.querySelector(".leaflet-control-zoom");
                    if (zc) {
                        zc.style.top = "10px";
                        zc.style.left = "auto";
                        zc.style.right = "10px";
                    }
                });
                </script>
                """
                m.get_root().html.add_child(Element(move_zoom_js))
                st_folium(m, width=950, height=650)

                # --- PowerPoint Export ---
                uploaded_image = st.file_uploader("\U0001F5BC\ufe0f Optional: Upload Map Screenshot for PowerPoint", type=["png","jpg","jpeg"])
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
                                table.cell(i, 0).text = str(int(row["Centre Number"]))
                                table.cell(i, 1).text = row["Addresses"]
                                table.cell(i, 2).text = f"{row.get('City', '')}, {row.get('State', '')} {row.get('Zipcode', '')}".strip(", ")
                                table.cell(i, 3).text = row["Format - Type of Centre"]
                                table.cell(i, 4).text = row["Transaction Milestone Status"]
                                table.cell(i, 5).text = f"{row['Distance (miles)']:.2f}"
                                for col_idx in range(cols):
                                    for p in table.cell(i, col_idx).text_frame.paragraphs:
                                        p.font.size = Pt(12)
                        rows = closest.to_dict(orient="records")
                        for i in range(0, len(rows), 4):
                            add_centres_to_slide_table(rows[i:i+4])
                        pptx_path = os.path.join(tempfile.gettempdir(), "ClosestCentres.pptx")
                        prs.save(pptx_path)
                        with open(pptx_path, "rb") as f:
                            st.download_button("\u2B07\uFE0F Download PowerPoint", f, file_name="ClosestCentres.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
                    except Exception as pptx_error:
                        st.error("\u274C PowerPoint export failed.")
                        st.text(str(pptx_error))

    except Exception as e:
        st.error(f"\u274C Unexpected error: {e}")
        st.error(traceback.format_exc())
else:
    st.info("Please enter an address above to begin.")
