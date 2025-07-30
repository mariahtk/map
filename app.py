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

                def has_valid_address(val):
                    return False if pd.isna(val) or (isinstance(val, str) and val.strip() == "") else True
                dupe_centre_nums = combined_data["Centre Number"][combined_data["Centre Number"].duplicated(keep=False)].unique()
                condition = combined_data["Centre Number"].isin(dupe_centre_nums) & (~combined_data["Addresses"].apply(has_valid_address))
                combined_data = combined_data[~condition]

                priority_order = {"Comps":0,"Active Centre":1,"Centre Opened":2}
                combined_data["Sheet Priority"] = combined_data["Source Sheet"].map(priority_order)
                data = combined_data.sort_values(by="Sheet Priority").drop_duplicates(subset=["Centre Number"],keep="first").drop(columns=["Sheet Priority"])

                active_centre_df = pd.read_excel(file_path, sheet_name="Active Centre", engine="openpyxl")
                active_centre_df["Centre Number"] = active_centre_df["Centre Number"].apply(normalize_centre_number)
                active_status_map = active_centre_df.dropna(subset=["Centre Number","Transaction Milestone Status"]).set_index("Centre Number")["Transaction Milestone Status"].to_dict()
                def replace_transaction_status(row):
                    return active_status_map[row["Centre Number"]] if row["Centre Number"] in active_status_map else row["Transaction Milestone Status"]
                data["Transaction Milestone Status"] = data.apply(replace_transaction_status, axis=1)
                data = filter_duplicates(data)
                for col in ["City","State","Zipcode"]:
                    if col not in data.columns: data[col] = ""
                data["Distance (miles)"] = data.apply(lambda row: geodesic(input_coords, (row["Latitude"], row["Longitude"])).miles, axis=1)
                data_sorted = data.sort_values("Distance (miles)").reset_index(drop=True)

                selected_centres, seen_distances, seen_centre_numbers = [], [], set()
                for _, row in data_sorted.iterrows():
                    d = row["Distance (miles)"]
                    centre_num = row["Centre Number"]
                    if centre_num in seen_centre_numbers: continue
                    if all(abs(d-x) >= 0.005 for x in seen_distances):
                        selected_centres.append(row)
                        seen_centre_numbers.add(centre_num)
                        seen_distances.append(d)
                    if len(selected_centres) == 5: break
                closest = pd.DataFrame(selected_centres)

                # Map with default zoom control
                m = folium.Map(location=input_coords, zoom_start=14, zoom_control=True, control_scale=True)
                folium.Marker(location=input_coords, popup=f"Your Address: {input_address}", icon=folium.Icon(color="green")).add_to(m)
                def get_marker_color(ftype):
                    return {"Regus":"blue","HQ":"darkblue","Signature":"purple","Spaces":"black","Non-Standard Brand":"gold"}.get(ftype,"red")
                distance_text = ""
                for _, row in closest.iterrows():
                    dest_coords = (row["Latitude"],row["Longitude"])
                    folium.PolyLine([input_coords,dest_coords], color="blue", weight=2.5).add_to(m)
                    color = get_marker_color(row["Format - Type of Centre"])
                    label = f"#{int(row['Centre Number'])} - ({row['Distance (miles)']:.2f} mi)"
                    folium.Marker(location=dest_coords,
                                  popup=(f"#{int(row['Centre Number'])} - {row['Addresses']} | {row.get('City','')}, {row.get('State','')} {row.get('Zipcode','')} | {row['Format - Type of Centre']} | {row['Transaction Milestone Status']} | {row['Distance (miles)']:.2f} mi"),
                                  tooltip=folium.Tooltip(f"<div style='font-size:16px;font-weight:bold'>{label}</div>", permanent=True, direction='right'),
                                  icon=folium.Icon(color=color)).add_to(m)
                    distance_text += f"Centre #{int(row['Centre Number'])} - {row['Addresses']}, {row.get('City','')}, {row.get('State','')} {row.get('Zipcode','')} - Format: {row['Format - Type of Centre']} - Milestone: {row['Transaction Milestone Status']} - {row['Distance (miles)']:.2f} miles\n"

                radius_miles = {"CBD":1,"Suburb":5,"Rural":10}
                radius_meters = radius_miles.get(area_type,5) * 1609.34
                folium.Circle(location=input_coords, radius=radius_meters, color="green", fill=True, fill_opacity=0.2).add_to(m)

                legend_template = f"""
                    {{% macro html(this, kwargs) %}}
                    <div style='position: absolute; top: 10px; left: 10px; width: 170px; z-index: 9999;
                                background-color: white; padding: 10px; border: 2px solid gray;
                                border-radius: 5px; font-size: 14px;'>
                        <b>Radius</b><br>
                        <span style='color:green;'>&#x25CF;</span> {radius_miles.get(area_type,5)}-mile Zone
                    </div>
                    {{% endmacro %}}
                """
                legend = MacroElement()
                legend._template = Template(legend_template)
                m.get_root().add_child(legend)

                # --- Force Zoom Control Below Legend ---
                custom_zoom_js = """
                <script>
                document.addEventListener("DOMContentLoaded", function() {
                    var zoomControl = document.querySelector(".leaflet-control-zoom");
                    if (zoomControl) {
                        zoomControl.style.position = "absolute";
                        zoomControl.style.top = "90px";   // just below legend
                        zoomControl.style.left = "10px";  // align with legend
                        zoomControl.style.zIndex = "1000";
                        zoomControl.style.display = "block";
                        zoomControl.style.visibility = "visible";
                        zoomControl.style.opacity = "1";
                    }
                });
                </script>
                """
                m.get_root().html.add_child(Element(custom_zoom_js))

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

                uploaded_image = st.file_uploader("\U0001F5BC\ufe0f Optional: Upload Map Screenshot for PowerPoint", type=["png","jpg","jpeg"])

                # --- POWERPOINT EXPORT FEATURE ---
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
                                    p.font.size = Pt(10)

                            for r_idx, (_, row) in enumerate(centres_subset.iterrows(), start=1):
                                table.cell(r_idx, 0).text = str(row["Centre Number"])
                                table.cell(r_idx, 1).text = row["Addresses"]
                                table.cell(r_idx, 2).text = f"{row.get('City','')}, {row.get('State','')} {row.get('Zipcode','')}"
                                table.cell(r_idx, 3).text = row["Format - Type of Centre"]
                                table.cell(r_idx, 4).text = row["Transaction Milestone Status"]
                                table.cell(r_idx, 5).text = f"{row['Distance (miles)']:.2f}"

                        add_centres_to_slide_table(closest, "Closest Centres Summary")
                        pptx_path = os.path.join(tempfile.gettempdir(), "closest_centres.pptx")
                        prs.save(pptx_path)

                        with open(pptx_path, "rb") as file:
                            st.download_button(label="Download PowerPoint", data=file, file_name="closest_centres.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
                        st.success("PowerPoint exported successfully!")
                    except Exception as e:
                        st.error(f"Error exporting PowerPoint: {str(e)}")
                        traceback.print_exc()

