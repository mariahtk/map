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
    .st-emotion-cache-1dp5vir {
        display: none !important;
    }
    div.block-container {
        padding-top: 1rem !important;
        padding-bottom: 1rem !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- JavaScript to remove "Manage App" ---
components.html("""
<script>
const killFloaters = () => {
    const floaters = document.querySelectorAll('div[aria-label*="Manage"], div[role="complementary"], a[href*="streamlit.app"]');
    floaters.forEach(el => {
        el.style.display = "none";
    });
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
        "new york","manhattan","brooklyn","queens","bronx","staten island",
        "los angeles","chicago","houston","phoenix","philadelphia","san antonio","san diego","dallas",
        "toronto","scarborough","etobicoke","north york","montreal","vancouver","calgary","ottawa"
    ]
    if any(city in formatted_str for city in big_cities_keywords):
        return "CBD"
    if any(key in components for key in ["village","hamlet","town"]):
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
            st.error(f"❌ API Error: {response.status_code}. Try again.")
        elif not data.get('results'):
            st.error("❌ Address not found. Try again.")
        else:
            location = data['results'][0]
            input_coords = (location['geometry']['lat'], location['geometry']['lng'])
            area_type = infer_area_type(location)
            st.write(f"Area type detected: **{area_type}**")

            # --- Load Excel data ---
            file_path = "Database IC.xlsx"
            df_comps = pd.read_excel(file_path, sheet_name="Comps", engine="openpyxl")
            df_active = pd.read_excel(file_path, sheet_name="Active Centre", engine="openpyxl")
            df_opened = pd.read_excel(file_path, sheet_name="Centre Opened", engine="openpyxl")

            # --- Normalize Centre Numbers (remove leading zeros) ---
            for df in [df_comps, df_active, df_opened]:
                df["Centre Number"] = df["Centre Number"].astype(str).str.lstrip("0").str.strip()

            # --- Keep only Centre Numbers that are present in all three tabs ---
            active_nums = set(df_active["Centre Number"])
            opened_nums = set(df_opened["Centre Number"])
            comps_nums = set(df_comps["Centre Number"])
            valid_nums = comps_nums & active_nums & opened_nums

            # --- Filter to only valid centres ---
            df_comps = df_comps[df_comps["Centre Number"].isin(valid_nums)]
            df_active = df_active[df_active["Centre Number"].isin(valid_nums)]
            df_opened = df_opened[df_opened["Centre Number"].isin(valid_nums)]

            # --- Combine, keeping Comps rows if duplicates exist ---
            df_comps["Source Sheet"] = "Comps"
            df_active["Source Sheet"] = "Active Centre"
            df_opened["Source Sheet"] = "Centre Opened"
            combined = pd.concat([df_comps, df_active, df_opened])
            priority_order = {"Comps": 0, "Active Centre": 1, "Centre Opened": 2}
            combined["Sheet Priority"] = combined["Source Sheet"].map(priority_order)
            data = (combined.sort_values("Sheet Priority")
                            .drop_duplicates(subset=["Centre Number"], keep="first")
                            .drop(columns=["Sheet Priority"]))

            # --- Replace Transaction Milestone Status with Active Centre if exists ---
            active_status_map = df_active.dropna(subset=["Centre Number","Transaction Milestone Status"]) \
                                         .set_index("Centre Number")["Transaction Milestone Status"].to_dict()
            data["Transaction Milestone Status"] = data.apply(
                lambda row: active_status_map.get(row["Centre Number"], row.get("Transaction Milestone Status", "")), axis=1)

            # --- Ensure required columns exist ---
            for col in ["City", "State", "Zipcode"]:
                if col not in data.columns:
                    data[col] = ""

            # --- Drop rows with missing coordinates ---
            data = data.dropna(subset=["Latitude", "Longitude"])

            # --- Calculate distance ---
            data["Distance (miles)"] = data.apply(
                lambda row: geodesic(input_coords, (row["Latitude"], row["Longitude"])).miles,
                axis=1
            )
            data_sorted = data.sort_values("Distance (miles)").reset_index(drop=True)

            # --- Select 5 closest unique centres ---
            selected_centres, seen_centre_numbers = [], set()
            for _, row in data_sorted.iterrows():
                if row["Centre Number"] not in seen_centre_numbers:
                    selected_centres.append(row)
                    seen_centre_numbers.add(row["Centre Number"])
                if len(selected_centres) == 5:
                    break
            closest = pd.DataFrame(selected_centres)

            # --- Folium Map ---
            m = folium.Map(location=input_coords, zoom_start=14)
            folium.Marker(location=input_coords, popup=f"Your Address: {input_address}", icon=folium.Icon(color="green")).add_to(m)

            def get_marker_color(ftype):
                return {"Regus":"blue","HQ":"darkblue","Signature":"purple","Spaces":"black","Non-Standard Brand":"gold"}.get(ftype,"red")

            distance_text = ""
            for _, row in closest.iterrows():
                dest_coords = (row["Latitude"], row["Longitude"])
                folium.PolyLine([input_coords, dest_coords], color="blue", weight=2.5).add_to(m)
                color = get_marker_color(row["Format - Type of Centre"])
                label = f"#{row['Centre Number']} - ({row['Distance (miles)']:.2f} mi)"
                folium.Marker(
                    location=dest_coords,
                    popup=(f"#{row['Centre Number']} - {row['Addresses']} | {row.get('City','')}, {row.get('State','')} {row.get('Zipcode','')} | {row['Format - Type of Centre']} | {row['Transaction Milestone Status']} | {row['Distance (miles)']:.2f} mi"),
                    tooltip=folium.Tooltip(f"<div style='font-size:16px;font-weight:bold'>{label}</div>", permanent=True, direction='right'),
                    icon=folium.Icon(color=color)
                ).add_to(m)
                distance_text += f"Centre #{row['Centre Number']} - {row['Addresses']}, {row.get('City','')}, {row.get('State','')} {row.get('Zipcode','')} - Format: {row['Format - Type of Centre']} - Milestone: {row['Transaction Milestone Status']} - {row['Distance (miles)']:.2f} miles\n"

            # Radius circle
            radius_miles = {"CBD":1,"Suburb":5,"Rural":10}
            folium.Circle(location=input_coords, radius=radius_miles.get(area_type,5)*1609.34, color="green", fill=True, fill_opacity=0.2).add_to(m)

            # Legend
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

            col1, col2 = st.columns([5,2])
            with col1:
                st_folium(m, width=950, height=650)
                st.markdown(f"<div style='font-size:20px; line-height:1.6; padding: 10px 0; font-weight: bold;'>{distance_text.replace(chr(10),'<br>')}</div>", unsafe_allow_html=True)

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

    except Exception as e:
        st.error("An error occurred:")
        st.text(str(e))
        st.text(traceback.format_exc())
else:
    st.info("Please enter an address above to get started.")
