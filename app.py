import pandas as pd
from geopy.distance import geodesic
import streamlit as st
import folium
from streamlit_folium import st_folium
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

# --- Function to infer area type ---
def infer_area_type(location):
    components = location.get("components", {})
    formatted_str = location.get("formatted", "").lower()

    big_cities_keywords = [
        "new york", "los angeles", "chicago", "houston", "phoenix", "philadelphia", "san antonio", "san diego",
        "dallas", "san jose", "austin", "jacksonville", "fort worth", "columbus", "charlotte", "san francisco",
        "indianapolis", "seattle", "denver", "washington", "boston", "el paso", "nashville", "detroit",
        "oklahoma city", "portland", "las vegas", "memphis", "louisville", "baltimore", "milwaukee", "albuquerque",
        "tucson", "fresno", "sacramento", "kansas city", "long beach", "atlanta", "colorado springs", "raleigh",
        "miami", "cleveland", "minneapolis", "honolulu", "pittsburgh", "st. louis", "cincinnati", "orlando",
        "toronto", "montreal", "vancouver", "calgary", "edmonton", "ottawa", "winnipeg", "quebec city",
        "hamilton", "kitchener", "london", "victoria", "halifax", "windsor", "saskatoon", "regina", "st. john's",
        "mexico city", "guadalajara", "monterrey", "puebla", "tijuana", "le\u00f3n", "cd ju\u00e1rez", "zapopan", "toluca",
        "quer\u00e9taro", "m\u00e9rida", "chihuahua", "hermosillo", "saltillo", "cuernavaca"
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

# --- MAIN APP ---
st.set_page_config(page_title="Closest Centres Map", layout="wide")
st.title("\ud83d\udccd Find 5 Closest Centres")

api_key = "edd4cb8a639240daa178b4c6321a60e6"
input_address = st.text_input("Enter an address:")

if input_address:
    try:
        encoded_address = urllib.parse.quote(input_address)
        url = f"https://api.opencagedata.com/geocode/v1/json?q={encoded_address}&key={api_key}"
        response = requests.get(url)
        data = response.json()

        if response.status_code != 200:
            st.error(f"\u274c API Error: {response.status_code}. Try again.")
        elif not data.get('results'):
            st.error("\u274c Address not found. Try again.")
        else:
            location = data['results'][0]
            input_coords = (location['geometry']['lat'], location['geometry']['lng'])

            area_type = infer_area_type(location)
            st.write(f"Area type detected: **{area_type}**")

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

            m = folium.Map(location=input_coords, zoom_start=14)
            folium.Marker(location=input_coords, popup=f"Your Address: {input_address}", icon=folium.Icon(color="green")).add_to(m)

            def get_marker_color(format_type):
                return {
                    "Regus": "blue",
                    "HQ": "darkblue",
                    "Signature": "purple",
                    "Spaces": "black",
                    "Non-Standard Brand": "gold"
                }.get(format_type, "red" if pd.isna(format_type) or format_type == "" else "gray")

            distance_text = ""
            distance_text += "Closest Centres (Distances in miles):\n"

            for _, row in closest.iterrows():
                dest_coords = (row["Latitude"], row["Longitude"])
                folium.PolyLine([input_coords, dest_coords], color="blue", weight=2.5).add_to(m)
                marker_color = get_marker_color(row["Format - Type of Centre"])
                label_text = f"#{int(row['Centre Number'])} - ({row['Distance (miles)']:.2f} mi)"
                folium.Marker(
                    location=dest_coords,
                    popup=(f"#{int(row['Centre Number'])} - {row['Addresses']} | {row.get('City', 'N/A')}, {row.get('State', 'N/A')} {row.get('Zipcode', 'N/A')} | {row['Format - Type of Centre']} | {row['Transaction Milestone Status']} | {row['Distance (miles)']:.2f} mi"),
                    tooltip=folium.Tooltip(f"<div style='font-size: 16px; font-weight: bold;'>{label_text}</div>", permanent=True, direction='right'),
                    icon=folium.Icon(color=marker_color)
                ).add_to(m)

                distance_text += (
                    f"Centre #{int(row['Centre Number'])} - {row['Addresses']}, "
                    f"{row.get('City', 'N/A')}, {row.get('State', 'N/A')} {row.get('Zipcode', 'N/A')} - "
                    f"Format: {row['Format - Type of Centre']} - Milestone: {row['Transaction Milestone Status']} - "
                    f"{row['Distance (miles)']:.2f} miles\n"
                )

            radius_miles = {"CBD": 1, "Suburb": 5, "Rural": 10}
            radius_meters = radius_miles.get(area_type, 5) * 1609.34
            folium.Circle(location=input_coords, radius=radius_meters, color="green", fill=True, fill_opacity=0.2).add_to(m)

            legend_html = f"""
                <div style="position: absolute; top: 10px; left: 10px; width: 220px; height: auto;
                            border: 2px solid grey; z-index:9999; font-size:14px;
                            background-color:white; opacity: 0.95; padding: 10px;">
                    <b>Map Legend</b><br>
                    <span style='color:green;'>&#x25A0;</span> Your Address<br>
                    <span style='color:blue;'>&#x25A0;</span> Centre<br>
                    <span style='color:green;'>&#x25CF;</span> {radius_miles.get(area_type, 5)}-mile Radius
                </div>
            """
            m.get_root().html.add_child(folium.Element(legend_html))

            col1, col2 = st.columns([5, 2])
            with col1:
                st_folium(m, width=950, height=650)
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

            st.subheader("Distances from Your Address to the Closest Centres:")
            styled_text = f"<div style='font-size:16px; line-height:1.6;'><b>{distance_text.replace(chr(10), '<br>')}</b></div>"
            st.markdown(styled_text, unsafe_allow_html=True)

    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("Please enter an address above to get started.")
