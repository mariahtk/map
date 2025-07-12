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

# MUST BE FIRST Streamlit call
st.set_page_config(page_title="Closest Centres Map", layout="wide")

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
    big_cities_keywords = ["new york", "los angeles", "toronto", "vancouver", "calgary", "mexico city", "houston"]  # shorten for brevity
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

            m = folium.Map(location=input_coords, zoom_start=14)
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
        st.error("An error occurred:")
        st.text(str(e))
        st.text(traceback.format_exc())
else:
    st.info("Please enter an address above to get started.")
