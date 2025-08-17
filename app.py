import pandas as pd
from geopy.distance import geodesic
import streamlit as st
import folium
from streamlit_folium import st_folium
import requests
import urllib.parse
from branca.element import Template, MacroElement
import streamlit.components.v1 as components
import base64
import json
import os

st.set_page_config(page_title="Closest Centres Map", layout="wide")

# Hide Streamlit UI elements
st.markdown("""
    <style>
    #MainMenu, footer, header, a[href*="github.com"], .viewerBadge_container__1QSob, .stAppViewerBadge {
        display: none !important;
    }
    </style>
""", unsafe_allow_html=True)

components.html("""
<script>
setInterval(() => {
  const badges = document.querySelectorAll(
    '.viewerBadge_container__1QSob, .stAppViewerBadge, a[href*="github.com"]');
  badges.forEach(badge => badge.style.display = 'none');
}, 500);
</script>
""", height=0)

# --- Login system ---
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
            st.rerun()
        else:
            st.error("Invalid email or password.")

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    login()
    st.stop()

# --- Helper functions ---
def infer_area_type(location):
    formatted_str = location.get("formatted", "").lower()
    cbd_keywords = ["new york","manhattan","brooklyn","queens","bronx","staten island",
                    "chicago","los angeles","san francisco","boston","washington","philadelphia",
                    "houston","seattle","miami","atlanta","dallas","phoenix","detroit",
                    "san diego","minneapolis","denver","austin","portland","nashville",
                    "new orleans","las vegas","toronto","montreal","vancouver","ottawa",
                    "calgary","edmonton","winnipeg","halifax","victoria","quebec city",
                    "mexico city","guadalajara","monterrey","tijuana"]
    suburb_keywords = ["westmount","laval","longueuil","brossard","côte-saint-luc","ndg",
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
                      "zapopan","tlajomulco","santa catarina","san nicolas de los garza"]
    if any(city in formatted_str for city in cbd_keywords):
        return "CBD"
    elif any(nhood in formatted_str for nhood in suburb_keywords):
        return "Suburb"
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

@st.cache_data
def load_data(file_path="Database IC.xlsx"):
    sheets = ["Comps", "Active Centre", "Centre Opened"]
    all_data = []
    for sheet in sheets:
        df = pd.read_excel(file_path, sheet_name=sheet, engine="openpyxl")
        df["Centre Number"] = df["Centre Number"].apply(normalize_centre_number)
        if sheet in ["Active Centre", "Centre Opened"]:
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
    priority_order = {"Comps": 0, "Active Centre": 1, "Centre Opened": 2}
    combined_data["Sheet Priority"] = combined_data["Source Sheet"].map(priority_order)
    data = combined_data.sort_values(by="Sheet Priority").drop_duplicates(subset=["Centre Number"], keep="first").drop(columns=["Sheet Priority"])
    active_centre_df = pd.read_excel(file_path, sheet_name="Active Centre", engine="openpyxl")
    active_centre_df["Centre Number"] = active_centre_df["Centre Number"].apply(normalize_centre_number)
    active_status_map = active_centre_df.dropna(subset=["Centre Number","Transaction Milestone Status"]).set_index("Centre Number")["Transaction Milestone Status"].to_dict()
    def replace_transaction_status(row):
        return active_status_map[row["Centre Number"]] if row["Centre Number"] in active_status_map else row["Transaction Milestone Status"]
    data["Transaction Milestone Status"] = data.apply(replace_transaction_status, axis=1)
    data = filter_duplicates(data)
    for col in ["City","State","Zipcode"]:
        if col not in data.columns:
            data[col] = ""
    return data

# --- Read and embed logo as Base64 for PDF ---
logo_b64 = ""
logo_path_options = ["IWG Logo.jpg", "IWG Logo.png", "IWG_Logo.jpg", "IWG_Logo.png"]
for p in logo_path_options:
    if os.path.exists(p):
        with open(p, "rb") as f:
            logo_b64 = base64.b64encode(f.read()).decode("utf-8")
        logo_ext = p.split(".")[-1].upper()
        logo_type = "PNG" if logo_ext == "PNG" else "JPEG"
        break
else:
    logo_type = "JPEG"  # default if not found; PDF will just omit logo

st.title("\U0001F4CD Find 5 Closest Centres")
api_key = "edd4cb8a639240daa178b4c6321a60e6"
input_address = st.text_input("Enter an address:")

if input_address:
    try:
        with st.spinner("Loading, please wait..."):
            encoded_address = urllib.parse.quote(input_address)
            url = f"https://api.opencagedata.com/geocode/v1/json?q={encoded_address}&key={api_key}"
            response = requests.get(url)
            data_geo = response.json()
            if response.status_code != 200:
                st.error(f"\u274C API Error: {response.status_code}. Try again.")
            elif not data_geo.get('results'):
                st.error("\u274C Address not found. Try again.")
            else:
                location = data_geo['results'][0]
                input_coords = (location['geometry']['lat'], location['geometry']['lng'])
                area_type = infer_area_type(location)
                st.write(f"Area type detected: **{area_type}**")

                data = load_data()
                data["Distance (miles)"] = data.apply(lambda row: geodesic(input_coords,(row["Latitude"],row["Longitude"])).miles, axis=1)
                data_sorted = data.sort_values("Distance (miles)").reset_index(drop=True)

                selected_centres, seen_distances, seen_centre_numbers = [], [], set()
                for _, row in data_sorted.iterrows():
                    d = row["Distance (miles)"]
                    centre_num = row["Centre Number"]
                    if centre_num in seen_centre_numbers:
                        continue
                    if all(abs(d-x)>=0.005 for x in seen_distances):
                        selected_centres.append(row)
                        seen_centre_numbers.add(centre_num)
                        seen_distances.append(d)
                    if len(selected_centres) == 5:
                        break
                closest = pd.DataFrame(selected_centres)

                m = folium.Map(location=input_coords, zoom_start=14, zoom_control=True, control_scale=True)
                folium.Marker(location=input_coords, popup=f"Your Address: {input_address}", icon=folium.Icon(color="green")).add_to(m)

                def get_marker_color(ftype):
                    return {"Regus":"blue","HQ":"darkblue","Signature":"purple","Spaces":"black","Non-Standard Brand":"gold"}.get(ftype,"red")

                distance_text = ""
                label_positions = []

                for idx, row in closest.iterrows():
                    dest_coords = (row["Latitude"], row["Longitude"])
                    label_positions.append({"lat":dest_coords[0],"lng":dest_coords[1]})
                    folium.PolyLine([input_coords,dest_coords], color="blue", weight=2.5).add_to(m)
                    color = get_marker_color(row["Format - Type of Centre"])
                    label = f"#{int(row['Centre Number'])} - ({row['Distance (miles)']:.2f} mi)"

                    # Original marker icon
                    folium.Marker(
                        location=dest_coords,
                        icon=folium.Icon(color=color),
                        popup=f"#{int(row['Centre Number'])} - {row['Addresses']}, {row.get('City','')} {row.get('State','')} {row.get('Zipcode','')} | {row['Format - Type of Centre']} | {row['Transaction Milestone Status']} | {row['Distance (miles)']:.2f} mi"
                    ).add_to(m)

                    # Draggable label beside marker
                    html_label = f"""
                    <div style="
                        background-color:white;
                        padding:4px 8px;
                        border:1px solid gray;
                        border-radius:4px;
                        font-weight:bold;
                        color:black;
                        white-space:nowrap;
                        font-size:14px;
                        display:inline-block;
                        text-align:left;
                        ">
                        {label}
                    </div>
                    """
                    icon = folium.DivIcon(html=html_label)
                    folium.Marker(location=(dest_coords[0]+0.00005,dest_coords[1]+0.00005), icon=icon, draggable=True).add_to(m)

                    distance_text += (
                        f"Centre #{int(row['Centre Number'])} - {row['Addresses']}, "
                        f"{row.get('City','')}, {row.get('State','')} {row.get('Zipcode','')} - "
                        f"Format: {row['Format - Type of Centre']} - "
                        f"Milestone: {row['Transaction Milestone Status']} - "
                        f"{row['Distance (miles)']:.2f} miles\n"
                    )

                radius_miles = {"CBD":1,"Suburb":5,"Rural":10}
                radius_m = radius_miles.get(area_type,5)*1609.34
                folium.Circle(location=input_coords,radius=radius_m,color="green",fill=True,fill_opacity=0.2).add_to(m)

                # Radius legend
                legend_html = f"""
                    {{% macro html(this, kwargs) %}}
                    <div style='position:absolute;top:70px;left:10px;width:180px;z-index:9999;
                                background-color:white;padding:10px;border:2px solid gray;
                                border-radius:5px;font-size:14px;color:black;text-shadow:1px 1px 2px white;'>
                        <b>Radius</b><br>
                        <span style='color:green;'>&#x25CF;</span> {radius_miles.get(area_type,5)}-mile Zone
                    </div>
                    {{% endmacro %}}
                """
                legend = MacroElement()
                legend._template = Template(legend_html)
                m.get_root().add_child(legend)

                col1,col2 = st.columns([5,2])
                with col1:
                    st_folium(m, width=950, height=650)
                    st.markdown(
                        f"<div style='font-size:18px;line-height:1.5;font-weight:bold;padding-top:8px;'>"
                        f"{distance_text.replace(chr(10),'<br>')}</div>",
                        unsafe_allow_html=True
                    )

                    # ---------- Download as PDF Button (Map + Address + Centres) ----------
                    # Prepare safe JS strings
                    details_js = json.dumps(distance_text)  # preserves newlines safely
                    address_js = json.dumps(input_address)
                    logo_data_url = ""
                    if logo_b64:
                        logo_data_url = f"data:image/{'png' if logo_type=='PNG' else 'jpeg'};base64,{logo_b64}"
                    logo_js = json.dumps(logo_data_url)

                    components.html(f"""
                        <button id="downloadPdfBtn" style="
                            margin-top:15px;
                            padding:10px 20px;
                            background-color: #007BFF;
                            color: white;
                            border: none;
                            border-radius: 8px;
                            font-size: 16px;
                            cursor: pointer;
                        ">📄 Download Map + Centres as PDF</button>

                        <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
                        <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
                        <script>
                        const LOGO_DATA_URL = {logo_js};
                        const DETAILS_TEXT = {details_js};
                        const ADDRESS_TEXT = {address_js};

                        function findFoliumIframe() {{
                            // Try specific srcdoc iframe first
                            let ifr = window.parent.document.querySelector('iframe[srcdoc]');
                            if (ifr) return ifr;
                            // Fallback: last iframe on the page
                            const all = window.parent.document.querySelectorAll('iframe');
                            return all.length ? all[all.length - 1] : null;
                        }}

                        async function captureMapCanvas() {{
                            const ifr = findFoliumIframe();
                            if (!ifr) throw new Error("Map iframe not found.");

                            // Access the iframe's document (srcdoc, same-origin)
                            const doc = ifr.contentDocument || ifr.contentWindow.document;
                            const target = doc.body;

                            // Give tiles a moment to load
                            await new Promise(r => setTimeout(r, 800));

                            // Capture using html2canvas
                            return await html2canvas(target, {{useCORS: true}});
                        }}

                        document.getElementById("downloadPdfBtn").onclick = async function() {{
                            try {{
                                const canvas = await captureMapCanvas();
                                const imgData = canvas.toDataURL("image/png");

                                const {{ jsPDF }} = window.jspdf;
                                const pdf = new jsPDF("p", "mm", "a4");

                                // Margins and layout
                                const marginX = 15;
                                let cursorY = 20;

                                // Add Logo if available
                                if (LOGO_DATA_URL) {{
                                    try {{
                                        // Fit logo within width 40mm, keep aspect ratio (assume ~1:1 to keep simple)
                                        pdf.addImage(LOGO_DATA_URL, "{'PNG' if logo_type=='PNG' else 'JPEG'}", marginX, cursorY - 10, 40, 20);
                                    }} catch(e) {{}}
                                }}

                                // Title
                                pdf.setFontSize(16);
                                pdf.text("Closest Centres Report", marginX, cursorY);
                                cursorY += 8;

                                // Address line
                                pdf.setFontSize(12);
                                const addrLabel = "Address: ";
                                const addr = ADDRESS_TEXT || "";
                                const dateStr = new Date().toLocaleDateString();
                                pdf.text(addrLabel + addr, marginX, cursorY);
                                cursorY += 6;
                                pdf.text("Date: " + dateStr, marginX, cursorY);
                                cursorY += 6;

                                // Map image (fit to page width)
                                const maxImgW = 180; // A4 width - margins
                                const imgW = maxImgW;
                                const imgH = canvas.height * imgW / canvas.width;
                                pdf.addImage(imgData, "PNG", marginX, cursorY, imgW, imgH);
                                cursorY += imgH + 8;

                                // Centres details with wrapping + pagination
                                const maxWidth = 180;
                                const lines = pdf.splitTextToSize(DETAILS_TEXT, maxWidth);
                                const lineHeight = 6;
                                for (let i = 0; i < lines.length; i++) {{
                                    if (cursorY > 287) {{ // page break threshold
                                        pdf.addPage();
                                        cursorY = 20;
                                    }}
                                    pdf.text(lines[i], marginX, cursorY);
                                    cursorY += lineHeight;
                                }}

                                pdf.save("closest_centres.pdf");
                            }} catch (err) {{
                                alert("Could not generate PDF: " + err.message);
                                console.error(err);
                            }}
                        }}
                        </script>
                    """, height=140)

                with col2:
                    st.markdown("""
                        <div style="
                            background-color:white;
                            padding:10px;
                            border:2px solid grey;
                            border-radius:10px;
                            width:100%;
                            margin-top:20px;
                            color:black;
                            text-shadow:1px 1px 2px white;
                            font-weight:bold;
                            font-size:14px;
                        ">
                            Centre Type Legend<br>
                            <i style="background-color: lightgreen; padding: 5px;">&#9724;</i> Proposed Address<br>
                            <i style="background-color: lightblue; padding: 5px;">&#9724;</i> Regus<br>
                            <i style="background-color: darkblue; padding: 5px;">&#9724;</i> HQ<br>
                            <i style="background-color: purple; padding: 5px;">&#9724;</i> Signature<br>
                            <i style="background-color: black; padding: 5px;">&#9724;</i> Spaces<br>
                            <i style="background-color: gold; padding: 5px;">&#9724;</i> Non-Standard Brand
                        </div>
                    """, unsafe_allow_html=True)

    except Exception as ex:
        st.error(f"Unexpected error: {ex}")
