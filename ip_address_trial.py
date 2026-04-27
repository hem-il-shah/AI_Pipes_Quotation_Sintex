import streamlit as st
import requests
from streamlit_js_eval import get_geolocation

# Set page config for a cleaner look
st.set_page_config(page_title="User Insights", layout="centered")

def get_public_ip():
    try:
        # Fetches the public IP the internet sees
        response = requests.get('https://api.ipify.org?format=json', timeout=5)
        return response.json()['ip']
    except Exception:
        return "Unable to fetch Public IP"

st.title("📍 User Connection & Location")

# --- SECTION 1: IP ADDRESS ---
st.subheader("Network Information")
public_ip = get_public_ip()
st.info(f"**Your Public IP:** {public_ip}")

st.divider()

# --- SECTION 2: GEO-LOCATION (LAT/LONG) ---
st.subheader("Device Location")
st.write("Click 'Allow' in your browser to share precise coordinates.")

# This triggers the browser's native Geolocation API
location = get_geolocation()

if location:
    # Extract coordinates
    lat = location['coords']['latitude']
    lon = location['coords']['longitude']
    acc = location['coords']['accuracy']
    
    # Display coordinates in columns
    col1, col2 = st.columns(2)
    col1.metric("Latitude", f"{lat}")
    col2.metric("Longitude", f"{lon}")
    
    st.write(f"🎯 **Accuracy:** {acc} meters")
    
    # Map Display
    # Note: If you get the CSS error again here, use a hard refresh (Ctrl+F5)
    st.map(data={'lat': [lat], 'lon': [lon]})
else:
    st.warning("Waiting for location permission... (Ensure your browser location is ON)")

st.caption("Note: IP fetching works automatically, but Lat/Long requires user consent per browser security policies.")