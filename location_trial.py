import streamlit as st
from streamlit_js_eval import get_geolocation

st.title("User Geolocation")

# This will trigger the browser's "Allow Location" prompt
location = get_geolocation()

if location:
    lat = location['coords']['latitude']
    lon = location['coords']['longitude']
    accuracy = location['coords']['accuracy']
    
    st.success(f"Location Found!")
    st.write(f"**Latitude:** {lat}")
    st.write(f"**Longitude:** {lon}")
    st.write(f"**Accuracy:** within {accuracy} meters")
    
    # Show on a map
    st.map(data={'lat': [lat], 'lon': [lon]})
else:
    st.info("Please allow location access in your browser to see coordinates.")