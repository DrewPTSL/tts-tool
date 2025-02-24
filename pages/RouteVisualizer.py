import streamlit as st
import folium
from folium import plugins
import requests
import pandas as pd
from streamlit_folium import st_folium

st.set_page_config(page_title="Route Visualizer", page_icon="🗺️",layout="wide")

@st.cache_data(show_spinner="Loading zone data...")
def load_zones_data(data_choice):
    if data_choice == "2006 Zones":
        zones_df = pd.read_csv('2006Zones.csv')
        zone_col = 'gta06'
        region_col = 'region'
    else:
        zones_df = pd.read_csv('2022Zones.csv')
        zone_col = 'TTS2022'
        region_col = 'Reg_name'
    return zones_df, zone_col, region_col

def parse_coordinates(coord_string):
    try:
        lat, lon = map(float, coord_string.strip().split(','))
        return lat, lon
    except:
        return None

def get_route(start_coords, end_coords):
    url = f"http://router.project-osrm.org/route/v1/driving/{start_coords[1]},{start_coords[0]};{end_coords[1]},{end_coords[0]}?overview=full&geometries=geojson"
    
    try:
        response = requests.get(url)
        response.raise_for_status()
        route_json = response.json()
        
        if "routes" in route_json and len(route_json["routes"]) > 0:
            route = route_json["routes"][0]
            return {
                "geometry": route["geometry"]["coordinates"],
                "distance": route["distance"],
                "duration": route["duration"]
            }
        else:
            return None
    except:
        return None

def create_map(start_coords, end_coords, route_data):
    center_lat = (start_coords[0] + end_coords[0]) / 2
    center_lon = (start_coords[1] + end_coords[1]) / 2
    m = folium.Map(location=[center_lat, center_lon], zoom_start=12,width='100%')

    # Add Draw plugin with only circle drawing enabled
    draw = plugins.Draw(
        draw_options={
            'polyline': False,
            'rectangle': False,
            'polygon': False,
            'circle': True,
            'marker': False,
            'circlemarker': False
        },
        edit_options={'edit': True}
    )
    m.add_child(draw)
    
    
    folium.Marker(
        start_coords,
        popup="Start",
        icon=folium.Icon(color='black', icon='info-sign')
    ).add_to(m)
    
    folium.Marker(
        end_coords,
        popup="End",
        icon=folium.Icon(color='red', icon='info-sign')
    ).add_to(m)
    
    coordinates = [[coord[1], coord[0]] for coord in route_data["geometry"]]
    folium.PolyLine(
        coordinates,
        weight=4,
        color='blue',
        opacity=0.8
    ).add_to(m)
    
    for coord in route_data["geometry"]:
        folium.CircleMarker(
            location=[coord[1], coord[0]],
            radius=3,
            color='red',
            fill=True,
            popup=f"Node: {coord[1]:.4f}, {coord[0]:.4f}"
        ).add_to(m)
    
    
    return m

def main():
   
    st.title("Route Visualizer")
    
    col1, col2= st.columns(2)
    with col1:

        start_coords_input = st.text_input(
            "Start Point Coordinates",
            value=""
        )
    with col2:
        end_coords_input = st.text_input(
            "End Point Coordinates",
            value=""
        )

    st.markdown("#### Zone Coordinates Lookup")



    data_choice = st.radio(
    "Select Data Year:",
    options=["2006 Zones", "2022 Zones"],
    horizontal=True,
    index=None)

    if data_choice:
        zones_df, zone_col, region_col = load_zones_data(data_choice)
    else:
        st.warning("Please select a data year")


    if not data_choice:
        query_zone = st.selectbox("Zone Lookup", [], disabled=True)
    elif data_choice == "2006 Zones":
        query_zone = st.selectbox("Zone Lookup", zones_df['GTA06'])
        zone_lat = zones_df.loc[zones_df['GTA06'] == query_zone, 'Latitude'].iloc[0]
        zone_long = zones_df.loc[zones_df['GTA06'] == query_zone, 'Longitude'].iloc[0]
        st.write(f"Coordinates based on zone: {zone_lat}, {zone_long}")
    else:
        query_zone = st.selectbox("Zone Lookup", zones_df['TTS2022'])
        zone_lat = zones_df.loc[zones_df['TTS2022'] == query_zone, 'Latitude'].iloc[0]
        zone_long = zones_df.loc[zones_df['TTS2022'] == query_zone, 'Longitude'].iloc[0]
        st.write(f"Coordinates based on zone: {zone_lat}, {zone_long}")

    
    if st.button("Calculate Route", type="primary"):
        start_parsed = parse_coordinates(start_coords_input)
        end_parsed = parse_coordinates(end_coords_input)
        
        if not start_parsed or not end_parsed:
            st.error("Invalid coordinates format. Please use: latitude, longitude")
            return
        
        start_coords = list(start_parsed)
        end_coords = list(end_parsed)
        
        with st.spinner("Calculating route..."):
            route_data = get_route(start_coords, end_coords)
            
            if route_data:
                m = create_map(start_coords, end_coords, route_data)
                
                st.subheader("Route Map")
                with st.container():
                    st_folium(m, height=600, use_container_width=True, returned_objects=[])

            else:
                st.error("Unable to calculate route. Please check your coordinates and try again.")

if __name__ == "__main__":
    main()
