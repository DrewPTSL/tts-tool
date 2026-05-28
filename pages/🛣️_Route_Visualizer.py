import streamlit as st
import folium
from folium import plugins
import requests
import pandas as pd
from streamlit_folium import st_folium

st.set_page_config(page_title="Route Visualizer", page_icon="🗺️", layout="wide", initial_sidebar_state="auto")

st.sidebar.title("🛣️ Route Visualizer")

if "start_coords_val" not in st.session_state:
    st.session_state["start_coords_val"] = ""
if "end_coords_val" not in st.session_state:
    st.session_state["end_coords_val"] = ""
if "route_data" not in st.session_state:
    st.session_state["route_data"] = None
if "route_coords" not in st.session_state:
    st.session_state["route_coords"] = None

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
    url = (
        f"http://router.project-osrm.org/route/v1/driving/"
        f"{start_coords[1]},{start_coords[0]};{end_coords[1]},{end_coords[0]}"
        f"?overview=full&geometries=geojson"
    )
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        route_json = response.json()
        if route_json.get("routes"):
            route = route_json["routes"][0]
            return {
                "geometry": route["geometry"]["coordinates"],
                "distance": route["distance"],
                "duration": route["duration"],
            }
        st.error("OSRM returned no routes for these coordinates.")
        return None
    except requests.exceptions.Timeout:
        st.error("Route request timed out. Try again or check your coordinates.")
        return None
    except requests.exceptions.RequestException as e:
        st.error(f"Routing request failed: {e}")
        return None

def create_map(start_coords, end_coords, route_data):
    center_lat = (start_coords[0] + end_coords[0]) / 2
    center_lon = (start_coords[1] + end_coords[1]) / 2
    m = folium.Map(
        location=[center_lat, center_lon],
        zoom_start=12,
        tiles="CartoDB positron",
        width='100%'
    )

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

def render_zone_lookup():
    st.markdown("#### Zone Coordinates Lookup")

    data_choice = st.radio(
        "Select Data Year:",
        options=["2006 Zones", "2022 Zones"],
        horizontal=True,
        index=None
    )

    if data_choice:
        zones_df, zone_col, region_col = load_zones_data(data_choice)
    else:
        st.warning("Please select a data year")
        st.selectbox("Zone Lookup", [], disabled=True)
        return

    if data_choice == "2006 Zones":
        query_zone = st.selectbox("Zone Lookup", zones_df['GTA06'])
        zone_lat = zones_df.loc[zones_df['GTA06'] == query_zone, 'Latitude'].iloc[0]
        zone_long = zones_df.loc[zones_df['GTA06'] == query_zone, 'Longitude'].iloc[0]
    else:
        query_zone = st.selectbox("Zone Lookup", zones_df['TTS2022'])
        zone_lat = zones_df.loc[zones_df['TTS2022'] == query_zone, 'Latitude'].iloc[0]
        zone_long = zones_df.loc[zones_df['TTS2022'] == query_zone, 'Longitude'].iloc[0]

    st.write(f"Coordinates based on zone: {zone_lat}, {zone_long}")

    col_a, col_b = st.columns(2)
    with col_a:
        if st.button("Use as Start"):
            st.session_state["start_coords_val"] = f"{zone_lat}, {zone_long}"
            st.rerun()
    with col_b:
        if st.button("Use as End"):
            st.session_state["end_coords_val"] = f"{zone_lat}, {zone_long}"
            st.rerun()

def render_circle_info(map_data):
    if map_data and map_data.get("all_drawings"):
        drawings = map_data["all_drawings"]
        circles = [
            d for d in drawings
            if d["geometry"]["type"] == "Point"
            and "radius" in d.get("properties", {})
        ]
        if circles:
            latest = circles[-1]
            c_lat = latest["geometry"]["coordinates"][1]
            c_lon = latest["geometry"]["coordinates"][0]
            c_radius = latest["properties"]["radius"]
            st.session_state["circle_centre"] = (c_lat, c_lon)
            st.session_state["circle_radius_m"] = c_radius

    if "circle_centre" in st.session_state:
        c_lat, c_lon = st.session_state["circle_centre"]
        c_radius = st.session_state["circle_radius_m"]

        st.subheader("Drawn Circle")
        col1, col2, col3 = st.columns(3)
        col1.metric("Centre Lat", f"{c_lat:.5f}")
        col2.metric("Centre Lon", f"{c_lon:.5f}")
        col3.metric("Radius", f"{c_radius:.0f} m")

        if st.button("Clear Circle"):
            del st.session_state["circle_centre"]
            del st.session_state["circle_radius_m"]
            st.rerun()

def render_map():
    if st.session_state["route_data"] is None:
        return

    m = create_map(
        st.session_state["route_coords"][0],
        st.session_state["route_coords"][1],
        st.session_state["route_data"]
    )

    st.subheader("Route Map")
    with st.container():
        map_data = st_folium(m, height=600, use_container_width=True, returned_objects=["all_drawings"])
        render_circle_info(map_data)

def main():
    st.title("Route Visualizer")

    col1, col2, col3 = st.columns([5, 5, 1])
    with col1:
        start_coords_input = st.text_input(
            "Start Point Coordinates",
            value=st.session_state["start_coords_val"]
        )
    with col2:
        end_coords_input = st.text_input(
            "End Point Coordinates",
            value=st.session_state["end_coords_val"]
        )
    with col3:
        st.write("")
        st.write("")
        if st.button("🔄", help="Reverse start and end points"):
            st.session_state["start_coords_val"] = end_coords_input
            st.session_state["end_coords_val"] = start_coords_input
            st.rerun()

    render_zone_lookup()

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
                st.session_state["route_data"] = route_data
                st.session_state["route_coords"] = (start_coords, end_coords)
                st.rerun()

    render_map()

if __name__ == "__main__":
    main()