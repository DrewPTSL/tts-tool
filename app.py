import streamlit as st
import pandas as pd
import folium
from polyline import decode
from geopy.distance import geodesic
import re
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
import requests
from streamlit_folium import folium_static
import io
import plotly.express as px
import matplotlib.pyplot as plt

# Set page config
st.set_page_config(
    page_title="TTS Route Analysis Tool",
    page_icon="🚗",
    layout="wide"
)

# Initialize session states
if 'pois' not in st.session_state:
    st.session_state.pois = []
    
if 'poi_count' not in st.session_state:
    st.session_state.poi_count = 0
    
if 'processing_started' not in st.session_state:
    st.session_state.processing_started = False
    
if 'results_df' not in st.session_state:
    st.session_state.results_df = None

def reset_app():
    st.session_state.pois = []
    st.session_state.poi_count = 0
    st.session_state.processing_started = False
    st.session_state.results_df = None
    st.experimental_rerun()

# Title and description
st.title("TTS Route Analysis Tool")

# Add reset button in the top right
col1, col2 = st.columns([6, 1])
with col2:
    if st.button("Reset App"):
        reset_app()

with col1:
    st.markdown("Upload your TTS file and analyze routes through points of interest.")

# File Upload Section
uploaded_file = st.file_uploader("Upload your TTS file", type=['txt'])

# Configuration Section
st.markdown("### Site Configuration")
col1, col2 = st.columns(2)

with col1:
    site_zone = st.number_input(
        "Site Zone",
        min_value=1,
        value=5241,
        help="Enter the site zone number"
    )

with col2:
    coords_input = st.text_input(
        "Site Coordinates (Latitude, Longitude)",
        value="43.212342441568325, -79.77345451007199",
        help="Enter coordinates in format: latitude, longitude"
    )

try:
    site_lat, site_lon = map(float, coords_input.replace(" ", "").split(","))
    valid_coords = True
except ValueError:
    st.error("Invalid coordinates format. Please use: latitude, longitude")
    valid_coords = False

# POI Management Section
st.markdown("### Points of Interest")
st.markdown("Add or remove points of interest for route analysis.")

# Add new POI
col1, col2, col3, col4 = st.columns([2, 2, 1, 1])
with col1:
    new_poi_name = st.text_input("POI Name", key="new_poi_name")
with col2:
    new_poi_coords = st.text_input("POI Coordinates (Lat, Lon)", key="new_poi_coords")
with col3:
    new_poi_threshold = st.number_input("Threshold (m)", min_value=1, value=50, key="new_poi_threshold")
with col4:
    if st.button("Add POI"):
        try:
            lat, lon = map(float, new_poi_coords.replace(" ", "").split(","))
            new_poi = {
                'id': f'POI_{st.session_state.poi_count + 1}',
                'name': new_poi_name,
                'coordinates': (lat, lon),
                'threshold': new_poi_threshold / 1000  # Convert meters to kilometers
            }
            st.session_state.pois.append(new_poi)
            st.session_state.poi_count += 1
            st.success(f"Added POI: {new_poi_name}")
        except ValueError:
            st.error("Invalid coordinates format")
        except Exception as e:
            st.error(f"Error adding POI: {str(e)}")

# Display and manage existing POIs
if st.session_state.pois:
    st.markdown("#### Current POIs")
    for idx, poi in enumerate(st.session_state.pois):
        col1, col2, col3, col4 = st.columns([2, 2, 1, 1])
        with col1:
            st.text(poi['name'])
        with col2:
            st.text(f"({poi['coordinates'][0]}, {poi['coordinates'][1]})")
        with col3:
            st.text(f"{int(poi['threshold'] * 1000)}m")
        with col4:
            if st.button("Delete", key=f"delete_{idx}"):
                st.session_state.pois.pop(idx)
                st.experimental_rerun()

# Main Processing Section
try:
    zones_df = pd.read_csv('Zones.csv')
    
    # Validate site zone exists in zones.csv
    if not zones_df[zones_df['GTA06'] == site_zone].empty and valid_coords:
        if uploaded_file is not None and len(st.session_state.pois) > 0:
            # Add start button
            start_button = st.button("Start Processing")
            
            if start_button:
                st.session_state.processing_started = True
            
            # Only show results if processing has started
            if st.session_state.processing_started:
                # Create a progress bar and status text
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                try:
                    def get_route(origin_lat, origin_lon, dest_lat, dest_lon):
                        url = f'http://router.project-osrm.org/route/v1/driving/{origin_lon},{origin_lat};{dest_lon},{dest_lat}?overview=full'
                        response = requests.get(url)
                        if response.status_code == 200:
                            data = response.json()
                            if data['code'] == 'Ok':
                                return data['routes'][0]['geometry']
                        return None

                    def passes_through(route_geometry, poi_list, threshold=0.1):
                        route_coords = decode(route_geometry)
                        matching_pois = []
                        
                        for route_point in route_coords:
                            for poi in poi_list:
                                poi_threshold = poi.get('threshold', threshold)
                                distance = geodesic(route_point, poi['coordinates']).km
                                if distance <= poi_threshold:
                                    matching_pois.append({
                                        'id': poi['id'],
                                        'name': poi['name'],
                                        'coordinates': poi['coordinates'],
                                        'threshold': poi_threshold,
                                        'actual_distance': distance
                                    })
                        
                        # Remove duplicate POIs
                        matching_pois = [dict(t) for t in {tuple(d.items()) for d in matching_pois}]
                        
                        return {
                            'passes': bool(matching_pois),
                            'num_pois_intersected': len(matching_pois),
                            'intersected_pois': matching_pois
                        }

                    def process_tts_file(content, zones_df, progress_callback=None, status_callback=None):
                        # Extract table from text file
                        table_pattern = re.compile(r"^\s*(\d+)\s+(\d+)\s+(\d+)\s*$", re.MULTILINE)
                        matches = table_pattern.findall(content)
                        df_origins = pd.DataFrame(matches, columns=["gta06_orig", "gta06_dest", "total"])
                        df_origins = df_origins.astype({"gta06_orig": int, "gta06_dest": int, "total": int})
                        
                        # Get site zone coordinates
                        site_zone_row = zones_df[zones_df['GTA06'] == site_zone]
                        site_zone_lat = site_zone_row['Latitude'].values[0]
                        site_zone_lon = site_zone_row['Longitude'].values[0]
                        
                        results = []
                        total_rows = len(df_origins)
                        
                        for idx, row in df_origins.iterrows():
                            if progress_callback:
                                progress = (idx + 1) / total_rows * 100  # Convert to percentage
                                progress_callback(progress)
                            if status_callback:
                                status_callback(f"Processing route {idx + 1} of {total_rows}")
                                
                            origin_id = row['gta06_orig']
                            dest_id = row['gta06_dest']
                            
                            # Check if zones exist
                            origin_row = zones_df[zones_df['GTA06'] == origin_id]
                            dest_row = zones_df[zones_df['GTA06'] == dest_id]
                            
                            if origin_row.empty or dest_row.empty:
                                continue
                            
                            try:
                                # Case 1: Both origin and destination are the site zone
                                if origin_id == site_zone and dest_id == site_zone:
                                    # Route from site zone to site (origin_to_site)
                                    route1 = get_route(site_zone_lat, site_zone_lon, site_lat, site_lon)
                                    
                                    # Route from site to site zone (site_to_destination)
                                    route2 = get_route(site_lat, site_lon, site_zone_lat, site_zone_lon)
                                    
                                    # Process origin_to_site route
                                    if route1:
                                        poi_check_result = passes_through(route1, st.session_state.pois)
                                        results.append({
                                            'origin_id': origin_id, 
                                            'dest_id': dest_id,
                                            'route_type': 'origin_to_site',
                                            'passes': poi_check_result['passes'], 
                                            'num_pois_intersected': poi_check_result['num_pois_intersected'],
                                            'intersected_pois': poi_check_result['intersected_pois'],
                                            'total': row['total']
                                        })
                                    
                                    # Process site_to_destination route
                                    if route2:
                                        poi_check_result = passes_through(route2, st.session_state.pois)
                                        results.append({
                                            'origin_id': origin_id, 
                                            'dest_id': dest_id,
                                            'route_type': 'site_to_destination',
                                            'passes': poi_check_result['passes'], 
                                            'num_pois_intersected': poi_check_result['num_pois_intersected'],
                                            'intersected_pois': poi_check_result['intersected_pois'],
                                            'total': row['total']
                                        })
                                    continue
                                
                                # Case 2: Origin is not the site zone (route from origin to site)
                                if origin_id != site_zone and dest_id == site_zone:
                                    origin_lat = origin_row['Latitude'].values[0]
                                    origin_lon = origin_row['Longitude'].values[0]
                                    
                                    route = get_route(origin_lat, origin_lon, site_lat, site_lon)
                                    
                                    if route:
                                        poi_check_result = passes_through(route, st.session_state.pois)
                                        results.append({
                                            'origin_id': origin_id, 
                                            'dest_id': dest_id,
                                            'route_type': 'origin_to_site',
                                            'passes': poi_check_result['passes'], 
                                            'num_pois_intersected': poi_check_result['num_pois_intersected'],
                                            'intersected_pois': poi_check_result['intersected_pois'],
                                            'total': row['total']
                                        })
                                
                                # Case 3: Origin is the site zone (route from site to destination)
                                if origin_id == site_zone and dest_id != site_zone:
                                    dest_lat = dest_row['Latitude'].values[0]
                                    dest_lon = dest_row['Longitude'].values[0]
                                    
                                    route = get_route(site_lat, site_lon, dest_lat, dest_lon)
                                    
                                    if route:
                                        poi_check_result = passes_through(route, st.session_state.pois)
                                        results.append({
                                            'origin_id': origin_id, 
                                            'dest_id': dest_id,
                                            'route_type': 'site_to_destination',
                                            'passes': poi_check_result['passes'], 
                                            'num_pois_intersected': poi_check_result['num_pois_intersected'],
                                            'intersected_pois': poi_check_result['intersected_pois'],
                                            'total': row['total']
                                        })
                                
                            except Exception as e:
                                st.error(f"Error processing route {origin_id} to {dest_id}: {str(e)}")
                                continue
                        
                        return pd.DataFrame(results)

                    def update_progress(progress):
                        progress_bar.progress(int(progress))
                        
                    def update_status(status):
                        status_text.text(status)
                    
                    # Process the file
                    content = uploaded_file.getvalue().decode()
                    st.session_state.results_df = process_tts_file(content, zones_df, update_progress, update_status)
                    
                    if st.session_state.results_df is not None and not st.session_state.results_df.empty:
                        status_text.text("Processing complete!")
                        
                        # Process the dataframe to show POI names
                        display_df = st.session_state.results_df.copy()
                        display_df['POI'] = display_df['intersected_pois'].apply(
                            lambda x: ', '.join(sorted(set([poi['name'] for poi in x]))) if x else ''
                        )
                        
                        # Select columns to display
                        display_df = display_df[['origin_id', 'dest_id', 'route_type', 'passes', 'POI', 'total']]
                        
                        # Display summary statistics
                        col1, col2 = st.columns(2)
                        with col1:
                            st.metric("Total Routes", len(st.session_state.results_df))
                        with col2:
                            st.metric("Routes with POI Matches", st.session_state.results_df['passes'].sum())
                        
                        # Create POI summaries by route type
                        st.subheader("POI Traffic Distribution")
                        
                        # Filter for routes that pass through POIs
                        poi_df = display_df[display_df['POI'] != '']
                        
                        # Create two columns for the pie charts
                        col1, col2 = st.columns(2)
                        
                        # Split the data by route type
                        origin_to_site = poi_df[poi_df['route_type'] == 'origin_to_site']
                        site_to_destination = poi_df[poi_df['route_type'] == 'site_to_destination']
                        
                        # Create summary for origin_to_site
                        if not origin_to_site.empty:
                            with col1:
                                origin_summary = origin_to_site.groupby('POI')['total'].sum()
                                total_traffic = origin_summary.sum()
                                # Calculate percentages
                                origin_percentages = (origin_summary / total_traffic * 100).round(1)
                                
                                # Create interactive pie chart
                                fig1 = px.pie(
                                    values=origin_percentages.values,
                                    names=origin_percentages.index,
                                    custom_data=[origin_summary.values],  # For formatting
                                    title="Origin to Site"
                                )
                                fig1.update_traces(
                                    textposition='inside',
                                    hovertemplate="<b>%{label}</b><br>" +
                                                "Percentage: %{percent}<br>" +
                                                "Total Traffic: %{customdata[0]}<extra></extra>"
                                )
                                fig1.update_layout(
                                    showlegend=True,
                                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=1),
                                    height=400
                                )
                                st.plotly_chart(fig1, use_container_width=True)
                        
                        # Create summary for site_to_destination
                        if not site_to_destination.empty:
                            with col2:
                                dest_summary = site_to_destination.groupby('POI')['total'].sum()
                                total_traffic = dest_summary.sum()
                                # Calculate percentages
                                dest_percentages = (dest_summary / total_traffic * 100).round(1)
                                
                                # Create interactive pie chart
                                fig2 = px.pie(
                                    values=dest_percentages.values,
                                    names=dest_percentages.index,
                                    custom_data=[dest_summary.values],  # For formatting
                                    title="Site to Destination"
                                )
                                fig2.update_traces(
                                    textposition='inside',
                                    hovertemplate="<b>%{label}</b><br>" +
                                                "Percentage: %{percent}<br>" +
                                                "Total Traffic: %{customdata[0]}<extra></extra>"
                                )
                                fig2.update_layout(
                                    showlegend=True,
                                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=1),
                                    height=400
                                )
                                st.plotly_chart(fig2, use_container_width=True)
                        
                        # Display results in a table
                        st.subheader("Route Analysis Results")
                        st.dataframe(display_df)
                        
                        # Generate Excel file for download
                        excel_buffer = io.BytesIO()
                        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                            display_df.to_excel(writer, sheet_name='Results', index=False)
                        
                        st.download_button(
                            label="Download Results as Excel",
                            data=excel_buffer.getvalue(),
                            file_name="tts_analysis_results.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                except Exception as e:
                    status_text.text(f"Error during processing: {str(e)}")
                    progress_bar.progress(0)
                    st.error(f"An error occurred: {str(e)}")
        elif len(st.session_state.pois) == 0:
            st.warning("Please add at least one Point of Interest before processing")
        else:
            st.warning("Please upload a TTS file to process")
    else:
        if not valid_coords:
            st.error("Please enter valid coordinates")
        else:
            st.error("Site zone does not exist in zones.csv")
except Exception as e:
    st.error(f"An error occurred: {str(e)}")
