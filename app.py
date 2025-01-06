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
from streamlit_folium import st_folium
import io
import plotly.express as px
import matplotlib.pyplot as plt
import geopandas as gpd
from shapely.geometry import Point

zones_df = pd.read_csv('Zones.csv')

gdf = gpd.read_file("Polygons.geojson")
if gdf.crs.to_epsg() != 4326:
    gdf = gdf.to_crs(epsg=4326)

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

# Title and description
st.title("TTS Route Analysis Tool")

# Add reset button in the top right
col1, col2 = st.columns([6, 1])

with col1:
    st.markdown("Upload your TTS file and analyze routes through points of interest.")

# File Upload Section
uploaded_file = st.file_uploader("Upload your TTS file", type=['txt'])

# Configuration Section
st.markdown("### Site Configuration")
col1, col2 = st.columns(2)

with col1:
    site_zone = st.selectbox(
        "Site Zone",
        zones_df['GTA06'],
        help="Enter the site zone number"
    )

with col2:
    coords_input = st.text_input(
        "Site Coordinates (Latitude, Longitude)",
        value="43.4109, -80.3142",
        help="Enter coordinates in format: latitude, longitude"
    )

try:
    site_lat, site_lon = map(float, coords_input.replace(" ", "").split(","))
    valid_coords = True
except ValueError:
    st.error("Invalid coordinates format. Please use: latitude, longitude")
    valid_coords = False

## Site Zone Matching

if site_lon:
    point = Point(site_lon, site_lat)
    matching_polygon = gdf[gdf.contains(point)]

    if not matching_polygon.empty:
        # Safely access the first matching polygon's 'gta06' column
        suggested_gta06_zone = matching_polygon.iloc[0]['gta06']
        st.write(f"Recommended zone based on coordinates: {suggested_gta06_zone}")
    else:
        st.write("No matching polygon found for the given coordinates.")


# POI Management Section
st.markdown("### Points of Interest")

# Add button to create new row
if st.button("Add New Row"):
    if 'num_rows' not in st.session_state:
        st.session_state.num_rows = 1
    else:
        st.session_state.num_rows += 1

if 'num_rows' not in st.session_state:
    st.session_state.num_rows = 1

# Display POI input rows
for i in range(st.session_state.num_rows):
   col1, col2, col3, col4 = st.columns([2, 2, 1, 0.5])
   with col1:
       name = st.text_input("POI Name", key=f"name_{i}", help="Enter name (e.g. North via Greenhouse Road)")
   with col2:
       coords = st.text_input("Coordinates (Lat, Lon)", key=f"coords_{i}",help="Enter coordinates in format: latitude, longitude")
   with col3:
       threshold = st.slider("Threshold (m)", min_value=1, max_value=500, value=50, key=f"threshold_{i}",help = "Select the threshold redius around the POI")
   with col4:
       st.write("")
       st.write("")
       if st.button("🗑️", key=f"delete_{i}", help = "Delete POI"):
           st.session_state.num_rows -= 1
           st.rerun()

# Process filled rows into POIs list before analysis
st.session_state.pois = []
for i in range(st.session_state.num_rows):
    name = st.session_state.get(f"name_{i}")
    coords = st.session_state.get(f"coords_{i}")
    threshold = st.session_state.get(f"threshold_{i}")
    
    if name and coords:
        try:
            lat, lon = map(float, coords.replace(" ", "").split(","))
            st.session_state.pois.append({
                'id': f'POI_{i + 1}',
                'name': name,
                'coordinates': (lat, lon),
                'threshold': threshold / 1000
            })
        except ValueError:
            st.error(f"Invalid coordinates format in row {i + 1}")

# After the site coordinates input section, add the map visualization
if valid_coords:
    try:          
        # Add a toggle button above the map
        show_map = st.toggle('Show Map', value=False)

        # Only display the map if toggle is on
        if show_map:
            m = folium.Map(location=[site_lat, site_lon], zoom_start=12, width='100%')

            # Add a download button
            if st.button("Download Map"):
                # Save map to HTML
                m.save("map.html")
                st.success("Map downloaded successfully!")

            # Add site zone marker
            sitezone_layer = folium.FeatureGroup(name="Site Zone Marker", show=True)
            folium.Marker(
                location=(site_lat, site_lon),
                popup=f"Site Zone {site_zone}",
                icon=folium.Icon(color='black', icon='home')
            ).add_to(sitezone_layer)

            # Add POIs with tooltips
            poi_layer = folium.FeatureGroup(name="POI Marker", show=True)
            poithreshold_layer = folium.FeatureGroup(name="POI Threshold Buffer", show=True)
            for poi in st.session_state.pois:
                folium.CircleMarker(
                    location=poi['coordinates'],
                    radius=5,
                    popup=f"POI ID: {poi['id']}<br>Name: {poi['name']}<br>Threshold: {poi['threshold']} km",
                    color='orange',
                    fill=True,
                    fillColor='orange',
                    fillOpacity=0.7
                ).add_to(poi_layer)

                # Add proximity circle for POIs with individual thresholds
                folium.Circle(
                    location=poi['coordinates'],
                    radius=poi['threshold'] * 1000,  # Convert km to meters
                    color='orange',
                    fill=True,
                    fillOpacity=0.2,
                    popup=f"{poi['name']}<br>Threshold: {poi['threshold']} km",
                ).add_to(poi_layer)

            def highlight_style_function(feature):
                return {
                    'fillColor': 'yellow',
                    'color': 'red',
                    'weight': 3,
                    'fillOpacity': 0.7
                }

            # Style function for regular polygons
            def style_function(feature):
                return {
                    'fillColor': 'white',
                    'color': 'black',
                    'weight': 2,
                    'fillOpacity': 0.5
                }

            # Highlighted polygon layer
            if site_zone == suggested_gta06_zone:
                selectedzone_layer = folium.FeatureGroup(name="Selected Zone", show=True)
                for _, row in matching_polygon.iterrows():
                    folium.GeoJson(
                        row.geometry,
                        style_function=style_function,
                        tooltip=f"gta06: {row['gta06']}, Region: {row['region']}"
                    ).add_to(selectedzone_layer)
            else:
                suggestedzone_layer = folium.FeatureGroup(name="Suggested Zone", show=True)
                for _, row in matching_polygon.iterrows():
                    folium.GeoJson(
                        row.geometry,
                        style_function=style_function,
                        tooltip=f"gta06: {row['gta06']}, Region: {row['region']}"
                    ).add_to(suggestedzone_layer)
                
                selectedzone_layer = folium.FeatureGroup(name="Selected Zone", show=True)
                for _, row in gdf.iterrows():
                    if row['gta06'] == site_zone:  # Check if the gta06 matches the site_zone
                        folium.GeoJson(
                            row.geometry,
                            style_function=highlight_style_function,
                            tooltip=f"gta06: {row['gta06']}, Region: {row['region']}"
                        ).add_to(selectedzone_layer)
                suggestedzone_layer.add_to(m)

            # Add layer control
            selectedzone_layer.add_to(m)
            sitezone_layer.add_to(m)
            poi_layer.add_to(m)
            folium.LayerControl(collapsed=False).add_to(m)
            
            # Display map in Streamlit
            st.subheader("Site and POI Map")
            with st.container():
                st_folium(m, height=600, use_container_width=True,returned_objects=[])
        
    except Exception as e:
        st.error(f"Error creating map: {str(e)}")

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
                        def generate_formatted_excel():
                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                # Format Route Results sheet
                                export_df = display_df[['origin_id', 'dest_id', 'route_type', 'POI', 'total']].copy()
                                export_df.to_excel(writer, sheet_name='Route Results', index=False)

                                # Create POI Summary DataFrame
                                poi_summary_df = pd.DataFrame([{
                                    'POI Name': poi['name'],
                                    'POI_ID': poi['id'],
                                    'Latitude': poi['coordinates'][0],
                                    'Longitude': poi['coordinates'][1],
                                    'Threshold (km)': poi['threshold']
                                } for poi in st.session_state.pois])

                                # Create Site Summary DataFrame
                                site_summary_df = pd.DataFrame([{
                                    'Site Zone': site_zone,
                                    'Latitude': site_lat,
                                    'Longitude': site_lon
                                }])

                                # Write sheets and apply formatting
                                poi_summary_df.to_excel(writer, sheet_name='Location Details', startrow=0, startcol=0, index=False)
                                site_summary_df.to_excel(writer, sheet_name='Location Details', startrow=0, startcol=poi_summary_df.shape[1] + 2, index=False)

                                
                                # Apply Excel formatting
                                workbook = writer.book
                                for sheet_name in ['Route Results', 'Location Details']:
                                    sheet = writer.sheets[sheet_name]
                                    if sheet_name == 'Location Details':
                                        apply_header_formatting(sheet, exclude_columns=[6,7])
                                    else:
                                        apply_header_formatting(sheet)
                                    autofit_columns(sheet)

                                # Create and format POI Traffic Analysis sheet
                                calc_sheet = create_poi_analysis_sheet(workbook, st.session_state.pois)
                                format_poi_analysis_sheet(calc_sheet, len(st.session_state.pois))

                                # Create Raw Text sheet if content available
                                create_raw_text_sheet(workbook, content)

                            return output.getvalue()

                        def apply_header_formatting(sheet, exclude_columns=None):
                            header_font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
                            header_fill = PatternFill(start_color='005295', end_color='005295', fill_type='solid')
                            border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                                        top=Side(style='thin'), bottom=Side(style='thin'))
                            if exclude_columns is None:
                                exclude_columns = []
                            for cell in sheet[1]:
                                if cell.column not in exclude_columns:
                                    cell.font = header_font
                                    cell.fill = header_fill
                                    cell.border = border
                                    cell.alignment = Alignment(horizontal='center')

                        def autofit_columns(sheet):
                            for column in sheet.columns:
                                max_length = 0
                                column = [cell for cell in column]
                                for cell in column:
                                    try:
                                        max_length = max(max_length, len(str(cell.value)))
                                    except:
                                        pass
                                adjusted_width = max_length + 2
                                sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

                        def create_poi_analysis_sheet(workbook, pois):
                            calc_sheet = workbook.create_sheet(title='POI Traffic Analysis')
                            total_row = len(pois) + 3
                            
                            calc_sheet['B2'] = 'From/To'
                            calc_sheet.merge_cells('C2:D2')
                            calc_sheet['C2'] = 'In'
                            calc_sheet.merge_cells('E2:F2')
                            calc_sheet['E2'] = 'Out'
                            
                            for idx, poi in enumerate(pois, start=3):
                                calc_sheet[f'B{idx}'] = poi['name']
                                calc_sheet[f'C{idx}'] = f'=SUMIFS(\'Route Results\'!E:E,\'Route Results\'!C:C,"origin_to_site",\'Route Results\'!D:D,"{poi["name"]}")'
                                calc_sheet[f'E{idx}'] = f'=SUMIFS(\'Route Results\'!E:E,\'Route Results\'!C:C,"site_to_destination",\'Route Results\'!D:D,"{poi["name"]}")'
                                calc_sheet[f'D{idx}'] = f'=IF(C{idx}>0,C{idx}/C{total_row},0)'
                                calc_sheet[f'F{idx}'] = f'=IF(E{idx}>0,E{idx}/E{total_row},0)'

                            calc_sheet[f'B{total_row}'] = "Total"
                            calc_sheet[f'C{total_row}'] = f'=SUM(C3:C{total_row-1})'
                            calc_sheet[f'D{total_row}'] = f'=SUM(D3:D{total_row-1})'
                            calc_sheet[f'E{total_row}'] = f'=SUM(E3:E{total_row-1})'
                            calc_sheet[f'F{total_row}'] = f'=SUM(F3:F{total_row-1})'
                            
                            return calc_sheet

                        def format_poi_analysis_sheet(sheet, num_pois):
                            total_row = num_pois + 3
                            
                            # Styles
                            header_font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
                            header_fill = PatternFill(start_color='005295', end_color='005295', fill_type='solid')
                            cell_font = Font(name='Arial', size=11)
                            total_fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
                            border_style = Border(left=Side(style='thin'), right=Side(style='thin'), 
                                        top=Side(style='thin'), bottom=Side(style='thin'))

                            # Format headers
                            for col in ['B', 'C', 'D', 'E', 'F']:
                                cell = sheet[f'{col}2']
                                cell.font = header_font
                                cell.fill = header_fill
                                cell.alignment = Alignment(horizontal='center')
                                cell.border = border_style

                            # Format POI rows
                            for row in range(3, total_row):
                                sheet[f'B{row}'].font = header_font
                                sheet[f'B{row}'].fill = header_fill
                                for col in ['C', 'D', 'E', 'F']:
                                    cell = sheet[f'{col}{row}']
                                    cell.font = cell_font
                                    cell.alignment = Alignment(horizontal='right')
                                    cell.border = border_style
                                    if col in ['D', 'F']:
                                        cell.number_format = '0.00%'

                            # Format the "Total" row
                            sheet[f'B{total_row}'].font = header_font
                            sheet[f'B{total_row}'].fill = header_fill

                            for col in ['C', 'D', 'E', 'F']:
                                total_cell = sheet[f'{col}{total_row}']
                                total_cell.font = Font(name='Arial', size=11, bold=True)
                                total_cell.fill = total_fill
                                total_cell.alignment = Alignment(horizontal='right')
                                total_cell.border = border_style
                                if col in ['D', 'F']:  # Ensure percentage format for Total row in columns D and F
                                    total_cell.number_format = '0.00%'

                            # Apply "all borders" to the entire table
                            for row in range(2, total_row + 1):
                                for col in ['B', 'C', 'D', 'E', 'F']:
                                    cell = sheet[f'{col}{row}']
                                    cell.border = border_style

                            # Autofit the width of column B
                            column_letter = 'B'
                            max_length = 0

                            # Iterate through all rows in column B to find the longest content
                            for row in sheet.iter_rows(min_col=2, max_col=2, min_row=2, max_row=total_row):
                                for cell in row:
                                    if cell.value:  # Ensure the cell has a value
                                        max_length = max(max_length, len(str(cell.value)))

                            # Adjust the column width (adding a little extra for padding)
                            sheet.column_dimensions[column_letter].width = max_length + 2

                        def create_raw_text_sheet(workbook, content):
                            raw_sheet = workbook.create_sheet(title='Raw Text')
                            row_idx = 1
                            split_mode = False
                            raw_sheet.sheet_view.show_grid_lines = False
                            
                            for line in content.split('\n'):
                                stripped_line = line.strip()
                                if stripped_line.startswith("gta06_orig"):
                                    split_mode = True
                                    
                                if split_mode and stripped_line:
                                    parts = stripped_line.split()
                                    for col_idx, value in enumerate(parts, start=1):
                                        raw_sheet.cell(row=row_idx, column=col_idx, value=value)
                                else:
                                    raw_sheet.cell(row=row_idx, column=1, value=stripped_line)
                                
                                row_idx += 1
                            
                        
                        excel_data = generate_formatted_excel()
                        st.download_button(
                            label="Download Results as Excel",
                            data=excel_data,
                            file_name="tts_analysis_results.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                        if st.session_state.results_df is not None and not st.session_state.results_df.empty:
                             # Add a toggle button above the map
                            show_route_map = st.toggle('Show Route Map', value=False)
                            if show_route_map:
                                
                                route_map = folium.Map(location=[site_lat, site_lon], zoom_start=10)

                                # Add a download button
                                if st.button("Download Map"):
                                    # Save map to HTML
                                    route_map.save("route_map.html")
                                    st.success("Map downloaded successfully!")
                                
                                # Create feature groups
                                site_layer = folium.FeatureGroup(name="Site Location", show=True)
                                origin_routes = folium.FeatureGroup(name="Origin to Site Routes", show=True)
                                dest_routes = folium.FeatureGroup(name="Site to Destination Routes", show=True)
                                poi_layer = folium.FeatureGroup(name="Points of Interest", show=True)
                                
                                # Add site marker
                                folium.Marker(
                                    location=[site_lat, site_lon],
                                    popup=folium.Popup("Site Location",max_width=200),
                                    icon=folium.Icon(color='black', icon='home')
                                ).add_to(site_layer)
                                
                                # Add routes
                                for _, row in st.session_state.results_df.iterrows():
                                    if row['passes']:
                                        origin_id = row['origin_id']
                                        dest_id = row['dest_id']
                                        route_type = row['route_type']
                                        
                                        if route_type == 'origin_to_site':
                                            origin_row = zones_df[zones_df['GTA06'] == origin_id].iloc[0]
                                            route = get_route(origin_row['Latitude'], origin_row['Longitude'], site_lat, site_lon)
                                            if route:
                                                coords = decode(route)
                                                folium.PolyLine(coords, weight=2, color='blue', opacity=0.8,
                                                    popup=folium.Popup(f"Origin Zone: {origin_id}<br> Traffic: {row['total']}",max_width=200)
                                                ).add_to(origin_routes)
                                                
                                        elif route_type == 'site_to_destination':
                                            dest_row = zones_df[zones_df['GTA06'] == dest_id].iloc[0]
                                            route = get_route(site_lat, site_lon, dest_row['Latitude'], dest_row['Longitude'])
                                            if route:
                                                coords = decode(route)
                                                folium.PolyLine(coords, weight=2, color='red', opacity=0.8,
                                                    popup=folium.Popup(f"Destination Zone: {dest_id}<br> Traffic: {row['total']}",max_width=200)
                                                ).add_to(dest_routes)
                                
                                # Add POI markers
                                for poi in st.session_state.pois:
                                    folium.CircleMarker(
                                        location=poi['coordinates'],
                                        radius=5,
                                        popup=f"POI ID: {poi['id']}<br>Name: {poi['name']}<br>Threshold: {poi['threshold']} km",
                                        color='orange',
                                        fill=True,
                                        fillColor='orange',
                                        fillOpacity=0.7
                                    ).add_to(poi_layer)

                                    # Add proximity circle for POIs with individual thresholds
                                    folium.Circle(
                                        location=poi['coordinates'],
                                        radius=poi['threshold'] * 1000,  # Convert km to meters
                                        color='orange',
                                        fill=True,
                                        fillOpacity=0.2,
                                        popup=f"{poi['name']}<br>Threshold: {poi['threshold']} km",
                                    ).add_to(poi_layer)

                                # Add zone nodes layer
                                origin_nodes = folium.FeatureGroup(name="Origin Zone Locations", show=True)
                                dest_nodes = folium.FeatureGroup(name="Destination Zone Locations", show=True)

                                # Add zone markers with enhanced popups
                                for _, row in st.session_state.results_df.iterrows():
                                    if row['passes']:
                                        # Get POI names and total trips
                                        poi = row['intersected_pois'][0]['name']
                                        total_trips = row['total']
                                        
                                        if row['route_type'] == 'origin_to_site':
                                            zone_id = row['origin_id']
                                            zone_row = zones_df[zones_df['GTA06'] == zone_id].iloc[0]
                                            popup_text = (f"Origin Zone: {zone_id}<br>"
                                                            f"POI: {poi}<br>"
                                                            f"Total Trips: {total_trips}")
                                            
                                            folium.Marker(
                                                location=[zone_row['Latitude'], zone_row['Longitude']],
                                                popup=folium.Popup(popup_text, max_width=200),
                                                icon=folium.Icon(color='cadetblue', icon ='car',prefix='fa')
                                            ).add_to(origin_nodes)
                                            
                                        else:  # site_to_destination
                                            zone_id = row['dest_id']
                                            zone_row = zones_df[zones_df['GTA06'] == zone_id].iloc[0]
                                            popup_text = (f"Destination Zone: {zone_id}<br>"
                                                            f"POI: {poi}<br>"
                                                            f"Total Trips: {total_trips}")
                                            
                                            folium.Marker(
                                                location=[zone_row['Latitude'], zone_row['Longitude']],
                                                popup=folium.Popup(popup_text, max_width=200),
                                                icon=folium.Icon(color='darkred', icon ='car-side',prefix='fa')
                                            ).add_to(dest_nodes)                     
                                
                                # Add all layers to map
                                site_layer.add_to(route_map)
                                origin_routes.add_to(route_map)
                                dest_routes.add_to(route_map)
                                poi_layer.add_to(route_map)
                                origin_nodes.add_to(route_map)
                                dest_nodes.add_to(route_map)
                                
                                # Add layer control
                                folium.LayerControl(collapsed=False).add_to(route_map)
                                
                                st.subheader("Route Map")
                                st_folium(route_map, height=600, width=None,returned_objects=[])

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

