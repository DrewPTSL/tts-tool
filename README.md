# TTS Route Analysis Tool 🚗
A Streamlit-based web application for analyzing Transportation Tomorrow Survey (TTS) routes through specified points of interest. This tool helps determine route distributions in relation to key locations.
## Features
### Data Support

* Supports both 2006 and 2022 TTS zone systems
* Handles zone-to-zone trip tables in text format
* Integrates with GeoJSON zone boundaries

### Site Configuration

* Interactive zone selection
* Coordinate-based site location input
* Automatic zone suggestion based on coordinates
* Visual validation through interactive mapping

### Points of Interest (POI) Management

* Dynamic addition and removal of POIs
* Customizable threshold distances for each POI
* Coordinate-based POI placement
* Visual representation on interactive maps

### Route Analysis

* Origin-to-site and site-to-destination route analysis
* Route intersection detection with POIs
* Traffic volume calculations
* Comprehensive results visualization

### Visualization

* Interactive maps showing:
     - Site location
     - POI locations with customizable buffer zones
     - Zone boundaries
     - Route traces for matching trips
     - Origin and destination markers
* Traffic distribution pie charts
* Detailed results tables

### Export Capabilities

* Excel export with multiple sheets:
     - Route analysis results
     - Location details
     - POI traffic analysis
     - Raw input data
* Map downloads in HTML format
* Formatted tables with proper styling

## Requirements
```
streamlit
pandas
folium
polyline
geopy
openpyxl
requests
streamlit_folium
plotly
matplotlib
geopandas
shapely
```

## Required Data Files

- `2006Zones.csv`: Contains 2006 TTS zone system information
- `2022Zones.csv`: Contains 2022 TTS zone system information
- `2006Polygons.geojson`: Contains 2006 zone boundary geometries
- `2022Polygons.geojson`: Contains 2022 zone boundary geometries

## Usage
1.  Upload TTS File
     - Upload your TTS trip table text file
2. Select Data Year
     - Choose between 2006 and 2022 TTS zone systems
3 Configure Site Location
     - Enter site zone number
     - Input site coordinates (latitude, longitude)
4. Add Points of Interest
     - Click "Add New Row" to create new POIs
     - Enter POI name, coordinates, and threshold distance
     - Use delete button to remove unwanted POIs
5. Process
     - Click "Start Processing" to begin analysis
6. View Results
     - Examine interactive maps
     - Review traffic distribution charts
     - Analyze detailed route information
     - Export results to Excel for more detailed view



## Map Features
### Site and POI Map

- Shows site location
- Displays POI locations with buffer zones
- Visualizes zone boundaries
- Supports layer toggling

### Route Map

- Displays matched routes
- Color-coded origin and destination paths
- Interactive popups with trip information
- Customizable layer visibility

## Output Format
The Excel export includes:

1. **Route Results**: Detailed trip analysis
2. **Location Details**: Site and POI information
3. **POI Traffic Analysis**: Traffic distribution calculations
4. **Raw Text**: Original input data

## Notes

- The tool uses OSRM for route calculations
- All distance calculations use geodesic measurements
- POI thresholds are applied in kilometers
- Map visualizations support both overview and detailed views

## Acknowledgments

- Built with Streamlit
- Uses OpenStreetMap routing through OSRM
- Incorporates Folium for mapping capabilities
- Utilizes the Transportation Tomorrow Survey (TTS) data structure
