import folium
import geopandas as gpd
from shapely.geometry import Point

# Load GeoJSON into a GeoDataFrame
gdf = gpd.read_file("C:\Temp\TTS\Polygons.geojson")

# Convert to WGS 84 (EPSG:4326) if necessary
if gdf.crs.to_epsg() != 4326:
    gdf = gdf.to_crs(epsg=4326)

# Define the target coordinate (latitude, longitude)
coordinate = (43.21323612898636, -79.77276508050248)
point = Point(coordinate[1], coordinate[0])  # Note: Point takes (longitude, latitude)

# Find the polygon containing the point
matching_polygon = gdf[gdf.contains(point)]

# Create a base Folium map centered on the point
m = folium.Map(location=[coordinate[0], coordinate[1]], zoom_start=12)

# Style function for regular polygons
def style_function(feature):
    return {
        'fillColor': 'white',
        'color': 'black',
        'weight': 2,
        'fillOpacity': 0.5
    }

# Style function for the highlighted polygon
def highlight_style_function(feature):
    return {
        'fillColor': 'yellow',
        'color': 'red',
        'weight': 3,
        'fillOpacity': 0.7
    }

# Regular polygons layer
regular_layer = folium.FeatureGroup(name="Other Zones")
for _, row in gdf.iterrows():
    if row.geometry not in matching_polygon.geometry.values:
        folium.GeoJson(
            row.geometry,
            style_function=style_function,
            tooltip=f"gta06: {row['gta06']}, Region: {row['region']}"
        ).add_to(regular_layer)

# Highlighted polygon layer
highlight_layer = folium.FeatureGroup(name="Highlighted Zone", show=True)
for _, row in matching_polygon.iterrows():
    folium.GeoJson(
        row.geometry,
        style_function=highlight_style_function,
        tooltip=f"gta06: {row['gta06']}, Region: {row['region']}"
    ).add_to(highlight_layer)

# Add the layers to the map
regular_layer.add_to(m)
highlight_layer.add_to(m)

# Add a marker for the point
folium.Marker(
    location=[coordinate[0], coordinate[1]],
    popup="Site",
    icon=folium.Icon(color="red", icon="info-sign")
).add_to(m)

# Add LayerControl for toggling
folium.LayerControl(collapsed=False).add_to(m)

# Save and display the map
m.save("Zone_Map.html")
print("Done")


