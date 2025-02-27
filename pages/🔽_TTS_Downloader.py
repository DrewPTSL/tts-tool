import streamlit as st
import pandas as pd
import geopandas as gpd
from shapely.geometry import Point
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from pathlib import Path
import tempfile
import os

# Cache functions for loading zone data
@st.cache_data(show_spinner="Loading zone data...")
def load_zones_data(data_choice):
    """Load zones data based on selected year"""
    if data_choice == "2006 Zones":
        zones_df = pd.read_csv('2006Zones.csv')
        zone_col = 'gta06'
        region_col = 'region'
    else:
        zones_df = pd.read_csv('2022Zones.csv')
        zone_col = 'TTS2022'
        region_col = 'Reg_name'
    return zones_df, zone_col, region_col

@st.cache_data(show_spinner="Loading polygons data...")
def load_geojson_data(data_choice):
    """Load GeoJSON data based on selected year"""
    if data_choice == "2006 Zones":
        file_path = "2006Polygons.geojson"
    else:
        file_path = "2022Polygons.geojson"
    gdf = gpd.read_file(file_path)
    if gdf.crs.to_epsg() != 4326:
        gdf = gdf.to_crs(epsg=4326)
    return gdf


def run_webscraper(site_zones, time_periods, data_choice, custom_time=None, headless=True):
    # Set up Chrome options
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--window-size=1920,1080")
    
    download_dir = str(Path.home() / "Downloads")  # Make sure this path exists
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    
    if headless:
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--disable-gpu")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)

    # Add Streamlit status containers
    status_container = st.empty()
    progress_bar = st.progress(0)

    def update_status(message):
        status_container.text(message)

    try:
        # Initialize the Chrome WebDriver using webdriver_manager
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), 
                                  options=chrome_options)
        
        # Login process
        update_status("Logging into TTS system...")
        driver.get("https://drs.dmg.utoronto.ca/idrs/drsQuery/tts")
        time.sleep(1)

        username_field = driver.find_element(By.ID, "username")
        username_field.send_keys(st.secrets["USERNAME"])
        
        password_field = driver.find_element(By.ID, "password")
        password_field.send_keys(st.secrets["PASSWORD"])
        
        send_button = driver.find_element(By.ID, "send")
        send_button.click()
        time.sleep(1)

        # Set zone type based on data_choice
        if data_choice == "2006 Zones":
            origin_zone_type = "2006 GTA zone of origin"
            dest_zone_type = "2006 GTA zone of destination"
        else:
            origin_zone_type = "2022 TTS zone of origin"
            dest_zone_type = "2022 TTS zone of destination"

        zones_str = ", ".join(str(zone) for zone in site_zones)

        # Process time periods
        time_ranges = []
        for period in time_periods:
            if period == "AM Peak":
                time_ranges.append("700-930")
            elif period == "PM Peak":
                time_ranges.append("1600-1830")
            elif period == "All Day":
                time_ranges.append("400-2800")
            elif period == "Other" and custom_time:
                custom_ranges = [time_range.strip() for time_range in custom_time.split(',')]
                time_ranges.extend(custom_ranges)

        total_iterations = len(time_ranges)
        current_iteration = 0

        # Navigate to cross tabulation page
        driver.get("https://drs.dmg.utoronto.ca/idrs/ttsForm/Cros/trip/2022")
        time.sleep(1)

        # Set up common parameters
        update_status("Setting up query parameters...")
        
        # Row variable
        row_variable = driver.find_element(By.XPATH, "//span[text()='Pick a Row Attribute']")
        row_variable.click()
        driver.switch_to.active_element.send_keys(origin_zone_type + Keys.RETURN)
        time.sleep(0.1)

        # Column variable
        column_variable = driver.find_element(By.XPATH, "//span[text()='Pick a Column Attribute']")
        column_variable.click()
        driver.switch_to.active_element.send_keys(dest_zone_type + Keys.RETURN)
        time.sleep(0.1)

        # Add filters
        add_button = driver.find_element(By.CLASS_NAME, "add")
        
        # Origin zone filter
        add_button.click()
        filter_1 = driver.find_element(By.XPATH, "//span[text()='Regional municipality of household']")
        filter_1.click()
        driver.switch_to.active_element.send_keys(origin_zone_type + Keys.RETURN)
        time.sleep(0.1)

        # Destination zone filter
        add_button.click()
        filter_2 = driver.find_element(By.XPATH, "//span[text()='Regional municipality of household']")
        filter_2.click()
        driver.switch_to.active_element.send_keys(dest_zone_type + Keys.RETURN)
        time.sleep(0.1)

        # Time filter
        add_button.click()
        filter_3 = driver.find_element(By.XPATH, "//span[text()='Regional municipality of household']")
        filter_3.click()
        driver.switch_to.active_element.send_keys("Start time of trip" + Keys.RETURN)
        time.sleep(0.1)

        # Fill in zone filter values (these don't change between iterations)
        zone_textboxes = driver.find_elements(By.XPATH, 
            '//input[@class="valuehtml ui-autocomplete-input" and @style="width:300px"]')
        
        if len(zone_textboxes) >= 3:
            zone_textboxes[0].send_keys(str(zones_str))
            zone_textboxes[1].send_keys(str(zones_str))
            time.sleep(0.1)

        # Set OR operator
        operator = driver.find_element(By.XPATH, "//span[text()='And']")
        operator.click()
        driver.switch_to.active_element.send_keys("Or" + Keys.RETURN)
        time.sleep(0.1)

        # Toggle checkboxes
        checkboxes = driver.find_elements(By.XPATH, '//input[@type="checkbox" and @class="toggle"]')
        if len(checkboxes) >= 3:
            checkboxes[-3].click()
            checkboxes[-2].click()
            time.sleep(0.1)

        # Set output format
        radio_button = driver.find_element(By.ID, "emmeFormat")
        radio_button.click()

        # Now loop through time ranges and only change the time value
        for time_range in time_ranges:
            update_status(f"Processing zone{'' if len(site_zones) == 1 else 's'} {zones_str} for time period {time_range}")
            
            # Only update the time range value
            if len(zone_textboxes) >= 3:
                # Clear previous time value
                zone_textboxes[2].clear()
                time.sleep(0.1)
                # Enter new time value
                zone_textboxes[2].send_keys(time_range)
                time.sleep(0.1)

            
            # Execute query for this time range
            time.sleep(1)
            update_status(f"Executing query for time period {time_range} and downloading results...")
            execute_button = driver.find_element(By.CLASS_NAME, "submitCrosstab")
            execute_button.click()

            # Save results
            try:
                # Wait for Execute Query button to be clickable (indicating the query is complete)
                WebDriverWait(driver, 30).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//button[@class='submitCrosstab' and text()='Execute Query']")
                    )
                )
                # Click the save button as before
                save_button = driver.find_element(By.CLASS_NAME, "saveAs")
                save_button.click()
                
                # Wait briefly for the file to download
                time.sleep(1)
                
                # Find the most recently downloaded file in the download directory
                download_dir = str(Path.home() / "Downloads")
                list_of_files = Path(download_dir).glob('*')
                latest_file = max(list_of_files, key=os.path.getctime)
                
                # Read the file content
                with open(latest_file, 'rb') as f:
                    file_content = f.read()
                
                # Create a temporary file we can access later
                with tempfile.NamedTemporaryFile(delete=False, suffix=f"_{time_range}.txt") as tmp:
                    tmp.write(file_content)
                    temp_path = tmp.name
                
                # Store the path for later download via Streamlit
                if 'download_files' not in st.session_state:
                    st.session_state.download_files = []
                st.session_state.download_files.append((f"tts_data_{time_range}.txt", temp_path))
                
                update_status(f"Downloaded data for time period {time_range}!")

                
                            
            except Exception as e:
                st.error(f"Error downloading for zones {zones_str}, time {time_range}: {str(e)}")
                continue

            current_iteration += 1
            progress_bar.progress(current_iteration / total_iterations)

        update_status("All files processed successfully!")
        return True

    except Exception as e:
        st.error(f"Error occurred: {str(e)}")
        return False
    finally:
        if 'driver' in locals():
            driver.quit()

## Streamlit UI

st.set_page_config(page_title="TTS Downloader", page_icon="🔽", layout="wide", initial_sidebar_state="auto")
st.sidebar.title("🔽 TTS Downloader")
st.title("TTS Downloader")

# Sidebar for credentials
# with st.sidebar.form("Credentials"):
#     username = st.text_input("TTS Username", key="username")
#     password = st.text_input("TTS Password", type="password", key="password")
#     submitted = st.form_submit_button("Log in")

# Main content
Zone_Year = ["2006 Zones", "2022 Zones"]
data_choice = st.pills("Select Data Year:", Zone_Year, selection_mode="single")

if data_choice:
    zones_df, zone_col, region_col = load_zones_data(data_choice)
    gdf = load_geojson_data(data_choice)

    col1, col2 = st.columns(2)
    
    with col1:
        site_zone = st.multiselect(
            "Site Zone(s)", 
            zones_df['GTA06'] if data_choice == "2006 Zones" else zones_df['TTS2022']
        )

    with col2:
        coords_input = st.text_input(
            "Site Coordinates (Latitude, Longitude) (Optional for looking up site zone)",
            value="",
            help="Enter coordinates in format: latitude, longitude"
        )
    if coords_input and data_choice:  # Add data_choice check
        site_lat, site_lon = map(float, coords_input.replace(" ", "").split(","))
        point = Point(site_lon,site_lat)
        matching_polygon = gdf[gdf.contains(point)]

        if not matching_polygon.empty:
            if data_choice == "2006 Zones":
                suggested_zone = matching_polygon.iloc[0]['gta06']
            else:
                suggested_zone = matching_polygon.iloc[0]['TTS2022'] 
            st.write(f"Recommended zone based on coordinates: {suggested_zone}")

    # Time period selection
    col1, col2 = st.columns(2)
    with col1:
        time_period = ["AM Peak", "PM Peak", "All Day", "Other"]
        time_choice = st.pills("Select Time Period:", time_period, selection_mode="multi")
    
    with col2:
        custom_time = None
        if "Other" in time_choice:
            custom_time = st.text_input(
                "Enter time range(s)",
                value="",
                help="e.g. 1200-1400 for 12 p.m. to 2 p.m. (Seperate multiple ranges with commas)"
            )

    # Download button
    if st.button("Processs Files"):
        if not site_zone:
            st.error("Please select at least one site zone")
        elif not time_choice:
            st.error("Please select at least one time period")
        else:
            with st.spinner("Processing files..."):
                success = run_webscraper(
                    site_zones=site_zone,
                    time_periods=time_choice,
                    data_choice=data_choice,
                    custom_time=custom_time)


    if 'download_files' in st.session_state and st.session_state.download_files:
        st.success("Files processed successfully! Click below to download:")
        
        for i, (filename, filepath) in enumerate(st.session_state.download_files):
            with open(filepath, "rb") as file:
                btn = st.download_button(
                    label=f"Download {filename}",
                    data=file,
                    file_name=filename,
                    mime="text/plain",
                    key=f"download_btn_{i}"  
                )
else:
    st.warning("Please select a data year")
