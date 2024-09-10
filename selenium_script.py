import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from openlocationcode import openlocationcode as olc
import time

# Step 1: Set up the Selenium WebDriver
# Make sure to have the appropriate WebDriver for your browser (e.g., ChromeDriver for Chrome)
driver_path = '/Users/nombauser/Desktop/chromedriver-mac-x64/chromedriver'
driver = webdriver.Chrome()

# Step 2: Read data from XLSX file
file_path = "/Users/nombauser/Downloads/Copy_of_plus_codes-plus_codes copy.csv"
df = pd.read_csv(file_path)

# Step 3: Extract the addresses from the DataFrame
address_column = 'businessAddress'  # Replace with the name of your address column
addresses = df[address_column].tolist()

# Function to get lat and longitude using Selenium with Google Maps
def get_lat_long(address):
    try:
        # Open Google Maps
        driver.get('https://www.google.com/maps')

        # Find the search box, enter the address, and submit the search
        search_box = driver.find_element(By.ID, 'searchboxinput')
        search_box.clear()
        search_box.send_keys(address)
        search_box.send_keys(Keys.ENTER)

        # Wait for the results to load
        time.sleep(0.5)

        # Extract URL with lat and long coordinates from the current page's URL
        current_url = driver.current_url
        if '/@' in current_url:
            parts = current_url.split('/@')[1].split(',')
            lat = parts[0]
            lon = parts[1]
            return lat, lon
        else:
            return None, None
    except Exception as e:
        print(f"Error retrieving coordinates for address {address}: {e}")
        return None, None

# Step 4: Create new columns for Latitude and Longitude
df['Latitude'] = None
df['Longitude'] = None
df['Plus_Code'] = None

# Step 5: Geocode each address
for i, address in enumerate(addresses):
    lat, lon = get_lat_long(address)
    df.at[i, 'Latitude'] = lat
    df.at[i, 'Longitude'] = lon
    if lat is None or lon is None:
        print("Warning: lat or Long is None. Skipping conversion.")
        df.at[i, 'Plus_Code'] = None
    # Step 2: Validate that inputs are numbers
    elif not isinstance(lat, (int, float)) or not isinstance(lon, (int, float)):
        print("Warning: lat and Long must be numeric values. Skipping conversion.")
        df.at[i, 'Plus_Code'] = None
    # Step 3: Validate that the lat and lon are within valid ranges
    elif lat < -90 or lat > 90:
        print("Warning: lat must be between -90 and 90. Skipping conversion.")
        df.at[i, 'Plus_Code'] = None
    elif lon < -180 or lon > 180:
        print("Warning: Long must be between -180 and 180. Skipping conversion.")
        df.at[i, 'Plus_Code'] = None
    else:
        df.at[i, 'Plus_Code'] = olc.encode(lat, lon)

# Print the DataFrame with the new lat and Longitude columns
print(df)

# Step 6: Save the DataFrame to a new Excel file
df.to_excel('output_with_coordinates.xlsx', index=False)

# Close the browser after completing the process
driver.quit()