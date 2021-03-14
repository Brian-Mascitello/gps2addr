from geopy.extra.rate_limiter import RateLimiter
from geopy.geocoders import Nominatim
import pandas as pd
from tqdm import tqdm

# tqdm is a simple progress bar for pandas.
tqdm.pandas()

# Create a geocoder with Nominatim.
geolocator = Nominatim(user_agent='myGeocoder', timeout=10)

# Following the specifications at https://operations.osmfoundation.org/policies/nominatim/, the rate is limited to one
# request per second. Forward takes address to coordinates, reverse takes coordinates to addresses.
forward_geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1)
reverse_geocode = RateLimiter(geolocator.reverse, min_delay_seconds=1)

# Get and format the Excel filename.
filename = input('Input the Excel filename: ')
if '.xlsx' not in filename:
    filename = filename + '.xlsx'

# Read the Excel file into a Dataframe df with pandas.
df = pd.read_excel(filename)

# Print the column names, the sample file has 'Location Name', and 'Address'.
# print(df.columns.values)

# Create the column 'Forward Location' to store the reverse geocode results in the Dataframe.
df['Forward Location'] = df['Address'].progress_apply(forward_geocode)

# Create the 'Latitude', 'Longitude', and 'Altitude' columns from the 'Location' information.
df['Point'] = df['Forward Location'].apply(lambda x: tuple(x.point) if x else None)
df[['Latitude', 'Longitude', 'Altitude']] = pd.DataFrame(df['Point'].tolist(), index=df.index)

# Create the column 'Coordinates' that combines the 'Latitude' and 'Longitude' columns into the proper format.
df['Coordinates'] = df['Latitude'].map(str) + ', ' + df['Longitude'].map(str)

# Remove any unnecessary columns.
df = df.drop(columns=['Point', 'Latitude', 'Longitude', 'Altitude'])

# Create the column 'Reversed Location' to store the reverse geocode results in the Dataframe.
# If done right, this should be very similar to the 'Forward Location' column.
df['Reversed Location'] = df['Coordinates'].progress_apply(reverse_geocode)

# Save the results in an output Excel file.
df.to_excel('output.xlsx', index=False)
