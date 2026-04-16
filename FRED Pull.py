import pandas as pd
from fredapi import Fred
import time

# --- CONFIGURATION ---
# Replace with your actual FRED API key
API_KEY = 'acbfcc7ed5ba2f611a3b0eff80a17233'
fred = Fred(api_key=API_KEY)

# Geographies: Name -> BLS Area Code (5 digits)
GEOGRAPHIES = {
    "US Total": "00000",
    "Atlanta-Sandy Springs-Roswell, GA": "12060",
    "Boston-Cambridge-Newton, MA-NH": "14460",
    "Chicago-Naperville-Elgin, IL-IN": "16980",
    "Dallas-Fort Worth-Arlington, TX": "19100",
    "Denver-Aurora-Centennial, CO": "19740",
    "Detroit-Warren-Dearborn, MI": "19820",
    "Houston-Pasadena-The Woodlands, TX": "26420",
    "Los Angeles-Long Beach-Anaheim, CA": "31080",
    "Miami-Fort Lauderdale-West Palm Beach, FL": "33100",
    "Minneapolis-St. Paul-Bloomington, MN-WI": "33460",
    "New York-Newark-Jersey City, NY-NJ": "35620",
    "Orlando-Kissimmee-Sanford, FL": "36740",
    "Philadelphia-Camden-Wilmington, PA-NJ-DE-MD": "37980",
    "Phoenix-Mesa-Chandler, AZ": "38060",
    "Riverside-San Bernardino-Ontario, CA": "40140",
    "San Diego-Chula Vista-Carlsbad, CA": "41740",
    "San Francisco-Oakland-Fremont, CA": "41860",
    "Seattle-Tacoma-Bellevue, WA": "42660",
    "Tampa-St. Petersburg-Clearwater, FL": "45300",
    "Washington-Arlington-Alexandria, DC-VA-MD-WV": "47900"
}

# Industries: Label -> Industry Code (8 digits)
# Many NAICS are grouped into Supersectors in FRED's MSA series
INDUSTRIES = {
    "21-Mining": "10000000",
    "22-Utilities": "44000000",
    "23-Construction": "20000000",
    "31-33-Manufacturing": "30000000",
    "42-Wholesale": "41000000",
    "44-45-Retail": "42000000",
    "48-49-Transport/Whse": "43000000",
    "51-Information": "50000000",
    "52-Finance/Insurance": "55000000", # Financial Activities
    "53-Real Estate": "55000000",        # Financial Activities
    "54-Prof/Sci/Tech": "60000000",      # Prof & Business Services
    "55-Management": "60000000",         # Prof & Business Services
    "56-Admin/Support": "60000000",      # Prof & Business Services
    "61-Education": "65000000",          # Education & Health
    "62-Health Care": "65000000",        # Education & Health
    "71-Arts/Ent": "70000000",           # Leisure & Hospitality
    "72-Accommodation": "70000000",      # Leisure & Hospitality
    "81-Other Svcs": "80000000",
    "92-Public Admin": "90000000"        # Government
}

def pull_fred_data():
    all_data = []
    
    for geo_name, geo_code in GEOGRAPHIES.items():
        print(f"Fetching data for {geo_name}...")
        for ind_label, ind_code in INDUSTRIES.items():
            # Series ID: SMU + Area + Industry + Data Type (01 = All Employees)
            series_id = f"SMU{geo_code}{ind_code}01"
            
            # National total series often use PAYEMS or specific CES IDs
            if geo_code == "00000":
                series_id = f"CES{ind_code[:2]}{ind_code[2:4]}000001" # National format variant
            
            try:
                data = fred.get_series(series_id, observation_start='2000-01-01')
                temp_df = data.reset_index()
                temp_df.columns = ['Date', 'Value']
                temp_df['Geography'] = geo_name
                temp_df['Industry'] = ind_label
                temp_df['SeriesID'] = series_id
                all_data.append(temp_df)
                time.sleep(0.5) # Avoid hitting rate limits
            except Exception as e:
                print(f"  Could not pull {series_id} ({ind_label}): {e}")

    final_df = pd.concat(all_data, ignore_index=True)
    
    # Pivot to wide format: Rows = Dates, Columns = Geography/Industry combinations
    pivot_df = final_df.pivot_table(index='Date', columns=['Geography', 'Industry'], values='Value')
    
    # Save to Excel
    pivot_df.to_excel("Employment_Time_Series_2000_Present.xlsx")
    print("\nFile saved successfully: Employment_Time_Series_2000_Present.xlsx")

if __name__ == "__main__":
    pull_fred_data()

