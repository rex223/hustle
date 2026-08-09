import pandas as pd
from pathlib import Path
import requests
import math
import time

# --- Config ---
EXCEL_PATH = Path(__file__).parent.parent / "sheets" / "forms_v5.xlsm"
TRAIN_CSV = Path(__file__).parent.parent / "sheets" / "train_data.csv"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"}

def haversine(lat1, lon1, lat2, lon2):
    R = 6371.0
    lat1, lon1, lat2, lon2 = map(math.radians, [lat1, lon1, lat2, lon2])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = math.sin(dlat/2)**2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon/2)**2
    return R * 2 * math.asin(math.sqrt(a))

def overpass_query(lat, lon, tag_filter, radius=15000):
    query = f"""
    [out:json];
    (
      node[{tag_filter}](around:{radius},{lat},{lon});
      way[{tag_filter}](around:{radius},{lat},{lon});
      relation[{tag_filter}](around:{radius},{lat},{lon});
    );
    out center;
    """
    for attempt in range(3):
        try:
            resp = requests.post("https://overpass-api.de/api/interpreter", data={"data": query}, headers=HEADERS, timeout=10)
            if resp.status_code == 429:
                time.sleep(5)
                continue
            resp.raise_for_status()
            return resp.json().get("elements", [])
        except Exception:
            time.sleep(2)
    return []

def nearest_name(elements, lat, lon):
    best_name, best_dist = "", float("inf")
    for el in elements:
        el_lat = el.get("lat") or el.get("center", {}).get("lat")
        el_lon = el.get("lon") or el.get("center", {}).get("lon")
        if el_lat is None: continue
        dist = haversine(lat, lon, float(el_lat), float(el_lon))
        name = (el.get("tags") or {}).get("name") or (el.get("tags") or {}).get("name:en", "")
        if dist < best_dist and name:
            best_dist, best_name = dist, name
    return best_name

def find_train(station_name):
    if not station_name: return ""
    try:
        url = f"https://erail.in/rail/getStations.aspx?StationName={station_name}&_=1"
        resp = requests.get(url, headers=HEADERS, timeout=5)
        if resp.status_code == 200:
            data = resp.json()
            if data:
                return f"From Sealdah to {station_name} ({data[0].get('StationCode', '')})"
    except Exception:
        pass
    return f"From Sealdah to {station_name}"

def main():
    if not EXCEL_PATH.exists():
        print("Base Excel file not found. Run Phase 1 first.")
        return
        
    print("Loading Excel...")
    import warnings
    warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
    df = pd.read_excel(EXCEL_PATH, sheet_name="Database", engine="openpyxl", dtype=str).fillna("")
    
    # We only process rows where Station is empty
    mask = df['Nearest Railway Station'].str.strip() == ""
    missing_df = df[mask]
    print(f"Found {len(missing_df)} row(s) needing Train/Station data.")
    
    results = []
    
    # Simplified geocoding stub: In a real environment, we'd geocode the Village/PIN first.
    # For now, if we have coordinates we use them, otherwise we skip.
    # Wait, previously we used Nominatim to geocode the Village + PIN.
    # Let's add Nominatim back for geocoding.
    
    for idx, row in missing_df.iterrows():
        code = row['CODE']
        village = row['VILLAGE']
        pin = row['PIN']
        
        print(f"[{code}] Finding station for {village} (PIN: {pin})...")
        
        # 1. Geocode
        lat, lon = None, None
        try:
            loc_url = f"https://nominatim.openstreetmap.org/search?q={village},{pin},West+Bengal&format=json"
            loc_resp = requests.get(loc_url, headers=HEADERS, timeout=5).json()
            if loc_resp:
                lat, lon = float(loc_resp[0]['lat']), float(loc_resp[0]['lon'])
            time.sleep(1.1) # respect Nominatim limit
        except Exception:
            pass
            
        if not lat:
            print("  Could not geocode location.")
            continue
            
        # 2. Find Nearest Station
        els = overpass_query(lat, lon, '"railway"="station"')
        if not els:
            els = overpass_query(lat, lon, '"railway"="halt"')
        station = nearest_name(els, lat, lon)
        
        # 3. Find Train
        train = find_train(station) if station else ""
        
        print(f"  Station: {station}")
        print(f"  Train  : {train}")
        
        results.append({
            "CODE": code,
            "Nearest Railway Station": station,
            "Train from Sealdah": train
        })
        
    if results:
        TRAIN_CSV.parent.mkdir(parents=True, exist_ok=True)
        pd.DataFrame(results).to_csv(TRAIN_CSV, index=False)
        print(f"\nSaved {len(results)} train records to {TRAIN_CSV.name}")
    else:
        print("No new train data found.")

if __name__ == "__main__":
    main()
