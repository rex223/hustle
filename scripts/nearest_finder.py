"""
nearest_finder.py
-----------------
For every row in extracted_data.csv that lacks transport data, this script:
1. Geocodes the madrasa location using PIN + Village via Nominatim (OSM, free, no key)
2. Queries Overpass API for the nearest railway station within 7km
3. Queries Overpass API for the nearest auto stand within 7km
4. Queries Overpass API for the nearest market/bazaar within 7km
5. Queries erail.in to find trains between Sealdah and the nearest station
6. Writes results back to extracted_data.csv

Usage:
    python scripts/nearest_finder.py
"""

import time
import math
import requests
import pandas as pd
from pathlib import Path

# ── Config ───────────────────────────────────────────────────────────────────
CSV_PATH = Path(__file__).parent.parent / "extracted_data.csv"
ORIGIN_STATION = "Sealdah"
SEARCH_RADIUS_M = 7000      # 7 km
NOMINATIM_DELAY = 1.1       # seconds between requests (OSM policy: max 1/sec)

HEADERS = {
    "User-Agent": "MadrasaDBProject/1.0 (research; contact: admin@example.com)"
}


# ── Maths ────────────────────────────────────────────────────────────────────
def haversine(lat1, lon1, lat2, lon2):
    """Distance in metres between two lat/lon points."""
    R = 6_371_000
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlam = math.radians(lon2 - lon1)
    a = math.sin(dphi/2)**2 + math.cos(phi1)*math.cos(phi2)*math.sin(dlam/2)**2
    return R * 2 * math.asin(math.sqrt(a))


def geocode(village: str, pin: str) -> tuple[float, float] | None:
    """Return (lat, lon) for a village/pincode, or None if not found, using ArcGIS."""
    queries = [
        f"{village}, {pin}, West Bengal, India",
        f"{pin}, West Bengal, India",
        f"{village}, West Bengal, India",
    ]
    url = "https://geocode.arcgis.com/arcgis/rest/services/World/GeocodeServer/findAddressCandidates"
    for q in queries:
        try:
            resp = requests.get(
                url,
                params={"singleLine": q, "f": "json", "maxLocations": 1},
                headers=HEADERS, timeout=10
            )
            data = resp.json()
            candidates = data.get("candidates", [])
            if candidates:
                loc = candidates[0].get("location", {})
                if 'y' in loc and 'x' in loc:
                    return float(loc['y']), float(loc['x'])
        except Exception:
            pass
        time.sleep(0.5)
    return None


# ── Overpass Queries ─────────────────────────────────────────────────────────
import time
def overpass_query(lat: float, lon: float, tag_filter: str, radius: int = 15000) -> list[dict]:
    query = f"""
    [out:json];
    (
      node[{tag_filter}](around:{radius},{lat},{lon});
      way[{tag_filter}](around:{radius},{lat},{lon});
      relation[{tag_filter}](around:{radius},{lat},{lon});
    );
    out center;
    """
    for attempt in range(4):
        try:
            resp = requests.post(
                "https://overpass-api.de/api/interpreter",
                data={"data": query}, headers=HEADERS, timeout=20
            )
            if resp.status_code == 429:
                time.sleep(5 * (attempt + 1))
                continue
            resp.raise_for_status()
            return resp.json().get("elements", [])
        except Exception:
            time.sleep(2)
            pass
    return []


def nearest_name(elements: list[dict], lat: float, lon: float) -> str:
    """Return name of the closest element to (lat, lon)."""
    best_name, best_dist = "", float("inf")
    for el in elements:
        el_lat = el.get("lat") or el.get("center", {}).get("lat")
        el_lon = el.get("lon") or el.get("center", {}).get("lon")
        if el_lat is None:
            continue
        dist = haversine(lat, lon, float(el_lat), float(el_lon))
        name = (el.get("tags") or {}).get("name") or (el.get("tags") or {}).get("name:en", "")
        if dist < best_dist and name:
            best_dist, best_name = dist, name
    return best_name


def find_nearest_station(lat: float, lon: float) -> str:
    els = overpass_query(lat, lon, '"railway"="station"')
    if not els:
        els = overpass_query(lat, lon, '"railway"="halt"')
    return nearest_name(els, lat, lon)


def find_nearest_auto_stand(lat: float, lon: float) -> str:
    els = overpass_query(lat, lon, '"amenity"="taxi"')
    if not els:
        els = overpass_query(lat, lon, '"highway"="bus_stop"', radius=3000)
    return nearest_name(els, lat, lon)


def find_nearest_bazaar(lat: float, lon: float) -> str:
    # Try market first, fallback to supermarket or mall
    for tag in ['"landuse"="retail"', '"shop"="supermarket"', '"amenity"="marketplace"']:
        els = overpass_query(lat, lon, tag)
        name = nearest_name(els, lat, lon)
        if name:
            return name
    return ""


# ── Train Lookup ─────────────────────────────────────────────────────────────
def find_train(destination_station: str) -> str:
    """
    Attempt to find a train from Sealdah to destination_station via erail.in search.
    Returns a short string like "Sealdah–Diamond Harbour Exp (34121)" or "" if not found.
    """
    if not destination_station:
        return ""
    try:
        # erail.in station search
        url = f"https://erail.in/rail/getStations.aspx?StationName={destination_station}&_=1"
        resp = requests.get(url, headers=HEADERS, timeout=10)
        data = resp.json()
        if not data:
            return f"From Sealdah to {destination_station}"
        dest_code = data[0].get("StationCode", "")
        if not dest_code:
            return f"From Sealdah to {destination_station}"

        # Search trains between Sealdah (SDAH) and destination
        train_url = (
            f"https://erail.in/rail/getTrains.aspx?"
            f"Station1=SDAH&Station2={dest_code}&RadioBtnSearchBy=Train&_=1"
        )
        tresp = requests.get(train_url, headers=HEADERS, timeout=10)
        trains = tresp.json()
        if trains:
            t = trains[0]
            return f"{t.get('TrainName','?')} ({t.get('TrainNo','?')})"
        return f"From Sealdah to {destination_station}"
    except Exception:
        return f"From Sealdah to {destination_station}"


# ── Main ─────────────────────────────────────────────────────────────────────
def main():
    if not CSV_PATH.exists():
        print("ERROR: extracted_data.csv not found. Run populate_data.py first.")
        return

    df = pd.read_csv(CSV_PATH, dtype=str).fillna("")

    transport_cols = ['Nearest Railway Station', 'Train from Sealdah', 'Auto Stand', 'Main Bazaar']
    for col in transport_cols:
        if col not in df.columns:
            df[col] = ""

    # Only process rows where transport data is missing or marked as '?' from a failed attempt
    needs_fill = df[
        (df['Nearest Railway Station'].str.strip() == "") | 
        (df['Nearest Railway Station'].str.strip() == "?")
    ].index.tolist()
    print(f"{len(needs_fill)} row(s) need transport data.\n")

    filled = 0
    for idx in needs_fill:
        row = df.loc[idx]
        village = str(row.get('VILLAGE', '')).strip()
        pin = str(row.get('PIN', '')).strip()
        name = str(row.get('NAME OF MADRASAH', '')).strip()

        print(f"[{idx+1}] {name} (PIN: {pin}, Village: {village})")

        coords = geocode(village, pin)
        if not coords:
            print(f"  Could not geocode — marking as '?'")
            df.loc[idx, 'Nearest Railway Station'] = "?"
            df.loc[idx, 'Train from Sealdah'] = "?"
            df.loc[idx, 'Auto Stand'] = "?"
            df.loc[idx, 'Main Bazaar'] = "?"
            continue

        lat, lon = coords
        print(f"  Coordinates: {lat:.4f}, {lon:.4f}")

        station = find_nearest_station(lat, lon)
        print(f"  Nearest station : {station or 'Not found'}")

        train = find_train(station) if station else ""
        print(f"  Train from Sealdah: {train or 'Not found'}")

        auto = find_nearest_auto_stand(lat, lon)
        print(f"  Auto stand      : {auto or 'Not found'}")

        bazaar = find_nearest_bazaar(lat, lon)
        print(f"  Main bazaar     : {bazaar or 'Not found'}")

        df.loc[idx, 'Nearest Railway Station'] = station or "?"
        df.loc[idx, 'Train from Sealdah'] = train or "?"
        df.loc[idx, 'Auto Stand'] = auto or "?"
        df.loc[idx, 'Main Bazaar'] = bazaar or "?"

        filled += 1
        time.sleep(NOMINATIM_DELAY)  # be polite to OSM

    df.to_csv(CSV_PATH, index=False)
    print(f"\nDone. Filled transport data for {filled}/{len(needs_fill)} rows.")
    print(f"Rows marked '?' need manual review.")


if __name__ == "__main__":
    main()
