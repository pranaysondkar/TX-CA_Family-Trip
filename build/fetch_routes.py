"""Fetches driving routes (OpenStreetMap / OSRM) for trip_data.ROUTES into routes.json.

Run once, or again after changing ROUTES:  python3 fetch_routes.py
"""
import json
import os
import subprocess
import time

import trip_data as T

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "routes.json")

def thin(coords, keep=350):
    step = max(1, len(coords) // keep)
    pts = coords[::step] + [coords[-1]]
    return [[round(lat, 4), round(lon, 4)] for lon, lat in pts]


routes = {}
for date, stops in T.ROUTES.items():
    coords = ";".join(f"{T.P[s][2]},{T.P[s][1]}" for s in stops)
    url = f"https://router.project-osrm.org/route/v1/driving/{coords}?overview=full&geometries=geojson"
    raw = subprocess.run(["curl", "-s", "-m", "30", url], capture_output=True, check=True).stdout
    r = json.loads(raw)["routes"][0]
    routes[date] = {
        "line": thin(r["geometry"]["coordinates"]),
        "miles": round(r["distance"] / 1609.34),
        "minutes": round(r["duration"] / 60),
        "legs": [{"miles": round(l["distance"] / 1609.34), "minutes": round(l["duration"] / 60)} for l in r["legs"]],
    }
    print(date, routes[date]["miles"], "mi", routes[date]["minutes"], "min", len(routes[date]["line"]), "pts")
    time.sleep(1)
json.dump(routes, open(OUT, "w"))
