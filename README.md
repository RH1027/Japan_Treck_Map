# Japan Trek Food Map

Restaurant recommendations for the McCombs Japan Trek, shown on a map.

- `index.html` — the page (MapLibre + OpenFreeMap tiles)
- `data/shops.json` — the restaurant list the page reads (generated)
- `data/locations.json` — address and map position for each restaurant
- `tools/build_data.py` — turns the recommendations Excel file into `data/shops.json`

## Updating the list

1. Add rows to the Excel file (one sheet per city: Tokyo, Kyoto, Osaka, Hiroshima, ...).
   Optional columns: `Tabelog`, `Google Map`, `Address`, `Google Rating`, `Tabelog Rating`, `Rating Date`.
2. Run `python3 tools/build_data.py path/to/Japan_Treck_recommendations_list.xlsx`.
3. Check `data/locations.json` for any new or missing positions, then commit and push.

Map tiles: OpenFreeMap · © OpenMapTiles · Data © OpenStreetMap contributors. Labels use English names where available.
