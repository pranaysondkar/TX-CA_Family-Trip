# Family Trip · Dec 6, 2026 – Jan 4, 2027

Live day-by-day itinerary (Austin & Hill Country + California road trip) as a static site.

- `index.html` – the page (open it, or host on GitHub Pages)
- `data.js` – generated trip data (don't edit by hand)
- `img/` – photos from Wikimedia Commons under free licenses (credits listed on the page)
- `build/trip_data.py` – **edit plans and prices here**, then run:

```
python3 build/build_site.py      # regenerates data.js
python3 build/build_excel.py     # regenerates the Excel plan (needs openpyxl)
```

The page shows a countdown or "Day N of 30", highlights today, pulls live forecasts
(Open-Meteo, about 16 days ahead) and live National Weather Service alerts, and saves
check-offs in the browser. Use **Save as PDF** for an offline copy.

## Live bookings & actual costs
One Google Sheet (the uploaded `Family_Trip_Dec2026_Plan.xlsx`) is the shared record. Its **Bookings** tab is read and
written by the website through a small Google Apps Script web app (`build/apps_script.gs`):
tick *Booked* and type the actual amount on the website, enter the family PIN once, and it saves for everyone.
Put the web-app URL in `BOOKINGS_API` (and the sheet link in `BOOKINGS_EDIT`) in `build/trip_data.py`, then rebuild.

## Publish on GitHub Pages
1. Create a repo (e.g. `family-trip-2026`) and push this folder's contents.
2. Repo → Settings → Pages → Source: *Deploy from a branch* → `main` / root.
3. The site appears at `https://<your-username>.github.io/family-trip-2026/`.
