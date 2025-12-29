
# PT Substations — KMZ & GPKG Exporter

A Python tool to process substation list in Excel extracted from E-Redes Open Data, **normalize coordinates** (supports Decimal Degrees and **DMS** formats), and export:
- A **KMZ** file (with custom icon and rich popups),
- An optional **GeoPackage (GPKG)** layer (if GeoPandas/Shapely are installed),
- A new Excel with **decimal coordinates**,
- A CSV report of **invalid coordinates**.

This is designed for datasets like `PT_SUBS_FINAL_UPDATED.xlsx` containing Portuguese substations, but it can be adapted to other countries/contexts.

---

## Features

- 🧭 **Coordinate parsing**:
  - Accepts Decimal Degrees using either `.` or `,` decimals
  - Accepts **DMS** (Degrees-Minutes-Seconds) with optional hemisphere (N/S/E/W) indicators
- 🖼️ **KMZ export** with a **custom icon** (blue circle + lightning bolt)
- 🧩 **Rich popup tables** in KML describing all row attributes
- 🧱 **GeoPackage export** (EPSG:4326) using optional GeoPandas/Shapely
- 📓 **Excel output** with Decimal Degrees appended
- 🚨 **Invalid coordinates report** (`.csv`) for rows that couldn’t be parsed
- 🔁 **Automatic name selection** using a priority list:
  - `Substation`, `Name`, `Substation Name`, `substation_name` (fallback to `Site {index}`)

---

## Requirements

### Python
- Python 3.9+ recommended

### Packages (required)
- `pandas`
- `Pillow` (PIL)
- `openpyxl` (for Excel writing)

### Optional (for GPKG export)
- `geopandas`
- `shapely`
- (these may require `fiona`, `pyproj`, GDAL stack depending on your environment)

Install the core dependencies:

```bash
pip install pandas pillow openpyxl
