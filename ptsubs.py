import re
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
from PIL import Image, ImageDraw
import pandas as pd
from pathlib import Path


BASE_DIR = Path(r"")
INPUT_XLSX = BASE_DIR / "PT_SUBS_FINAL_UPDATED.xlsx"     
OUTPUT_KMZ = BASE_DIR / "PT_SUBS_FINAL_UPDATED_ALL.kmz"
OUTPUT_GPKG = BASE_DIR / "PT_SUBS_FINAL_UPDATED_ALL.gpkg"
OUTPUT_EXCEL_DEC = BASE_DIR / "PT_SUBS_FINAL_WITH_DECIMAL.xlsx"
INVALID_CSV = BASE_DIR / "PT_SUBS_INVALID_COORDS.csv"
KML_FILENAME = "doc.kml"
ICON_FILENAME = "substation.png"
NAME_COL_PRIORITY = ["Substation", "Name", "Substation Name", "substation_name"]


def try_float(val):
    if pd.isna(val):
        return None
    s = str(val).strip()
    # se tiver sinais de DMS, não é decimal puro
    if any(ch in s for ch in ["°", "'", "’", "N", "S", "E", "W", "º", "″", "”", "′"]):
        return None
    s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None


DMS_PATTERN = re.compile(
    r"""
    (?P<hem1>[NnSsEeWw])?\s*
    (?P<deg>\d{1,3})\s*(?:°|º|\s)\s*
    (?P<min>\d{1,2})\s*(?:'|’|′|\s)\s*
    (?P<sec>\d{1,2}(?:[\.,]\d+)?)?\s*(?:"|”|″)?\s*
    (?P<hem2>[NnSsEeWw])?
    """,
    re.VERBOSE
)

def dms_to_decimal(deg, minutes, seconds, hemisphere):
    deg = float(deg)
    minutes = float(minutes) if minutes is not None else 0.0
    seconds = float(str(seconds).replace(",", ".")) if seconds not in (None, "") else 0.0
    dec = deg + minutes/60 + seconds/3600
    if hemisphere and hemisphere.upper() in ["S", "W"]:
        dec = -dec
    return dec

def parse_coord(value):
    """Devolve graus decimais a partir de decimal (., ,) ou DMS; caso contrário None."""
    if pd.isna(value):
        return None
    s = str(value).strip()
    
    f = try_float(s)
    if f is not None:
        return f
    
    m = DMS_PATTERN.search(s)
    if m:
        hem = m.group("hem1") or m.group("hem2") or ""
        return dms_to_decimal(m.group("deg"), m.group("min"), m.group("sec"), hem)
    return None

def build_icon_png(size=96) -> BytesIO:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.ellipse((4, 4, size-4, size-4), fill=(30, 144, 255, 230), outline=(0, 0, 0, 180), width=4)
    bolt = [
        (size*0.55, size*0.18),
        (size*0.38, size*0.52),
        (size*0.55, size*0.52),
        (size*0.45, size*0.84),
        (size*0.68, size*0.48),
        (size*0.52, size*0.48),
    ]
    draw.polygon(bolt, fill=(255, 215, 0, 255))
    bio = BytesIO()
    img.save(bio, format="PNG")
    bio.seek(0)
    return bio

def esc(val):
    return str(val).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def write_gpkg(valid_df, name_col, lon_field="_lon_", lat_field="_lat_"):
    try:
        import geopandas as gpd
        from shapely.geometry import Point
    except Exception:
        print("Aviso: GeoPackage não gerado (geopandas/shapely não instalados).")
        print("Instale com: pip install geopandas shapely pyproj fiona")
        return

    gdf = valid_df.copy()
    
    gdf["geometry"] = gdf.apply(lambda r: Point(float(r[lon_field]), float(r[lat_field])), axis=1)
    gdf = gpd.GeoDataFrame(gdf, geometry="geometry", crs="EPSG:4326")

    
    drop_cols = [c for c in ["_lat_", "_lon_"] if c in gdf.columns]
    gdf = gdf.drop(columns=drop_cols, errors="ignore")

    OUTPUT_GPKG.parent.mkdir(parents=True, exist_ok=True)
    gdf.to_file(OUTPUT_GPKG, layer="substations", driver="GPKG")
    print(f"GPKG gravado: {OUTPUT_GPKG}")

def main():
    if not INPUT_XLSX.exists():
        raise FileNotFoundError(f"Ficheiro não encontrado: {INPUT_XLSX}")

    df = pd.read_excel(INPUT_XLSX)

    
    lat_candidates = [c for c in df.columns if str(c).strip().lower() in ["latitude", "lat"]]
    lon_candidates = [c for c in df.columns if str(c).strip().lower() in ["longitude", "lon", "long"]]
    lat_col = lat_candidates[0] if lat_candidates else df.columns[4]   # E
    lon_col = lon_candidates[0] if lon_candidates else df.columns[5]   # F

    
    name_col = next((c for c in NAME_COL_PRIORITY if c in df.columns), None)

    
    df["Latitude_decimal"]  = df[lat_col].apply(parse_coord)
    df["Longitude_decimal"] = df[lon_col].apply(parse_coord)

    
    valid_mask = df["Latitude_decimal"].notna() & df["Longitude_decimal"].notna()
    valid = df[valid_mask].copy()
    invalid = df[~valid_mask].copy()

    
    icon_bytes = build_icon_png()
    
    cols_for_desc = list(df.columns)

    kml_parts = []
    kml_parts.append(f"""<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Document>
    <name>PT Substations (All)</name>
    <Style id="substationStyle">
      <IconStyle>
        <scale>1.1</scale>
        <Icon><href>{ICON_FILENAME}</href></Icon>
      </IconStyle>
      <LabelStyle><scale>0.9</scale></LabelStyle>
    </Style>
    <Folder><name>Substations</name>
""")

    for idx, row in valid.iterrows():
        lat = row["Latitude_decimal"]
        lon = row["Longitude_decimal"]
        name = str(row[name_col]) if name_col else f"Site {idx+1}"

        html_rows = ["<table border='1' cellpadding='3' cellspacing='0'>"]
        for col in cols_for_desc:
            val = row.get(col, "")
            html_rows.append(f"<tr><th align='left'>{esc(col)}</th><td>{esc(val)}</td></tr>")
        html_rows.append("</table>")
        desc_html = "".join(html_rows)

        kml_parts.append(f"""
      <Placemark>
        <name>{esc(name)}</name>
        <styleUrl>#substationStyle</styleUrl>
        <description><![CDATA[{desc_html}]]></description>
        <Point><coordinates>{lon},{lat},0</coordinates></Point>
      </Placemark>
""")

    kml_parts.append("    </Folder>\n  </Document>\n</kml>\n")
    kml_bytes = "".join(kml_parts).encode("utf-8")

    OUTPUT_KMZ.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(OUTPUT_KMZ, "w", ZIP_DEFLATED) as zf:
        zf.writestr(KML_FILENAME, kml_bytes)
        zf.writestr(ICON_FILENAME, icon_bytes.read())
    print(f"KMZ gravado: {OUTPUT_KMZ}")

    
    valid["_lat_"] = valid["Latitude_decimal"]
    valid["_lon_"] = valid["Longitude_decimal"]
    write_gpkg(valid, name_col)


    with pd.ExcelWriter(OUTPUT_EXCEL_DEC, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Substations_Decimal")
    print(f"Excel com coordenadas decimais gravado: {OUTPUT_EXCEL_DEC}")

  
    if not invalid.empty:
        rep_cols = []
        if name_col:
            rep_cols.append(name_col)
        rep_cols.extend([lat_col, lon_col, "Latitude_decimal", "Longitude_decimal"])
        rep_cols = [c for c in rep_cols if c in df.columns or c in ["Latitude_decimal","Longitude_decimal"]]
        invalid[rep_cols].to_csv(INVALID_CSV, index=False, encoding="utf-8-sig")
        print(f"CSV inválidos: {INVALID_CSV}")

    print(f"Linhas totais: {len(df)} | Incluídas: {len(valid)} | Excluídas: {len(invalid)}")

if __name__ == "__main__":
    main()
