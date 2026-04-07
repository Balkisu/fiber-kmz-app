import io
import re
import time
import zipfile
from collections import defaultdict
from typing import Any, Dict, List, Tuple

import folium
import pandas as pd
import requests
import streamlit as st
from folium.plugins import Draw
from streamlit_folium import st_folium

st.set_page_config(page_title="Fiber Route → KMZ", layout="wide")
st.title("Fiber Backbone Route → KMZ Generator")
st.caption("Upload an Excel/CSV, enter data manually, or draw points on the map.")

ORS_API_KEY = "eyJvcmciOiI1YjNjZTM1OTc4NTExMTAwMDFjZjYyNDgiLCJpZCI6IjUyYzhkNWVkYjFlNDQzZmNiYTVmNWE3MDJmNjcxZGQwIiwiaCI6Im11cm11cjY0In0="
ORS_URL = "https://api.openrouteservice.org/v2/directions/driving-car/geojson"
MAX_WAYPOINTS = 45

ROUTE_COLORS = [
    "ff0000ff", "ff00cc00", "ff00aaff", "ffff00ff",
    "ff00ffff", "ffff8800", "ff0088ff", "ff8800ff",
]

# ── Init session state ────────────────────────────────────────────────────────
for key in ["routes", "kmz_bytes", "kmz_filename", "debug_log", "df_upload", "upload_filename", "df_draw"]:
    if key not in st.session_state:
        st.session_state[key] = None

# ── Coordinate parsing ────────────────────────────────────────────────────────
_DMS_RE = re.compile(
    r"""^\s*(?P<deg>\d{1,3})\s*[°d]\s*
        (?P<min>\d{1,2})\s*[']\s*
        (?P<sec>\d{1,2}(?:\.\d+)?)\s*["]?\s*
        (?P<dir>[NSEW])\s*$""",
    re.VERBOSE | re.IGNORECASE,
)

def to_decimal(coord) -> float:
    if coord is None or (isinstance(coord, float) and pd.isna(coord)):
        raise ValueError("Empty coordinate")
    if isinstance(coord, (int, float)):
        return float(coord)
    s = str(coord).strip()
    try:
        return float(s)
    except ValueError:
        pass
    m = _DMS_RE.match(s)
    if not m:
        raise ValueError(f"Cannot parse: {s!r}")
    dec = float(m.group("deg")) + float(m.group("min")) / 60 + float(m.group("sec")) / 3600
    if m.group("dir").upper() in ("S", "W"):
        dec = -dec
    return dec

# ── ORS routing ───────────────────────────────────────────────────────────────
def ors_segment(points_lonlat: List[Tuple[float, float]], log: list) -> Tuple[List[Tuple[float, float]], float]:
    headers = {"Authorization": ORS_API_KEY, "Content-Type": "application/json"}
    payload = {
        "coordinates": [[lon, lat] for lon, lat in points_lonlat],
        "instructions": False,
        "radiuses": [-1] * len(points_lonlat),
    }
    log.append(f"  → Calling ORS with {len(points_lonlat)} waypoints...")
    try:
        r = requests.post(ORS_URL, json=payload, headers=headers, timeout=60)
        log.append(f"  → Status: {r.status_code}")
        if r.status_code != 200:
            log.append(f"  → Error body: {r.text[:300]}")
            return list(points_lonlat), 0.0
        data = r.json()
        if "features" not in data or not data["features"]:
            log.append(f"  → Unexpected response: {str(data)[:300]}")
            return list(points_lonlat), 0.0
        coords = [(float(c[0]), float(c[1])) for c in data["features"][0]["geometry"]["coordinates"]]
        dist = float(data["features"][0]["properties"]["summary"]["distance"])
        log.append(f"  → OK: {len(coords)} road points, {dist/1000:.2f} km")
        return coords, dist
    except Exception as e:
        log.append(f"  → Exception: {e}")
        return list(points_lonlat), 0.0

def ors_route(points_lonlat: List[Tuple[float, float]], log: list) -> Tuple[List[Tuple[float, float]], float]:
    if len(points_lonlat) < 2:
        return list(points_lonlat), 0.0
    chunks, i = [], 0
    while i < len(points_lonlat) - 1:
        end = min(i + MAX_WAYPOINTS, len(points_lonlat))
        chunks.append(points_lonlat[i:end])
        i = end - 1
    all_coords: List[Tuple[float, float]] = []
    total_dist = 0.0
    for chunk in chunks:
        seg_coords, seg_dist = ors_segment(chunk, log)
        total_dist += seg_dist
        if not all_coords:
            all_coords.extend(seg_coords)
        else:
            all_coords.extend(seg_coords[1:] if seg_coords else [])
        time.sleep(0.5)
    return all_coords, total_dist

# ── KML / KMZ ─────────────────────────────────────────────────────────────────
def _x(s): return (s or "").replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def build_kml(routes: Dict[str, Dict[str, Any]]) -> str:
    lines = ['<?xml version="1.0" encoding="UTF-8"?>',
             '<kml xmlns="http://www.opengis.net/kml/2.2">',
             "<Document>", "  <name>Proposed Fiber Backbone Routes</name>"]
    for route_name, info in routes.items():
        sid, color = info["style_id"], info["color"]
        lines += [f'  <Style id="{sid}_n"><LineStyle><color>{color}</color><width>3</width></LineStyle><PolyStyle><fill>0</fill></PolyStyle></Style>',
                  f'  <Style id="{sid}_h"><LineStyle><color>{color}</color><width>5</width></LineStyle><PolyStyle><fill>0</fill></PolyStyle></Style>',
                  f'  <StyleMap id="{sid}"><Pair><key>normal</key><styleUrl>#{sid}_n</styleUrl></Pair><Pair><key>highlight</key><styleUrl>#{sid}_h</styleUrl></Pair></StyleMap>']
    for route_name, info in routes.items():
        sid = info["style_id"]
        dist_km = info.get("distance_km")
        coord_str = " ".join(f"{lon},{lat},0" for lon, lat in info["line"])
        title = f"{route_name} ({dist_km:.2f} km)" if dist_km else route_name
        lines += ["  <Folder>", f"    <name>{_x(route_name)}</name>",
                  "    <Placemark>", f"      <name>{_x(title)}</name>",
                  f"      <styleUrl>#{sid}</styleUrl>",
                  "      <LineString><tessellate>1</tessellate>",
                  f"        <coordinates>{coord_str}</coordinates>",
                  "      </LineString>", "    </Placemark>"]
        for s in info["sites"]:
            label = f'{s["SiteID"]} - {s["Site Name"]}'
            lines += ["    <Placemark>", f"      <name>{_x(label)}</name>",
                      "      <Style><IconStyle><Icon><href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png</href></Icon></IconStyle>",
                      "      <LabelStyle><scale>0.9</scale></LabelStyle></Style>",
                      f'      <Point><coordinates>{s["Longitude"]},{s["Latitude"]},0</coordinates></Point>',
                      "    </Placemark>"]
        lines.append("  </Folder>")
    lines += ["</Document>", "</kml>"]
    return "\n".join(lines)

def to_kmz(kml: str) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("doc.kml", kml)
    return buf.getvalue()

# ── Column normalisation ──────────────────────────────────────────────────────
def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    col_map = {c.strip().lower(): c for c in df.columns}
    def pick(*names):
        for n in names:
            if n in col_map: return col_map[n]
        return None
    needed = {
        "SiteID":     pick("siteid","site id","site_id"),
        "Site Name":  pick("site name","sitename","site_name"),
        "Longitude":  pick("longitude","lon","long"),
        "Latitude":   pick("latitude","lat"),
        "Route Name": pick("route name","routename","route_name"),
    }
    missing = [k for k,v in needed.items() if v is None]
    if missing: raise ValueError(f"Missing columns: {missing}")
    out = df[[needed[k] for k in needed]].copy()
    out.columns = list(needed.keys())
    return out

def df_to_routes(df: pd.DataFrame) -> Dict[str, Dict[str, Any]]:
    df = df.dropna(subset=["Route Name"]).copy()
    df["Route Name"] = df["Route Name"].astype(str).str.strip()
    df = df[df["Route Name"] != ""]
    if df.empty: raise ValueError("No valid rows with a Route Name.")
    df["Longitude"] = df["Longitude"].apply(to_decimal)
    df["Latitude"]  = df["Latitude"].apply(to_decimal)

    grouped: Dict[str, List[dict]] = defaultdict(list)
    for _, row in df.iterrows():
        grouped[str(row["Route Name"])].append({
            "SiteID":    str(row["SiteID"]).strip(),
            "Site Name": str(row["Site Name"]).strip(),
            "Longitude": float(row["Longitude"]),
            "Latitude":  float(row["Latitude"]),
        })

    routes: Dict[str, Dict[str, Any]] = {}
    route_names = list(grouped.keys())
    log = []
    progress = st.progress(0, text="Starting...")

    for i, (route_name, sites) in enumerate(grouped.items()):
        progress.progress(i / len(route_names), text=f"Routing: {route_name}...")
        log.append(f"\n=== Route {i+1}/{len(route_names)}: {route_name} ({len(sites)} sites) ===")
        points = [(s["Longitude"], s["Latitude"]) for s in sites]
        line, dist_m = ors_route(points, log)
        style_id = "r_" + re.sub(r"[^a-zA-Z0-9]", "_", route_name)[:50]
        routes[route_name] = {
            "style_id":    style_id,
            "color":       ROUTE_COLORS[i % len(ROUTE_COLORS)],
            "sites":       sites,
            "line":        line,
            "distance_km": round(dist_m / 1000, 2),
        }

    progress.progress(1.0, text="Done!")
    st.session_state.debug_log = "\n".join(log)
    return routes

# ── Tabs ──────────────────────────────────────────────────────────────────────
tab_upload, tab_manual, tab_draw = st.tabs(["Upload Excel / CSV", "Manual Entry", "Draw on Map"])

df_manual_out = None
df_draw = None

with tab_upload:
    uploaded = st.file_uploader("Upload .xlsx or .csv", type=["xlsx", "csv"], key="file_uploader")
    if uploaded is not None:
        try:
            raw = pd.read_csv(uploaded) if uploaded.name.lower().endswith(".csv") else pd.read_excel(uploaded)
            st.session_state.df_upload = normalize_df(raw)
            st.session_state.upload_filename = uploaded.name
        except Exception as e:
            st.error(str(e))

    if st.session_state.df_upload is not None:
        st.success(f"File loaded: {st.session_state.get('upload_filename', '')} — {len(st.session_state.df_upload)} rows")
        st.dataframe(st.session_state.df_upload, use_container_width=True)
        if st.button("Clear uploaded file"):
            st.session_state.df_upload = None
            st.session_state.upload_filename = None
            st.rerun()

    with tab_manual:
        if "manual_df" not in st.session_state:
        st.session_state.manual_df = pd.DataFrame([
            {"SiteID": "T5424", "Site Name": "Biu Road",     "Longitude": 12.1731,   "Latitude": 10.6150,  "Route Name": "Biu-Little Gombi"},
            {"SiteID": "T5153", "Site Name": "New layout",   "Longitude": 12.5669,   "Latitude": 10.3989,  "Route Name": "Biu-Little Gombi"},
            {"SiteID": "T5062", "Site Name": "BUK Old Site", "Longitude": 8.47875,   "Latitude": 11.97625, "Route Name": "BUK-Funtua"},
            {"SiteID": "T5076", "Site Name": "Gwarzo",       "Longitude": 8.653306,  "Latitude": 12.176472,"Route Name": "BUK-Funtua"},
        ])

        st.write("Enter sites below in route order.")

        df_manual_out = st.data_editor(
        st.session_state.manual_df,
        num_rows="dynamic",
        use_container_width=True,
        key="manual_ed",
        column_config={
            "Longitude": st.column_config.NumberColumn("Longitude", format="%.6f"),
            "Latitude":  st.column_config.NumberColumn("Latitude",  format="%.6f"),
        }
    )



with tab_draw:
    if "draw_pts" not in st.session_state:
        st.session_state.draw_pts = []
    route_draw = st.text_input("Route Name", value="Drawn Route")
    prefix = st.text_input("SiteID prefix", value="PT")
    m = folium.Map(location=[9.08, 8.67], zoom_start=6, control_scale=True)
    Draw(export=False,
         draw_options={"marker": True, "polyline": False, "polygon": False,
                       "rectangle": False, "circle": False, "circlemarker": False},
         edit_options={"edit": False, "remove": True}).add_to(m)
    for i, p in enumerate(st.session_state.draw_pts, 1):
        folium.Marker([p["lat"], p["lon"]], tooltip=f"{prefix}{i:03d}").add_to(m)
    map_out = st_folium(m, width="100%", height=500)
    if map_out and map_out.get("last_active_drawing"):
        geom = map_out["last_active_drawing"].get("geometry", {})
        if geom.get("type") == "Point":
            lon, lat = geom["coordinates"]
            last = st.session_state.draw_pts
            if not last or last[-1] != {"lon": lon, "lat": lat}:
                st.session_state.draw_pts.append({"lon": float(lon), "lat": float(lat)})
                st.rerun()
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Clear drawn points"):
            st.session_state.draw_pts = []
            st.rerun()
    with col2:
        st.write(f"Points placed: **{len(st.session_state.draw_pts)}**")
    if st.session_state.draw_pts:
        st.session_state.df_draw = pd.DataFrame([
            {"SiteID": f"{prefix}{i:03d}", "Site Name": f"Point {i}",
             "Longitude": p["lon"], "Latitude": p["lat"], "Route Name": route_draw}
            for i, p in enumerate(st.session_state.draw_pts, 1)
        ])
        st.dataframe(st.session_state.df_draw, use_container_width=True)
    else:
        st.session_state.df_draw = None

# ── Generate ──────────────────────────────────────────────────────────────────
st.divider()

# Auto-select data source: Upload > Draw > Manual
if st.session_state.df_upload is not None:
    df_in_auto = st.session_state.df_upload
    source_label = f"Uploaded file ({st.session_state.get('upload_filename', '')})"
elif st.session_state.df_draw is not None:
    df_in_auto = st.session_state.df_draw
    source_label = "Drawn points"
else:
    df_in_auto = normalize_df(df_manual_out) if df_manual_out is not None else None
    source_label = "Manual entry"

st.info(f"Data source: **{source_label}**")

filename = st.text_input("Output filename", value="fiber_routes.kmz")
if not filename.lower().endswith(".kmz"):
    filename += ".kmz"

if st.button("Generate KMZ", type="primary"):
    st.session_state.routes = None
    st.session_state.kmz_bytes = None
    st.session_state.debug_log = None
    try:
        df_in = df_in_auto

        if df_in is None or df_in.empty:
            st.error("No data to process.")
        else:
            routes = df_to_routes(df_in)
            kml = build_kml(routes)
            st.session_state.routes = routes
            st.session_state.kmz_bytes = to_kmz(kml)
            st.session_state.kmz_filename = filename

    except Exception as e:
        st.error(f"Error: {e}")
        import traceback
        st.session_state.debug_log = traceback.format_exc()

# ── Results (persistent) ──────────────────────────────────────────────────────
if st.session_state.debug_log:
    with st.expander("Debug log", expanded=False):
        st.code(st.session_state.debug_log)

if st.session_state.routes:
    routes = st.session_state.routes

    st.subheader("Route Preview")
    pm = folium.Map(location=[9.08, 8.67], zoom_start=6)
    for rname, info in routes.items():
        for s in info["sites"]:
            folium.CircleMarker(
                [s["Latitude"], s["Longitude"]], radius=5,
                tooltip=f'{s["SiteID"]} - {s["Site Name"]}',
                fill=True, fill_opacity=0.9,
            ).add_to(pm)
        folium.PolyLine(
            [(lat, lon) for lon, lat in info["line"]],
            weight=3,
            tooltip=f'{rname} ({info["distance_km"]} km)',
        ).add_to(pm)
    st_folium(pm, width="100%", height=520)

    st.subheader("Summary")
    st.dataframe(pd.DataFrame([
        {"Route": r, "Sites": len(v["sites"]), "Distance (km)": v["distance_km"]}
        for r, v in routes.items()
    ]), use_container_width=True)

if st.session_state.kmz_bytes:
    st.download_button(
        label="⬇️ Download KMZ",
        data=st.session_state.kmz_bytes,
        file_name=st.session_state.kmz_filename,
        mime="application/vnd.google-earth.kmz",
    )
