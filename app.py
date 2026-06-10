import os
import html
from datetime import date, timedelta
from typing import List, Dict, Any

from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, JSONResponse
from openpyxl import load_workbook

EXCEL_FILE = "AIRPORT_ZONE.xlsx"
SHEET_NAME = "AIRPORT_ZONE"
HUB = "DXB"
HUB_ZONE = "Z09"

app = FastAPI(title="Airport - Zone OD Calculator")

AIRPORT_ZONE: Dict[str, str] = {}
AIRPORTS: List[str] = []
STARTUP_ERROR = ""
HIT_COUNT = 0


def increment_hit_count() -> int:
    global HIT_COUNT
    HIT_COUNT += 1
    return HIT_COUNT


def get_hit_count() -> int:
    return HIT_COUNT


def norm(v) -> str:
    return str(v or "").strip().upper()


def zone_num(z: str) -> int:
    try:
        return int(str(z).upper().replace("Z", ""))
    except Exception:
        return 0


def direction(z1: str, z2: str) -> int:
    a = zone_num(z1)
    b = zone_num(z2)
    if b > a:
        return 1
    if b < a:
        return -1
    return 0


def od_pair(o: str, d: str) -> str:
    return f"{o}{d}"


def sector_pair(s: Dict[str, Any]) -> str:
    return f"{s['origin']}{s['destination']}"


def load_airport_zone():
    global AIRPORT_ZONE, AIRPORTS, STARTUP_ERROR

    AIRPORT_ZONE = {}
    AIRPORTS = []
    STARTUP_ERROR = ""

    if not os.path.exists(EXCEL_FILE):
        STARTUP_ERROR = f"Airport zone mapping not found. Please place {EXCEL_FILE} in same folder as app.py"
        return

    wb = load_workbook(EXCEL_FILE, data_only=True)

    if SHEET_NAME not in wb.sheetnames:
        STARTUP_ERROR = f"Sheet {SHEET_NAME} not found in {EXCEL_FILE}"
        return

    ws = wb[SHEET_NAME]
    headers = [norm(c.value) for c in ws[1]]

    for required in ["AIRLINE", "ZONE", "AIRPORT"]:
        if required not in headers:
            STARTUP_ERROR = "Excel must have columns AIRLINE, ZONE, AIRPORT"
            return

    zone_idx = headers.index("ZONE")
    airport_idx = headers.index("AIRPORT")

    for row in ws.iter_rows(min_row=2, values_only=True):
        airport = norm(row[airport_idx])
        zone = norm(row[zone_idx])
        if airport and zone:
            AIRPORT_ZONE[airport] = zone

    if HUB not in AIRPORT_ZONE:
        AIRPORT_ZONE[HUB] = HUB_ZONE

    AIRPORTS = sorted(AIRPORT_ZONE.keys())


load_airport_zone()


def get_zone(airport: str) -> str:
    airport = norm(airport)
    if airport not in AIRPORT_ZONE:
        raise ValueError(f"Airport {airport} not found in AIRPORT_ZONE Excel mapping.")
    return AIRPORT_ZONE[airport]


def default_rows():
    return [{
        "sno": 1,
        "departure_date": date.today().isoformat(),
        "origin": "",
        "destination": "",
        "travel_class": "Y",
    }]


def rows_from_itinerary_string(text: str) -> List[Dict[str, Any]]:
    clean = norm(text)
    clean = clean.replace("➜", "").replace("->", "").replace("-", "")
    clean = clean.replace(",", " ").replace("/", " ").replace("|", " ")
    clean = clean.replace("\n", " ").replace("\t", " ").replace(" ", "")

    rows = []
    if len(clean) >= 6:
        idx = 1
        for pos in range(0, len(clean), 6):
            sector = clean[pos:pos + 6]
            if len(sector) < 6:
                continue

            rows.append({
                "sno": idx,
                "departure_date": date.today().isoformat(),
                "origin": sector[:3],
                "destination": sector[3:6],
                "travel_class": "Y",
            })
            idx += 1

    return rows or default_rows()


def build_sectors(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    sectors = []

    for r in rows:
        o = norm(r.get("origin"))
        d = norm(r.get("destination"))

        if not o or not d:
            continue

        oz = get_zone(o)
        dz = get_zone(d)

        sectors.append({
            "sno": r.get("sno"),
            "date": r.get("departure_date"),
            "origin": o,
            "destination": d,
            "origin_zone": oz,
            "destination_zone": dz,
            "class": norm(r.get("travel_class")) or "Y",
        })

    return sectors


def group_by_class(sectors: List[Dict[str, Any]]) -> List[List[Dict[str, Any]]]:
    groups = []
    current = []

    for s in sectors:
        if not current:
            current.append(s)
        elif s["class"] == current[-1]["class"]:
            current.append(s)
        else:
            groups.append(current)
            current = [s]

    if current:
        groups.append(current)

    return groups


def find_side_trip_indexes(sectors: List[Dict[str, Any]]) -> set:
    side_indexes = set()

    if len(sectors) < 5:
        return side_indexes

    trip_origin = sectors[0]["origin"]

    if trip_origin == HUB:
        return side_indexes

    hub_loops = []
    i = 0

    while i < len(sectors):
        if sectors[i]["origin"] == HUB and sectors[i]["destination"] != HUB:
            start = i
            end = None

            for j in range(i + 1, len(sectors)):
                if sectors[j]["origin"] != HUB and sectors[j]["destination"] == HUB:
                    end = j
                    break

            if end is not None:
                hub_loops.append((start, end))
                i = end + 1
            else:
                i += 1
        else:
            i += 1

    if not hub_loops:
        return side_indexes

    main_loop = max(hub_loops, key=lambda pair: pair[1] - pair[0])

    for start, end in hub_loops:
        if (start, end) == main_loop:
            continue
        for idx in range(start, end + 1):
            side_indexes.add(idx)

    return side_indexes


def build_virtual_legs(sectors: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    legs = []
    i = 0

    while i < len(sectors):
        s = sectors[i]

        if (
            i + 1 < len(sectors)
            and s["destination"] == HUB
            and sectors[i + 1]["origin"] == HUB
            and s["origin"] != sectors[i + 1]["destination"]
        ):
            n = sectors[i + 1]
            legs.append({
                "origin": s["origin"],
                "destination": n["destination"],
                "origin_zone": s["origin_zone"],
                "destination_zone": n["destination_zone"],
                "class": s["class"],
                "component_sectors": [sector_pair(s), sector_pair(n)],
            })
            i += 2
        else:
            legs.append({
                "origin": s["origin"],
                "destination": s["destination"],
                "origin_zone": s["origin_zone"],
                "destination_zone": s["destination_zone"],
                "class": s["class"],
                "component_sectors": [sector_pair(s)],
            })
            i += 1

    return legs


def classify_return(ods: List[Dict[str, Any]]) -> str:
    if len(ods) <= 1:
        return "ONE WAY"

    first = ods[0]
    last = ods[-1]

    if first["origin"] == last["destination"] and first["destination"] == last["origin"]:
        return "RETURN MIRRORED JOURNEY"

    return "RETURN NON MIRRORED JOURNEY"


def make_od_from_leg(leg: Dict[str, Any], turnaround="NO", reason="Sector retained as OD."):
    return {
        "origin": leg["origin"],
        "destination": leg["destination"],
        "origin_zone": leg["origin_zone"],
        "destination_zone": leg["destination_zone"],
        "class": leg["class"],
        "component_sectors": leg["component_sectors"],
        "turnaround": turnaround,
        "reasoning": reason,
    }


def calculate_main_ods(sectors: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    if not sectors:
        return []

    trip_origin = sectors[0]["origin"]
    legs = build_virtual_legs(sectors)

    if all(s["origin"] != HUB and s["destination"] != HUB for s in sectors):
        airports = [sectors[0]["origin"]] + [s["destination"] for s in sectors]
        zones = [sectors[0]["origin_zone"]] + [s["destination_zone"] for s in sectors]

        non_zero_signs = [
            direction(zones[i - 1], zones[i])
            for i in range(1, len(zones))
            if direction(zones[i - 1], zones[i]) != 0
        ]

        if not non_zero_signs:
            return [
                {
                    "origin": s["origin"],
                    "destination": s["destination"],
                    "origin_zone": s["origin_zone"],
                    "destination_zone": s["destination_zone"],
                    "class": s["class"],
                    "component_sectors": [sector_pair(s)],
                    "turnaround": "NO",
                    "reasoning": "Same-zone non-hub sectors retained sector-wise.",
                }
                for s in sectors
            ]

        if non_zero_signs[0] != non_zero_signs[-1]:
            turn_index = len(airports) // 2

            return [
                {
                    "origin": airports[0],
                    "destination": airports[turn_index],
                    "origin_zone": zones[0],
                    "destination_zone": zones[turn_index],
                    "class": sectors[0]["class"],
                    "component_sectors": [sector_pair(s) for s in sectors[:turn_index]],
                    "turnaround": "YES",
                    "reasoning": "Zone direction changed; turnaround point created OD break.",
                },
                {
                    "origin": airports[turn_index],
                    "destination": airports[-1],
                    "origin_zone": zones[turn_index],
                    "destination_zone": zones[-1],
                    "class": sectors[0]["class"],
                    "component_sectors": [sector_pair(s) for s in sectors[turn_index:]],
                    "turnaround": "YES",
                    "reasoning": "Inbound journey after turnaround point.",
                },
            ]

        return [{
            "origin": airports[0],
            "destination": airports[-1],
            "origin_zone": zones[0],
            "destination_zone": zones[-1],
            "class": sectors[0]["class"],
            "component_sectors": [sector_pair(s) for s in sectors],
            "turnaround": "NO",
            "reasoning": "Continuous non-hub direction; sectors combined into one OD.",
        }]

    if trip_origin == HUB:
        pure_hub_loops = False

        if len(sectors) % 2 == 0:
            pure_hub_loops = True

            for i in range(0, len(sectors), 2):
                if i + 1 >= len(sectors):
                    pure_hub_loops = False
                    break

                outbound = sectors[i]
                inbound = sectors[i + 1]

                if not (
                    outbound["origin"] == HUB
                    and outbound["destination"] != HUB
                    and inbound["origin"] != HUB
                    and inbound["destination"] == HUB
                ):
                    pure_hub_loops = False
                    break

        if pure_hub_loops:
            return [
                {
                    "origin": s["origin"],
                    "destination": s["destination"],
                    "origin_zone": s["origin_zone"],
                    "destination_zone": s["destination_zone"],
                    "class": s["class"],
                    "component_sectors": [sector_pair(s)],
                    "turnaround": "YES" if len(sectors) > 1 else "NO",
                    "reasoning": "Pure repeated hub loops retained sector based.",
                }
                for s in sectors
            ]

    if trip_origin != HUB and len(legs) == 2:
        return [
            make_od_from_leg(
                l,
                "YES",
                "Hub transfer eliminated; return journey split into outbound and inbound OD.",
            )
            for l in legs
        ]

    ods = []
    start = legs[0]

    current_origin = start["origin"]
    current_origin_zone = start["origin_zone"]
    current_class = start["class"]
    current_components = []

    prev_sign = 0
    last_dest = start["destination"]
    last_dest_zone = start["destination_zone"]

    for idx, leg in enumerate(legs):
        sign = direction(leg["origin_zone"], leg["destination_zone"])

        if idx == 0:
            prev_sign = sign
            current_components.extend(leg["component_sectors"])
            last_dest = leg["destination"]
            last_dest_zone = leg["destination_zone"]
            continue

        if prev_sign != 0 and sign != 0 and sign != prev_sign:
            ods.append({
                "origin": current_origin,
                "destination": last_dest,
                "origin_zone": current_origin_zone,
                "destination_zone": last_dest_zone,
                "class": current_class,
                "component_sectors": current_components,
                "turnaround": "YES",
                "reasoning": "Zone direction changed; turnaround point created OD break.",
            })

            current_origin = leg["origin"]
            current_origin_zone = leg["origin_zone"]
            current_class = leg["class"]
            current_components = list(leg["component_sectors"])
        else:
            current_components.extend(leg["component_sectors"])

        if sign != 0:
            prev_sign = sign

        last_dest = leg["destination"]
        last_dest_zone = leg["destination_zone"]

    ods.append({
        "origin": current_origin,
        "destination": last_dest,
        "origin_zone": current_origin_zone,
        "destination_zone": last_dest_zone,
        "class": current_class,
        "component_sectors": current_components,
        "turnaround": "NO" if len(ods) == 0 else "YES",
        "reasoning": "Continuous zone sequence; sectors combined into one OD.",
    })

    return ods


def calculate_group(group: List[Dict[str, Any]], subgroup: str) -> List[Dict[str, Any]]:
    side_indexes = find_side_trip_indexes(group)

    main_sectors = []
    side_sectors = []

    for idx, s in enumerate(group):
        if idx in side_indexes:
            side_sectors.append(s)
        else:
            main_sectors.append(s)

    main_ods = calculate_main_ods(main_sectors)
    main_classification = classify_return(main_ods)

    results = []

    for od in main_ods:
        results.append({
            "subgroup": subgroup,
            "class": od["class"],
            "od_pair": od_pair(od["origin"], od["destination"]),
            "origin_zone": od["origin_zone"],
            "destination_zone": od["destination_zone"],
            "turnaround": od["turnaround"],
            "classification": main_classification,
            "component_sectors": ", ".join(od["component_sectors"]),
            "reasoning": od["reasoning"],
        })

    for s in side_sectors:
        results.append({
            "subgroup": subgroup,
            "class": s["class"],
            "od_pair": od_pair(s["origin"], s["destination"]),
            "origin_zone": s["origin_zone"],
            "destination_zone": s["destination_zone"],
            "turnaround": "YES",
            "classification": "SIDE TRIP",
            "component_sectors": sector_pair(s),
            "reasoning": "Return journey exists inside another return itinerary; sector marked as side trip.",
        })

    return results


def calculate_all(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    sectors = build_sectors(rows)
    groups = group_by_class(sectors)

    final = []
    sno = 1

    for i, group in enumerate(groups, start=1):
        group_result = calculate_group(group, f"Group {i}")

        for r in group_result:
            r["sno"] = sno
            final.append(r)
            sno += 1

    return final


def parse_rows_from_form(form) -> List[Dict[str, Any]]:
    row_count = int(form.get("row_count", 1))
    rows = []

    for i in range(1, row_count + 1):
        rows.append({
            "sno": i,
            "departure_date": form.get(f"departure_date_{i}", date.today().isoformat()),
            "origin": norm(form.get(f"origin_{i}", "")),
            "destination": norm(form.get(f"destination_{i}", "")),
            "travel_class": norm(form.get(f"travel_class_{i}", "Y")),
        })

    return rows


def airport_select(name: str, selected: str) -> str:
    options = ['<option value=""></option>']

    for a in AIRPORTS:
        sel = "selected" if a == selected else ""
        options.append(f'<option value="{a}" {sel}>{a}</option>')

    return f'<select name="{name}">{"".join(options)}</select>'


def class_select(name: str, selected: str) -> str:
    options = []

    for c in ["F", "J", "W", "Y"]:
        sel = "selected" if c == selected else ""
        options.append(f'<option value="{c}" {sel}>{c}</option>')

    return f'<select name="{name}">{"".join(options)}</select>'


def render_results(results: List[Dict[str, Any]], message: str = "") -> str:
    if message:
        return f'<span class="error">{html.escape(message)}</span>'

    if not results:
        return "No results yet."

    rows_html = ""

    for r in results:
        rows_html += f"""
        <tr>
            <td>{r.get("sno", "")}</td>
            <td>{html.escape(str(r.get("subgroup", "")))}</td>
            <td>{html.escape(str(r.get("class", "")))}</td>
            <td class="tag">{html.escape(str(r.get("od_pair", "")))}</td>
            <td>{html.escape(str(r.get("origin_zone", "")))}</td>
            <td>{html.escape(str(r.get("destination_zone", "")))}</td>
            <td>{html.escape(str(r.get("turnaround", "")))}</td>
            <td>{html.escape(str(r.get("classification", "")))}</td>
            <td>{html.escape(str(r.get("component_sectors", "")))}</td>
            <td>{html.escape(str(r.get("reasoning", "")))}</td>
        </tr>
        """

    return f"""
    <div class="table-wrap">
    <table class="result-table">
    <thead>
    <tr>
        <th>Sno</th>
        <th>Subgroup</th>
        <th>Class</th>
        <th>OD Pair</th>
        <th>Origin Zone</th>
        <th>Destination Zone</th>
        <th>Turnaround</th>
        <th>Classification</th>
        <th>Component Sectors</th>
        <th>Reasoning</th>
    </tr>
    </thead>
    <tbody>{rows_html}</tbody>
    </table>
    </div>
    """


def render_page(rows: List[Dict[str, Any]], results=None, error_message: str = "") -> HTMLResponse:
    results = results or []

    startup_html = ""
    if STARTUP_ERROR:
        startup_html = f'<div class="error-box">{html.escape(STARTUP_ERROR)}</div>'

    row_html = ""

    for idx, r in enumerate(rows, start=1):
        dep_date = html.escape(str(r.get("departure_date") or date.today().isoformat()))
        origin = norm(r.get("origin"))
        destination = norm(r.get("destination"))
        travel_class = norm(r.get("travel_class")) or "Y"

        row_html += f"""
        <tr>
            <td><input name="sno_{idx}" value="{idx}" readonly></td>
            <td><input type="date" name="departure_date_{idx}" value="{dep_date}"></td>
            <td>{airport_select(f"origin_{idx}", origin)}</td>
            <td>{airport_select(f"destination_{idx}", destination)}</td>
            <td>{class_select(f"travel_class_{idx}", travel_class)}</td>
            <td>
                <button type="submit" name="action" value="delete_{idx}" class="clear-btn">
                    Delete
                </button>
            </td>
        </tr>
        """

    html_page = f"""
<!DOCTYPE html>
<html>
<head>
<title>EK Loyalty OD Calculator</title>

<style>
body {{ margin:0; font-family:Arial,sans-serif; background:#f7f5f2; }}
header {{
    background:linear-gradient(90deg,#7a0019,#b8860b);
    color:white;
    padding:20px 30px;
    min-height:80px;
    display:flex;
    align-items:center;
    gap:20px;
}}
.header-wrap {{ flex:1; }}
.container {{ padding:24px; }}
.card {{ background:white; border-radius:14px; padding:20px; margin-bottom:22px; box-shadow:0 4px 16px rgba(0,0,0,.08); }}
.error-box {{ background:#ffe0e0; color:#990000; padding:14px; margin:20px; border-radius:10px; font-weight:bold; }}
.table-wrap {{ overflow-x:auto; }}
table {{ width:100%; border-collapse:collapse; font-size:14px; }}
th {{ background:#7a0019; color:white; padding:10px; white-space:nowrap; }}
td {{ padding:8px; border-bottom:1px solid #ddd; vertical-align:middle; }}
input,select {{ width:100%; padding:8px; border:1px solid #bbb; border-radius:6px; box-sizing:border-box; text-transform:uppercase; }}
button {{ background:#7a0019; color:white; border:none; padding:11px 18px; border-radius:8px; cursor:pointer; font-weight:bold; margin-top:14px; margin-right:8px; }}
.gold-btn {{ background:#b8860b; }}
.clear-btn {{ background:#c00020; }}
.result-table th {{ background:#b8860b; }}
.tag {{ font-weight:bold; color:#7a0019; white-space:nowrap; }}
.error {{ color:red; font-weight:bold; }}

.hit-counter {{
    margin-left:auto;
    background:rgba(255,255,255,.18);
    border:1px solid rgba(255,255,255,.35);
    padding:10px 16px;
    border-radius:999px;
    font-weight:bold;
    white-space:nowrap;
}}

.route-flow {{
    display:flex;
    flex-wrap:wrap;
    align-items:center;
    gap:8px;
    padding:4px 0 8px 0;
}}

.route-empty {{
    color:#777;
    font-size:13px;
    padding:4px 0;
}}

.itin-load-box {{
    display:flex;
    gap:8px;
    margin-bottom:14px;
    align-items:center;
}}

.itin-load-box input {{
    flex:1;
}}

.itin-load-box button {{
    background:#c00020;
    margin-top:0;
    white-space:nowrap;
}}
</style>
</head>

<body>

{startup_html}

<header>
    <div class="header-wrap">
        <h1>Loyalty Origin and Destination (OnD) Engine</h1>
        <p>
            Intelligent Journey Analysis • Turnaround Detection •
            Return Journey Classification • Side Trip Recognition •
            Zig-Zag Evaluation • Hub DXB • Zone Z01–Z17
        </p>
        <p>
        <b>By Syed Abbas Rizvi</b>
        </p>
    </div>

    <div class="hit-counter">
        Hits : {get_hit_count()}
    </div>
</header>

<div class="container">

<div class="card">
<h2>Itinerary Input</h2>

<form method="post" action="/submit">
<input type="hidden" name="row_count" value="{len(rows)}">

<div class="itin-load-box">
    <input
        type="text"
        name="itinerary_string"
        placeholder="Paste itinerary string e.g. DXBSIN SINMEL MELDXB DXBLHR"
    >
    <button type="submit" name="action" value="load_string">
        Load Itinerary
    </button>
</div>

<div class="table-wrap">
<table>
<thead>
<tr>
<th>Sno</th>
<th>Departure Date</th>
<th>Origin</th>
<th>Destination</th>
<th>Class</th>
<th>Action</th>
</tr>
</thead>
<tbody>
{row_html}
</tbody>
</table>
</div>

<button class="gold-btn" type="submit" name="action" value="add">Add Row</button>
<button type="submit" name="action" value="calculate">Calculate OD</button>
<button class="clear-btn" type="submit" name="action" value="clear">Clear</button>
</form>
</div>

<div class="card">
<h2>Calculated OD Results</h2>
<div id="results">{render_results(results, error_message)}</div>
</div>

</div>

</body>
</html>
"""

    return HTMLResponse(html_page)


@app.get("/", response_class=HTMLResponse)
def home():
    return render_page(default_rows())


@app.post("/submit", response_class=HTMLResponse)
async def submit(request: Request):
    form = await request.form()
    action = form.get("action", "")

    if action == "clear":
        return render_page(default_rows())

    rows = parse_rows_from_form(form)

    if action == "load_string":
        itinerary_text = form.get("itinerary_string", "")
        rows = rows_from_itinerary_string(itinerary_text)

        for i, r in enumerate(rows, start=1):
            r["sno"] = i

        return render_page(rows)

    if str(action).startswith("delete_"):
        try:
            delete_idx = int(str(action).replace("delete_", ""))
            rows = [r for i, r in enumerate(rows, start=1) if i != delete_idx]
        except Exception:
            pass

        if not rows:
            rows = default_rows()

        for i, r in enumerate(rows, start=1):
            r["sno"] = i

        return render_page(rows)

    if action == "add":
        last_date = rows[-1].get("departure_date") or date.today().isoformat()

        try:
            next_date = (date.fromisoformat(last_date) + timedelta(days=1)).isoformat()
        except Exception:
            next_date = date.today().isoformat()

        rows.append({
            "sno": len(rows) + 1,
            "departure_date": next_date,
            "origin": "",
            "destination": "",
            "travel_class": "Y",
        })

        return render_page(rows)

    if action == "calculate":
        if STARTUP_ERROR:
            return render_page(rows, [], STARTUP_ERROR)

        valid_rows = [r for r in rows if r.get("origin") and r.get("destination")]

        if not valid_rows:
            return render_page(rows, [], "Please enter at least one valid itinerary row.")

        try:
            results = calculate_all(rows)
            return render_page(rows, results)
        except Exception as e:
            return render_page(rows, [], str(e))

    return render_page(rows)


@app.get("/api/hit-count")
def api_hit_count():
    return {
        "rest_api_hits": get_hit_count()
    }


@app.post("/api/calculate-od")
async def api_calculate_od(request: Request):
    increment_hit_count()

    if STARTUP_ERROR:
        return JSONResponse(
            status_code=500,
            content={
                "success": False,
                "error": STARTUP_ERROR,
                "rest_api_hits": get_hit_count()
            }
        )

    try:
        payload = await request.json()
        rows = payload.get("rows", [])

        if not rows:
            return JSONResponse(
                status_code=400,
                content={
                    "success": False,
                    "error": "rows are required",
                    "rest_api_hits": get_hit_count()
                }
            )

        results = calculate_all(rows)

        return {
            "success": True,
            "rest_api_hits": get_hit_count(),
            "results": results
        }

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "success": False,
                "error": str(e),
                "rest_api_hits": get_hit_count()
            }
        )
