"""
Build a single JS payload from the raw Excel files in Vacancy_Data/ and Arrears_Data/.
Outputs: raw_data_compiled.js (drops window.RAW_VACANCY / window.RAW_ARREARS).

This is lightweight and avoids external packages: it reads XLSX via zipfile + ElementTree.
Usage (from Z:\\Dashboards\\TV):
    python build_raw_data.py
"""
import json
import re
import zipfile
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parent
VAC_DIR = ROOT / "Vacancy_Data"
ARR_DIR = ROOT / "Arrears_Data"
OUT_FILE = ROOT / "raw_data_compiled.js"
PROP_INFO = json.loads((ROOT / "data_property_info.json").read_text(encoding="utf-8"))
AREA_PORTFOLIOS = json.loads((ROOT / "data_area_portfolios.json").read_text(encoding="utf-8"))


def letters_to_idx(letters: str) -> int:
    """Convert column letters (A, B, AA) to zero-based index."""
    val = 0
    for ch in letters:
        val = val * 26 + (ord(ch.upper()) - 64)
    return val - 1


def parse_shared_strings(zf: zipfile.ZipFile):
    try:
        data = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(data)
    strings = []
    for si in root.findall(".//{*}si"):
        text = "".join(t.text or "" for t in si.findall(".//{*}t"))
        strings.append(text)
    return strings


def parse_sheet_values(zf: zipfile.ZipFile, sheet_path: str, shared_strings):
    xml_bytes = zf.read(sheet_path)
    root = ET.fromstring(xml_bytes)
    rows = []
    for row in root.findall(".//{*}row"):
        cells = {}
        for c in row.findall("{*}c"):
            ref = c.attrib.get("r", "")
            m = re.match(r"([A-Z]+)", ref)
            col_idx = letters_to_idx(m.group(1)) if m else None
            cell_type = c.attrib.get("t")
            v = c.find("{*}v")
            text_val = v.text if v is not None else None
            if cell_type == "s":  # shared string
                try:
                    val = shared_strings[int(text_val)]
                except Exception:
                    val = text_val
            elif cell_type == "inlineStr":
                val = "".join(t.text or "" for t in c.findall(".//{*}t"))
            else:
                val = text_val
            if val is not None:
                # Attempt numeric conversion; keep original if it fails
                try:
                    if "." in val:
                        num = float(val)
                        val = int(num) if num.is_integer() else num
                    else:
                        val = int(val)
                except Exception:
                    pass
            cells[col_idx] = val
        # Convert sparse row to list
        if cells:
            max_idx = max(cells.keys())
            row_vals = [None] * (max_idx + 1)
            for idx, val in cells.items():
                row_vals[idx] = val
            rows.append(row_vals)
    return rows


def pick_sheet_path(zf: zipfile.ZipFile, preferred_names=None):
    try:
        wb_xml = zf.read("xl/workbook.xml")
    except KeyError:
        return None
    preferred = [n.lower() for n in (preferred_names or [])]
    root = ET.fromstring(wb_xml)
    sheets = []
    ns = {"r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    for sh in root.findall(".//{*}sheet"):
        name = (sh.attrib.get("name") or "").strip()
        r_id = sh.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        sheets.append((name, r_id))
    # map rId -> target
    rel_map = {}
    try:
        rel_xml = zf.read("xl/_rels/workbook.xml.rels")
        rel_root = ET.fromstring(rel_xml)
        for rel in rel_root.findall("{*}Relationship"):
            rel_map[rel.attrib.get("Id")] = rel.attrib.get("Target")
    except KeyError:
        pass
    def resolve_target(rid):
        target = rel_map.get(rid, "")
        if target.startswith("/"):
            target = target.lstrip("/")
        if not target.startswith("xl/"):
            target = "xl/" + target
        return target
    if sheets:
        # try preferred names first
        for pname in preferred:
            for nm, rid in sheets:
                if nm.lower() == pname and rid:
                    return resolve_target(rid)
        # fallback to first sheet
        nm, rid = sheets[0]
        if rid:
            return resolve_target(rid)
    return None


def read_xlsx_sheet(path: Path, preferred_names=None):
    with zipfile.ZipFile(path, "r") as zf:
        sheet_path = pick_sheet_path(zf, preferred_names=preferred_names)
        if not sheet_path:
            sheet_paths = sorted([n for n in zf.namelist() if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")])
            if not sheet_paths:
                return []
            sheet_path = sheet_paths[0]
        shared = parse_shared_strings(zf)
        return parse_sheet_values(zf, sheet_path, shared)


def row_to_dict(rows, header_keywords):
    """Find header row containing any keyword and return list of dict rows from that point."""
    header_idx = None
    headers = None
    for i, row in enumerate(rows):
        values = [str(v).strip() if v is not None else "" for v in row]
        if any(k in values for k in header_keywords):
            header_idx = i
            headers = values
            break
    if headers is None:
        return []
    data_dicts = []
    for row in rows[header_idx + 1 :]:
        vals = row + [None] * max(0, len(headers) - len(row))
        # stop before footer/total rows
        if any(isinstance(v, str) and v.strip().lower().startswith("total") for v in vals):
            break
        if all(v in (None, "", "None") for v in vals):
            continue
        entry = {headers[j]: vals[j] if j < len(vals) else None for j in range(len(headers))}
        data_dicts.append(entry)
    return data_dicts


def load_property_lookup():
    name_to_team = {}
    code_to_team = {}
    for area in AREA_PORTFOLIOS:
        for b in area["buildings"]:
            name_to_team[b] = area["key"]
    for bname, info in PROP_INFO.items():
        code = str(info.get("code") or "")
        if code and bname in name_to_team:
            code_to_team[code] = name_to_team[bname]
    return name_to_team, code_to_team


def parse_date_from_filename(path: Path):
    # Try YYYY_MM_DD then MM_DD_YYYY
    m = re.search(r"(\d{4}_\d{2}_\d{2})", path.name)
    if m:
        return datetime.strptime(m.group(1), "%Y_%m_%d")
    m = re.search(r"(\d{2}_\d{2}_\d{4})", path.name)
    if m:
        return datetime.strptime(m.group(1), "%m_%d_%Y")
    return None


def process_vacancy():
    files = sorted(VAC_DIR.glob("Vacancy-*.xlsx"))
    name_to_team, code_to_team = load_property_lookup()
    latest_date = None
    latest_rows = []
    trend = {}
    for f in files:
        dt = parse_date_from_filename(f)
        if not dt:
            continue
        rows_raw = read_xlsx_sheet(f, preferred_names=["data", "Data", "sheet1", "Sheet1"])
        dicts = row_to_dict(rows_raw, header_keywords=["Building", "Tenant Status"])
        # Normalize and compute team
        vacant_count_by_team = {}
        for r in dicts:
            building = str(r.get("Building") or r.get("Name") or "").split("-")[0].strip()
            team = name_to_team.get(building, "Unassigned")
            r["team"] = team
            r["building"] = building
            status = str(r.get("Tenant Status") or "").lower()
            is_vacant = "vacant" in status
            if is_vacant:
                vacant_count_by_team[team] = vacant_count_by_team.get(team, 0) + 1
        trend[dt.isoformat()] = vacant_count_by_team
        if latest_date is None or dt > latest_date:
            latest_date = dt
            latest_rows = dicts
    return latest_date, latest_rows, trend


def process_arrears():
    files = sorted(ARR_DIR.glob("AgingSummary_*.xlsx"))
    name_to_team, code_to_team = load_property_lookup()
    latest_date = None
    latest_rows = []
    trend = {}
    for f in files:
        dt = parse_date_from_filename(f)
        if not dt:
            continue
        rows_raw = read_xlsx_sheet(f, preferred_names=["data", "Data", "Sheet2", "sheet2", "Sheet1", "sheet1"])
        dicts = row_to_dict(rows_raw, header_keywords=["Tenant", "Current Owed"])
        sum_by_team = {}
        for r in dicts:
            prop_code = str(r.get("Property") or "").strip()
            team = code_to_team.get(prop_code, "Unassigned")
            r["team"] = team
            try:
                current = float(r.get("Current Owed") or 0)
            except Exception:
                current = 0.0
            sum_by_team[team] = sum_by_team.get(team, 0.0) + current
        trend[dt.isoformat()] = sum_by_team
        if latest_date is None or dt > latest_date:
            latest_date = dt
            latest_rows = dicts
    # Build per-team arrays for charting
    trend_by_team = {}
    for iso, data in trend.items():
        ts = int(datetime.fromisoformat(iso).timestamp() * 1000)
        total_val = sum(data.values())
        for team in set(list(data.keys()) + ["RS1", "RS2", "RS3", "RS4", "All"]):
            val = total_val if team == "All" else data.get(team, 0.0)
            trend_by_team.setdefault(team, []).append({"ts": ts, "v": val})
    for arr in trend_by_team.values():
        arr.sort(key=lambda x: x["ts"])
    return latest_date, latest_rows, trend_by_team


def main():
    vac_date, vac_rows, vac_trend = process_vacancy()
    arr_date, arr_rows, arr_trend = process_arrears()
    payload = {
        "RAW_VACANCY": {
            "latestDate": vac_date.isoformat() if vac_date else None,
            "rows": vac_rows,
            "vacantTrendByTeam": vac_trend,
        },
        "RAW_ARREARS": {
            "latestDate": arr_date.isoformat() if arr_date else None,
            "rows": arr_rows,
            "trendByTeam": arr_trend,
        },
    }
    js = (
        "// Auto-generated by build_raw_data.py\n"
        "window.RAW_VACANCY = " + json.dumps(payload["RAW_VACANCY"]) + ";\n"
        "window.RAW_ARREARS = " + json.dumps(payload["RAW_ARREARS"]) + ";\n"
    )
    OUT_FILE.write_text(js, encoding="utf-8")
    print(f"Generated {OUT_FILE.name}")
    if vac_date:
        print(f"Vacancy latest: {vac_date.date()} rows={len(vac_rows)}")
    if arr_date:
        print(f"Arrears latest: {arr_date.date()} rows={len(arr_rows)}")


if __name__ == "__main__":
    main()
