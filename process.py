"""
India Grid Carbon Tracker — Data Processor
-------------------------------------------
Usage (single file):
    python process.py "PSP Reports\26_05_26_NLDC_PSP_536.xls"

Usage (full folder — recommended first time):
    python process.py --folder "PSP Reports"

Writes data.json in the same folder as this script.
Requires: pip install pandas xlrd openpyxl
"""

import sys, json, re, subprocess, os
from datetime import datetime
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("Run: pip install pandas xlrd openpyxl")
    sys.exit(1)

EF_COAL  = 0.95
EF_GAS   = 0.45
CEA_BASE = 0.716   # CEA annual baseline EF

# ── LibreOffice path (Windows) ─────────────────────────────────────────────
SOFFICE_PATHS = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "soffice",  # if in PATH
]

def find_soffice():
    for p in SOFFICE_PATHS:
        if Path(p).exists() or p == "soffice":
            return p
    return None

def convert_xls(xls_path):
    """Convert .xls to .xlsx, return path to .xlsx"""
    xls_path = Path(xls_path).resolve()
    out_dir   = Path(os.environ.get("TEMP", "/tmp"))
    out_path  = out_dir / (xls_path.stem + ".xlsx")

    soffice = find_soffice()
    if not soffice:
        raise RuntimeError(
            "LibreOffice not found.\n"
            "Download from https://www.libreoffice.org/ and install."
        )

    result = subprocess.run(
        [soffice, "--headless", "--norestore", "--convert-to", "xlsx",
         str(xls_path), "--outdir", str(out_dir)],
        capture_output=True, text=True
    )
    if not out_path.exists():
        raise RuntimeError(f"Conversion failed: {result.stderr}")
    return str(out_path)

def to_f(v):
    try:
        return float(str(v).replace(",","").strip())
    except:
        return 0.0

def parse_date(filename):
    m = re.search(r"(\d{2})[._](\d{2})[._](\d{2})", Path(filename).stem)
    if m:
        dd, mm, yy = m.groups()
        months = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        return f"20{yy}-{mm}-{dd}", f"{dd} {months[int(mm)]} 20{yy}"
    return datetime.today().strftime("%Y-%m-%d"), datetime.today().strftime("%d %b %Y")

# ── Process one XLS file ───────────────────────────────────────────────────
def process_file(xls_path):
    print(f"  Processing {Path(xls_path).name} ...", end=" ")
    xlsx = convert_xls(xls_path)
    report_date, report_date_display = parse_date(xls_path)

    # TimeSeries — auto-detect format:
    # APRIL format: skiprows=4, 13 cols, total_gen at col 11
    # MAY+  format: skiprows=3, 15 cols, row 0 is sub-headers, data from row 1, total_gen at col 12
    # Discriminate by column count (reliable across all files):
    # APRIL format = 13 cols: total_gen at col 11
    # MAY+  format = 15 cols: total_gen at col 12, storage at col 9
    raw = pd.read_excel(xlsx, sheet_name="TimeSeries", header=None, skiprows=4)
    raw.columns = range(raw.shape[1])

    if raw.shape[1] <= 13:
        # APRIL format — 13 cols
        # TIME(0) FREQ(1) DEMAND(2) NUCLEAR(3) WIND(4) SOLAR(5) HYDRO(6) GAS(7) THERMAL(8) OTHERS(9) NET_DEMAND(10) TOTAL_GEN(11) NET_TRANS(12)
        df = raw.copy()
        col_map = {0:"time",1:"freq",2:"demand",3:"nuclear",4:"wind",5:"solar",
                   6:"hydro",7:"gas",8:"thermal",9:"others",10:"net_demand",
                   11:"total_gen",12:"net_trans"}
    else:
        # MAY+ format — 15 cols, skiprows=4 puts sub-header at row 0, data from row 1
        # Row 0: nan | ' ' | (A) | (B) | (C) | (D) | (E) | (F) | (G) | (H) | (I) | (J) | (K) | nan | (L)
        # TIME(0) FREQ(1) DEMAND(2) NUCLEAR(3) WIND(4) SOLAR(5) HYDRO(6) GAS(7) THERMAL(8) STORAGE(9) OTHERS(10) NET_DEMAND(11) TOTAL_GEN(12) NaN(13) NET_TRANS(14)
        df = raw.iloc[1:].reset_index(drop=True).copy()
        df.columns = range(df.shape[1])
        col_map = {0:"time",1:"freq",2:"demand",3:"nuclear",4:"wind",5:"solar",
                   6:"hydro",7:"gas",8:"thermal",9:"storage",10:"others",
                   11:"net_demand",12:"total_gen",14:"net_trans"}

    df.rename(columns={k:v for k,v in col_map.items() if k < df.shape[1]}, inplace=True)

    for c in ["freq","demand","nuclear","wind","solar","hydro","gas","thermal","storage","total_gen"]:
        if c not in df.columns:
            df[c] = 0.0
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

    df = df[df["thermal"] > 0].copy()
    df["time"] = df["time"].astype(str).str.strip()

    k = 0.25/1000.0
    mu = {}
    for c in ["thermal","solar","wind","hydro","nuclear","gas","storage","total_gen"]:
        mu[c] = float((df[c]*k).sum()) if c in df.columns else 0.0
    mu["total"] = mu.pop("total_gen")

    co2_thermal = mu["thermal"]*1000*EF_COAL
    co2_gas     = mu["gas"]    *1000*EF_GAS
    co2_total   = co2_thermal + co2_gas
    re_avoided  = (mu["solar"]+mu["wind"])*1000*EF_COAL
    total_mwh   = mu["total"] * 1000
    grid_ef     = round(co2_total / total_mwh, 4) if total_mwh > 10 else 0.0
    # Sanity check — EF should be between 0.3 and 1.2 for India grid
    if grid_ef > 1.2 or grid_ef < 0.3:
        grid_ef = 0.0
    re_share    = (mu["solar"]+mu["wind"]+mu["hydro"])/mu["total"]*100 if mu["total"] > 0 else 0
    clean_share = (mu["solar"]+mu["wind"]+mu["hydro"]+mu["nuclear"])/mu["total"]*100 if mu["total"] > 0 else 0

    # Timeseries rows
    timeseries = []
    for _, r in df.iterrows():
        tgen = to_f(r.get("total_gen", 0))
        th   = to_f(r.get("thermal",  0))
        ga   = to_f(r.get("gas",      0))
        ef   = (th*EF_COAL + ga*EF_GAS)/tgen if tgen > 0 else 0
        timeseries.append({
            "t":       r["time"],
            "demand":  round(to_f(r.get("demand",  0))/1000, 1),
            "thermal": round(th/1000, 1),
            "solar":   round(to_f(r.get("solar",   0))/1000, 1),
            "wind":    round(to_f(r.get("wind",    0))/1000, 1),
            "hydro":   round(to_f(r.get("hydro",   0))/1000, 1),
            "nuclear": round(to_f(r.get("nuclear", 0))/1000, 1),
            "gas":     round(ga/1000, 1),
            "storage": round(to_f(r.get("storage", 0))/1000, 1),
            "total":   round(tgen/1000, 1),
            "freq":    round(to_f(r.get("freq", 50)), 2),
            "ef":      round(ef, 4),
        })

    # MOP_E regions — parse by exact row content
    regions = {}
    try:
        mop = pd.read_excel(xlsx, sheet_name="MOP_E", header=None)
        rkeys = ["NR","WR","SR","ER","NER","Total"]

        def extract_nums(row_vals, n=6):
            """Extract up to n positive floats from a row, treating '-' as 0."""
            nums = []
            for v in row_vals[1:]:
                s = str(v).strip()
                if s in ["-","—",""]:
                    nums.append(0.0)
                else:
                    try: nums.append(float(s.replace(",","")))
                    except: pass
                if len(nums) >= n: break
            return nums

        for i, row in mop.iterrows():
            vals = [str(v).strip() for v in row]
            line = " ".join(vals)
            # Clean vals for keyword matching
            kv = [v for v in vals if v not in ["nan","None",""]]
            if not kv: continue

            if "Energy Met (MU)" in line:
                nums = extract_nums(kv)
                for rk,v in zip(rkeys, nums): regions.setdefault(rk,{})["energy_mu"] = v

            elif "Peak Shortage" in line and "(MW)" in line:
                nums = extract_nums(kv)
                for rk,v in zip(rkeys, nums): regions.setdefault(rk,{})["peak_shortage_mw"] = v

            elif "Maximum Demand Met" in line and "SCADA" in line:
                nums = extract_nums(kv)
                for rk,v in zip(rkeys, nums): regions.setdefault(rk,{})["peak_demand_mw"] = v

            elif "Solar Gen" in line:
                nums = extract_nums(kv)
                for rk,v in zip(rkeys, nums): regions.setdefault(rk,{})["solar_mu"] = v

            elif "Wind Gen" in line:
                nums = extract_nums(kv)
                for rk,v in zip(rkeys, nums): regions.setdefault(rk,{})["wind_mu"] = v

            elif "Hydro Gen" in line:
                nums = extract_nums(kv)
                for rk,v in zip(rkeys, nums): regions.setdefault(rk,{})["hydro_mu"] = v

    except Exception as e:
        print(f"(MOP_E skipped: {e})", end=" ")

    # Frequency — Row format: All India | FVI | <49.7 | 49.7-49.8 | 49.8-49.9 | <49.9 | 49.9-50.05 | >50.05
    freq_profile = {}
    try:
        for i, row in mop.iterrows():
            vals = [str(v).strip() for v in row if str(v).strip() not in ["nan","None",""]]
            if vals and vals[0] == "All India" and len(vals) >= 7:
                freq_profile = {
                    "fvi":         to_f(vals[1]),
                    "below_49_7":  to_f(vals[2]) if len(vals)>2 else 0,
                    "band_49_7_8": to_f(vals[3]) if len(vals)>3 else 0,
                    "band_49_8_9": to_f(vals[4]) if len(vals)>4 else 0,
                    "below_49_9":  to_f(vals[5]) if len(vals)>5 else 0,
                    "normal":      to_f(vals[6]) if len(vals)>6 else 0,
                    "above_50_05": to_f(vals[7]) if len(vals)>7 else 0,
                }
                break
    except: pass

    # Section G — Sourcewise generation by region (MU, gross)
    # Rows: Coal, Lignite, Hydro, Nuclear, Gas/Naphtha/Diesel, RES, Total
    # Columns: NR, WR, SR, ER, NER, All India, % Share
    regions_sourcewise = {}
    try:
        rkeys_sw = ["NR", "WR", "SR", "ER", "NER", "All India"]
        src_map = {
            "Coal":      "coal_mu",
            "Lignite":   "lignite_mu",
            "Hydro":     "hydro_mu",
            "Nuclear":   "nuclear_mu",
            "Gas":       "gas_mu",   # matches "Gas, Naptha & Diesel"
            "RES":       "res_mu",   # matches "RES (Wind, Solar, Biomass & Others)"
            "Total":     "total_mu",
        }
        for i, row in mop.iterrows():
            vals = [str(v).strip() for v in row if str(v).strip() not in ["nan", "None", ""]]
            if not vals:
                continue
            key = None
            first = vals[0]
            if first == "Coal":
                key = "coal_mu"
            elif first == "Lignite":
                key = "lignite_mu"
            elif first == "Hydro" and "Hydro Gen" not in " ".join(vals):
                key = "hydro_mu"
            elif first == "Nuclear":
                key = "nuclear_mu"
            elif first.startswith("Gas"):
                key = "gas_mu"
            elif first.startswith("RES"):
                key = "res_mu"
            elif first == "Total" and len(vals) >= 6:
                # Guard: only the sourcewise total row has 6+ numeric cols
                try:
                    float(vals[1].replace(",", ""))
                    key = "total_mu"
                except:
                    pass

            if key:
                nums = []
                for v in vals[1:]:
                    s2 = v.replace(",", "").strip()
                    try:
                        nums.append(float(s2))
                    except:
                        pass
                # nums order: NR, WR, SR, ER, NER, All India, %share
                for idx, rk in enumerate(rkeys_sw):
                    if idx < len(nums):
                        regions_sourcewise.setdefault(rk, {})[key] = round(nums[idx], 2)

        # Derive implied CO₂ EF per region: (coal+lignite)*0.95 + gas*0.45) / total*1000
        for rk, d in regions_sourcewise.items():
            total = d.get("total_mu", 0)
            if total > 0:
                co2 = (d.get("coal_mu", 0) + d.get("lignite_mu", 0)) * 0.95 + \
                      d.get("gas_mu", 0) * 0.45
                d["implied_ef"] = round(co2 / total, 4)
                fossil = d.get("coal_mu", 0) + d.get("lignite_mu", 0) + d.get("gas_mu", 0)
                d["fossil_pct"] = round(fossil / total * 100, 1)
                d["res_pct"]    = round(d.get("res_mu", 0) / total * 100, 1)
                d["clean_pct"]  = round((d.get("hydro_mu", 0) + d.get("nuclear_mu", 0) +
                                         d.get("res_mu", 0)) / total * 100, 1)
            else:
                d["implied_ef"] = 0.0
                d["fossil_pct"] = 0.0
                d["res_pct"]    = 0.0
                d["clean_pct"]  = 0.0

    except Exception as e:
        print(f"(Section G skipped: {e})", end=" ")

    # Cross border
    cross_border = {}
    try:
        cb = pd.read_excel(xlsx, sheet_name="CrossBorder", header=None)
        net_section = False
        for i, row in cb.iterrows():
            vals = [str(v).strip() for v in row if str(v).strip() not in ["nan","None",""]]
            if not vals: continue
            if "Net from India" in " ".join(vals):
                net_section = True; continue
            if net_section and vals[0] in ["Bhutan","Nepal","Bangladesh","Myanmar"]:
                nums = [to_f(v) for v in vals[1:]]
                cross_border[vals[0]] = {"net_mu": nums[-1] if nums else 0}
    except: pass

    summary = {
        "total_gen_mu":   round(mu["total"],   2),
        "thermal_mu":     round(mu["thermal"], 2),
        "solar_mu":       round(mu["solar"],   2),
        "wind_mu":        round(mu["wind"],    2),
        "hydro_mu":       round(mu["hydro"],   2),
        "nuclear_mu":     round(mu["nuclear"], 2),
        "gas_mu":         round(mu["gas"],     2),
        "co2_thermal_kt": round(co2_thermal/1000, 1),
        "co2_gas_kt":     round(co2_gas/1000,     1),
        "co2_total_kt":   round(co2_total/1000,   1),
        "re_avoided_kt":  round(re_avoided/1000,  1),
        "grid_ef":        round(grid_ef,  4),
        "re_share_pct":   round(re_share, 2),
        "clean_share_pct":round(clean_share, 2),
        "peak_demand_mw": round(float(df["demand"].max()), 0),
        "peak_solar_mw":  round(float(df["solar"].max()),  0),
    }

    print(f"CO₂: {summary['co2_total_kt']} kt · EF: {summary['grid_ef']}")

    return {
        "meta": {
            "report_date":         report_date,
            "report_date_display": report_date_display,
            "source_file":         Path(xls_path).name,
            "processed_at":        datetime.now().isoformat(timespec="seconds"),
            "ef_coal":   EF_COAL,
            "ef_gas":    EF_GAS,
            "ef_source": "CEA CO2 Baseline Database 2023-24",
            "note": "Derived from NLDC SCADA. Rooftop RE excluded.",
        },
        "summary":            summary,
        "regions":            regions,
        "regions_sourcewise": regions_sourcewise,
        "freq_profile":       freq_profile,
        "cross_border":       cross_border,
        "timeseries":         timeseries,
    }

# ── History summary (one row per day) ─────────────────────────────────────
def make_history(all_days):
    """Light summary array for trend charts — one entry per day."""
    history = []
    for d in all_days:
        s = d["summary"]
        history.append({
            "date":       d["meta"]["report_date"],
            "date_disp":  d["meta"]["report_date_display"],
            "co2_kt":     s["co2_total_kt"],
            "ef":         s["grid_ef"],
            "avoided_kt": s["re_avoided_kt"],
            "re_pct":     s["re_share_pct"],
            "total_mu":   s["total_gen_mu"],
            "solar_mu":   s["solar_mu"],
            "wind_mu":    s["wind_mu"],
            "peak_mw":    s["peak_demand_mw"],
        })
    history.sort(key=lambda x: x["date"])
    return history

# ── Entry point ────────────────────────────────────────────────────────────
def main():
    script_dir = Path(__file__).parent
    out_path   = script_dir / "data.json"

    # Load existing data.json if present (to preserve history)
    existing_history = []
    existing_dates   = set()
    if out_path.exists():
        try:
            existing = json.loads(out_path.read_text())
            existing_history = existing.get("history", [])
            existing_dates   = {h["date"] for h in existing_history}
        except: pass

    # ── Folder mode: process all XLS in PSP Reports/ ──────────────────────
    if "--folder" in sys.argv:
        idx = sys.argv.index("--folder")
        folder = Path(sys.argv[idx+1]) if idx+1 < len(sys.argv) else script_dir / "PSP Reports"
        # Use case-insensitive deduplication for Windows
        seen = set()
        xls_files = []
        for f in sorted(folder.iterdir()):
            if f.suffix.lower() == ".xls" and f.name.lower() not in seen:
                seen.add(f.name.lower())
                xls_files.append(f)
        xls_files.sort(key=lambda f: f.name)
        if not xls_files:
            print(f"No XLS files found in {folder}")
            sys.exit(1)
        print(f"Found {len(xls_files)} files in {folder}\n")

        all_day_data = []
        latest = None
        for f in xls_files:
            try:
                d = process_file(f)
                all_day_data.append(d)
                if latest is None or d["meta"]["report_date"] > latest["meta"]["report_date"]:
                    latest = d
            except Exception as e:
                print(f"  SKIP {f.name}: {e}")

        if not all_day_data:
            print("\nNo files could be processed. Install LibreOffice first:")
            print("  https://www.libreoffice.org/download/download/")
            sys.exit(1)

        new_history = make_history(all_day_data)
        out = dict(latest)
        out["history"] = new_history

    # ── Single file mode ──────────────────────────────────────────────────
    else:
        if len(sys.argv) < 2:
            print(__doc__)
            sys.exit(1)
        xls_path = sys.argv[1]
        d = process_file(xls_path)

        # Merge with existing history
        new_row = make_history([d])[0]
        if new_row["date"] not in existing_dates:
            existing_history.append(new_row)
        existing_history.sort(key=lambda x: x["date"])

        out = d.copy()
        out["history"] = existing_history

    # Write
    out_path.write_text(json.dumps(out, indent=2))
    s = out["summary"]
    print(f"\n✓ data.json updated → {out_path}")
    print(f"  Date      : {out['meta']['report_date_display']}")
    print(f"  Total gen : {s['total_gen_mu']} MU")
    print(f"  CO₂       : {s['co2_total_kt']} kt")
    print(f"  Grid EF   : {s['grid_ef']} tCO₂/MWh")
    print(f"  RE avoided: {s['re_avoided_kt']} kt")
    print(f"  History   : {len(out['history'])} days")
    print(f"\nNow: git add data.json && git commit -m 'data: {out['meta']['report_date']}' && git push")

if __name__ == "__main__":
    main()
