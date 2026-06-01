# India Grid Carbon Tracker

Daily derived carbon emission monitor for India's power grid.
Public dashboard → https://your-username.github.io/india-grid-carbon/

---

## What this shows

- Daily CO₂ derived from NLDC SCADA generation data
- Intraday grid emission factor (tCO₂/MWh)
- RE avoided emissions vs coal baseline
- Regional power supply snapshot
- Cross-border energy exchange

**Emissions are derived, not directly measured.**
Methodology: Thermal (MU) × 0.95 tCO₂/MWh + Gas (MU) × 0.45 tCO₂/MWh
Source: CEA CO₂ Baseline Database 2023-24

---

## Daily update workflow

### 1. Install dependencies (first time only)
```bash
pip install pandas xlrd openpyxl
```
You also need LibreOffice installed (for XLS conversion):
- **Linux**: `sudo apt install libreoffice`
- **macOS**: Download from https://www.libreoffice.org/

### 2. Download the NLDC PSP report
Go to → https://grid-india.in → Reports → Power Supply Position
Download the daily Excel file (format: `DD_MM_YY_NLDC_PSP_NNN.xls`)

### 3. Run the processing script
```bash
python process.py /path/to/26_05_26_NLDC_PSP_536.xls
```

This writes `data.json` in the repo folder.

### 4. Commit and push
```bash
git add data.json
git commit -m "data: 26 May 2026"
git push
```

### 5. Refresh the dashboard
Wait ~30 seconds for GitHub Pages to update, then refresh.

---

## Files

```
index.html      ← dashboard (reads data.json via fetch)
process.py      ← run this after downloading the XLS
data.json       ← auto-generated, committed to repo
README.md       ← this file
```

---

## Data sources

| Source | What | URL |
|--------|------|-----|
| NLDC PSP | 15-min SCADA generation by fuel | grid-india.in |
| CEA | CO₂ emission factors | cea.nic.in |

---

Created by **Sanket**
