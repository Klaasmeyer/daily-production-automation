# Daily Production Automation

Automates the daily production spreadsheet workflow for Klaasmeyer.

## What It Does

- Reads crew daily report `.xlsx` files from the `Daily's` subfolder
- Extracts footage values per work type (Bore, Plow, Trench, Aerial, Cable, Drops, etc.)
- Updates the master `DAILY PRODUCTION TEMPLATE` spreadsheet with:
  - Per-job footage values
  - Daily, weekly, monthly, quarterly, and yearly totals
  - Day-of-week breakdown table (Mon–Fri)
  - Dollar value calculations based on rate table
- Generates a crew/subcontractor contribution report (`Crew Report MM-DD-YYYY.xlsx`)
- Saves a dated copy of the production template each day

## Requirements

- Python 3.8+
- `openpyxl` library

Install dependencies:
```
pip install openpyxl
```

## Usage

```bash
# Run for today's date
python daily_production_auto.py

# Run for a specific date
python daily_production_auto.py --date 03-24-2026

# Preview without writing any files
python daily_production_auto.py --dry-run
```

## File Structure

```
Daily Production/
├── daily_production_auto.py          # Main automation script
├── DAILY PRODUCTION TEMPLATE (7).xlsx  # Master template (updated daily)
└── Daily's/                          # Drop crew report files here
    ├── ATT LG 01-1742-25.xlsx
    └── ...
```

## Notes

- Place all crew `.xlsx` report files in the `Daily's` subfolder before running
- The script auto-detects two different report layouts
- Duplicate files for the same job are handled automatically (most recent wins)
- Cell `A2` in the template tracks the last run date to prevent double-counting
