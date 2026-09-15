#!/usr/bin/env python3
"""
ANTEX File Parser & Summary Generator
======================================
Parses all .atx files in the ATXfiles directory structure and generates:
  1. An Excel spreadsheet (antex_summary.xlsx) with all extracted metadata
  2. A SQLite database (antex_database.db) for future updates and queries

Extracted fields per file:
  - Antenna name (including radome code)
  - Serial number
  - Calibration method (ROBOT, CHAMBER, etc.)
  - Calibration institution (IfE, Geopp, etc.)
  - Number of individual calibrations
  - Calibration date(s)
  - PCC resolution: DAZI, ZEN1, ZEN2, DZEN
  - Frequencies (e.g., G01, G02, R01, R02, ...)
  - Full file path and ATX file name

Author: Amr Fawzy, Institut für Erdmessung (IfE), Leibniz Universität Hannover
Date: 2026-04-01
"""

import os
import re
import sqlite3
from pathlib import Path
from datetime import datetime

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False
    print("[WARNING] openpyxl not installed. Run: pip install openpyxl")
    print("          Will generate CSV fallback instead.\n")


# ── Configuration ──────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
ATX_ROOT = BASE_DIR / "Summary_ATX" / "Summary_ATX" / "ATXfiles"
OUTPUT_XLSX = BASE_DIR / "antex_summary.xlsx"
OUTPUT_CSV = BASE_DIR / "antex_summary.csv"
OUTPUT_DB = BASE_DIR / "antex_database.db"


# ── Date Normalization ────────────────────────────────────────────────────────
def _normalize_date(date_str: str) -> str:
    """
    Normalize date strings from ANTEX files to YYYY-MM-DD format.
    
    Handles formats like:
      '01-Sep-20'   -> '2020-09-01'
      '03-MAR-12'   -> '2012-03-03'
      '2000-05-28'  -> '2000-05-28'  (already correct)
      ''            -> ''
    """
    if not date_str or not date_str.strip():
        return ""
    
    date_str = date_str.strip()
    
    # Already in YYYY-MM-DD format
    if len(date_str) == 10 and date_str[4] == "-" and date_str[7] == "-":
        return date_str
    
    # Try DD-Mon-YY or DD-MON-YY (e.g. '01-Sep-20', '03-MAR-12')
    from datetime import datetime as dt
    for fmt in ("%d-%b-%y", "%d-%B-%Y", "%Y-%m-%d"):
        try:
            parsed = dt.strptime(date_str, fmt)
            return parsed.strftime("%Y-%m-%d")
        except ValueError:
            continue
    
    # Return original if no format matches
    return date_str


# ── ANTEX Record Parsing ──────────────────────────────────────────────────────
def parse_antex_file(filepath: Path) -> list[dict]:
    """
    Parse a single ANTEX file and return a list of antenna records.
    
    An ANTEX file can contain multiple antenna blocks (START OF ANTENNA /
    END OF ANTENNA), though in practice most files contain exactly one.
    Each block yields one record in the output.
    """
    records = []
    
    try:
        with open(filepath, "r", encoding="utf-8", errors="replace") as f:
            lines = f.readlines()
    except Exception as e:
        print(f"  [ERROR] Could not read {filepath}: {e}")
        return records
    
    # Known ANTEX label strings for robust matching.
    # Some files have non-standard column widths where data bleeds past col 60,
    # so we search for the label anywhere in the line suffix.
    # ANTEX 1.4 and 2.0 name several records differently. Without the 2.0 names
    # a v2.0 file parsed only half-way: METH, DAZI and ZEN came through (those
    # records are identical in both versions) while the antenna type, serial
    # number and frequencies stayed blank.
    #
    #   ANTEX 1.4              ANTEX 2.0
    #   TYPE / SERIAL NO       TYPE / SN   (receiver antennas)
    #                          TYPE / SVN / SAT ID   (satellite antennas)
    #   # OF FREQUENCIES       # OF PHV
    #   START OF FREQUENCY     START OF PHV   (may list SEVERAL frequencies)
    LABEL_MAP = {
        "START OF ANTENNA": "START OF ANTENNA",
        "END OF ANTENNA": "END OF ANTENNA",
        "TYPE / SERIAL NO": "TYPE / SERIAL NO",
        "TYPE / SVN / SAT ID": "TYPE / SVN / SAT ID",
        "TYPE / SN": "TYPE / SN",
        "METH / BY / # / DATE": "METH / BY / # / DATE",
        "DAZI": "DAZI",
        "ZEN1 / ZEN2 / DZEN": "ZEN1 / ZEN2 / DZEN",
        "# OF FREQUENCIES": "# OF FREQUENCIES",
        "# OF PHV": "# OF PHV",
        "START OF FREQUENCY": "START OF FREQUENCY",
        "START OF PHV": "START OF PHV",
    }
    
    def identify_label(raw_line: str) -> str:
        """Identify the ANTEX record label from a line, robust to column shifts."""
        # Check from longest to shortest label to avoid partial matches
        for label in sorted(LABEL_MAP.keys(), key=len, reverse=True):
            if label in raw_line:
                return label
        return ""
    
    # ANTEX version from the header, e.g. "     1.4            M" on the
    # "ANTEX VERSION / SYST" record. It applies to the whole file, so it is
    # read once and attached to every antenna record.
    antex_version = ""
    for line in lines[:20]:
        if "ANTEX VERSION" in line:
            antex_version = line[:8].strip()
            break

    # State machine for parsing
    in_antenna = False
    current = None

    for line in lines:
        # Strip trailing whitespace / CR/LF but preserve leading spaces (fixed-width)
        raw = line.rstrip("\r\n")
        
        # Identify the record label (robust to column shifts)
        label = identify_label(raw)
        
        # Extract content: everything before the label starts
        if label:
            label_pos = raw.index(label)
            content = raw[:label_pos]
        else:
            content = raw
        
        # ── Antenna block boundaries ──
        if label == "START OF ANTENNA":
            in_antenna = True
            current = {
                "antex_version": antex_version,
                "antenna_type": "",
                "radome": "",
                "serial_number": "",
                "method": "",
                "institution": "",
                "num_calibrations": "",
                "calibration_date": "",
                "dazi": "",
                "zen1": "",
                "zen2": "",
                "dzen": "",
                "num_frequencies": "",
                "frequencies": [],
                "file_name": filepath.name,
                "file_path": str(filepath),
                "relative_path": "",
            }
            # Compute relative path from ATX_ROOT
            try:
                current["relative_path"] = str(filepath.relative_to(ATX_ROOT))
            except ValueError:
                current["relative_path"] = str(filepath)
            continue
        
        if label == "END OF ANTENNA":
            if current is not None:
                # Join frequencies into a comma-separated string
                current["frequencies_str"] = ", ".join(current["frequencies"])
                current["num_frequencies_detected"] = str(len(current["frequencies"]))
                records.append(current)
            in_antenna = False
            current = None
            continue
        
        if not in_antenna or current is None:
            continue
        
        # ── Parse individual records ──
        if label == "TYPE / SVN / SAT ID":
            # ANTEX 2.0, satellite antennas: type in cols 1-30, SVN in 41-50,
            # COSPAR ID in 51-60. The SVN takes the serial-number slot, matching
            # how v1.4 satellite blocks are shown.
            current["antenna_type"] = content[:30].strip()
            current["radome"] = "NONE"
            current["serial_number"] = content[40:50].strip()

        elif label in ("TYPE / SERIAL NO", "TYPE / SN"):
            # Cols 1-20: antenna type (including radome in cols 17-20 typically)
            # Cols 21-40: serial number
            type_field = content[:20].strip()
            serial_field = content[20:40].strip()
            
            # Parse antenna type with radome
            # Standard ANTEX: cols 1-16 = antenna type, cols 17-20 = radome code
            antenna_name_part = content[:16].strip()
            radome_part = content[16:20].strip()
            
            if not radome_part:
                radome_part = "NONE"
            
            current["antenna_type"] = f"{antenna_name_part} {radome_part}".strip()
            current["radome"] = radome_part
            current["serial_number"] = serial_field
        
        elif label == "METH / BY / # / DATE":
            # Format: METHOD(cols 1-20) AGENCY(cols 21-40) #(cols 41-45) DATE(cols 46-end)
            method = content[:20].strip()
            agency = content[20:40].strip()
            num_cal = content[40:46].strip()
            date_str = content[46:].strip()
            
            current["method"] = method.upper()        # Normalize: ROBOT/robot -> ROBOT
            current["institution"] = agency
            current["num_calibrations"] = num_cal
            current["calibration_date"] = _normalize_date(date_str)
        
        elif label == "DAZI":
            dazi = content.strip()
            current["dazi"] = dazi
        
        elif label == "ZEN1 / ZEN2 / DZEN":
            parts = content.split()
            if len(parts) >= 3:
                current["zen1"] = parts[0]
                current["zen2"] = parts[1]
                current["dzen"] = parts[2]
            elif len(parts) == 2:
                current["zen1"] = parts[0]
                current["zen2"] = parts[1]
            elif len(parts) == 1:
                current["zen1"] = parts[0]
        
        elif label in ("# OF FREQUENCIES", "# OF PHV"):
            current["num_frequencies"] = content.strip()

        elif label in ("START OF FREQUENCY", "START OF PHV"):
            # v1.4 names one frequency per record; v2.0 lists every frequency
            # that shares the pattern, e.g. "   G01   S01   E01   C01   J01".
            for freq_code in content.split():
                if freq_code not in current["frequencies"]:
                    current["frequencies"].append(freq_code)
    
    # Handle case where file ends without END OF ANTENNA
    if current is not None and current["antenna_type"]:
        current["frequencies_str"] = ", ".join(current["frequencies"])
        current["num_frequencies_detected"] = str(len(current["frequencies"]))
        records.append(current)
    
    return records


# ── Directory Scanner ─────────────────────────────────────────────────────────
def scan_atx_files(root: Path) -> list[dict]:
    """Recursively find and parse all .atx files under root."""
    all_records = []
    atx_files = sorted(root.rglob("*.atx"))
    
    print(f"Found {len(atx_files)} ANTEX files to process...\n")
    
    for i, fpath in enumerate(atx_files, 1):
        if i % 50 == 0 or i == len(atx_files):
            print(f"  Processing file {i}/{len(atx_files)}: {fpath.name}")
        
        records = parse_antex_file(fpath)
        all_records.extend(records)
    
    print(f"\nTotal antenna records extracted: {len(all_records)}")
    return all_records


# ── Excel Output ──────────────────────────────────────────────────────────────
COLUMNS = [
    ("ANTEX Version", "antex_version"),
    ("Antenna Type (with Radome)", "antenna_type"),
    ("Serial Number", "serial_number"),
    ("Radome Code", "radome"),
    ("Calibration Method", "method"),
    ("Calibration Institution", "institution"),
    ("# Calibrations", "num_calibrations"),
    ("Calibration Date", "calibration_date"),
    ("DAZI [°]", "dazi"),
    ("ZEN1 [°]", "zen1"),
    ("ZEN2 [°]", "zen2"),
    ("DZEN [°]", "dzen"),
    ("# Frequencies", "num_frequencies_detected"),
    ("Frequencies", "frequencies_str"),
    ("ATX File Name", "file_name"),
    ("Relative Path", "relative_path"),
]


def write_excel(records: list[dict], output_path: Path):
    """Write records to a formatted Excel workbook with filtering enabled."""
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ANTEX Summary"
    
    # ── Styles ──
    header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    data_font = Font(name="Calibri", size=10)
    data_alignment = Alignment(vertical="center", wrap_text=False)
    
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    
    alt_fill = PatternFill(start_color="E8EDF5", end_color="E8EDF5", fill_type="solid")
    
    # ── Write header row ──
    for col_idx, (header_text, _) in enumerate(COLUMNS, 1):
        cell = ws.cell(row=1, column=col_idx, value=header_text)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
    
    # ── Write data rows ──
    for row_idx, record in enumerate(records, 2):
        for col_idx, (_, key) in enumerate(COLUMNS, 1):
            value = record.get(key, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = data_font
            cell.alignment = data_alignment
            cell.border = thin_border
            
            # Alternate row shading
            if row_idx % 2 == 0:
                cell.fill = alt_fill
    
    # ── Auto-fit column widths ──
    for col_idx, (header_text, key) in enumerate(COLUMNS, 1):
        max_len = len(header_text)
        for record in records:
            val = str(record.get(key, ""))
            max_len = max(max_len, len(val))
        # Clamp width between 10 and 60 characters
        adjusted_width = min(max(max_len + 2, 10), 60)
        ws.column_dimensions[get_column_letter(col_idx)].width = adjusted_width
    
    # ── Enable auto-filter ──
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}{len(records) + 1}"
    
    # ── Freeze header row ──
    ws.freeze_panes = "A2"
    
    # ── Save ──
    wb.save(output_path)
    print(f"\n[OK] Excel file saved: {output_path}")


def write_csv(records: list[dict], output_path: Path):
    """Fallback: write records to CSV if openpyxl is not available."""
    import csv
    
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=[key for _, key in COLUMNS],
            extrasaction="ignore",
        )
        # Write header with human-readable names
        writer.writerow({key: header for header, key in COLUMNS})
        writer.writerows(records)
    
    print(f"\n[OK] CSV file saved: {output_path}")


# ── SQLite Database ───────────────────────────────────────────────────────────
def write_database(records: list[dict], db_path: Path):
    """
    Create/update a SQLite database with all antenna records.
    
    The database uses a single table 'antennas' that can be extended
    for future ANTEX file imports. A unique constraint on
    (antenna_type, serial_number, calibration_date, file_name) prevents
    duplicate entries when re-running.
    """
    conn = sqlite3.connect(str(db_path))
    cursor = conn.cursor()
    
    # Create table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS antennas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            antenna_type TEXT,
            serial_number TEXT,
            radome TEXT,
            method TEXT,
            institution TEXT,
            num_calibrations TEXT,
            calibration_date TEXT,
            dazi TEXT,
            zen1 TEXT,
            zen2 TEXT,
            dzen TEXT,
            num_frequencies TEXT,
            frequencies TEXT,
            file_name TEXT,
            relative_path TEXT,
            file_path TEXT,
            imported_at TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(antenna_type, serial_number, calibration_date, file_name)
        )
    """)
    
    # Create useful indexes
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_antenna_type ON antennas(antenna_type)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_serial_number ON antennas(serial_number)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_method ON antennas(method)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_institution ON antennas(institution)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_calibration_date ON antennas(calibration_date)")
    
    # Insert records (REPLACE to handle re-runs)
    insert_sql = """
        INSERT OR REPLACE INTO antennas 
        (antenna_type, serial_number, radome, method, institution, 
         num_calibrations, calibration_date, dazi, zen1, zen2, dzen,
         num_frequencies, frequencies, file_name, relative_path, file_path, imported_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """
    
    now = datetime.now().isoformat()
    
    for record in records:
        cursor.execute(insert_sql, (
            record.get("antenna_type", ""),
            record.get("serial_number", ""),
            record.get("radome", ""),
            record.get("method", ""),
            record.get("institution", ""),
            record.get("num_calibrations", ""),
            record.get("calibration_date", ""),
            record.get("dazi", ""),
            record.get("zen1", ""),
            record.get("zen2", ""),
            record.get("dzen", ""),
            record.get("num_frequencies_detected", ""),
            record.get("frequencies_str", ""),
            record.get("file_name", ""),
            record.get("relative_path", ""),
            record.get("file_path", ""),
            now,
        ))
    
    conn.commit()
    
    # Print summary statistics
    cursor.execute("SELECT COUNT(*) FROM antennas")
    total = cursor.fetchone()[0]
    
    cursor.execute("SELECT COUNT(DISTINCT antenna_type) FROM antennas")
    unique_types = cursor.fetchone()[0]
    
    cursor.execute("SELECT COUNT(DISTINCT serial_number) FROM antennas")
    unique_serials = cursor.fetchone()[0]
    
    cursor.execute("SELECT COUNT(DISTINCT method) FROM antennas")
    unique_methods = cursor.fetchone()[0]
    
    cursor.execute("SELECT COUNT(DISTINCT institution) FROM antennas")
    unique_institutions = cursor.fetchone()[0]
    
    conn.close()
    
    print(f"\n[OK] Database saved: {db_path}")
    print(f"  Total records:             {total}")
    print(f"  Unique antenna types:      {unique_types}")
    print(f"  Unique serial numbers:     {unique_serials}")
    print(f"  Unique calibration methods:{unique_methods}")
    print(f"  Unique institutions:       {unique_institutions}")


# ── Print Summary Statistics ──────────────────────────────────────────────────
def print_summary(records: list[dict]):
    """Print a summary of the parsed data to the console."""
    
    print("\n" + "=" * 70)
    print("  ANTEX FILES SUMMARY")
    print("=" * 70)
    
    # Count unique values
    types = set(r["antenna_type"] for r in records)
    serials = set(r["serial_number"] for r in records)
    methods = set(r["method"] for r in records)
    institutions = set(r["institution"] for r in records)
    
    print(f"\n  Total ATX records:          {len(records)}")
    print(f"  Unique antenna types:       {len(types)}")
    print(f"  Unique serial numbers:      {len(serials)}")
    print(f"  Calibration methods:        {', '.join(sorted(methods))}")
    print(f"  Calibration institutions:   {', '.join(sorted(institutions))}")
    
    # Frequency distribution
    all_freqs = set()
    for r in records:
        all_freqs.update(r.get("frequencies", []))
    print(f"  Unique frequencies found:   {', '.join(sorted(all_freqs))}")
    
    # Method distribution
    print("\n  Records by calibration method:")
    method_counts = {}
    for r in records:
        m = r["method"]
        method_counts[m] = method_counts.get(m, 0) + 1
    for m, count in sorted(method_counts.items(), key=lambda x: -x[1]):
        print(f"    {m:<25s} {count:>5d}")
    
    # Institution distribution
    print("\n  Records by institution:")
    inst_counts = {}
    for r in records:
        inst = r["institution"]
        inst_counts[inst] = inst_counts.get(inst, 0) + 1
    for inst, count in sorted(inst_counts.items(), key=lambda x: -x[1]):
        print(f"    {inst:<35s} {count:>5d}")
    
    print("\n" + "=" * 70)


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    print("=" * 70)
    print("  ANTEX File Parser & Summary Generator")
    print("  IfE, Leibniz Universität Hannover")
    print("=" * 70)
    print(f"\n  Scanning directory: {ATX_ROOT}\n")
    
    if not ATX_ROOT.exists():
        print(f"[ERROR] ATX directory not found: {ATX_ROOT}")
        print("        Please ensure the ATXfiles folder exists.")
        return
    
    # 1. Scan and parse all files
    records = scan_atx_files(ATX_ROOT)
    
    if not records:
        print("\n[WARNING] No antenna records found. Check the directory structure.")
        return
    
    # 2. Print summary
    print_summary(records)
    
    # 3. Write Excel (or CSV fallback)
    if HAS_OPENPYXL:
        write_excel(records, OUTPUT_XLSX)
    else:
        write_csv(records, OUTPUT_CSV)
    
    # 4. Write SQLite database
    write_database(records, OUTPUT_DB)
    
    print("\n[OK] All done! Files generated:")
    if HAS_OPENPYXL:
        print(f"  - {OUTPUT_XLSX}")
    else:
        print(f"  - {OUTPUT_CSV}")
    print(f"  - {OUTPUT_DB}")
    print()


if __name__ == "__main__":
    main()
