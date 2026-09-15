ATX-Scanner - ANTEX File Summary Browser
=========================================

Version: 1.1.0
Release date: 2026-09-01

Desktop application for browsing, filtering, searching, and exporting
GNSS antenna calibration metadata from ANTEX files.

Part of the PCC software suite developed at the Institut für Erdmessung (IfE),
Leibniz Universität Hannover.


CITATION
--------
If you use this software, please cite:

  Kröger, J., Kersten, T. & Schön, S. PCC-Explorer: An open-source software
  tool to assess the impact of GNSS antenna phase center corrections on
  geodetic parameters. GPS Solut 30, 93 (2026).
  https://doi.org/10.1007/s10291-026-02056-2

A DOI for this program itself is planned with the v2.0 release of the suite;
until then please cite the paper above.


QUICK START
-----------
1. Double-click ATX-Scanner.exe to launch.
2. The application starts empty. Load a database or scan a folder to begin.
3. Click "Load Database" to open an existing .db file.
4. Click "Scan ATX Folder" to scan a folder (and ALL subfolders recursively)
   for .atx files. A progress dialog with a Cancel button is shown.
5. Click column headers to sort. Click a row to see full details.


FEATURES
--------
Filtering & Search:
  - Search box: type any text to filter across all fields
  - Dropdown filters: Antenna Name, Radome Code, Method, Institution
  - Frequency checkboxes: tick one or more frequency codes to filter
  - ANTEX version checkboxes: tick 1.4 and/or 2.0 to show only records from
    files of that format version. The version is read from the file header,
    shown as its own sortable column, and with no box ticked all versions
    are listed.
  - Reset Filters: clear all active filters

ANTEX 2.0 support:
  Both format versions are parsed. Version 2.0 names the same information
  differently, and all of it is read: TYPE / SN (receiver antennas),
  TYPE / SVN / SAT ID (satellite antennas), # OF PHV and START OF PHV. One
  START OF PHV record may list several frequencies; these are split into
  individual entries.

  Note: "Frequencies" means phase-pattern frequencies, exactly as in v1.4.
  An ANTEX 2.0 antenna carrying only gain patterns (GDV/GAIN) therefore
  shows none, which is correct - counting those would present an antenna
  as phase-calibrated when it is not.

Data Sources:
  - Load Database: open any .db file
  - Scan ATX Folder: select a folder; all subfolders are searched recursively
    for .atx files. A Cancel button is available during scanning.

Export:
  - Export to Excel: save filtered records as formatted .xlsx file
  - Export to CSV: save filtered records as .csv file
  - Export Database: save filtered records as SQLite .db file to a
    user-defined name and path


DATABASE COMPATIBILITY
----------------------
The .db files use standard SQLite format and can be opened with:
  - DB Browser for SQLite (free) - https://sqlitebrowser.org
  - DBeaver (free) - https://dbeaver.io
  - Python sqlite3 module - built into Python
  - SQLiteStudio (free) - https://sqlitestudio.pl
  - DataGrip (commercial) - https://www.jetbrains.com/datagrip


FIELDS EXTRACTED
----------------
  ANTEX Version       Format version of the file: 1.4 or 2.0
  Antenna Type        Antenna name including radome code
  Serial Number       Individual antenna serial number
  Radome Code         4-character radome identifier
  Calibration Method  e.g., ROBOT, CHAMBER, FIELD
  Institution         Calibrating institution
  # Calibrations      Number of individual calibrations
  Calibration Date    Date of calibration (YYYY-MM-DD)
  DAZI                Azimuth increment [deg]
  ZEN1                Zenith start angle [deg]
  ZEN2                Zenith end angle [deg]
  DZEN                Zenith increment [deg]
  # Frequencies       Number of frequency bands
  Frequencies         Frequency codes (e.g., G01, G02, R01)
  File Name           ATX file name
  Relative Path       Path relative to scanned root


ADDING NEW ANTEX FILES
----------------------
Option A: Scan a folder directly
  1. Click "Scan ATX Folder" in the toolbar
  2. Select the folder containing .atx files
  3. All subfolders are searched recursively

Option B: Rebuild the database
  1. Place new .atx files in the ATXfiles directory
  2. Run: python parse_antex.py
  3. Click "Load Database" to reload


PACKAGE CONTENTS
----------------
  ATX-Scanner.exe      Standalone application (no installation required)
  antex_database.db    Pre-built database with all parsed ANTEX records
  parse_antex.py       Python script to re-parse ATX files
  README.md            Documentation (Markdown)
  README.txt           Documentation (plain text, this file)
  LICENSE.txt          Short license notice (GPLv3)
  LICENSE              Full GNU General Public License v3 text


SYSTEM REQUIREMENTS
-------------------
  Windows 10 or later
  No additional software installation required


LICENSE
-------
This project is licensed under the GNU General Public License v3.0 or later.
See LICENSE and LICENSE.txt


DEVELOPERS
----------
  Amr Fawzy, M.Sc.
  Dr.-Ing. Johannes Kröger

  Institut für Erdmessung (IfE)
  Leibniz Universität Hannover
  Schneiderberg 50, D-30167 Hannover

  Contact: kroeger@ife.uni-hannover.de
  Web: https://www.ife.uni-hannover.de

STAY INFORMED
-------------

  The Institut für Erdmessung runs a moderated mailing list for its GNSS
  software. It announces new releases and warns you about changes that can
  break your work, for example when a server for satellite orbit products
  moves to a new address. Every program in the suite offers this once when it
  first starts, and the Contact dialog can open it again at any time.

  Subscribe (web form):
    https://listserv.uni-hannover.de/cgi-bin/wa?SUBED1=SOFTWARE-IFE&A=1

  Subscribe by e-mail:
    send the single line   subscribe software-ife
    to                     listserv@listserv.uni-hannover.de

  Send that command line on its own. LISTSERV reads the message body line by
  line, so a signature added by your mail program can stop it.

  List address: SOFTWARE-IFE@LISTSERV.UNI-HANNOVER.DE

  LISTSERV answers with a confirmation mail. The subscription becomes active
  only after you reply to it and a moderator approves the request. Subscribing
  is voluntary and you can leave the list at any time.
