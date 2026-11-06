# ATX-Scanner

Find out which GNSS antenna calibrations you actually hold.

> **Version 1.1.1, released 1 November 2026.** Download the Windows package
> from [Releases](https://github.com/J-kroeger/atx-scanner/releases/latest), unzip it and start
> `ATX-Scanner.exe`. No installation and no Python required. The full source code is in
> this repository.

## What it does

Point it at a folder of ANTEX files and it reads every calibration inside them
into one searchable table. Answering "do we have a robot calibration for this
antenna, and which ANTEX version is it in?" stops being a matter of opening
files one by one.

Each record can be filtered by antenna name, radome code, calibration method,
institution and ANTEX version, and the whole table exports to CSV.

## Why it exists

Calibration files accumulate. They arrive from different facilities, in
different ANTEX versions, with different serial numbers and radome codes, and
the same antenna type is often calibrated several times. Before you can compare
two calibrations you have to know what you have, and that is a surprisingly
tedious question to answer by hand.

## Features

- Scans whole folders of ANTEX files, including subfolders
- One table of every calibration found, with the file each came from
- Filters by antenna, radome, calibration method, institution and ANTEX version
- Full-text search across all fields
- Shows the frequencies each calibration covers and its zenith grid
- Exports to CSV, or saves the scan as a database to reopen later
- Opens the source file in Explorer, or copies its path


## Running from source

The program is written in Python and was built with Python 3.12.

```
python -m pip install -r requirements.txt
python atx_scanner.py
```

On Windows, `launch_gui.bat` does the same when no packaged executable is next to it.
The full user guide is in [`README.txt`](README.txt).

## Part of PCC-Suite

This program is one of seven released together as
[PCC-Suite](https://github.com/J-kroeger/pcc-suite), a collection of open-source programs for GNSS antenna
calibration values from the Institut für Erdmessung (IfE), Leibniz Universität
Hannover. Each is a standalone Windows executable, released and versioned
separately, so you can take only the one you need. No installation, no Python
required.

Archived releases and DOIs are gathered in the Zenodo community
[Open Source Software Packages for GNSS Data Processing](https://zenodo.org/communities/gnss-open-source-solutions).

## Licence

GNU General Public License v3.0 or later. Free to use, share and modify. See
[LICENSE](LICENSE).

## Citation

The method behind the suite and its validation against real PPP solutions are
described in:

> Kröger, J., Kersten, T. & Schön, S. (2026). PCC-Explorer: an open-source
> software tool to assess the impact of GNSS antenna phase center corrections on
> geodetic parameters. *GPS Solutions* **30**, 93. [https://doi.org/10.1007/s10291-026-02056-2](https://doi.org/10.1007/s10291-026-02056-2)

## Stay informed

Release announcements, and warnings when an external data source moves, go to
the institute software mailing list:

```
ife-software@lists.uni-hannover.de
```

Subscribe at <https://lists.uni-hannover.de/sympa/subscribe/ife-software>, or send
the line `subscribe ife-software` and, on the next line, `QUIT` to
<sympa@lists.uni-hannover.de>. The program offers to subscribe you on first
start. You can decline, and you can ask not to be reminded again.

## Contact

**Dr.-Ing. Johannes Kröger**
Institut für Erdmessung (IfE), Leibniz Universität Hannover
Schneiderberg 50, D-30167 Hannover

Email: [kroeger@ife.uni-hannover.de](mailto:kroeger@ife.uni-hannover.de)
Web: [www.ife.uni-hannover.de](https://www.ife.uni-hannover.de)
