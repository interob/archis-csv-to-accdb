# archis-csv-to-accdb

Stream a delimited CSV file into a Microsoft Access (`.accdb`) table, with
column mapping, type coercion, and geometry extraction driven by a simple
`.ini` schema file.

Built for loading exports from [ARCHIS](https://archis.cultureelerfgoed.nl/)
(the Dutch national archaeological information system) into an Access
database, but the tool itself is generic — nothing in `csv_to_accdb.py`
depends on ARCHIS specifically. `archis.ini` is included as a real-world
example schema covering seven ARCHIS export tables (`onderzoeksmeldingen`,
`complexen`, `sporen`, `structuren`, `vindplaatsen`, `vondsten`,
`vondstmeldingen`).

## Features

- Column mapping and type coercion defined declaratively in an `.ini` file,
  one section per target table.
- Supported types: `TEXT`, `INTEGER`, `BIGINT`, `FLOAT`, `DATETIME`.
- Geometry helpers that pull the first coordinate pair out of a WKT geometry
  string (`WKT_FIRST_X` / `WKT_FIRST_Y`) or a GeoJSON-style
  `FeatureCollection` array (`MULTIPOINT_FIRST_X` / `MULTIPOINT_FIRST_Y`).
- Target table is created automatically if it doesn't exist; if it does,
  its contents are truncated before the new data is loaded.
- Bad rows are skipped (and optionally logged to a CSV) rather than aborting
  the whole run, unless `--abort-on-error` is passed.
- `--dry-run` validates and coerces every row without touching the database.

## Requirements

- Python 3.10+
- [`pyodbc`](https://pypi.org/project/pyodbc/)
- The Microsoft Access Database Engine ODBC driver (Microsoft Access
  Database Engine 2016 Redistributable, or an Office install that provides
  `{Microsoft Access Driver (*.mdb, *.accdb)}`). This is a Windows-only
  driver.

```bash
pip install pyodbc
```

## Running in an OSGeo4W environment (Windows)

[OSGeo4W](https://trac.osgeo.org/osgeo4w/) is a convenient way to get a
self-contained Python 3 environment on Windows without touching a system
Python install. The script has no GDAL/geospatial dependency — OSGeo4W is
just used here as the Python runtime + shell.

1. **Install OSGeo4W**, if not already present. Run the
   [OSGeo4W installer](https://trac.osgeo.org/osgeo4w/) and, under
   "Select Packages", make sure `python3-core` (and `python3-pip`) are
   selected — these are included by default in the Express Install, and
   also come bundled with a full QGIS install. Prefer the 64-bit installer.

2. **Open the OSGeo4W Shell** (Start Menu → OSGeo4W → OSGeo4W Shell). This
   sets up `PATH`/`PYTHONHOME` so `python3` resolves to the OSGeo4W
   interpreter.

3. **Install `pyodbc`** into that environment:

   ```bat
   python3 -m pip install pyodbc
   ```

4. **Install the Microsoft Access Database Engine ODBC driver.** This is a
   separate Microsoft component, not part of OSGeo4W — download the
   [Access Database Engine 2016 Redistributable](https://www.microsoft.com/en-us/download/details.aspx?id=54920)
   and run it.
   - Match the driver's bitness to the OSGeo4W Python (`python3 -c "import struct; print(struct.calcsize('P')*8)"`
     — 64 in almost all current installs), i.e. install `AccessDatabaseEngine_X64.exe`.
   - If 32-bit Microsoft Office is also installed on the machine, the
     64-bit driver install will refuse to proceed. Either install the
     matching 32-bit driver instead (and use a 32-bit OSGeo4W/Python), or
     force the 64-bit driver alongside 32-bit Office with:
     ```bat
     AccessDatabaseEngine_X64.exe /quiet
     ```
   - Verify it registered: open **ODBC Data Sources (64-bit)**
     (`C:\Windows\System32\odbcad32.exe`) → *Drivers* tab → confirm
     `Microsoft Access Driver (*.mdb, *.accdb)` is listed.

5. **Run the script** from the OSGeo4W Shell, `cd`'d into this repo:

   ```bat
   cd C:\path\to\archis-csv-to-accdb
   python3 csv_to_accdb.py --csv vindplaatsen_20260414.csv --db ARCHIS.accdb --table vindplaatsen --config archis.ini
   ```

## Usage

```bash
python csv_to_accdb.py \
    --csv   data.csv \
    --db    database.accdb \
    --table TableName \
    --config schema.ini
```

### Options

| Flag | Description |
|---|---|
| `--csv FILE` | Source CSV file (required) |
| `--db FILE` | Target `.accdb` database file (required) |
| `--table NAME` | Target table name; must match a section in `--config` (required) |
| `--config FILE` | `.ini` file with column → type definitions (required) |
| `--delimiter CHAR` | Field delimiter (default: `;`) |
| `--encoding ENC` | CSV file encoding (default: `utf-8-sig`) |
| `--dry-run` | Validate and coerce every row without writing to the database |
| `--abort-on-error` | Stop on the first bad row instead of skipping it |
| `--error-log FILE` | Write details of all skipped rows to this CSV file |
| `-v`, `--verbose` | Enable DEBUG-level logging |

### Example: loading an ARCHIS export

```bash
python csv_to_accdb.py \
    --csv vindplaatsen_20260414.csv \
    --db  ARCHIS.accdb \
    --table vindplaatsen \
    --config archis.ini
```

## Schema (`.ini`) format

One section per target table. Each entry maps a target column to a source
CSV column and a type:

```ini
[TableName]
col_a     = src_a::TEXT
col_b     = src_b::INTEGER
col_c     = src_c::FLOAT
col_d     = src_d::DATETIME
locatie_x = geometry_wkt::WKT_FIRST_X
locatie_y = geometry_wkt::WKT_FIRST_Y
```

See [`archis.ini`](archis.ini) for a complete example.

## License

[MIT](LICENSE)
