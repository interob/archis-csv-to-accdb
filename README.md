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
