#!/usr/bin/env python3
"""
csv_to_accdb.py  —  Stream a delimited CSV file into a Microsoft Access (.accdb) table.

The column-to-type mapping is read from an .ini file whose section name matches
the target table.

Each entry has the form:
    target_column = source_column::TYPE

where TYPE is one of: TEXT, INTEGER, FLOAT, DATETIME,
MULTIPOINT_FIRST_X, MULTIPOINT_FIRST_Y.

The last two pseudo-types parse a GeoJSON FeatureCollection-style JSON array
from *source_column* and extract the X (longitude / easting) or Y
(latitude / northing) coordinate of the first object's coordinates array.

Table handling
--------------
  • If the table does not exist it is created automatically from the schema.
  • If the table already exists its contents are truncated before loading.

Usage
-----
    python csv_to_accdb.py \\
        --csv   data.csv        \\
        --db    database.accdb  \\
        --table TableName       \\
        --config schema.ini     \\
        [options]

INI format
----------
    [TableName]
    col_a    = src_a::TEXT
    col_b    = src_b::INTEGER
    col_c    = src_c::FLOAT
    col_d    = src_d::DATETIME
    locatie_x = display_coordinates::MULTIPOINT_FIRST_X
    locatie_y = display_coordinates::MULTIPOINT_FIRST_Y

Driver requirement
------------------
    Requires the Microsoft Access Database Engine ODBC driver
    (Microsoft Access Database Engine 2016 Redistributable or Office installation).
    On Windows this is typically available as:
        {Microsoft Access Driver (*.mdb, *.accdb)}
"""

import argparse
import configparser
import csv
import json
import logging
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

import pyodbc

# ── constants ─────────────────────────────────────────────────────────────────

SUPPORTED_TYPES: set[str] = {
    "TEXT", "INTEGER", "BIGINT", "FLOAT", "DATETIME",
    "MULTIPOINT_FIRST_X", "MULTIPOINT_FIRST_Y",
    "WKT_FIRST_X", "WKT_FIRST_Y",
}

# Matches the first coordinate pair in any WKT geometry string.
# Works for POINT(x y), POLYGON((x y, ...)), MULTIPOLYGON(((x y, ...))), etc.
_WKT_COORD_RE = re.compile(r'(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)')

# Tried in order; first match wins.
DATETIME_FORMATS: list[str] = [
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%dT%H:%M:%S",
    "%d-%m-%Y %H:%M:%S",
    "%d/%m/%Y %H:%M:%S",
    "%Y-%m-%d",
    "%d-%m-%Y",
    "%d/%m/%Y",
]

CSV_FIELD_SIZE_LIMIT = 10_000_000  # bytes; needed for large geometry/JSON fields

log = logging.getLogger(__name__)


# ── type coercion ─────────────────────────────────────────────────────────────

def coerce(raw: str, col_type: str) -> Any:
    """
    Convert a stripped CSV string to the appropriate Python type.

    Returns None for empty strings (→ NULL in the database).
    Raises ValueError if conversion fails.
    """
    if raw == "":
        return None
    if col_type == "TEXT":
        return raw
    if col_type in ("INTEGER", "BIGINT"):
        # Validate as a whole number, then return as string.
        # The Access ODBC driver raises HYC00 when binding a Python int outside
        # the INT32 range (it escalates to SQL_BIGINT, which it does not
        # support via SQLBindParameter).  Passing the value as a string lets
        # Access do the implicit conversion to BIGINT on insert, bypassing
        # the driver limitation entirely.
        v = int(raw.replace(" ", "").replace(".", ""))
        return str(v)
    if col_type == "FLOAT":
        # Tolerate both decimal comma (EU locale) and decimal point.
        return float(raw.replace(" ", "").replace(",", "."))
    if col_type == "DATETIME":
        for fmt in DATETIME_FORMATS:
            try:
                dt = datetime.strptime(raw, fmt)
                # Return as string: the Access ODBC driver cannot bind Python
                # datetime objects via SQLBindParameter (raises HYC00).
                return dt.strftime("%Y-%m-%d %H:%M:%S")
            except ValueError:
                continue
        raise ValueError(f"unrecognised datetime value {raw!r}")
    if col_type in ("WKT_FIRST_X", "WKT_FIRST_Y"):
        m = _WKT_COORD_RE.search(raw)
        if not m:
            raise ValueError(f"no coordinate pair found in WKT value: {raw[:80]!r}")
        return float(m.group(1) if col_type == "WKT_FIRST_X" else m.group(2))
    if col_type in ("MULTIPOINT_FIRST_X", "MULTIPOINT_FIRST_Y"):
        try:
            points = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise ValueError(f"invalid JSON: {exc}") from exc
        if not isinstance(points, list) or len(points) == 0:
            raise ValueError("expected a non-empty JSON array")
        coords = points[0].get("coordinates") if isinstance(points[0], dict) else None
        if not isinstance(coords, list) or len(coords) < 2:
            raise ValueError(f"coordinates array missing or too short: {coords!r}")
        idx = 0 if col_type == "MULTIPOINT_FIRST_X" else 1
        return float(coords[idx])
    raise ValueError(f"unknown type {col_type!r}")  # unreachable after schema validation


# ── config / schema loading ───────────────────────────────────────────────────

def load_schema(config_path: Path, table: str) -> dict[str, tuple[str, str]]:
    """
    Read the .ini file and return an ordered mapping for *table*:

        { target_column: (source_column, TYPE) }

    Each INI value must be in the form  source_column::TYPE.

    configparser lowercases keys by default; that is disabled here so column
    names are preserved exactly as written in the .ini file.
    """
    cfg = configparser.RawConfigParser()
    cfg.optionxform = str           # preserve original case
    if not cfg.read(config_path, encoding="utf-8"):
        raise SystemExit(f"[error] Cannot read config file: {config_path}")
    if not cfg.has_section(table):
        available = ", ".join(f"[{s}]" for s in cfg.sections()) or "(none)"
        raise SystemExit(
            f"[error] Section [{table}] not found in {config_path}.\n"
            f"        Available sections: {available}"
        )
    schema: dict[str, tuple[str, str]] = {}
    for target_col, raw_value in cfg.items(table):
        if "::" not in raw_value:
            raise SystemExit(
                f"[error] Column '{target_col}': value must be 'source_column::TYPE', "
                f"got '{raw_value}'."
            )
        source_col, raw_type = raw_value.split("::", maxsplit=1)
        source_col = source_col.strip()
        typ = raw_type.strip().upper()
        if not source_col:
            raise SystemExit(
                f"[error] Column '{target_col}': source column name is empty."
            )
        if typ not in SUPPORTED_TYPES:
            raise SystemExit(
                f"[error] Column '{target_col}': unsupported type '{raw_type}'.\n"
                f"        Allowed types: {', '.join(sorted(SUPPORTED_TYPES))}"
            )
        schema[target_col] = (source_col, typ)
    if not schema:
        raise SystemExit(f"[error] Section [{table}] in {config_path} has no columns.")
    return schema


# ── CSV validation ────────────────────────────────────────────────────────────

def validate_header(
    csv_fields: list[str],
    schema: dict[str, tuple[str, str]],
) -> None:
    """
    Ensure every *source* column referenced in the schema is present in the
    CSV header.  Raises SystemExit with a descriptive message if any are missing.
    """
    present = {f.strip() for f in csv_fields}
    # A source column may be referenced by more than one target; deduplicate
    # but preserve a deterministic order for the error message.
    seen: set[str] = set()
    missing: list[str] = []
    for source_col, _ in schema.values():
        if source_col not in seen:
            seen.add(source_col)
            if source_col not in present:
                missing.append(source_col)
    if missing:
        raise SystemExit(
            f"[error] {len(missing)} source column(s) absent from CSV header:\n"
            + "".join(f"          · {c}\n" for c in missing)
        )


# ── database helpers ──────────────────────────────────────────────────────────

def accdb_connect(db_path: Path) -> pyodbc.Connection:
    conn_str = (
        r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={db_path.resolve()};"
    )
    try:
        conn = pyodbc.connect(conn_str, autocommit=True)
        log.debug("Connected to %s", db_path)
        return conn
    except pyodbc.Error as exc:
        raise SystemExit(f"[error] Cannot connect to {db_path}:\n  {exc}") from exc


def table_exists(conn: pyodbc.Connection, table: str) -> bool:
    return any(
        row.table_name.lower() == table.lower()
        for row in conn.cursor().tables(tableType="TABLE")
    )


def ensure_table(
    conn: pyodbc.Connection,
    table: str,
    schema: dict[str, tuple[str, str]],
) -> None:
    """
    Always drop (if present) and recreate the table from the current schema.
    This guarantees column types match the schema on every run, which matters
    when the schema changes between runs (e.g. INTEGER → BIGINT).
    """
    # INTEGER and BIGINT both create a BIGINT column.  Values are passed to the
    # driver as strings so Access performs the implicit conversion on insert,
    # avoiding the HYC00 raised when binding SQL_BIGINT via SQLBindParameter.
    type_map = {
        "TEXT":               "LONGTEXT",
        "INTEGER":            "BIGINT",
        "BIGINT":             "BIGINT",
        "FLOAT":              "DOUBLE",
        "DATETIME":           "DATETIME",
        "MULTIPOINT_FIRST_X": "DOUBLE",
        "MULTIPOINT_FIRST_Y": "DOUBLE",
        "WKT_FIRST_X":        "DOUBLE",
        "WKT_FIRST_Y":        "DOUBLE",
    }
    col_defs = ", ".join(
        f"[{target_col}] {type_map[typ]}"
        for target_col, (_, typ) in schema.items()
    )
    if table_exists(conn, table):
        # Drop and recreate so column types always reflect the current schema.
        log.info("Table [%s] exists — dropping and recreating.", table)
        conn.execute(f"DROP TABLE [{table}]")
    conn.execute(f"CREATE TABLE [{table}] ({col_defs})")
    log.debug("DDL: CREATE TABLE [%s] (%s)", table, col_defs)
    log.info("Table [%s] ready.", table)


def build_insert_sql(table: str, columns: list[str]) -> str:
    col_list     = ", ".join(f"[{c}]" for c in columns)
    placeholders = ", ".join("?" * len(columns))
    return f"INSERT INTO [{table}] ({col_list}) VALUES ({placeholders})"


# ── core pipeline ─────────────────────────────────────────────────────────────

def process(args: argparse.Namespace) -> None:
    # 1. Schema  { target_col: (source_col, TYPE) }
    schema: dict[str, tuple[str, str]] = load_schema(args.config, args.table)
    ordered_cols = list(schema.keys())
    log.info("Schema loaded: %d column(s) for table [%s]", len(ordered_cols), args.table)

    # 2. CSV sanity check
    csv.field_size_limit(CSV_FIELD_SIZE_LIMIT)
    if not args.csv.exists():
        raise SystemExit(f"[error] CSV file not found: {args.csv}")

    # 3. Database connection & table preparation
    conn: pyodbc.Connection | None = None
    if not args.dry_run:
        if not args.db.exists():
            raise SystemExit(f"[error] Database file not found: {args.db}")
        conn = accdb_connect(args.db)
        ensure_table(conn, table=args.table, schema=schema)

    insert_sql = build_insert_sql(args.table, ordered_cols)
    log.debug("INSERT SQL: %s", insert_sql)

    # 4. Stream, validate, and insert row by row
    inserted = 0
    skipped  = 0
    error_rows: list[tuple[int, str, str]] = []   # (line_no, first_field_value, reason)

    with args.csv.open(newline="", encoding=args.encoding) as fh:
        reader = csv.DictReader(
            fh,
            delimiter=args.delimiter,
            quotechar='"',
            skipinitialspace=True,
        )

        if reader.fieldnames is None:
            raise SystemExit("[error] CSV file appears to be empty.")

        validate_header(list(reader.fieldnames), schema)
        log.info("CSV header OK — all mapped source columns present.")

        cursor = conn.cursor() if conn is not None else None
        if cursor is not None:
            # The Access ODBC driver cannot auto-size Memo (LONGTEXT) parameters
            # and raises HYC00 for large values (e.g. multi-megabyte geometry
            # JSON).  Declaring the full Access Memo capacity (2^30 - 1 chars)
            # pre-allocates a sufficient buffer for every TEXT column.
            cursor.setinputsizes([
                (pyodbc.SQL_WLONGVARCHAR, 1_073_741_823, 0)
                if schema[col][1] == "TEXT" else None
                for col in ordered_cols
            ])

        try:
            for line_no, raw_row in enumerate(reader, start=2):   # header occupies line 1
                record: list[Any] = []
                fail_reason: str | None = None

                for target_col in ordered_cols:
                    source_col, col_type = schema[target_col]
                    raw_val = (raw_row.get(source_col) or "").strip()
                    try:
                        record.append(coerce(raw_val, col_type))
                    except (ValueError, TypeError) as exc:
                        fail_reason = (
                            f"target '{target_col}' ← source '{source_col}' "
                            f"({col_type}): {exc}"
                        )
                        break

                if fail_reason is not None:
                    first_source = schema[ordered_cols[0]][0]
                    first_field = raw_row.get(first_source, "")
                    error_rows.append((line_no, first_field, fail_reason))
                    skipped += 1
                    log.warning("Line %d skipped — %s", line_no, fail_reason)
                    if args.abort_on_error:
                        raise SystemExit(
                            f"[error] Aborted at line {line_no}: {fail_reason}"
                        )
                    continue

                if cursor is not None:
                    cursor.execute(insert_sql, record)
                inserted += 1

                if inserted % 1000 == 0:
                    log.info("  … %d rows inserted", inserted)

        except Exception:
            # Autocommit means we cannot roll back.  Truncating the table
            # leaves it empty rather than in a partially-loaded state.
            if conn is not None:
                log.warning(
                    "Error after %d row(s) — truncating [%s] to avoid partial data.",
                    inserted, args.table,
                )
                conn.execute(f"DELETE FROM [{args.table}]")
            raise

    # 5. Close
    if conn is not None:
        conn.close()

    prefix = "[DRY RUN] " if args.dry_run else ""
    log.info("%sDone.  Inserted: %d  |  Skipped: %d", prefix, inserted, skipped)

    # 6. Error log
    if error_rows:
        if args.error_log:
            with args.error_log.open("w", encoding="utf-8", newline="") as ef:
                writer = csv.writer(ef)
                writer.writerow(["line", "first_field", "reason"])
                writer.writerows(error_rows)
            log.info("Error details written to %s", args.error_log)
        else:
            log.warning(
                "%d row(s) skipped. Use --error-log FILE to save details.", skipped
            )


# ── CLI ───────────────────────────────────────────────────────────────────────

def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        prog="csv_to_accdb",
        description=(
            "Stream a delimited CSV into a Microsoft Access (.accdb) table.\n"
            "Column mapping and types are defined in an .ini configuration file.\n\n"
            "The target table is created automatically if absent;\n"
            "if it already exists its data is truncated before loading."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    req = p.add_argument_group("required arguments")
    req.add_argument("--csv",    required=True, type=Path, metavar="FILE",
                     help="Source CSV file")
    req.add_argument("--db",     required=True, type=Path, metavar="FILE",
                     help="Target .accdb database file")
    req.add_argument("--table",  required=True, metavar="NAME",
                     help="Target table name (must match a section in --config)")
    req.add_argument("--config", required=True, type=Path, metavar="FILE",
                     help=".ini file with column → type definitions")

    csv_opts = p.add_argument_group("CSV options")
    csv_opts.add_argument("--delimiter", default=";", metavar="CHAR",
                          help="Field delimiter (default: ';')")
    csv_opts.add_argument("--encoding",  default="utf-8-sig", metavar="ENC",
                          help="CSV file encoding (default: utf-8-sig)")

    run_opts = p.add_argument_group("run options")
    run_opts.add_argument("--dry-run", action="store_true",
                          help="Validate and coerce every row without writing to the database")
    run_opts.add_argument("--abort-on-error", action="store_true",
                          help=(
                              "Stop on the first bad row (default: skip bad rows and continue); "
                              "successful records in target table are retained."
                          ))
    run_opts.add_argument("--error-log", type=Path, metavar="FILE",
                          help="Write details of all skipped rows to this CSV file")
    run_opts.add_argument("--verbose", "-v", action="store_true",
                          help="Enable DEBUG-level logging")

    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> None:
    args = parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)-8s %(message)s",
        stream=sys.stderr,
    )
    process(args)


if __name__ == "__main__":
    main()
