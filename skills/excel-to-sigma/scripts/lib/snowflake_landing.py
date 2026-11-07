"""Phase 1b — land an Excel workbook's data Tables in Snowflake.

Pure planning helpers (identifiers, type inference, CSV, ``snow`` argv generation); the CLI
in ``scripts/land-in-snowflake.py`` runs the argv lists. Nothing here touches the network.
"""
from __future__ import annotations

import csv
import datetime as dt
import json
import re
from pathlib import Path

NUMBER = "NUMBER(38,6)"
# Excel serial range covering 1970..2100: anything plausible as a date, nothing like a small count.
SERIAL_MIN, SERIAL_MAX = 25569, 73050
DATE_HEADER = re.compile(r"(date|month|period|day)", re.I)
EPOCH = dt.date(1899, 12, 30)
# Database / schema / role / table-prefix values go into SQL text unquoted, so each must be one plain
# unquoted Snowflake identifier: no dots, quotes, spaces, semicolons or a leading digit / '-'.
PLAIN_IDENTIFIER = re.compile(r"[A-Za-z_][A-Za-z0-9_$]*")


def validate_identifier(value, label: str, allow_empty: bool = False) -> str:
    """``value`` when it is one plain unquoted Snowflake identifier (``^[A-Za-z_][A-Za-z0-9_$]*$``); a
    table prefix may also be empty. Raises ``ValueError`` before any SQL is built."""
    if allow_empty and value == "":
        return value
    if not isinstance(value, str) or not PLAIN_IDENTIFIER.fullmatch(value):
        raise ValueError(f"{label} {value!r} is not a plain Snowflake identifier "
                         "(letters, digits, '_' and '$'; starting with a letter or '_')")
    return value


def _check_plan(plan: dict) -> None:
    """A plan read back from landing.json is checked again: every name that reaches SQL text."""
    validate_identifier(plan.get("database"), "database")
    validate_identifier(plan.get("schema"), "schema")
    for table in plan.get("tables", []):
        validate_identifier(table.get("snowflakeTable"), "table")


def snowflake_identifier(header: str) -> str:
    """UPPER_SNAKE ``[A-Z0-9_]`` identifier; a leading digit gets a ``C_`` prefix."""
    text = re.sub(r"[^A-Za-z0-9]+", "_", str(header)).strip("_").upper()
    if not text:
        raise ValueError(f"header {header!r} has no identifier characters")
    return "C_" + text if text[0].isdigit() else text


def snowflake_identifiers(headers: list[str]) -> list[str]:
    seen: dict[str, str] = {}
    out = []
    for header in headers:
        name = snowflake_identifier(header)
        if name in seen:
            raise ValueError(f"headers {seen[name]!r} and {header!r} both map to column {name}")
        seen[name] = header
        out.append(name)
    return out


def _is_num(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _is_date(v) -> bool:
    return isinstance(v, (dt.datetime, dt.date))


def infer_column_types(rows: list[dict], headers: list[str]) -> dict[str, str]:
    types = {}
    for header in headers:
        values = [r.get(header) for r in rows if r.get(header) not in (None, "")]
        if not values:
            types[header] = "VARCHAR"
        elif all(_is_date(v) for v in values):
            types[header] = "DATE"
        elif all(_is_num(v) for v in values):
            serials = (DATE_HEADER.search(str(header)) and
                       all(float(v).is_integer() and SERIAL_MIN <= v <= SERIAL_MAX for v in values))
            types[header] = "DATE" if serials else NUMBER
        else:
            types[header] = "VARCHAR"
    return types


def csv_cell(value, sql_type: str) -> str:
    if value is None or value == "":
        return ""
    if sql_type == "DATE":
        if isinstance(value, dt.datetime):
            return value.date().isoformat()
        if isinstance(value, dt.date):
            return value.isoformat()
        return (EPOCH + dt.timedelta(days=int(value))).isoformat()
    if sql_type == NUMBER:
        return repr(float(value)) if isinstance(value, float) else str(value)
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if _is_date(value):
        return value.isoformat()
    return str(value)


def read_table(source: Path, table: str) -> tuple[list[str], list[dict]]:
    from openpyxl import load_workbook
    wb = load_workbook(source, data_only=True)
    for ws in wb.worksheets:
        if table in ws.tables:
            ref = ws.tables[table].ref
            grid = [[c.value for c in row] for row in ws[ref]]
            headers = [str(h).strip() for h in grid[0]]
            rows = [dict(zip(headers, r)) for r in grid[1:] if any(v not in (None, "") for v in r)]
            return headers, rows
    raise ValueError(f"table {table!r} not found in {source}")


def snowflake_table_name(excel_table: str, prefix: str) -> str:
    base = excel_table[3:] if excel_table.startswith("tbl") else excel_table
    return prefix + snowflake_identifier(base)


def landing_plan(source: Path, tables: list[str], database: str, schema: str, prefix: str,
                 out_dir: Path | None = None) -> dict:
    """Plan + one CSV per table, written beside ``out_dir`` (default: a ``landing/`` dir next to the source)."""
    validate_identifier(database, "--database")
    validate_identifier(schema, "--schema")
    validate_identifier(prefix, "--table-prefix", allow_empty=True)
    source = Path(source)
    out_dir = Path(out_dir) if out_dir else source.parent / "landing"
    out_dir.mkdir(parents=True, exist_ok=True)
    plan = {"database": database, "schema": schema, "tables": []}
    for table in tables:
        headers, rows = read_table(source, table)
        names = snowflake_identifiers(headers)
        types = infer_column_types(rows, headers)
        target = snowflake_table_name(table, prefix)
        path = (out_dir / f"{target}.csv").resolve()
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(names)
            for row in rows:
                w.writerow([csv_cell(row.get(h), types[h]) for h in headers])
        plan["tables"].append({
            "excelTable": table, "snowflakeTable": target,
            "columns": [{"header": h, "name": n, "type": types[h]} for h, n in zip(headers, names)],
            "rowCount": len(rows), "csv": str(path)})
    return plan


def fqn(plan: dict, table: dict) -> str:
    return f"{plan['database']}.{plan['schema']}.{table['snowflakeTable']}"


def _sql(connection: str, statement: str, *, json_out: bool = False) -> list[str]:
    argv = ["snow", "sql", "--connection", connection]
    if json_out:
        argv += ["--format", "json"]
    return argv + ["-q", statement]


def exists_command(plan: dict, table: dict, connection: str) -> list[str]:
    _check_plan({**plan, "tables": [table]})
    return _sql(connection, f"SHOW TABLES LIKE '{table['snowflakeTable']}' IN SCHEMA "
                f"{plan['database']}.{plan['schema']}", json_out=True)


def count_command(plan: dict, table: dict, connection: str) -> list[str]:
    _check_plan({**plan, "tables": [table]})
    return _sql(connection, f"SELECT COUNT(*) AS N FROM {fqn(plan, table)}", json_out=True)


def snow_commands(plan: dict, connection: str, replace: bool, grant_role: str | None) -> list[list[str]]:
    """argv lists, per table: create, PUT, COPY INTO, optional per-table GRANT, then COUNT(*).

    Never creates a schema; ``CREATE OR REPLACE ... COPY GRANTS`` only when ``replace`` (a replaced table
    keeps the privileges granted on the one it replaces). Every identifier is validated first.
    """
    _check_plan(plan)
    if grant_role is not None:
        validate_identifier(grant_role, "--grant-select-to")
    commands = []
    for table in plan["tables"]:
        name = fqn(plan, table)
        csv_name = Path(table["csv"]).name
        stage = f"@{plan['database']}.{plan['schema']}.%{table['snowflakeTable']}"
        cols = ", ".join(f'"{c["name"]}" {c["type"]}' for c in table["columns"])
        create = (f"CREATE OR REPLACE TABLE {name} ({cols}) COPY GRANTS" if replace
                  else f"CREATE TABLE {name} ({cols})")
        commands.append(_sql(connection, create))
        commands.append(_sql(connection, f"PUT 'file://{table['csv']}' {stage} AUTO_COMPRESS=TRUE OVERWRITE=TRUE"))
        commands.append(_sql(
            connection,
            f"COPY INTO {name} FROM {stage} FILES = ('{csv_name}.gz') "
            "FILE_FORMAT = (TYPE = CSV SKIP_HEADER = 1 FIELD_OPTIONALLY_ENCLOSED_BY = '\"' "
            "EMPTY_FIELD_AS_NULL = TRUE NULL_IF = ('')) PURGE = TRUE ON_ERROR = ABORT_STATEMENT"))
        if grant_role:
            commands.append(_sql(connection, f"GRANT SELECT ON TABLE {name} TO ROLE {grant_role}"))
        commands.append(_sql(connection, f"SELECT COUNT(*) AS N FROM {name}", json_out=True))
    return commands


def sigma_sync_command(plan: dict, table: dict, connection_id: str, profile: str | None) -> list[str]:
    argv = ["sigma"] + (["-p", profile] if profile else [])
    return argv + ["api", "connections", "sync", "--params", json.dumps({"connectionId": connection_id}),
                   "--json", json.dumps({"path": [plan["database"], plan["schema"], table["snowflakeTable"]]})]
