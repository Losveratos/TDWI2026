#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Baut aus ZlatyLev_DWH.xlsx zwei DB-Artefakte des Sternschemas:

  ZlatyLev_DWH.sqlite        - fertige SQLite-Datenbank (Schema + Daten + FKs)
  dwh_postgres_dump.sql      - self-contained PostgreSQL-Dump (pg_dump-Stil,
                               COPY ... FROM stdin mit Inline-Daten), wiederher-
                               stellbar via:  psql -d zlatylev -f dwh_postgres_dump.sql

Snake_case-Namen wie in dwh_postgres.sql / dwh_fabric.sql.
"""
import datetime as dt
import os
import sqlite3

import openpyxl

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DWH = os.path.join(BASE, "ZlatyLev_DWH.xlsx")
SQLITE = os.path.join(BASE, "ZlatyLev_DWH.sqlite")
PG_DUMP = os.path.join(BASE, "dwh_postgres_dump.sql")

# (sheet, sql_table, [(sql_col, type_sqlite, type_pg)], pk-cols)
TABLES = [
    ("DimDate", "dim_date", [
        ("date", "TEXT", "date"), ("date_key", "INTEGER", "integer"),
        ("year", "INTEGER", "integer"), ("quarter", "INTEGER", "integer"),
        ("quarter_name", "TEXT", "text"), ("month_no", "INTEGER", "integer"),
        ("month_name", "TEXT", "text"), ("month_year", "TEXT", "text"),
        ("month_year_sort", "INTEGER", "integer"), ("month_start", "TEXT", "date"),
        ("day_of_month", "INTEGER", "integer"), ("day_name", "TEXT", "text"),
    ], ["date"]),
    ("DimProduct", "dim_product", [
        ("product_key", "INTEGER", "integer"), ("product_name", "TEXT", "text"),
        ("category", "TEXT", "text"), ("style", "TEXT", "text"),
        ("abv", "REAL", "numeric(3,1)"), ("launch_year", "INTEGER", "integer"),
    ], ["product_key"]),
    ("DimCustomer", "dim_customer", [
        ("customer_key", "INTEGER", "integer"), ("customer_name", "TEXT", "text"),
        ("customer_type", "TEXT", "text"), ("street", "TEXT", "text"),
        ("postal_code", "TEXT", "text"), ("city", "TEXT", "text"),
        ("region", "TEXT", "text"), ("country", "TEXT", "text"),
        ("customer_since", "INTEGER", "integer"), ("key_account_manager", "TEXT", "text"),
    ], ["customer_key"]),
    ("DimKennzahl", "dim_kennzahl", [
        ("account_key", "INTEGER", "integer"), ("account_name", "TEXT", "text"),
        ("account_group", "TEXT", "text"), ("unit", "TEXT", "text"),
        ("sort_order", "INTEGER", "integer"),
    ], ["account_key"]),
    ("FactActual", "fact_actual", [
        ("date", "TEXT", "date"), ("customer_key", "INTEGER", "integer"),
        ("product_key", "INTEGER", "integer"), ("account_key", "INTEGER", "integer"),
        ("value", "NUMERIC", "numeric(18,2)"),
    ], None),
]

# Fremdschluessel von fact_actual
FKS = [
    ("date", "dim_date", "date"),
    ("customer_key", "dim_customer", "customer_key"),
    ("product_key", "dim_product", "product_key"),
    ("account_key", "dim_kennzahl", "account_key"),
]


def read_sheet(wb, sheet):
    ws = wb[sheet]
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] is None:
            continue
        rows.append(list(r))
    return rows


def norm_val(v):
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


# ---------------------------------------------------------------------------
# SQLite
# ---------------------------------------------------------------------------

def build_sqlite(data):
    if os.path.exists(SQLITE):
        os.remove(SQLITE)
    con = sqlite3.connect(SQLITE)
    cur = con.cursor()
    cur.execute("PRAGMA foreign_keys = ON;")

    for sheet, tbl, cols, pk in TABLES:
        coldefs = []
        for name, st, _ in cols:
            coldefs.append(f"  {name} {st}")
        if pk:
            coldefs.append(f"  PRIMARY KEY ({', '.join(pk)})")
        if tbl == "fact_actual":
            for col, rt, rc in FKS:
                coldefs.append(
                    f"  FOREIGN KEY ({col}) REFERENCES {rt}({rc})")
        cur.execute(f"DROP TABLE IF EXISTS {tbl};")
        cur.execute(f"CREATE TABLE {tbl} (\n" + ",\n".join(coldefs) + "\n);")

    for sheet, tbl, cols, pk in TABLES:
        rows = [[norm_val(v) for v in row] for row in data[sheet]]
        ph = ", ".join("?" * len(cols))
        cur.executemany(f"INSERT INTO {tbl} VALUES ({ph})", rows)

    cur.execute("CREATE INDEX ix_fact_date ON fact_actual(date);")
    cur.execute("CREATE INDEX ix_fact_prod ON fact_actual(product_key);")
    cur.execute("CREATE INDEX ix_fact_cust ON fact_actual(customer_key);")
    con.commit()

    # Kontrolle
    cur.execute("""SELECT d.year, SUM(f.value) FROM fact_actual f
                   JOIN dim_date d ON f.date = d.date
                   WHERE f.account_key = 1 GROUP BY d.year ORDER BY d.year;""")
    chk = cur.fetchall()
    con.close()
    return chk


# ---------------------------------------------------------------------------
# PostgreSQL-Dump (pg_dump-Stil, self-contained)
# ---------------------------------------------------------------------------

def pg_escape_copy(v):
    if v is None:
        return r"\N"
    s = str(v)
    s = s.replace("\\", "\\\\").replace("\t", "\\t").replace("\n", "\\n").replace("\r", "\\r")
    return s


def build_pg_dump(data):
    L = []
    L.append("-- ZLATÝ LEV BREWERY a.s. — PostgreSQL-Dump (self-contained)")
    L.append("-- Wiederherstellen:  createdb zlatylev && psql -d zlatylev -f dwh_postgres_dump.sql")
    L.append("-- Erzeugt aus ZlatyLev_DWH.xlsx (scripts/gen_db.py).")
    L.append("")
    L.append("SET client_encoding = 'UTF8';")
    L.append("SET standard_conforming_strings = on;")
    L.append("SET client_min_messages = warning;")
    L.append("BEGIN;")
    L.append("")
    L.append("DROP TABLE IF EXISTS fact_actual CASCADE;")
    L.append("DROP TABLE IF EXISTS dim_date CASCADE;")
    L.append("DROP TABLE IF EXISTS dim_product CASCADE;")
    L.append("DROP TABLE IF EXISTS dim_customer CASCADE;")
    L.append("DROP TABLE IF EXISTS dim_kennzahl CASCADE;")
    L.append("")

    # CREATE TABLE
    for sheet, tbl, cols, pk in TABLES:
        coldefs = [f"    {name} {pgt}" for name, _, pgt in cols]
        L.append(f"CREATE TABLE {tbl} (")
        L.append(",\n".join(coldefs))
        L.append(");")
        L.append("")

    # COPY-Daten
    for sheet, tbl, cols, pk in TABLES:
        colnames = ", ".join(c[0] for c in cols)
        L.append(f"COPY {tbl} ({colnames}) FROM stdin;")
        for row in data[sheet]:
            vals = [pg_escape_copy(norm_val(v)) for v in row]
            L.append("\t".join(vals))
        L.append("\\.")
        L.append("")

    # Constraints/Indizes nach dem Load
    L.append("ALTER TABLE dim_date     ADD PRIMARY KEY (date);")
    L.append("ALTER TABLE dim_product  ADD PRIMARY KEY (product_key);")
    L.append("ALTER TABLE dim_customer ADD PRIMARY KEY (customer_key);")
    L.append("ALTER TABLE dim_kennzahl ADD PRIMARY KEY (account_key);")
    L.append("ALTER TABLE fact_actual")
    L.append("    ADD FOREIGN KEY (date)         REFERENCES dim_date(date),")
    L.append("    ADD FOREIGN KEY (customer_key) REFERENCES dim_customer(customer_key),")
    L.append("    ADD FOREIGN KEY (product_key)  REFERENCES dim_product(product_key),")
    L.append("    ADD FOREIGN KEY (account_key)  REFERENCES dim_kennzahl(account_key);")
    L.append("CREATE INDEX ix_fact_date ON fact_actual(date);")
    L.append("CREATE INDEX ix_fact_prod ON fact_actual(product_key);")
    L.append("CREATE INDEX ix_fact_cust ON fact_actual(customer_key);")
    L.append("")
    L.append("COMMIT;")
    L.append("")
    with open(PG_DUMP, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(L))


# ---------------------------------------------------------------------------

def main():
    wb = openpyxl.load_workbook(DWH, data_only=True)
    data = {sheet: read_sheet(wb, sheet) for sheet, *_ in TABLES}

    chk = build_sqlite(data)
    build_pg_dump(data)

    print("OK")
    for sheet, tbl, *_ in TABLES:
        print(f"  {tbl:<14} {len(data[sheet]):>6} Zeilen")
    print("  Volume hl je Jahr (SQLite-Kontrolle):")
    for y, v in chk:
        print(f"    {y}: {v:,.0f} hl")
    print(f"  -> {SQLITE}")
    print(f"  -> {PG_DUMP}")


if __name__ == "__main__":
    main()
