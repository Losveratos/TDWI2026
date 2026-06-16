# ZLATÝ LEV BREWERY — Demo-Sternschema (Power BI vs. Pyramid vs. SAC)

Realistischer Controlling-Case einer Prager Brauerei (gegr. 1882) für die
Tool-Demo. **Ist** kommt aus einer sauberen DB-Struktur (Sternschema),
**Plan** kommt als Kreuztabelle aus Excel — genau die typische Trennung
DWH ↔ Finance-Planung.

## Dateien

| Datei | Inhalt |
|---|---|
| `ZlatyLev_DWH.xlsx` | **Sternschema**, je Tabelle ein Blatt (als Excel-Tabelle/ListObject): `DimDate`, `DimProduct`, `DimCustomer`, `DimKennzahl`, `FactActual` |
| `ZlatyLev_Plan.xlsx` | **Plan-Kreuztabelle 2026** (Blatt `Plan 2026`): geplant je Großkunde × Produkt × Monat, Net Sales €. Bewusst „roh" wie aus Finance: Titelblock, Beschreibung, verbundene Zellen, Zwischensumme je Kunde, GRAND TOTAL, Kommentar — Power Query muss das entwirren. Datenstart = Zeile 11. |
| `dwh_csv/*.csv` | Jede DWH-Tabelle als CSV (Semikolon, UTF-8) — für DB-Load |
| `dwh_postgres.sql` / `dwh_fabric.sql` | DDL + Lade-Anweisung (CSV) für PostgreSQL bzw. Fabric/Azure SQL |
| `ZlatyLev_DWH.sqlite` | **Fertige SQLite-Datenbank** (Schema + Daten + FKs) — direkt verbindbar |
| `dwh_postgres_dump.sql` | **Self-contained PostgreSQL-Dump** (pg_dump-Stil, Daten inline) |
| `scripts/gen_star_schema.py` | Generator DWH+Plan (liest echte hl-Zahlen aus `ZlatyLev_Sales_Report_1.xlsx`) |
| `scripts/gen_plan_2026.py` | Generator der „gemeinen" Plan-Kreuztabelle 2026 |
| `scripts/gen_db.py` | Generator SQLite-DB + PostgreSQL-Dump (liest `ZlatyLev_DWH.xlsx`) |

Die Volumina (hl) im Fakt sind **exakt** die Jahres-/Monatswerte des
Original-Reports: 2023 = 90.740, 2024 = 98.360, 2025 = 103.120, 2026 = 41.225 hl
(2026 = lfd. Jahr bis April). Net Sales/COGS/Discounts sind daraus mit
realistischen Preisen/Kosten je Produkt und Jahr abgeleitet.

## Modell (Stern)

```
                 ┌─────────────┐
                 │  DimDate    │  (täglich 2023–2026, Datumstabelle)
                 └──────┬──────┘
                        │ Date
   ┌─────────────┐      │      ┌──────────────┐
   │ DimCustomer │──────┤──────│ DimProduct   │
   │ (Großkunde) │ Cust │ Prod │ (11 Biere)   │
   └─────────────┘      │      └──────────────┘
                 ┌──────┴───────┐
                 │  FactActual  │  Date · CustomerKey · ProductKey ·
                 │  (long)      │  AccountKey · Value
                 └──────┬───────┘
                        │ AccountKey
                 ┌──────┴───────┐
                 │ DimKennzahl  │  Volume(hl) · Net Sales · Discounts · COGS
                 └──────────────┘

   ZlatyLev_Plan.xlsx → Power Query (entwirren + unpivot) → FactPlan
        FactPlan: Date · CustomerKey · ProductKey · NetSalesPlan   (nur 2026)
        relate: DimDate[Date] 1:* · DimProduct[ProductKey] 1:* · DimCustomer[CustomerKey] 1:*
```

**Bewusste Designentscheidung:** `FactActual` ist im **Long-/Account-Format**
(eine Kennzahl je Zeile über `DimKennzahl`). Das ist die Mess­wert-Dimension,
wie SAC und Pyramid sie nativ nutzen — ideal, um die drei Tools direkt zu
vergleichen. **Keine** Szenario-Dimension: AC/PL/PY/YTD/MTD entstehen als
DAX-Measures (siehe unten). `FactPlan` ist **nur 2026** und je Großkunde ×
Produkt geplant (Key-Account-Planung) — kommt aber als „rohe" Kreuztabelle
mit verbundenen Zellen und Zwischensummen, die Power Query erst säubert.

## Tabellen & Schlüssel

| Tabelle | Schlüssel | Wichtige Spalten |
|---|---|---|
| `DimDate` | `Date` | `Year, Quarter, MonthNo, MonthName, MonthYear, MonthYearSort` |
| `DimProduct` | `ProductKey` | `ProductName, Category, Style, ABV, LaunchYear` |
| `DimCustomer` | `CustomerKey` | `CustomerName, CustomerType, Street, PostalCode, City, Region, Country, KeyAccountManager` |
| `DimKennzahl` | `AccountKey` | `AccountName (Volume/Net Sales/Discounts/COGS), Unit, SortOrder` |
| `FactActual` | — | `Date, CustomerKey, ProductKey, AccountKey, Value` |

---

## Power Query (M) — eine Source, dann Verweise

Anlegen als zwei **Parameter** und je eine Query pro Tabelle. So gibt es genau
*eine* Quelle (`Src_DWH`), alle Tabellen referenzieren sie.

```m
// Parameter: pDwhPath  (Text)   z.B. "C:\Users\...\Projekt\ZlatyLev_DWH.xlsx"
// Parameter: pPlanPath (Text)   z.B. "C:\Users\...\Projekt\ZlatyLev_Plan.xlsx"

// Query: Src_DWH  (die EINE Quelle)
let
    Src = Excel.Workbook(File.Contents(pDwhPath), null, true)
in
    Src
```

```m
// Query: DimDate   (Verweis auf Src_DWH)
let
    t = Src_DWH{[Item="tbl_DimDate", Kind="Table"]}[Data],
    typed = Table.TransformColumnTypes(t, {
        {"Date", type date}, {"DateKey", Int64.Type}, {"Year", Int64.Type},
        {"Quarter", Int64.Type}, {"QuarterName", type text},
        {"MonthNo", Int64.Type}, {"MonthName", type text},
        {"MonthYear", type text}, {"MonthYearSort", Int64.Type},
        {"MonthStart", type date}, {"DayOfMonth", Int64.Type}, {"DayName", type text}})
in
    typed
```

```m
// Query: DimProduct
let t = Src_DWH{[Item="tbl_DimProduct", Kind="Table"]}[Data] in t

// Query: DimCustomer
let t = Src_DWH{[Item="tbl_DimCustomer", Kind="Table"]}[Data] in t

// Query: DimKennzahl
let t = Src_DWH{[Item="tbl_DimKennzahl", Kind="Table"]}[Data] in t

// Query: FactActual
let
    t = Src_DWH{[Item="tbl_FactActual", Kind="Table"]}[Data],
    typed = Table.TransformColumnTypes(t, {
        {"Date", type date}, {"CustomerKey", Int64.Type},
        {"ProductKey", Int64.Type}, {"AccountKey", Int64.Type},
        {"Value", type number}})
in
    typed
```

### Plan: „gemeine" Kreuztabelle → FactPlan (entwirren + unpivot)

Das Blatt `Plan 2026` hat Titel/Beschreibung (Zeilen 1–10), Kunden-Kopfzeilen als
**verbundene Zellen**, je Kunde eine **Zwischensumme** und am Ende GRAND TOTAL +
Notizen. Trick: Top-Block per `Table.Skip(10)` weg, Kundenname aus den verbundenen
Kopfzeilen **nach unten füllen**, Subtotal-/Total-/Notiz-Zeilen filtern, dann
unpivotieren. Excel liefert generische Spalten `Column1` (Label), `Column2..13`
(Jan–Dez), `Column14` (Total).

```m
// Query: Src_Plan
let Src = Excel.Workbook(File.Contents(pPlanPath), null, true) in Src

// Query: FactPlan
let
    raw   = Src_Plan{[Item="Plan 2026", Kind="Sheet"]}[Data],
    cut   = Table.Skip(raw, 10),                       // Titel/Beschreibung/Header weg
    ren   = Table.RenameColumns(cut, {{"Column1", "Label"}}),
    // Kunden-Kopfzeile: Monate leer, enthält '·', kein Subtotal/GRAND, beginnt nicht mit '·'
    flag  = Table.AddColumn(ren, "CustomerRaw", each
                if [Column2] = null
                   and [Label] <> null
                   and Text.Contains([Label], "·")
                   and not Text.StartsWith([Label], "·")
                   and not Text.StartsWith(Text.Trim([Label]), "Subtotal")
                   and not Text.StartsWith(Text.Trim([Label]), "GRAND")
                then Text.Trim(Text.BeforeDelimiter([Label], "·"))
                else null, type text),
    fill  = Table.FillDown(flag, {"CustomerRaw"}),      // Kundenname nach unten füllen
    // nur echte Produktzeilen behalten
    prod  = Table.SelectRows(fill, each [Column2] <> null
                and not Text.StartsWith(Text.Trim([Label]), "Subtotal")
                and not Text.StartsWith(Text.Trim([Label]), "GRAND")),
    pname = Table.AddColumn(prod, "Product", each Text.Trim([Label]), type text),
    keep  = Table.SelectColumns(pname,
                {"CustomerRaw", "Product", "Column2","Column3","Column4","Column5",
                 "Column6","Column7","Column8","Column9","Column10","Column11",
                 "Column12","Column13"}),
    unpiv = Table.UnpivotOtherColumns(keep, {"CustomerRaw","Product"}, "Col", "NetSalesPlan"),
    mno   = Table.AddColumn(unpiv, "MonthNo",
                each Number.FromText(Text.AfterDelimiter([Col], "Column")) - 1, Int64.Type),
    dated = Table.AddColumn(mno, "Date", each #date(2026, [MonthNo], 1), type date),
    // Keys nachschlagen (Name → Key)
    jc = Table.NestedJoin(dated, {"CustomerRaw"}, DimCustomer, {"CustomerName"}, "c", JoinKind.LeftOuter),
    kc = Table.ExpandTableColumn(jc, "c", {"CustomerKey"}, {"CustomerKey"}),
    jp = Table.NestedJoin(kc, {"Product"}, DimProduct, {"ProductName"}, "p", JoinKind.LeftOuter),
    kp = Table.ExpandTableColumn(jp, "p", {"ProductKey"}, {"ProductKey"}),
    res = Table.SelectColumns(kp, {"Date","CustomerKey","ProductKey","NetSalesPlan"}),
    typed = Table.TransformColumnTypes(res, {
                {"Date", type date}, {"CustomerKey", Int64.Type},
                {"ProductKey", Int64.Type}, {"NetSalesPlan", type number}}),
    clean = Table.SelectRows(typed, each [NetSalesPlan] <> null and [NetSalesPlan] <> 0)
in
    clean
```

> Kontroll-Summe nach dem Laden: `FactPlan` muss **17.474.100 €** ergeben (GRAND TOTAL 2026).

### Später in die Datenbank? Nur `Src_DWH` tauschen.

Alle Dim-/Fakt-Queries referenzieren nur `Src_DWH` — der Rest bleibt. Statt
Excel die DB anzapfen (Tabellennamen sind `dim_date`, `fact_actual`, …):

```m
// PostgreSQL
let Src = PostgreSQL.Database("localhost", "zlatylev") in Src
//   dann z.B.:  DimDate = Src_DWH{[Schema="public", Item="dim_date"]}[Data]

// Fabric Warehouse / Azure SQL
let Src = Sql.Database("<server>.datawarehouse.fabric.microsoft.com", "<wh>") in Src
//   dann z.B.:  DimDate = Src_DWH{[Schema="dbo", Item="dim_date"]}[Data]

// SQLite (ZlatyLev_DWH.sqlite) — via ODBC-Treiber "SQLite3 ODBC Driver"
let Src = Odbc.DataSource("Driver=SQLite3 ODBC Driver;Database=" & pDwhSqlitePath & ";", [HierarchicalNavigation=true]) in Src
//   dann z.B.:  DimDate = Src_DWH{[Name="dim_date"]}[Data]
```

**DB-Artefakte bauen / wiederherstellen:**

```bash
# SQLite + Postgres-Dump aus der DWH-Excel erzeugen
python scripts/gen_db.py

# Postgres-Dump wiederherstellen (self-contained, kein CSV nötig)
createdb zlatylev
psql -d zlatylev -f dwh_postgres_dump.sql

# Alternativ: leere DB per DDL + CSV laden
psql -d zlatylev -f dwh_postgres.sql

# SQLite direkt prüfen
sqlite3 ZlatyLev_DWH.sqlite "SELECT d.year, SUM(f.value) FROM fact_actual f \
  JOIN dim_date d ON f.date=d.date WHERE f.account_key=1 GROUP BY d.year;"
```

`FactPlan` bleibt aus Excel (Plan wird von Finance dort gepflegt).

---

## Beziehungen & Modell-Settings

1. `DimDate[Date]` → `FactActual[Date]`  (1:*), Single.
2. `DimDate[Date]` → `FactPlan[Date]`    (1:*), Single.
3. `DimCustomer[CustomerKey]` → `FactActual[CustomerKey]` (1:*).
4. `DimCustomer[CustomerKey]` → `FactPlan[CustomerKey]`   (1:*).
5. `DimProduct[ProductKey]`  → `FactActual[ProductKey]`  (1:*).
6. `DimProduct[ProductKey]`  → `FactPlan[ProductKey]`    (1:*).
7. `DimKennzahl[AccountKey]` → `FactActual[AccountKey]`  (1:*).
8. `DimDate` **als Datumstabelle markieren** → Spalte `Date`.
9. `MonthName` → *Nach Spalte sortieren* → `MonthNo`.
   `MonthYear` → *Nach Spalte sortieren* → `MonthYearSort`.

---

## DAX-Measures (Tabelle `_Measures`)

### Basiswerte (über die Kennzahlen-Dimension)

```dax
Volume hl   = CALCULATE ( SUM ( FactActual[Value] ), DimKennzahl[AccountName] = "Volume" )
Net Sales   = CALCULATE ( SUM ( FactActual[Value] ), DimKennzahl[AccountName] = "Net Sales" )
Discounts   = CALCULATE ( SUM ( FactActual[Value] ), DimKennzahl[AccountName] = "Discounts" )
COGS        = CALCULATE ( SUM ( FactActual[Value] ), DimKennzahl[AccountName] = "COGS" )

Gross Sales   = [Net Sales] + [Discounts]
Gross Margin  = [Net Sales] - [COGS]
Gross Margin % = DIVIDE ( [Gross Margin], [Net Sales] )
```

### Plan (aus der Kreuztabelle)

```dax
Net Sales PL = SUM ( FactPlan[NetSalesPlan] )
```

### AC / PY / YTD / MTD — sauber für die Zeitachse (4 Jahre)

```dax
-- AC ist [Net Sales] selbst (aktueller Kontext)

Net Sales PY      = CALCULATE ( [Net Sales], SAMEPERIODLASTYEAR ( DimDate[Date] ) )

Net Sales YTD     = TOTALYTD ( [Net Sales], DimDate[Date] )
Net Sales PY YTD  = CALCULATE ( [Net Sales YTD], SAMEPERIODLASTYEAR ( DimDate[Date] ) )
Net Sales PL YTD  = TOTALYTD ( [Net Sales PL], DimDate[Date] )

Net Sales MTD     = TOTALMTD ( [Net Sales], DimDate[Date] )
-- Hinweis: Fakt-Grain ist monatlich → MTD = Monatswert. Bei täglichem Grain
--          (DB-Variante) liefert MTD echte Monats-Kumulation.
```

### Abweichungen (IBCS: ΔPL / ΔPY)

```dax
ΔPL  = [Net Sales] - [Net Sales PL]
ΔPL% = DIVIDE ( [Net Sales] - [Net Sales PL], [Net Sales PL] )
ΔPY  = [Net Sales] - [Net Sales PY]
ΔPY% = DIVIDE ( [Net Sales] - [Net Sales PY], [Net Sales PY] )
```

Für `Volume hl`, `COGS`, `Gross Margin` analog (PY/YTD/MTD mit demselben Muster).

---

## Mapping auf den Business-Chart-Builder

| Builder-Feld | Modell-Feld |
|---|---|
| Kategorie / Achse | `DimDate[MonthYear]` (sort `MonthYearSort`) oder `DimProduct[Category]` |
| Kennzahl Primär (AC) | `[Net Sales]` |
| Referenz (PL) | `[Net Sales PL]` |
| 2. Referenz (PY) | `[Net Sales PY]` |
| ΔPL / ΔPY | `[ΔPL]` · `[ΔPY]` (bzw. `%`) |

Tipp: Für eine 12-Monats-IBCS-Achse Seitenfilter `DimDate[Year]` auf ein Jahr
setzen und `DimDate[MonthName]` (sort `MonthNo`) als Kategorie nutzen.
