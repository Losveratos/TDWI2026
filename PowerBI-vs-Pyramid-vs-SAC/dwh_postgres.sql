-- =====================================================================
-- ZLATÝ LEV BREWERY a.s. — Demo-DWH (Sternschema) für PostgreSQL
-- Lädt die CSVs aus  dwh_csv/  (Semikolon, UTF-8 mit BOM).
-- Plan bleibt bewusst in Excel (ZlatyLev_Plan.xlsx) — wie von Finance gepflegt.
--
-- Ausführen z.B.:  psql -d zlatylev -f dwh_postgres.sql
-- Die \copy-Pfade ggf. an deinen Ablageort anpassen.
-- =====================================================================

DROP TABLE IF EXISTS fact_actual;
DROP TABLE IF EXISTS dim_date;
DROP TABLE IF EXISTS dim_product;
DROP TABLE IF EXISTS dim_customer;
DROP TABLE IF EXISTS dim_kennzahl;

CREATE TABLE dim_date (
    date            date PRIMARY KEY,
    date_key        integer      NOT NULL,
    year            integer      NOT NULL,
    quarter         integer      NOT NULL,
    quarter_name    text         NOT NULL,
    month_no        integer      NOT NULL,
    month_name      text         NOT NULL,
    month_year      text         NOT NULL,
    month_year_sort integer      NOT NULL,
    month_start     date         NOT NULL,
    day_of_month    integer      NOT NULL,
    day_name        text         NOT NULL
);

CREATE TABLE dim_product (
    product_key  integer PRIMARY KEY,
    product_name text    NOT NULL,
    category     text    NOT NULL,
    style        text,
    abv          numeric(3,1),
    launch_year  integer
);

CREATE TABLE dim_customer (
    customer_key        integer PRIMARY KEY,
    customer_name       text    NOT NULL,
    customer_type       text    NOT NULL,
    street              text,
    postal_code         text,
    city                text,
    region              text,
    country             text,
    customer_since      integer,
    key_account_manager text
);

CREATE TABLE dim_kennzahl (
    account_key   integer PRIMARY KEY,
    account_name  text    NOT NULL,
    account_group text,
    unit          text,
    sort_order    integer
);

CREATE TABLE fact_actual (
    date         date    NOT NULL REFERENCES dim_date(date),
    customer_key integer NOT NULL REFERENCES dim_customer(customer_key),
    product_key  integer NOT NULL REFERENCES dim_product(product_key),
    account_key  integer NOT NULL REFERENCES dim_kennzahl(account_key),
    value        numeric(18,2) NOT NULL
);
CREATE INDEX ix_fact_actual_date ON fact_actual(date);
CREATE INDEX ix_fact_actual_prod ON fact_actual(product_key);

-- ---- Daten laden (client-seitig; Header-Zeile überspringen) -----------
\copy dim_date     FROM 'dwh_csv/DimDate.csv'     WITH (FORMAT csv, HEADER true, DELIMITER ';', ENCODING 'UTF8');
\copy dim_product  FROM 'dwh_csv/DimProduct.csv'  WITH (FORMAT csv, HEADER true, DELIMITER ';', ENCODING 'UTF8');
\copy dim_customer FROM 'dwh_csv/DimCustomer.csv' WITH (FORMAT csv, HEADER true, DELIMITER ';', ENCODING 'UTF8');
\copy dim_kennzahl FROM 'dwh_csv/DimKennzahl.csv' WITH (FORMAT csv, HEADER true, DELIMITER ';', ENCODING 'UTF8');
\copy fact_actual  FROM 'dwh_csv/FactActual.csv'  WITH (FORMAT csv, HEADER true, DELIMITER ';', ENCODING 'UTF8');

-- Kontrolle: sollte 90740 / 98360 / 103120 / 41225 hl liefern
-- SELECT year, SUM(value) FROM fact_actual f JOIN dim_date d USING(date)
--   WHERE account_key = 1 GROUP BY year ORDER BY year;
