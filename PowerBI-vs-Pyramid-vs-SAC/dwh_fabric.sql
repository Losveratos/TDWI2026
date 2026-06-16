-- =====================================================================
-- ZLATÝ LEV BREWERY a.s. — Demo-DWH (Sternschema) für
-- Microsoft Fabric Warehouse / Azure SQL / SQL Server (T-SQL)
-- Plan bleibt in Excel (ZlatyLev_Plan.xlsx).
-- =====================================================================

DROP TABLE IF EXISTS fact_actual;
DROP TABLE IF EXISTS dim_date;
DROP TABLE IF EXISTS dim_product;
DROP TABLE IF EXISTS dim_customer;
DROP TABLE IF EXISTS dim_kennzahl;

CREATE TABLE dim_date (
    [date]            date         NOT NULL,
    date_key          int          NOT NULL,
    [year]            int          NOT NULL,
    quarter           int          NOT NULL,
    quarter_name      varchar(4)   NOT NULL,
    month_no          int          NOT NULL,
    month_name        varchar(8)   NOT NULL,
    month_year        varchar(16)  NOT NULL,
    month_year_sort   int          NOT NULL,
    month_start       date         NOT NULL,
    day_of_month      int          NOT NULL,
    day_name          varchar(4)   NOT NULL
);

CREATE TABLE dim_product (
    product_key  int          NOT NULL,
    product_name varchar(80)  NOT NULL,
    category     varchar(40)  NOT NULL,
    style        varchar(60),
    abv          decimal(3,1),
    launch_year  int
);

CREATE TABLE dim_customer (
    customer_key        int           NOT NULL,
    customer_name       varchar(120)  NOT NULL,
    customer_type       varchar(40)   NOT NULL,
    street              varchar(120),
    postal_code         varchar(20),
    city                varchar(60),
    region              varchar(40),
    country             varchar(40),
    customer_since      int,
    key_account_manager varchar(60)
);

CREATE TABLE dim_kennzahl (
    account_key   int          NOT NULL,
    account_name  varchar(40)  NOT NULL,
    account_group varchar(40),
    unit          varchar(10),
    sort_order    int
);

CREATE TABLE fact_actual (
    [date]       date          NOT NULL,
    customer_key int           NOT NULL,
    product_key  int           NOT NULL,
    account_key  int           NOT NULL,
    [value]      decimal(18,2) NOT NULL
);

-- ---- Daten laden -----------------------------------------------------
-- Variante A (SQL Server / Azure SQL, lokale Datei):
--   BULK INSERT dim_date FROM 'C:\...\dwh_csv\DimDate.csv'
--     WITH (FORMAT='CSV', FIRSTROW=2, FIELDTERMINATOR=';', ROWTERMINATOR='0x0a', CODEPAGE='65001');
--   (für jede Tabelle wiederholen)
--
-- Variante B (Fabric Warehouse / Synapse, CSV liegt im OneLake/Blob):
--   COPY INTO dim_date
--   FROM 'https://<storage>/dwh_csv/DimDate.csv'
--   WITH (FILE_TYPE='CSV', FIRSTROW=2, FIELDTERMINATOR=';', ENCODING='UTF8');
--   (für jede Tabelle wiederholen)

-- Schlüssel/Constraints nach dem Load (Fabric: NOT ENFORCED ist üblich):
ALTER TABLE dim_date     ADD CONSTRAINT pk_dim_date     PRIMARY KEY NONCLUSTERED ([date])      NOT ENFORCED;
ALTER TABLE dim_product  ADD CONSTRAINT pk_dim_product  PRIMARY KEY NONCLUSTERED (product_key) NOT ENFORCED;
ALTER TABLE dim_customer ADD CONSTRAINT pk_dim_customer PRIMARY KEY NONCLUSTERED (customer_key)NOT ENFORCED;
ALTER TABLE dim_kennzahl ADD CONSTRAINT pk_dim_kennzahl PRIMARY KEY NONCLUSTERED (account_key) NOT ENFORCED;

-- Kontrolle: 90740 / 98360 / 103120 / 41225
-- SELECT d.[year], SUM(f.[value]) FROM fact_actual f JOIN dim_date d ON f.[date]=d.[date]
--   WHERE f.account_key = 1 GROUP BY d.[year] ORDER BY d.[year];
