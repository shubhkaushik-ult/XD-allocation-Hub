# System Architecture & Technical Specifications

## Architectural Overview

The **XD Allocation Hub** is structured as a lightweight Flask web application serving both an intuitive single-page web dashboard and serverless API endpoints.

```
+-----------------------------------------------------------------------+
|                              User Web UI                              |
|                 (HTML5 / Vanilla CSS / Inter Font)                    |
+-----------------------------------------------------------------------+
                                    |
                                    v
+-----------------------------------------------------------------------+
|                          Vercel Serverless                            |
|                            (api/index.py)                             |
+-----------------------------------------------------------------------+
                                    |
                                    v
+-----------------------------------------------------------------------+
|                       xd_processor.py Engine                          |
|  - File Parser (Excel / CSV)                                          |
|  - Raw Indent Coalescing                                              |
|  - Add-on & Staples Processors (is_staples flag)                      |
|  - DB Metadata Enrichment (fetch_addon_metadata_from_db)             |
|  - Contract & Supplier ID Assignment (OU83946700 for BLR Staples)    |
|  - OpenPyXL Multi-Sheet Excel Builder & Zip Streamer                  |
+-----------------------------------------------------------------------+
```

## Data Input Pipeline & Pipelines

The application processes three types of inputs per city:

1. **Raw Indent File (`raw_indent_file`)**: Primary allocation file containing raw store indents, FSNs, vertical information, and quantities.
2. **Extra Add-ons / Pluckk File (`addon_file`)**: Add-on allocations processed via `process_addon_df(df, city_key, is_staples=False)`.
3. **Staples File (`staples_file`)**: Staples items processed via `process_addon_df(df, city_key, is_staples=True)`.

## City Supplier & Contract Mappings

When processing Add-on and Staples data, Supplier IDs and Contract IDs are assigned based on city, category tag (ambient vs chilled), and upload type:

| City | Upload Type / Category | Contract ID | Supplier ID |
| :--- | :--- | :--- | :--- |
| **Bengaluru** | Staples Upload | `SHR-OR-01021287` | **`OU83946700`** |
| **Bengaluru** | Add-on (Ambient) | `SHR-OR-01021287` | `OU83946715` |
| **Bengaluru** | Add-on (Chilled) | `SHR-OR-01021287` | `OU77187305` |
| **Mumbai** | Ambient | `SHR-OR-01062967` | `OU49532875` |
| **Mumbai** | Chilled | `SHR-OR-01062967` | `OU42311586` |
| **Chennai** | All Categories | `SHR-OR-01021287` | `OU56307764` |
| **Trichy / Coimbatore** | Default | `SHR-OR-01021287` | `OU83946715` / `OU77187305` |

## Excel Report Structure

The generated `.xlsx` workbook contains structured tabs formatted via OpenPyXL:
1. **Full PO Summary**: Itemized purchase order table with Store, Store Site ID, FSN, QTY, Supplier ID, Contract ID, Title, Brand, and Vertical.
2. **Category / Supplier Grouping**: Aggregated PO totals grouped by Supplier ID and Vertical.
3. **Audit & Log Data**: Source traceability records.
