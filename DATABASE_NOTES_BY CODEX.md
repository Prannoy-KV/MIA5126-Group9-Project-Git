# Ontario waste database

Build the SQLite file from the local `Datasets` directory:

```sh
python build_database.py
```

Requires Python 3.10+, `pandas`, and `openpyxl` (already listed in the workspace requirements). Default input/output paths are relative to the script, so it also works from a different working directory. Optional paths:

```sh
python build_database.py --data-dir /path/to/Datasets --output /path/to/ontario_waste.sqlite
```

Each run rebuilds the database from scratch in a temporary file, validates it, and replaces the output only after success. It does not change the datasets. Missing years/categories, duplicate program records, and unrecognised layouts stop the build.

## Analysis tables

RPRA tables use `(program_code, year)` as their primary key and cover 2011–2021 inclusive. `program_name` retains each year's reported name. `municipal_group` is NULL when the source workbook does not supply it. A program can represent a municipality, region, or recycling association.

| Table | Contents | Initial records |
| --- | --- | ---: |
| `municipal_programs` | Program codes and latest available names | 274 |
| `residential_waste` | Reported households, population, generated/diverted/disposed tonnes, rates, and component shares | 1,787 |
| `blue_box_tonnage` | Total and material-specific marketed tonnes, household counts | 2,651 |
| `organics` | Organics collected, separated by material and curbside/depot collection | 1,784 |
| `other_recyclables` | Non-Blue-Box materials collected | 1,787 |
| `accepted_materials` | Accepted materials, collection frequency, household counts, and garbage collection policies | 2,651 |
| `blue_box_financials` | Reported costs, revenue, and financial measures | 2,651 |
| `census_demographics` | Population/dwellings for 2016 and 2021, 2021 land area/density/ranks | 577 |

The financial and accepted-material tables retain distinct columns where source labels differ. For example, total costs and costs per tonne remain separate measures. Check `source_columns` before comparing fields whose definitions change. Fields absent in a year remain NULL; missing values are never replaced with zero. Yes/no flags become 1/0; original values remain in the source archive. Collection-frequency fields can contain text such as `Depot`.

The census file contains population/dwelling data, despite its `age_demographics` filename. Its primary key is `dguid`. It remains separate from RPRA: program service areas do not necessarily match census subdivisions, and names are not unique. No automatic geographic join or annual population interpolation is performed.

## Views and units

- `waste_indicators`: residential waste plus calculated diversion/disposal percentages and kg/person using the reported population **including calculated seasonal population**. A zero denominator produces NULL. Source-provided kg/person measures are retained separately in columns ending `_reported`.
- `waste_components`: one program/year/component per row, with its original fraction, denominator basis, and estimated component tonnes. These are estimates derived from reported shares, not additional measured tonnages or waste-composition audits.
- `dataset_coverage`: annual imported program counts by dataset, including source filenames.

`residential_waste` source rates and `_share` fields are **fractions from 0 to 1**. `waste_indicators` fields ending `_pct` are percentages from 0 to 100. Census change fields are also percentages. Mass fields ending `_tonnes` are tonnes; per-capita fields explicitly use kilograms.

The importer infers component-share denominators by comparing component sums against the total rate and 1. This matters because some printed headers are misleading: the 2015 workbook says `% of Generated`, but its component values sum to 1 and represent shares of diverted/disposed waste. Current source values indicate generated-waste shares in 2011–2013, and diverted/disposed-waste shares in 2014–2021. Ambiguous/mixed definitions stop the build.

## Provenance and checks

`source_files` stores relative filenames, years, categories, and SHA-256 hashes. `source_sheets` identifies the worksheet used for analysis. The first worksheet supplies program records; later worksheets repeat orderings, contain group summaries, or are empty. They are archived without duplicating analysis records.

`raw_sheet_rows` preserves nonempty worksheet rows as JSON arrays in original column order, including headers, notes, totals, and repeated sheets. These are Excel **cached values**, not formulas, formatting, or merged-cell metadata. `source_columns` maps analysis columns back to their original headers. Excel row/column numbers are 1-based. Each analysis record carries `source_sheet_id` and `source_excel_row`.

`quality_issues` records residential mass-balance, missing-value, negative-value, and diversion-rate anomalies. The initial build flags Guelph (program 36), 2012: generated = 43,734.99 tonnes, while diverted + disposed = 43,759.99 tonnes. The source values are preserved. SQLite integrity and foreign-key checks must pass before publishing the file.

## Example

```python
import sqlite3
import pandas as pd

with sqlite3.connect("ontario_waste.sqlite") as connection:
    trends = pd.read_sql_query(
        """
        SELECT program_code, program_name, year,
               diversion_rate_pct, disposed_kg_per_capita
        FROM waste_indicators
        WHERE program_code = ?
        ORDER BY year
        """,
        connection,
        params=(20,),  # Toronto
    )
print(trends)
```

Run the dataset integration checks from the project directory:

```sh
python -m unittest discover -s tests -v
```
