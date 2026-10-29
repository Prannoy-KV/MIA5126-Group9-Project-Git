"""Build ontario_waste.sqlite from the 2011–2021 files in Datasets.

Run ``python build_database.py``. Requires pandas and openpyxl.
Default paths are relative to this script, regardless of the working directory.
A failed build leaves the previous database intact.
"""

from __future__ import annotations

import argparse
from collections import defaultdict
from contextlib import closing
from datetime import datetime, timezone
import hashlib
import json
import logging
import math
from pathlib import Path
import re
import sqlite3
import tempfile
import warnings

import pandas as pd
from openpyxl import load_workbook

#==============================
# Part 1: Find the source files: It looks for the 2011–2021 RPRA Excel files and the census CSV in the expected folders. build_database 
# Details
#==============================
PROJECT_DIR = Path(__file__).resolve().parent
YEARS = range(2011, 2022)
CATEGORIES = (
    "accepted_materials", "blue_box_financials", "blue_box_tonnage",
    "organics", "other_recyclables", "residential_waste",
)
CODE_HEADERS = {"pc", "program_code"}
NAME_HEADERS = {
    "program_title", "program_name", "municipality", "municipal_program",
    "municial_title", "municipal_title",
}
GROUP_HEADERS = {"group", "municipal_group", "muni_group"}
DIVERSION_COMPONENTS = {
    "residential_deposit_return_program": "deposit_return_share",
    "residential_reuse": "reuse_share",
    "residential_on_property": "on_property_share",
    "residential_recyclables_diverted": "recyclables_share",
    "residential_organics_diverted": "organics_share",
    "residential_mhsw_treatment_reuse_recycling": "mhsw_share",
}
DISPOSAL_COMPONENTS = {
    "residential_efw": "energy_from_waste_share",
    "residential_hazardous_waste_disposal": "hazardous_disposal_share",
    "residential_landfill": "landfill_share",
}
CENSUS_FIELDS = {
    1: "population_2021", 2: "population_2016", 3: "population_change_pct",
    4: "private_dwellings_2021", 5: "private_dwellings_2016",
    6: "private_dwellings_change_pct", 7: "occupied_dwellings_2021",
    8: "occupied_dwellings_2016", 9: "occupied_dwellings_change_pct",
    10: "land_area_km2_2021", 11: "population_density_2021",
    12: "national_population_rank_2021", 13: "ontario_population_rank_2021",
}


def clean_text(value: object) -> str:
    return re.sub(r"\s+", " ", str(value)).strip() if value is not None else ""


def slug(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "_", clean_text(value).lower()).strip("_")


def header_key(value: object) -> str:
    # Remove footnote digits attached to words; preserve plastic-type numbers.
    text = re.sub(r"(?<=[A-Za-z])\d+\b", "", clean_text(value))
    return slug(text).replace("p_c", "pc")


def number(value: object) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, str):
        value = value.strip().replace(",", "")
    try:
        result = float(value)
    except (TypeError, ValueError):
        return None
    return result if math.isfinite(result) else None


def program_code(value: object) -> int | None:
    result = number(value)
    return int(result) if result is not None and result > 0 and result.is_integer() else None


def cell_value(value: object) -> object:
    if isinstance(value, str):
        value = clean_text(value)
        if value.lower() in {"", "n/a", "na", "..", "...", "-", "suppressed"}:
            return None
        if value.lower() in {"yes", "y"}:
            return 1
        if value.lower() in {"no", "n"}:
            return 0
        parsed = number(value)
        return parsed if parsed is not None else value
    return value


def category_for(path: Path) -> str:
    name = slug(path.stem)
    for term, category in [
        ("accepted_materials", "accepted_materials"), ("diversion", "residential_waste"),
        ("organics", "organics"), ("recyclables", "other_recyclables"),
        ("financial", "blue_box_financials"), ("cost", "blue_box_financials"),
        ("tonnage", "blue_box_tonnage"), ("tonnes", "blue_box_tonnage"),
    ]:
        if term in name:
            return category
    raise ValueError(f"Unrecognised RPRA workbook: {path.name}")


def find_header(rows: list[tuple]) -> tuple[int, int, int, int | None]:
    for row_index, row in enumerate(rows[:15]):
        keys = [header_key(value) for value in row]
        codes = [i for i, key in enumerate(keys) if key in CODE_HEADERS]
        names = [i for i, key in enumerate(keys) if key in NAME_HEADERS]
        groups = [i for i, key in enumerate(keys) if key in GROUP_HEADERS]
        if len(codes) == 1 and len(names) == 1:
            return row_index, codes[0], names[0], groups[0] if groups else None
    raise ValueError("No program-code/name header found in the first 15 rows")


def waste_columns(rows: list[tuple], header: int) -> dict[int, str]:
    """Find measures by labels: 2011 and 2014 use different column positions."""
    result = {}
    for col, value in enumerate(rows[header]):
        key = header_key(value)
        if "single_family_households" in key:
            result[col] = "single_family_households"
        elif key == "reported_multi_family_households":
            result[col] = "multi_family_households"
        elif key == "reported_seasonal_households":
            result[col] = "seasonal_households"
        elif key == "reported_population":
            result[col] = "population_reported"
        elif "population_calculated_seasonal_population" in key:
            result[col] = "population_with_seasonal"
        for activity in ("generated", "diverted", "disposed"):
            if key == f"total_residential_waste_{activity}":
                result[col] = f"waste_{activity}_tonnes"
                result[col + 1] = f"{activity}_kg_per_capita_reported"
    for row in rows[header:header + 2]:
        for col, value in enumerate(row):
            key = header_key(value)
            if key in DIVERSION_COMPONENTS:
                result[col] = DIVERSION_COMPONENTS[key]
            elif key in DISPOSAL_COMPONENTS:
                result[col] = DISPOSAL_COMPONENTS[key]
            elif key in {"total_residential_diversion_rate", "total_residential_waste_diversion_rate"}:
                result[col] = "diversion_rate_reported"
            elif key == "total_residential_disposal_rate":
                result[col] = "disposal_rate_reported"
    required = {
        "population_reported", "population_with_seasonal", "waste_generated_tonnes",
        "waste_diverted_tonnes", "waste_disposed_tonnes", "diversion_rate_reported",
        "disposal_rate_reported", *DIVERSION_COMPONENTS.values(), *DISPOSAL_COMPONENTS.values(),
    }
    if required - set(result.values()):
        raise ValueError(f"Missing waste headers: {sorted(required - set(result.values()))}")
    return result


def detail_columns(category: str, rows: list[tuple], header: int) -> dict[int, str]:
    result = {
        col: header_key(value) for col, value in enumerate(rows[header])
        if header_key(value) and header_key(value) not in CODE_HEADERS | NAME_HEADERS | GROUP_HEADERS
    }
    if category == "blue_box_tonnage":
        materials = (
            "printed_paper", "occ_obb", "mixed_paper", "polycoat", "pet", "hdpe",
            "plastic_film", "tubs_and_lids", "polystyrene", "mixed_plastic",
            "steel", "aluminum", "flint_glass", "coloured_glass",
        )
        paper_col = next(c for c, v in enumerate(rows[header]) if header_key(v) == "paper_tonnes")
        for offset, material in enumerate(materials):
            col = paper_col + offset
            label = header_key(rows[header + 1][col])
            expected = material.removesuffix("_glass")
            if not label.startswith(expected):
                raise ValueError(f"Unexpected Blue Box material order at column {col + 1}: {label}")
            result[col] = f"{material}_marketed_tonnes"
        for col, key in list(result.items()):
            if key.startswith("total_reported_and_or_calculated_marketed_tonnes"):
                result[col] = "total_marketed_tonnes"
            elif key in {"total_households_served", "total_households_serviced"}:
                result[col] = "households_serviced"
            elif key == "tonnes_household":
                result[col] = "marketed_tonnes_per_household"
    elif category == "organics":
        for col, key in list(result.items()):
            if key.startswith("total_"):
                result[col] = "total_organics_collected_tonnes"
            elif key.startswith(("curbside_", "depot_")):
                method = key.split("_")[0]
                for offset in range(5):
                    material = header_key(rows[header + 1][col + offset])
                    if material not in {"yard_waste", "leaves", "christmas_trees", "bulky_yard_waste", "household_organics"}:
                        raise ValueError(f"Unexpected organics subheader: {material}")
                    result[col + offset] = f"{method}_{material}_tonnes"
    elif category == "other_recyclables":
        for col, key in list(result.items()):
            if key.startswith(("total_other_recyclables", "total_non_blue_box_recyclables")):
                result[col] = "total_other_recyclables_collected_tonnes"
            elif key in {"other_please_describe_tonnes", "other_materials"}:
                result[col] = "other_materials_tonnes"
            elif key == "total_households_served":
                result[col] = "households_serviced"
    # Different financial and accepted-material definitions retain distinct labels.
    if len(result.values()) != len(set(result.values())):
        raise ValueError(f"Duplicate metric names in {category}")
    return result


def infer_basis(records: list[dict], components: dict[str, str], rate: str) -> str:
    """Distinguish shares of generated waste from shares of diverted/disposed waste."""
    alternate = "diverted" if rate.startswith("diversion") else "disposed"
    votes = {"generated": 0, alternate: 0}
    for record in records:
        values = [number(record.get(key)) for key in components.values()]
        reported_rate = number(record.get(rate))
        if any(v is None for v in values) or reported_rate is None:
            continue
        if not 0.01 < reported_rate < 0.99:
            continue  # Zero and 100% do not distinguish the denominators.
        total = sum(values)
        if math.isclose(total, reported_rate, abs_tol=0.002):
            votes["generated"] += 1
        if math.isclose(total, 1.0, abs_tol=0.002):
            votes[alternate] += 1
    basis = max(votes, key=votes.get)
    if votes[basis] == 0 or min(votes.values()) > 0:
        raise ValueError(f"Ambiguous component percentage basis: {votes}")
    return basis


SCHEMA = """
PRAGMA foreign_keys = ON;
CREATE TABLE build_metadata (key TEXT PRIMARY KEY, value TEXT NOT NULL);
CREATE TABLE source_files (
    source_file_id INTEGER PRIMARY KEY, relative_path TEXT NOT NULL UNIQUE,
    dataset_type TEXT NOT NULL, year INTEGER, sha256 TEXT NOT NULL
);
CREATE TABLE source_sheets (
    source_sheet_id INTEGER PRIMARY KEY,
    source_file_id INTEGER NOT NULL REFERENCES source_files,
    sheet_name TEXT NOT NULL, selected_for_analysis INTEGER NOT NULL,
    header_excel_row INTEGER, imported_records INTEGER NOT NULL DEFAULT 0,
    UNIQUE(source_file_id, sheet_name)
);
CREATE TABLE raw_sheet_rows (
    source_sheet_id INTEGER NOT NULL REFERENCES source_sheets,
    excel_row INTEGER NOT NULL, values_json TEXT NOT NULL,
    PRIMARY KEY(source_sheet_id, excel_row)
);
CREATE TABLE source_columns (
    source_sheet_id INTEGER NOT NULL REFERENCES source_sheets,
    excel_column INTEGER NOT NULL, database_column TEXT NOT NULL,
    original_header_json TEXT NOT NULL,
    PRIMARY KEY(source_sheet_id, excel_column)
);
CREATE TABLE municipal_programs (program_code INTEGER PRIMARY KEY, latest_program_name TEXT NOT NULL);
CREATE TABLE quality_issues (
    dataset_type TEXT NOT NULL, year INTEGER, program_code INTEGER,
    issue TEXT NOT NULL, details TEXT NOT NULL
);
"""


def add_source(conn: sqlite3.Connection, path: Path, root: Path, category: str, year: int | None) -> int:
    cursor = conn.execute(
        "INSERT INTO source_files(relative_path, dataset_type, year, sha256) VALUES (?, ?, ?, ?)",
        (path.relative_to(root).as_posix(), category, year, hashlib.sha256(path.read_bytes()).hexdigest()),
    )
    return cursor.lastrowid


def import_workbook(conn: sqlite3.Connection, path: Path, root: Path, year: int, category: str) -> list[dict]:
    source_id = add_source(conn, path, root, category, year)
    # First sheets contain the program records. Later sheets are repeated orderings
    # or summaries. Preserve every worksheet but import program records only once.
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Unknown extension is not supported.*", category=UserWarning)
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            records = []
            for sheet_index, sheet in enumerate(workbook.worksheets):
                rows = list(sheet.values)
                selected = sheet_index == 0
                cursor = conn.execute(
                    "INSERT INTO source_sheets(source_file_id, sheet_name, selected_for_analysis) VALUES (?, ?, ?)",
                    (source_id, sheet.title, int(selected)),
                )
                sheet_id = cursor.lastrowid
                raw_rows = []
                for row_number, row in enumerate(rows, 1):
                    values = list(row)
                    while values and values[-1] is None:
                        values.pop()
                    if values:
                        raw_rows.append((sheet_id, row_number, json.dumps(values, ensure_ascii=False, default=str)))
                conn.executemany("INSERT INTO raw_sheet_rows VALUES (?, ?, ?)", raw_rows)
                if not selected:
                    continue
                header, code_col, name_col, group_col = find_header(rows)
                columns = waste_columns(rows, header) if category == "residential_waste" else detail_columns(category, rows, header)
                depth = 3 if category == "residential_waste" else 2 if category in {"organics", "blue_box_tonnage"} else 1
                conn.executemany(
                    "INSERT INTO source_columns VALUES (?, ?, ?, ?)",
                    [(sheet_id, col + 1, key, json.dumps([row[col] if col < len(row) else None for row in rows[header:header + depth]], default=str))
                     for col, key in columns.items()],
                )
                seen = set()
                for row_number, row in enumerate(rows[header + depth:], header + depth + 1):
                    code = program_code(row[code_col])
                    name = clean_text(row[name_col])
                    if code is None:
                        continue  # Exclude titles, notes, group labels and subtotals.
                    if not name or name.lower().startswith(("total", "average")):
                        raise ValueError(f"Program code without valid name at Excel row {row_number}")
                    if code in seen:
                        raise ValueError(f"Duplicate program {code} in {year} {category}")
                    seen.add(code)
                    record = {
                        "program_code": code, "year": year, "program_name": name,
                        "municipal_group": program_code(row[group_col]) if group_col is not None else None,
                        "source_sheet_id": sheet_id, "source_excel_row": row_number,
                    }
                    record.update({key: cell_value(row[col]) for col, key in columns.items()})
                    records.append(record)
                if not records:
                    raise ValueError("No program records imported")
                if category == "residential_waste":
                    diversion_basis = infer_basis(records, DIVERSION_COMPONENTS, "diversion_rate_reported")
                    disposal_basis = infer_basis(records, DISPOSAL_COMPONENTS, "disposal_rate_reported")
                    for record in records:
                        record["diversion_component_basis"] = diversion_basis
                        record["disposal_component_basis"] = disposal_basis
                conn.execute(
                    "UPDATE source_sheets SET header_excel_row = ?, imported_records = ? WHERE source_sheet_id = ?",
                    (header + 1, len(records), sheet_id),
                )
            return records
        finally:
            workbook.close()


def write_program_table(conn: sqlite3.Connection, category: str, records: list[dict]) -> None:
    common = ["program_code", "year", "program_name", "municipal_group", "source_sheet_id", "source_excel_row"]
    columns = common + sorted(set().union(*(r.keys() for r in records)) - set(common))
    text_columns = {"program_name", "diversion_component_basis", "disposal_component_basis"}
    definitions = [f'"{col}" ' + ("TEXT" if col in text_columns else "INTEGER" if col in common else "NUMERIC") for col in columns]
    definitions.extend([
        "PRIMARY KEY (program_code, year)",
        "FOREIGN KEY (program_code) REFERENCES municipal_programs(program_code)",
        "FOREIGN KEY (source_sheet_id) REFERENCES source_sheets(source_sheet_id)",
        "CHECK (year BETWEEN 2011 AND 2021)",
    ])
    conn.execute(f'CREATE TABLE "{category}" ({", ".join(definitions)})')
    conn.executemany(
        f'INSERT INTO "{category}" VALUES ({", ".join("?" for _ in columns)})',
        [tuple(r.get(c) for c in columns) for r in records],
    )
    conn.execute(f'CREATE INDEX "idx_{category}_year" ON "{category}"(year)')


def import_census(conn: sqlite3.Connection, root: Path) -> None:
    path = root / "ontario_municipal_age_demographics.csv"
    if not path.is_file():
        raise FileNotFoundError(f"Missing census CSV: {path}")
    source_id = add_source(conn, path, root, "census_demographics", None)
    frame = pd.read_csv(path, dtype={"DGUID": str})
    names = {"DGUID": "dguid", "Municipality": "municipality", "GEO": "geographic_name"}
    for column in frame.columns:
        match = re.search(r"\[(\d+)\]$", column)
        if match and int(match[1]) in CENSUS_FIELDS:
            names[column] = CENSUS_FIELDS[int(match[1])]
    if set(names.values()) != {"dguid", "municipality", "geographic_name", *CENSUS_FIELDS.values()}:
        raise ValueError("Census CSV is missing expected population/dwelling columns")
    frame = frame[list(names)].rename(columns=names)
    conn.execute(
        "CREATE TABLE census_demographics (dguid TEXT PRIMARY KEY NOT NULL, municipality TEXT NOT NULL, "
        "geographic_name TEXT NOT NULL, " + ", ".join(f'{key} REAL' for key in CENSUS_FIELDS.values()) +
        ", source_file_id INTEGER NOT NULL REFERENCES source_files)"
    )
    columns = list(frame.columns)
    for key in CENSUS_FIELDS.values():
        frame[key] = pd.to_numeric(frame[key], errors="raise")
    rows = [tuple(None if pd.isna(v) else v for v in row) + (source_id,) for row in frame.itertuples(index=False, name=None)]
    conn.executemany(
        f'INSERT INTO census_demographics ({", ".join(columns)}, source_file_id) VALUES ({", ".join("?" for _ in range(len(columns) + 1))})', rows,
    )


def create_views(conn: sqlite3.Connection) -> None:
    conn.executescript("""
    CREATE VIEW waste_indicators AS
    SELECT w.*,
        100.0 * waste_diverted_tonnes / NULLIF(waste_generated_tonnes, 0) AS diversion_rate_pct,
        100.0 * waste_disposed_tonnes / NULLIF(waste_generated_tonnes, 0) AS disposal_rate_pct,
        1000.0 * waste_generated_tonnes / NULLIF(population_with_seasonal, 0) AS generated_kg_per_capita,
        1000.0 * waste_diverted_tonnes / NULLIF(population_with_seasonal, 0) AS diverted_kg_per_capita,
        1000.0 * waste_disposed_tonnes / NULLIF(population_with_seasonal, 0) AS disposed_kg_per_capita
    FROM residential_waste AS w;
    CREATE VIEW dataset_coverage AS
    SELECT f.year, f.dataset_type, s.sheet_name, s.imported_records, f.relative_path
    FROM source_files AS f JOIN source_sheets AS s USING(source_file_id)
    WHERE s.selected_for_analysis = 1;
    """)
    selects = []
    for components, basis in [(DIVERSION_COMPONENTS, "diversion_component_basis"), (DISPOSAL_COMPONENTS, "disposal_component_basis")]:
        for share in components.values():
            selects.append(f"""SELECT program_code, year, '{share.removesuffix('_share')}' AS component,
                {share} AS source_fraction, {basis} AS source_fraction_basis,
                {share} * CASE {basis}
                    WHEN 'generated' THEN waste_generated_tonnes
                    WHEN 'diverted' THEN waste_diverted_tonnes
                    WHEN 'disposed' THEN waste_disposed_tonnes END AS estimated_tonnes
                FROM residential_waste""")
    conn.execute("CREATE VIEW waste_components AS " + " UNION ALL ".join(selects))


def validate_database(conn: sqlite3.Connection) -> None:
    if conn.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
        raise ValueError("SQLite integrity check failed")
    if conn.execute("PRAGMA foreign_key_check").fetchall():
        raise ValueError("SQLite foreign-key check failed")
    # Keep source anomalies visible; do not invent replacements or drop programs.
    for code, year, generated, diverted, disposed, rate in conn.execute(
        "SELECT program_code, year, waste_generated_tonnes, waste_diverted_tonnes, "
        "waste_disposed_tonnes, diversion_rate_reported FROM residential_waste"
    ):
        measures = [generated, diverted, disposed, rate]
        if any(v is not None and not isinstance(v, (int, float)) for v in measures):
            raise ValueError(f"Non-numeric waste measure for program {code}, {year}")
        issues = []
        if any(v is None for v in measures):
            issues.append("missing_waste_measure")
        if any(v is not None and v < 0 for v in measures):
            issues.append("negative_waste_measure")
        if all(v is not None for v in [generated, diverted, disposed]) and not math.isclose(generated, diverted + disposed, abs_tol=0.1, rel_tol=1e-5):
            issues.append("mass_balance_mismatch")
        if rate is not None and not 0 <= rate <= 1:
            issues.append("diversion_rate_out_of_range")
        if generated and diverted is not None and rate is not None and not math.isclose(diverted / generated, rate, abs_tol=0.001):
            issues.append("diversion_rate_mismatch")
        for issue in issues:
            conn.execute("INSERT INTO quality_issues VALUES (?, ?, ?, ?, ?)", ("residential_waste", year, code, issue, json.dumps(measures)))


def build_database(data_dir: Path = PROJECT_DIR / "Datasets", output: Path = PROJECT_DIR / "ontario_waste.sqlite") -> Path:
    data_dir, output = Path(data_dir).resolve(), Path(output).resolve()
    rpra_dir = data_dir / "RPRA Datasets"
    if not rpra_dir.is_dir():
        raise FileNotFoundError(f"Missing RPRA dataset directory: {rpra_dir}")
    files = {}
    for path in sorted(rpra_dir.rglob("*.xlsx")):
        if path.name.startswith("~$"):
            continue
        match = re.match(r"(\d{4})", path.name)
        if match is None or int(match[1]) not in YEARS:
            raise ValueError(f"Expected a 2011–2021 workbook: {path}")
        key = (int(match[1]), category_for(path))
        if key in files:
            raise ValueError(f"Multiple workbooks for {key}: {files[key]} and {path}")
        files[key] = path
    missing = {(year, category) for year in YEARS for category in CATEGORIES} - files.keys()
    if missing:
        raise ValueError(f"Missing annual workbooks: {sorted(missing)}")
    if output in files.values() or output == data_dir / "ontario_municipal_age_demographics.csv":
        raise ValueError("Output path must not overwrite an input dataset")
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(prefix=f".{output.stem}_", suffix=".sqlite", dir=output.parent, delete=False) as temp:
        temp_path = Path(temp.name)
    try:
        with closing(sqlite3.connect(temp_path)) as conn:
            conn.executescript(SCHEMA)
            conn.executemany("INSERT INTO build_metadata VALUES (?, ?)", [
                ("built_at_utc", datetime.now(timezone.utc).isoformat()),
                ("schema_version", "1"), ("start_year", "2011"), ("end_year", "2021"),
                ("rates", "Residential source rates/shares are fractions (0–1); census change fields are percentages."),
                ("raw_rows", "JSON arrays contain cached Excel values in original column order; row/column numbers are 1-based."),
                ("worksheet_selection", "First worksheet supplies program records; all worksheets archived. Repeated views/summaries are not counted twice."),
                ("geography", "RPRA programs and census geographies remain separate; no name join or population interpolation."),
            ])
            datasets = defaultdict(list)
            latest_names = {}
            for (year, category), path in sorted(files.items()):
                try:
                    records = import_workbook(conn, path, data_dir, year, category)
                except (ValueError, IndexError, KeyError, StopIteration) as exc:
                    raise ValueError(f"{path.name}: {exc}") from exc
                datasets[category].extend(records)
                for record in records:
                    latest_names[record["program_code"]] = record["program_name"]
                logging.info("%s %s: %s program records", year, category, len(records))
            conn.executemany("INSERT INTO municipal_programs VALUES (?, ?)", sorted(latest_names.items()))
            for category in CATEGORIES:
                write_program_table(conn, category, datasets[category])
            import_census(conn, data_dir)
            create_views(conn)
            validate_database(conn)
            conn.commit()
            issue_count = conn.execute("SELECT COUNT(*) FROM quality_issues").fetchone()[0]
            logging.info("Imported %s RPRA workbooks; %s programs; %s quality flags", len(files), len(latest_names), issue_count)
        temp_path.replace(output)
    finally:
        temp_path.unlink(missing_ok=True)
    logging.info("Created %s", output)
    return output


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--data-dir", type=Path, default=PROJECT_DIR / "Datasets")
    parser.add_argument("--output", type=Path, default=PROJECT_DIR / "ontario_waste.sqlite")
    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    try:
        build_database(args.data_dir, args.output)
    except (ValueError, OSError, sqlite3.Error) as exc:
        parser.exit(1, f"Build failed: {exc}\n")


if __name__ == "__main__":
    main()
