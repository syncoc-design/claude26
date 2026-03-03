#!/usr/bin/env python3
"""
OSU Career Roadmap Employee Data Mapper

Downloads and parses OSU HR PDFs (job catalog + salary structure),
loads an employee Excel file, fuzzy-matches job profiles to the catalog,
calculates compa-ratios, and outputs an enriched Excel workbook.

Usage:
    python parse_osu_employee_data.py "path/to/OSU Staff Salary Play.xlsx"

    # Or with pre-downloaded PDFs:
    python parse_osu_employee_data.py "path/to/salary.xlsx" \
        --job-catalog path/to/osu-job-catalog-and-job-code-table.pdf \
        --salary-structure path/to/salary-structure-pay-ranges-slides.pdf
"""

import argparse
import os
import re
import sys
import tempfile

import pandas as pd
import requests
from rapidfuzz import fuzz, process

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
HOURS_PER_YEAR = 2080

JOB_CATALOG_URL = (
    "https://hr.osu.edu/wp-content/uploads/osu-job-catalog-and-job-code-table.pdf"
)
SALARY_STRUCTURE_URL = (
    "https://hr.osu.edu/wp-content/uploads/salary-structure-pay-ranges-slides.pdf"
)

# Career band letters and their full names
CAREER_BANDS = {
    "T": "Technical",
    "S": "Specialized",
    "M": "Managerial",
    "C": "Clinical",
    "E": "Executive",
}

# Regex to pull band letter + level digit from the tail of an 8-char job code
# The job code is 8 characters: Function(2) + Subfunction(4) + Band(1) + Level(1)
# e.g., "ADOFCAS1" -> Band=S, Level=1
BAND_LEVEL_RE = re.compile(r"([TSMCE])(\d)$", re.IGNORECASE)

DEFAULT_FUZZY_THRESHOLD = 85

# ---------------------------------------------------------------------------
# PDF Download
# ---------------------------------------------------------------------------

def download_pdf(url: str, dest: str) -> str:
    """Download a PDF from *url* and save it to *dest*. Returns the path."""
    print(f"  Downloading {url} ...")
    resp = requests.get(url, timeout=60)
    resp.raise_for_status()
    with open(dest, "wb") as f:
        f.write(resp.content)
    print(f"  Saved to {dest} ({len(resp.content):,} bytes)")
    return dest


# ---------------------------------------------------------------------------
# Job Catalog Parser
# ---------------------------------------------------------------------------

def _clean_text(text: str) -> str:
    """Normalize whitespace in extracted PDF text."""
    return " ".join(text.split()).strip()


def parse_job_catalog(pdf_path: str) -> pd.DataFrame:
    """
    Parse the OSU Job Catalog PDF and extract job codes + titles.

    The PDF contains tables with columns like:
        Job Code | Job Profile (Title) | Function | Subfunction | ...

    We look for 8-character alphanumeric codes and the adjacent title text.

    Returns a DataFrame with columns:
        job_code, job_title, career_band, career_level
    """
    import pdfplumber  # lazy import — heavy dependency

    print(f"\n--- Parsing Job Catalog: {pdf_path} ---")
    records = []
    job_code_pattern = re.compile(r"^[A-Z]{2}[A-Z0-9]{4}[TSMCE]\d$", re.IGNORECASE)

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            tables = page.extract_tables()
            for table in tables:
                if not table:
                    continue
                for row in table:
                    if not row:
                        continue
                    # Clean all cells
                    cells = [_clean_text(str(c)) if c else "" for c in row]

                    # Find the cell that looks like an 8-char job code
                    for i, cell in enumerate(cells):
                        # Remove any spaces that snuck in
                        candidate = cell.replace(" ", "").upper()
                        if len(candidate) == 8 and job_code_pattern.match(candidate):
                            # The title is usually in the next column
                            title = ""
                            for j in range(len(cells)):
                                if j != i and cells[j] and len(cells[j]) > 3:
                                    # Pick the longest non-code cell as the title
                                    if len(cells[j]) > len(title):
                                        title = cells[j]
                            if title:
                                band_match = BAND_LEVEL_RE.search(candidate)
                                band = band_match.group(1).upper() if band_match else ""
                                level = band_match.group(2) if band_match else ""
                                records.append({
                                    "job_code": candidate,
                                    "job_title": title,
                                    "career_band": band,
                                    "career_level": int(level) if level else None,
                                })
                            break

            # Also try text-based extraction as fallback for non-tabular pages
            text = page.extract_text() or ""
            for line in text.split("\n"):
                tokens = line.split()
                for token in tokens:
                    candidate = token.replace(" ", "").upper()
                    if len(candidate) == 8 and job_code_pattern.match(candidate):
                        # Already captured via table? Skip duplicates later
                        rest = line.replace(token, "").strip()
                        # Remove other short tokens (page numbers etc.)
                        rest = re.sub(r"\b\d{1,3}\b", "", rest).strip()
                        if len(rest) > 3:
                            band_match = BAND_LEVEL_RE.search(candidate)
                            band = band_match.group(1).upper() if band_match else ""
                            level = band_match.group(2) if band_match else ""
                            records.append({
                                "job_code": candidate,
                                "job_title": rest,
                                "career_band": band,
                                "career_level": int(level) if level else None,
                            })

    df = pd.DataFrame(records)
    if df.empty:
        print("  WARNING: No job codes extracted. Check PDF format.")
        return df

    # De-duplicate: keep first occurrence of each job_code
    df = df.drop_duplicates(subset="job_code", keep="first").reset_index(drop=True)
    print(f"  Extracted {len(df)} unique job codes")
    return df


# ---------------------------------------------------------------------------
# Salary Structure Parser
# ---------------------------------------------------------------------------

def _parse_dollar(val: str) -> float | None:
    """Convert a string like '$45,000' or '45000' to float."""
    if not val:
        return None
    cleaned = re.sub(r"[,$\s]", "", str(val))
    try:
        return float(cleaned)
    except ValueError:
        return None


def parse_salary_structure(pdf_path: str) -> pd.DataFrame:
    """
    Parse the OSU Salary Structure PDF and extract pay grades with
    minimum, midpoint, and maximum values.

    Returns a DataFrame with columns:
        pay_grade, range_min, range_mid, range_max
    """
    import pdfplumber  # lazy import — heavy dependency

    print(f"\n--- Parsing Salary Structure: {pdf_path} ---")
    records = []

    # Pay grade pattern: letter(s) + digits, e.g. A01, A12, B03, N05
    grade_pattern = re.compile(r"^[A-Z]{1,2}\d{1,2}$", re.IGNORECASE)

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            tables = page.extract_tables()
            for table in tables:
                if not table:
                    continue

                # Try to identify header row to find min/mid/max columns
                header_idx = None
                min_col = mid_col = max_col = grade_col = None

                for ri, row in enumerate(table):
                    if not row:
                        continue
                    cells = [_clean_text(str(c)).lower() if c else "" for c in row]
                    # Look for header keywords
                    for ci, cell in enumerate(cells):
                        if "grade" in cell or "pay grade" in cell:
                            grade_col = ci
                        if "min" in cell and "mid" not in cell:
                            min_col = ci
                        if "mid" in cell:
                            mid_col = ci
                        if "max" in cell:
                            max_col = ci
                    if min_col is not None and max_col is not None:
                        header_idx = ri
                        break

                if header_idx is None:
                    # Try positional approach: grade, min, mid, max in consecutive columns
                    for ri, row in enumerate(table):
                        if not row:
                            continue
                        cells = [_clean_text(str(c)) if c else "" for c in row]
                        for ci, cell in enumerate(cells):
                            candidate = cell.replace(" ", "").upper()
                            if grade_pattern.match(candidate) and ci + 3 < len(cells):
                                v1 = _parse_dollar(cells[ci + 1])
                                v2 = _parse_dollar(cells[ci + 2])
                                v3 = _parse_dollar(cells[ci + 3])
                                if v1 and v2 and v3 and v1 < v2 < v3:
                                    records.append({
                                        "pay_grade": candidate,
                                        "range_min": v1,
                                        "range_mid": v2,
                                        "range_max": v3,
                                    })
                    continue

                # We found a header — parse data rows below it
                if grade_col is None:
                    grade_col = 0  # default assumption

                for row in table[header_idx + 1:]:
                    if not row:
                        continue
                    cells = [_clean_text(str(c)) if c else "" for c in row]
                    if grade_col >= len(cells):
                        continue
                    candidate = cells[grade_col].replace(" ", "").upper()
                    if not grade_pattern.match(candidate):
                        continue

                    min_val = _parse_dollar(cells[min_col]) if min_col is not None and min_col < len(cells) else None
                    mid_val = _parse_dollar(cells[mid_col]) if mid_col is not None and mid_col < len(cells) else None
                    max_val = _parse_dollar(cells[max_col]) if max_col is not None and max_col < len(cells) else None

                    if min_val is not None or mid_val is not None or max_val is not None:
                        records.append({
                            "pay_grade": candidate,
                            "range_min": min_val,
                            "range_mid": mid_val,
                            "range_max": max_val,
                        })

            # Fallback: text-based extraction
            text = page.extract_text() or ""
            for line in text.split("\n"):
                tokens = line.split()
                for ti, token in enumerate(tokens):
                    candidate = token.replace(" ", "").upper()
                    if grade_pattern.match(candidate):
                        # Try to find 3 dollar amounts after the grade
                        nums = []
                        for t in tokens[ti + 1:]:
                            v = _parse_dollar(t)
                            if v is not None:
                                nums.append(v)
                            if len(nums) == 3:
                                break
                        if len(nums) == 3 and nums[0] < nums[1] < nums[2]:
                            records.append({
                                "pay_grade": candidate,
                                "range_min": nums[0],
                                "range_mid": nums[1],
                                "range_max": nums[2],
                            })

    df = pd.DataFrame(records)
    if df.empty:
        print("  WARNING: No pay grades extracted. Check PDF format.")
        return df

    df = df.drop_duplicates(subset="pay_grade", keep="first").reset_index(drop=True)
    print(f"  Extracted {len(df)} pay grades")
    return df


# ---------------------------------------------------------------------------
# Excel Loader
# ---------------------------------------------------------------------------

def load_employee_data(xlsx_path: str) -> pd.DataFrame:
    """
    Load the 'Combined' sheet from the employee Excel file.

    Expected columns (positional):
        A: Department, B: Unit, C: Job Profile, D: FTE,
        E: Pay Rate Type, F: Base Pay (Position),
        G: Standardized Hourly, H: Position Group
    """
    print(f"\n--- Loading Employee Data: {xlsx_path} ---")
    df = pd.read_excel(xlsx_path, sheet_name="Combined")

    # Standardize column names based on position if needed
    expected_cols = [
        "Department", "Unit", "Job Profile", "FTE",
        "Pay Rate Type", "Base Pay (Position)",
        "Standardized Hourly", "Position Group",
    ]

    # If the actual column names are different, try to map them
    # First, check if the expected names are close matches
    actual_cols = list(df.columns)
    col_map = {}
    for exp in expected_cols:
        # Try exact match first
        if exp in actual_cols:
            col_map[exp] = exp
            continue
        # Try case-insensitive partial match
        for act in actual_cols:
            if exp.lower() in str(act).lower() or str(act).lower() in exp.lower():
                col_map[exp] = act
                break

    # If we couldn't map enough columns, fall back to positional
    if len(col_map) < 6:
        print("  Column names don't match expected names, using positional mapping")
        if len(actual_cols) >= 8:
            rename_map = {actual_cols[i]: expected_cols[i] for i in range(8)}
            df = df.rename(columns=rename_map)
        else:
            print(f"  WARNING: Expected 8+ columns, got {len(actual_cols)}")
    else:
        rename_map = {v: k for k, v in col_map.items() if v != k}
        if rename_map:
            df = df.rename(columns=rename_map)

    print(f"  Loaded {len(df)} rows")
    print(f"  Columns: {list(df.columns)}")

    # Step 1: Fix blank Standardized Hourly values
    # Formula: Base Pay / (FTE * 2080)
    df["FTE"] = pd.to_numeric(df["FTE"], errors="coerce")
    df["Base Pay (Position)"] = pd.to_numeric(df["Base Pay (Position)"], errors="coerce")
    df["Standardized Hourly"] = pd.to_numeric(df["Standardized Hourly"], errors="coerce")

    blank_mask = df["Standardized Hourly"].isna()
    valid_calc = blank_mask & df["FTE"].notna() & (df["FTE"] > 0) & df["Base Pay (Position)"].notna()
    df.loc[valid_calc, "Standardized Hourly"] = (
        df.loc[valid_calc, "Base Pay (Position)"] / (df.loc[valid_calc, "FTE"] * HOURS_PER_YEAR)
    )
    filled = valid_calc.sum()
    still_blank = df["Standardized Hourly"].isna().sum()
    print(f"  Filled {filled} blank Standardized Hourly values")
    if still_blank > 0:
        print(f"  {still_blank} rows still have blank Standardized Hourly (missing FTE or Base Pay)")

    return df


# ---------------------------------------------------------------------------
# Fuzzy Matching
# ---------------------------------------------------------------------------

def fuzzy_match_jobs(
    employee_df: pd.DataFrame,
    catalog_df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Fuzzy-match each Job Profile in the employee data to the job catalog.

    Adds columns: Job Code, Career Band, Career Level, Match Confidence
    """
    print("\n--- Fuzzy Matching Job Profiles ---")

    if catalog_df.empty:
        print("  WARNING: Job catalog is empty, skipping matching")
        employee_df["Job Code"] = None
        employee_df["Career Band"] = None
        employee_df["Career Level"] = None
        employee_df["Match Confidence"] = 0
        return employee_df

    # Build the choices list from catalog titles
    catalog_titles = catalog_df["job_title"].tolist()
    title_to_row = {t: i for i, t in enumerate(catalog_titles)}

    job_codes = []
    bands = []
    levels = []
    confidences = []

    unique_profiles = employee_df["Job Profile"].dropna().unique()
    print(f"  Matching {len(unique_profiles)} unique job profiles against {len(catalog_titles)} catalog entries...")

    # Cache results for unique profiles
    match_cache = {}
    for profile in unique_profiles:
        result = process.extractOne(
            str(profile),
            catalog_titles,
            scorer=fuzz.token_sort_ratio,
            score_cutoff=0,
        )
        if result:
            matched_title, score, idx = result
            cat_row = catalog_df.iloc[idx]
            match_cache[profile] = {
                "job_code": cat_row["job_code"],
                "career_band": cat_row["career_band"],
                "career_level": cat_row["career_level"],
                "confidence": score,
            }
        else:
            match_cache[profile] = {
                "job_code": None,
                "career_band": None,
                "career_level": None,
                "confidence": 0,
            }

    # Apply cached results to all rows
    for _, row in employee_df.iterrows():
        profile = row.get("Job Profile")
        if pd.isna(profile) or profile not in match_cache:
            job_codes.append(None)
            bands.append(None)
            levels.append(None)
            confidences.append(0)
        else:
            m = match_cache[profile]
            job_codes.append(m["job_code"])
            bands.append(m["career_band"])
            levels.append(m["career_level"])
            confidences.append(m["confidence"])

    employee_df["Job Code"] = job_codes
    employee_df["Career Band"] = bands
    employee_df["Career Level"] = levels
    employee_df["Match Confidence"] = confidences

    matched = sum(1 for c in confidences if c >= DEFAULT_FUZZY_THRESHOLD)
    print(f"  Matched: {matched}, Unmatched (below {DEFAULT_FUZZY_THRESHOLD}): {len(confidences) - matched}")

    return employee_df


# ---------------------------------------------------------------------------
# Pay Range Lookup & Compa-Ratio
# ---------------------------------------------------------------------------

def apply_pay_ranges(
    employee_df: pd.DataFrame,
    salary_df: pd.DataFrame,
    catalog_df: pd.DataFrame,
    threshold: int = DEFAULT_FUZZY_THRESHOLD,
) -> pd.DataFrame:
    """
    Look up pay grade for each matched job and calculate compa-ratios.

    Adds columns: Pay Grade, Range Min, Range Mid, Range Max,
                  Compa Ratio, Below Min Flag
    """
    print("\n--- Applying Pay Ranges & Calculating Compa-Ratios ---")

    # Initialize new columns
    employee_df["Pay Grade"] = None
    employee_df["Range Min"] = None
    employee_df["Range Mid"] = None
    employee_df["Range Max"] = None
    employee_df["Compa Ratio"] = None
    employee_df["Below Min Flag"] = None

    if salary_df.empty:
        print("  WARNING: Salary structure is empty, skipping pay range lookup")
        return employee_df

    # Build pay grade lookup: pay_grade -> {min, mid, max}
    grade_lookup = {}
    for _, row in salary_df.iterrows():
        grade_lookup[row["pay_grade"]] = {
            "min": row["range_min"],
            "mid": row["range_mid"],
            "max": row["range_max"],
        }

    # The job catalog may include a pay grade column. If not, we need to
    # infer the pay grade from the band+level. The mapping between band/level
    # and pay grade isn't 1:1 (per OSU docs), but we'll try what we can.
    #
    # Strategy: If the catalog has a pay_grade column, use it directly.
    # Otherwise, we'll look for the grade in the catalog data or leave it
    # for the user to fill in manually.

    has_pay_grade_in_catalog = "pay_grade" in catalog_df.columns if not catalog_df.empty else False

    if has_pay_grade_in_catalog:
        code_to_grade = dict(zip(catalog_df["job_code"], catalog_df["pay_grade"]))
    else:
        code_to_grade = {}

    matched_count = 0
    for idx, row in employee_df.iterrows():
        job_code = row.get("Job Code")
        if not job_code or row.get("Match Confidence", 0) < threshold:
            continue

        # Look up pay grade
        pay_grade = code_to_grade.get(job_code)
        if not pay_grade:
            # Try to infer from band + level if the salary structure uses
            # band-level as grade identifiers (e.g., "T1", "S2")
            band = row.get("Career Band", "")
            level = row.get("Career Level", "")
            if band and level:
                candidate_grade = f"{band}{level}"
                if candidate_grade in grade_lookup:
                    pay_grade = candidate_grade

        if not pay_grade or pay_grade not in grade_lookup:
            continue

        grade_info = grade_lookup[pay_grade]
        range_min = grade_info["min"]
        range_mid = grade_info["mid"]
        range_max = grade_info["max"]

        # Convert annual ranges to hourly if they look like annual values (> 1000)
        if range_min and range_min > 1000:
            range_min_hourly = range_min / HOURS_PER_YEAR
        else:
            range_min_hourly = range_min

        if range_mid and range_mid > 1000:
            range_mid_hourly = range_mid / HOURS_PER_YEAR
        else:
            range_mid_hourly = range_mid

        if range_max and range_max > 1000:
            range_max_hourly = range_max / HOURS_PER_YEAR
        else:
            range_max_hourly = range_max

        employee_df.at[idx, "Pay Grade"] = pay_grade
        employee_df.at[idx, "Range Min"] = round(range_min_hourly, 4) if range_min_hourly else None
        employee_df.at[idx, "Range Mid"] = round(range_mid_hourly, 4) if range_mid_hourly else None
        employee_df.at[idx, "Range Max"] = round(range_max_hourly, 4) if range_max_hourly else None

        # Compa Ratio = Standardized Hourly / Midpoint Hourly
        std_hourly = row.get("Standardized Hourly")
        if pd.notna(std_hourly) and range_mid_hourly and range_mid_hourly > 0:
            compa = std_hourly / range_mid_hourly
            employee_df.at[idx, "Compa Ratio"] = round(compa, 4)

        # Below Min Flag
        if pd.notna(std_hourly) and range_min_hourly:
            if std_hourly < range_min_hourly:
                employee_df.at[idx, "Below Min Flag"] = "BELOW MIN"
            else:
                employee_df.at[idx, "Below Min Flag"] = "OK"

        matched_count += 1

    print(f"  Applied pay ranges to {matched_count} rows")

    below_min_count = (employee_df["Below Min Flag"] == "BELOW MIN").sum()
    if below_min_count > 0:
        print(f"  {below_min_count} employees are BELOW their pay range minimum")

    return employee_df


# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------

def write_output(
    employee_df: pd.DataFrame,
    output_path: str,
    threshold: int = DEFAULT_FUZZY_THRESHOLD,
) -> None:
    """
    Write the final Excel workbook with 3 sheets:
      1. Mapped Data — all original + new columns
      2. Summary — avg compa ratio, employee count, below-min count by Band+Level
      3. Unmatched — rows with confidence < threshold or no match
    """
    print(f"\n--- Writing Output: {output_path} ---")

    # Sheet 1: Mapped Data (all rows)
    mapped_df = employee_df.copy()

    # Sheet 3: Unmatched rows (confidence < threshold or no match)
    unmatched_df = employee_df[
        (employee_df["Match Confidence"] < threshold)
        | employee_df["Job Code"].isna()
    ].copy()

    # Sheet 2: Summary pivot
    matched_df = employee_df[
        (employee_df["Match Confidence"] >= threshold)
        & employee_df["Career Band"].notna()
        & employee_df["Career Level"].notna()
    ].copy()

    if not matched_df.empty:
        matched_df["Career Level"] = matched_df["Career Level"].astype(int)
        matched_df["Band_Level"] = matched_df["Career Band"] + matched_df["Career Level"].astype(str)

        summary = matched_df.groupby(["Career Band", "Career Level"]).agg(
            Avg_Compa_Ratio=("Compa Ratio", "mean"),
            Employee_Count=("Compa Ratio", "size"),
            Below_Min_Count=("Below Min Flag", lambda x: (x == "BELOW MIN").sum()),
        ).reset_index()

        summary = summary.sort_values("Avg_Compa_Ratio", ascending=True).reset_index(drop=True)
        summary.columns = [
            "Career Band", "Career Level",
            "Avg Compa Ratio", "Employee Count", "Below Min Count",
        ]
        # Round for readability
        summary["Avg Compa Ratio"] = summary["Avg Compa Ratio"].round(4)
    else:
        summary = pd.DataFrame(columns=[
            "Career Band", "Career Level",
            "Avg Compa Ratio", "Employee Count", "Below Min Count",
        ])

    # Write to Excel
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        mapped_df.to_excel(writer, sheet_name="Mapped Data", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)
        unmatched_df.to_excel(writer, sheet_name="Unmatched", index=False)

    print(f"  Sheet 'Mapped Data': {len(mapped_df)} rows")
    print(f"  Sheet 'Summary': {len(summary)} rows")
    print(f"  Sheet 'Unmatched': {len(unmatched_df)} rows")
    print(f"\n  Output saved to: {output_path}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="OSU Career Roadmap Employee Data Mapper"
    )
    parser.add_argument(
        "excel_file",
        help="Path to the employee Excel file (with 'Combined' sheet)",
    )
    parser.add_argument(
        "--job-catalog",
        default=None,
        help="Path to pre-downloaded job catalog PDF (otherwise downloads from OSU HR)",
    )
    parser.add_argument(
        "--salary-structure",
        default=None,
        help="Path to pre-downloaded salary structure PDF (otherwise downloads from OSU HR)",
    )
    parser.add_argument(
        "--output", "-o",
        default=None,
        help="Output file path (default: OSU_Career_Roadmap_Mapped.xlsx in same directory as input)",
    )
    parser.add_argument(
        "--threshold",
        type=int,
        default=DEFAULT_FUZZY_THRESHOLD,
        help=f"Fuzzy match confidence threshold (default: {DEFAULT_FUZZY_THRESHOLD})",
    )

    args = parser.parse_args()

    # Validate input file
    excel_path = os.path.expanduser(args.excel_file)
    if not os.path.isfile(excel_path):
        print(f"ERROR: Excel file not found: {excel_path}")
        sys.exit(1)

    # Set output path
    if args.output:
        output_path = args.output
    else:
        output_dir = os.path.dirname(os.path.abspath(excel_path))
        output_path = os.path.join(output_dir, "OSU_Career_Roadmap_Mapped.xlsx")

    threshold = args.threshold

    # --- Step 2: Download/load job catalog PDF ---
    if args.job_catalog:
        job_catalog_path = args.job_catalog
    else:
        job_catalog_path = os.path.join(tempfile.gettempdir(), "osu-job-catalog.pdf")
        if not os.path.isfile(job_catalog_path):
            try:
                download_pdf(JOB_CATALOG_URL, job_catalog_path)
            except Exception as e:
                print(f"  WARNING: Could not download job catalog: {e}")
                print("  You can download it manually and pass with --job-catalog")
                job_catalog_path = None

    # --- Step 3: Download/load salary structure PDF ---
    if args.salary_structure:
        salary_structure_path = args.salary_structure
    else:
        salary_structure_path = os.path.join(tempfile.gettempdir(), "osu-salary-structure.pdf")
        if not os.path.isfile(salary_structure_path):
            try:
                download_pdf(SALARY_STRUCTURE_URL, salary_structure_path)
            except Exception as e:
                print(f"  WARNING: Could not download salary structure: {e}")
                print("  You can download it manually and pass with --salary-structure")
                salary_structure_path = None

    # Parse PDFs
    catalog_df = pd.DataFrame()
    if job_catalog_path and os.path.isfile(job_catalog_path):
        catalog_df = parse_job_catalog(job_catalog_path)

    salary_df = pd.DataFrame()
    if salary_structure_path and os.path.isfile(salary_structure_path):
        salary_df = parse_salary_structure(salary_structure_path)

    # --- Step 1 & 3: Load employee data (fills blank Standardized Hourly) ---
    employee_df = load_employee_data(excel_path)

    # --- Step 4: Fuzzy match + enrichment ---
    employee_df = fuzzy_match_jobs(employee_df, catalog_df)
    employee_df = apply_pay_ranges(employee_df, salary_df, catalog_df, threshold=threshold)

    # --- Step 5: Output ---
    write_output(employee_df, output_path, threshold=threshold)

    print("\nDone!")


if __name__ == "__main__":
    main()
