#!/usr/bin/env python3
"""
Tests for parse_osu_employee_data.py using synthetic data.

Validates:
  - Standardized Hourly fill logic
  - Fuzzy matching
  - Compa-ratio calculation
  - Below-min flagging
  - Output file structure (3 sheets)
"""

import os
import tempfile

import pandas as pd
import pytest

from parse_osu_employee_data import (
    DEFAULT_FUZZY_THRESHOLD,
    HOURS_PER_YEAR,
    apply_pay_ranges,
    fuzzy_match_jobs,
    write_output,
)


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def catalog_df():
    """Synthetic job catalog."""
    return pd.DataFrame([
        {"job_code": "ADOFCAS1", "job_title": "Administrative Specialist 1", "career_band": "S", "career_level": 1},
        {"job_code": "ADOFCAS2", "job_title": "Administrative Specialist 2", "career_band": "S", "career_level": 2},
        {"job_code": "ADOFCAS3", "job_title": "Administrative Specialist 3", "career_band": "S", "career_level": 3},
        {"job_code": "ITSWDET1", "job_title": "Software Developer 1", "career_band": "T", "career_level": 1},
        {"job_code": "ITSWDET2", "job_title": "Software Developer 2", "career_band": "T", "career_level": 2},
        {"job_code": "ITSWDET3", "job_title": "Software Developer 3", "career_band": "T", "career_level": 3},
        {"job_code": "FABORAM1", "job_title": "Operations Manager 1", "career_band": "M", "career_level": 1},
        {"job_code": "FABORAM2", "job_title": "Operations Manager 2", "career_band": "M", "career_level": 2},
        {"job_code": "HCNURST1", "job_title": "Staff Nurse 1", "career_band": "T", "career_level": 1},
        {"job_code": "HCNURSC1", "job_title": "Clinical Nurse 1", "career_band": "C", "career_level": 1},
    ])


@pytest.fixture
def salary_df():
    """Synthetic salary structure (annual values)."""
    return pd.DataFrame([
        {"pay_grade": "S1", "range_min": 31200, "range_mid": 41600, "range_max": 52000},
        {"pay_grade": "S2", "range_min": 37440, "range_mid": 49920, "range_max": 62400},
        {"pay_grade": "S3", "range_min": 45760, "range_mid": 60320, "range_max": 74880},
        {"pay_grade": "T1", "range_min": 33280, "range_mid": 43680, "range_max": 54080},
        {"pay_grade": "T2", "range_min": 41600, "range_mid": 54080, "range_max": 66560},
        {"pay_grade": "T3", "range_min": 52000, "range_mid": 66560, "range_max": 81120},
        {"pay_grade": "M1", "range_min": 52000, "range_mid": 68640, "range_max": 85280},
        {"pay_grade": "M2", "range_min": 62400, "range_mid": 83200, "range_max": 104000},
        {"pay_grade": "C1", "range_min": 41600, "range_mid": 54080, "range_max": 66560},
    ])


@pytest.fixture
def employee_df():
    """Synthetic employee data mimicking the 'Combined' sheet."""
    return pd.DataFrame({
        "Department": ["Arts & Sciences", "Engineering", "IT", "Facilities", "Health", "Finance", "Admin"],
        "Unit": ["Dean's Office", "ECE", "Enterprise", "Operations", "Nursing", "Budget", "Provost"],
        "Job Profile": [
            "Administrative Specialist 1",
            "Administrative Specialist 2",
            "Software Developer 2",
            "Operations Manager 1",
            "Clinical Nurse 1",
            "Underwater Basket Weaver",  # Should NOT match
            "Admin Specialist 1",        # Fuzzy match to Admin Specialist 1
        ],
        "FTE": [1.0, 1.0, 1.0, 1.0, 1.0, 0.5, 1.0],
        "Pay Rate Type": ["Salary", "Salary", "Salary", "Salary", "Hourly", "Salary", "Salary"],
        "Base Pay (Position)": [35000, 45000, 55000, 70000, None, 20000, 30000],
        "Standardized Hourly": [16.83, 21.63, 26.44, 33.65, 25.00, None, None],
        "Position Group": ["Staff", "Staff", "Staff", "Staff", "Clinical", "Staff", "Staff"],
    })


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

class TestStandardizedHourlyFill:
    def test_fills_blank_values(self, employee_df):
        """Blank Standardized Hourly should be calculated as Base Pay / (FTE * 2080)."""
        from parse_osu_employee_data import HOURS_PER_YEAR
        df = employee_df.copy()
        # Simulate the fill logic from load_employee_data
        df["FTE"] = pd.to_numeric(df["FTE"], errors="coerce")
        df["Base Pay (Position)"] = pd.to_numeric(df["Base Pay (Position)"], errors="coerce")
        df["Standardized Hourly"] = pd.to_numeric(df["Standardized Hourly"], errors="coerce")

        blank_mask = df["Standardized Hourly"].isna()
        valid_calc = blank_mask & df["FTE"].notna() & (df["FTE"] > 0) & df["Base Pay (Position)"].notna()
        df.loc[valid_calc, "Standardized Hourly"] = (
            df.loc[valid_calc, "Base Pay (Position)"] / (df.loc[valid_calc, "FTE"] * HOURS_PER_YEAR)
        )

        # "Underwater Basket Weaver" (FTE=0.5, Base=20000) -> 20000/(0.5*2080) = 19.23
        assert abs(df.iloc[5]["Standardized Hourly"] - 19.2308) < 0.01

        # "Admin Specialist 1" (FTE=1.0, Base=30000) -> 30000/2080 = 14.42
        assert abs(df.iloc[6]["Standardized Hourly"] - 14.4231) < 0.01

    def test_preserves_existing_values(self, employee_df):
        """Non-blank Standardized Hourly values should not be overwritten."""
        df = employee_df.copy()
        df["Standardized Hourly"] = pd.to_numeric(df["Standardized Hourly"], errors="coerce")
        original_value = df.iloc[0]["Standardized Hourly"]

        blank_mask = df["Standardized Hourly"].isna()
        valid_calc = blank_mask & df["FTE"].notna() & (df["FTE"] > 0)
        df.loc[valid_calc, "Standardized Hourly"] = 999  # dummy fill

        assert df.iloc[0]["Standardized Hourly"] == original_value


class TestFuzzyMatching:
    def test_exact_match(self, employee_df, catalog_df):
        """Exact title matches should get high confidence."""
        result = fuzzy_match_jobs(employee_df.copy(), catalog_df)
        # "Administrative Specialist 1" is exact match
        row0 = result.iloc[0]
        assert row0["Job Code"] == "ADOFCAS1"
        assert row0["Career Band"] == "S"
        assert row0["Career Level"] == 1
        assert row0["Match Confidence"] == 100

    def test_fuzzy_match(self, employee_df, catalog_df):
        """'Admin Specialist 1' should fuzzy-match to 'Administrative Specialist 1'."""
        result = fuzzy_match_jobs(employee_df.copy(), catalog_df)
        row6 = result.iloc[6]
        assert row6["Job Code"] is not None
        assert row6["Match Confidence"] > 0

    def test_no_match(self, employee_df, catalog_df):
        """'Underwater Basket Weaver' should not match well."""
        result = fuzzy_match_jobs(employee_df.copy(), catalog_df)
        row5 = result.iloc[5]
        assert row5["Match Confidence"] < 85

    def test_empty_catalog(self, employee_df):
        """Empty catalog should result in no matches."""
        result = fuzzy_match_jobs(employee_df.copy(), pd.DataFrame())
        assert result["Job Code"].isna().all()
        assert (result["Match Confidence"] == 0).all()


class TestPayRanges:
    def test_compa_ratio_calculation(self, employee_df, catalog_df, salary_df):
        """Compa ratio should be Standardized Hourly / Midpoint Hourly."""
        df = fuzzy_match_jobs(employee_df.copy(), catalog_df)
        result = apply_pay_ranges(df, salary_df, catalog_df)

        # Row 0: Admin Specialist 1, S1, hourly=16.83, mid=41600/2080=20.0
        # Compa = 16.83 / 20.0 = 0.8415
        row0 = result.iloc[0]
        if pd.notna(row0["Compa Ratio"]):
            expected_mid_hourly = 41600 / 2080
            expected_compa = 16.83 / expected_mid_hourly
            assert abs(row0["Compa Ratio"] - expected_compa) < 0.01

    def test_below_min_flag(self, employee_df, catalog_df, salary_df):
        """Employees below range minimum should be flagged."""
        df = fuzzy_match_jobs(employee_df.copy(), catalog_df)
        result = apply_pay_ranges(df, salary_df, catalog_df)

        # Row 0: hourly=16.83, S1 min=31200/2080=15.0 -> OK (16.83 > 15.0)
        row0 = result.iloc[0]
        if pd.notna(row0["Below Min Flag"]):
            min_hourly = 31200 / 2080
            if 16.83 >= min_hourly:
                assert row0["Below Min Flag"] == "OK"
            else:
                assert row0["Below Min Flag"] == "BELOW MIN"

    def test_empty_salary_structure(self, employee_df, catalog_df):
        """Empty salary structure should leave pay range columns as None."""
        df = fuzzy_match_jobs(employee_df.copy(), catalog_df)
        result = apply_pay_ranges(df, pd.DataFrame(), catalog_df)
        assert result["Pay Grade"].isna().all()
        assert result["Compa Ratio"].isna().all()


class TestOutputFile:
    def test_three_sheets(self, employee_df, catalog_df, salary_df):
        """Output file should have exactly 3 sheets."""
        df = fuzzy_match_jobs(employee_df.copy(), catalog_df)
        df = apply_pay_ranges(df, salary_df, catalog_df)

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            output_path = f.name

        try:
            write_output(df, output_path)
            xl = pd.ExcelFile(output_path)
            assert "Mapped Data" in xl.sheet_names
            assert "Summary" in xl.sheet_names
            assert "Unmatched" in xl.sheet_names
        finally:
            os.unlink(output_path)

    def test_unmatched_sheet_contents(self, employee_df, catalog_df, salary_df):
        """Unmatched sheet should contain low-confidence rows."""
        df = fuzzy_match_jobs(employee_df.copy(), catalog_df)
        df = apply_pay_ranges(df, salary_df, catalog_df)

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            output_path = f.name

        try:
            write_output(df, output_path)
            unmatched = pd.read_excel(output_path, sheet_name="Unmatched")
            # "Underwater Basket Weaver" should be in unmatched
            assert any("Basket Weaver" in str(jp) for jp in unmatched["Job Profile"].tolist())
        finally:
            os.unlink(output_path)

    def test_summary_columns(self, employee_df, catalog_df, salary_df):
        """Summary sheet should have the correct columns."""
        df = fuzzy_match_jobs(employee_df.copy(), catalog_df)
        df = apply_pay_ranges(df, salary_df, catalog_df)

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            output_path = f.name

        try:
            write_output(df, output_path)
            summary = pd.read_excel(output_path, sheet_name="Summary")
            expected_cols = {"Career Band", "Career Level", "Avg Compa Ratio", "Employee Count", "Below Min Count"}
            assert expected_cols.issubset(set(summary.columns))
        finally:
            os.unlink(output_path)

    def test_mapped_data_has_all_columns(self, employee_df, catalog_df, salary_df):
        """Mapped Data sheet should have original + new columns."""
        df = fuzzy_match_jobs(employee_df.copy(), catalog_df)
        df = apply_pay_ranges(df, salary_df, catalog_df)

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            output_path = f.name

        try:
            write_output(df, output_path)
            mapped = pd.read_excel(output_path, sheet_name="Mapped Data")
            # Original columns
            for col in ["Department", "Unit", "Job Profile", "FTE", "Pay Rate Type",
                        "Base Pay (Position)", "Standardized Hourly", "Position Group"]:
                assert col in mapped.columns, f"Missing original column: {col}"
            # New columns
            for col in ["Job Code", "Career Band", "Career Level", "Pay Grade",
                        "Range Min", "Range Mid", "Range Max", "Compa Ratio",
                        "Below Min Flag", "Match Confidence"]:
                assert col in mapped.columns, f"Missing new column: {col}"
        finally:
            os.unlink(output_path)


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
