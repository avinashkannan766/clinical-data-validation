
from pathlib import Path

import pandas as pd


PROJECT_DIR = Path(__file__).resolve().parent.parent
REPORTS_DIR = PROJECT_DIR / "reports"


def test_clean_report_has_no_errors_or_reviews():
    """The clean synthetic dataset should have no validation findings."""
    report_path = REPORTS_DIR / "validation_report.xlsx"

    assert report_path.exists(), f"Missing report: {report_path}"

    sheets = pd.read_excel(report_path, sheet_name=None)

    assert "Errors" in sheets
    assert "Reviews" in sheets

    assert sheets["Errors"].empty, "Clean report contains errors"
    assert sheets["Reviews"].empty, "Clean report contains review findings"


def test_test_report_contains_expected_findings():
    """The intentionally flawed synthetic dataset should trigger findings."""
    report_path = REPORTS_DIR / "test_validation_report.xlsx"

    assert report_path.exists(), f"Missing report: {report_path}"

    sheets = pd.read_excel(report_path, sheet_name=None)

    assert "Errors" in sheets
    assert "Reviews" in sheets

    error_count = len(sheets["Errors"])
    review_count = len(sheets["Reviews"])

    assert error_count > 0, "Expected seeded defects to produce errors"
    assert review_count > 0, "Expected seeded defects to produce reviews"

    assert error_count == 19
    assert review_count == 5


def test_defect_summary_has_twelve_unique_matches():
    """Each seeded defect should have exactly one matching finding."""
    summary_path = REPORTS_DIR / "defect_detection_summary.xlsx"

    assert summary_path.exists(), f"Missing summary: {summary_path}"

    summary = pd.read_excel(
        summary_path,
        sheet_name="Defect_Summary"
    )

    required_columns = {
        "Defect_ID",
        "Match_Count",
        "Status",
        "Matched_Finding",
    }
    assert required_columns.issubset(summary.columns)

    assert len(summary) == 12
    assert summary["Defect_ID"].nunique() == 12
    assert (summary["Status"] == "PASS").all()
    assert (summary["Match_Count"] == 1).all()
    assert summary["Matched_Finding"].notna().all()
