
import pandas as pd

from scripts.validation_helpers import (
    missing,
    parse_dates,
    check_required_columns,
)


def test_missing_detects_blank_values():
    assert missing(None)
    assert missing(float("nan"))
    assert missing("")
    assert missing("   ")


def test_missing_accepts_valid_values():
    assert not missing("CT001")
    assert not missing(25)


def test_parse_dates_converts_valid_and_invalid_dates():
    df = pd.DataFrame({
        "Visit_Date": ["2026-01-01", "not-a-date"]
    })

    parse_dates(df, "Visit_Date")

    assert pd.notna(df.loc[0, "Visit_Date"])
    assert pd.isna(df.loc[1, "Visit_Date"])


def test_parse_dates_ignores_missing_column():
    df = pd.DataFrame({"Subject_ID": ["CT001"]})

    parse_dates(df, "Visit_Date")

    assert "Visit_Date" not in df.columns


def test_check_required_columns_passes_when_all_exist():
    df = pd.DataFrame({
        "Subject_ID": ["CT001"],
        "Age": [25],
    })

    assert check_required_columns(
        df, ["Subject_ID", "Age"]
    )


def test_check_required_columns_fails_when_one_is_missing():
    df = pd.DataFrame({"Subject_ID": ["CT001"]})

    assert not check_required_columns(
        df, ["Subject_ID", "Age"]
    )
