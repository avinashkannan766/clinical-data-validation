
import pandas as pd


def missing(value):
    if pd.isna(value):
        return True
    return isinstance(value, str) and value.strip() == ""


def parse_dates(df, column):
    if column in df.columns:
        df[column] = pd.to_datetime(
            df[column],
            errors="coerce",
        )


def check_required_columns(df, required_columns):
    return all(
        column in df.columns
        for column in required_columns
    )
