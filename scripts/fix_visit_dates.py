
from pathlib import Path
import pandas as pd

# --------------------------------------------------
# 1. File paths
# --------------------------------------------------

project_dir = Path(__file__).resolve().parent.parent

input_file = project_dir / "data" / "clinical_study_raw.xlsx"
output_file = project_dir / "data" / "clinical_study_corrected.xlsx"

print("Input workbook:", input_file)

if not input_file.exists():
    raise FileNotFoundError(f"Workbook not found: {input_file}")

# --------------------------------------------------
# 2. Read all workbook sheets
# --------------------------------------------------

sheets = pd.read_excel(input_file, sheet_name=None)

required_sheets = [
    "Participants",
    "Visits",
    "Laboratory",
    "Adverse_Events",
]

missing_sheets = [name for name in required_sheets if name not in sheets]

if missing_sheets:
    raise ValueError(f"Missing workbook sheets: {missing_sheets}")

participants = sheets["Participants"].copy()
visits = sheets["Visits"].copy()
lab = sheets["Laboratory"].copy()

print("\nOriginal record counts:")
for name, df in sheets.items():
    print(f"{name}: {len(df)} records")

# --------------------------------------------------
# 3. Check required columns
# --------------------------------------------------

required_columns = {
    "Participants": [
        "Subject_ID",
        "Consent_Date",
    ],
    "Visits": [
        "Subject_ID",
        "Visit_Name",
        "Visit_Date",
        "Visit_Status",
    ],
    "Laboratory": [
        "Lab_ID",
        "Subject_ID",
        "Visit_Name",
        "Test_Name",
        "Result_Value",
        "Unit",
        "Collection_Date",
    ],
}

for sheet_name, columns in required_columns.items():
    actual_columns = sheets[sheet_name].columns.tolist()
    missing_columns = [
        col for col in columns if col not in actual_columns
    ]

    if missing_columns:
        raise ValueError(
            f"{sheet_name} is missing columns: {missing_columns}\n"
            f"Actual columns: {actual_columns}"
        )

# --------------------------------------------------
# 4. Parse participant consent dates
# --------------------------------------------------

participants["Consent_Date"] = pd.to_datetime(
    participants["Consent_Date"],
    errors="coerce",
)

invalid_consent_dates = participants[
    participants["Consent_Date"].isna()
]

if not invalid_consent_dates.empty:
    print("\nERROR: Missing or invalid consent dates:")
    print(invalid_consent_dates.to_string(index=False))
    raise ValueError(
        "Cannot calculate visit dates until consent dates are corrected."
    )

# --------------------------------------------------
# 5. Inspect laboratory keys BEFORE correction
# --------------------------------------------------

print("\nChecking laboratory record keys...")

missing_keys = lab[
    lab["Subject_ID"].isna()
    | lab["Visit_Name"].isna()
]

if not missing_keys.empty:
    print("\nERROR: Laboratory records have missing Subject_ID or Visit_Name:")
    print(missing_keys.to_string(index=False))

known_subjects = set(participants["Subject_ID"].dropna())

unknown_subjects = lab[
    lab["Subject_ID"].notna()
    & ~lab["Subject_ID"].isin(known_subjects)
]

if not unknown_subjects.empty:
    print("\nERROR: Laboratory records contain unknown subjects:")
    print(unknown_subjects.to_string(index=False))

valid_visit_names = ["Baseline", "Week 4", "Week 8"]

invalid_visit_names = lab[
    lab["Visit_Name"].notna()
    & ~lab["Visit_Name"].isin(valid_visit_names)
]

if not invalid_visit_names.empty:
    print("\nERROR: Laboratory records have unexpected visit names:")
    print(invalid_visit_names.to_string(index=False))

# Stop safely if a laboratory key cannot be matched.
if (
    not missing_keys.empty
    or not unknown_subjects.empty
    or not invalid_visit_names.empty
):
    raise ValueError(
        "\nCorrection stopped. Review the laboratory records printed above. "
        "No corrected workbook was written."
    )

# --------------------------------------------------
# 6. Validate participant and visit identifiers
# --------------------------------------------------

if participants["Subject_ID"].isna().any():
    print("\nERROR: A participant has a missing Subject_ID.")
    raise ValueError("Cannot continue with missing participant IDs.")

if participants["Subject_ID"].duplicated().any():
    print("\nERROR: Duplicate participant IDs:")
    print(
        participants[
            participants["Subject_ID"].duplicated(keep=False)
        ].to_string(index=False)
    )
    raise ValueError("Cannot continue with duplicate participant IDs.")

if visits["Subject_ID"].isna().any():
    print("\nERROR: A visit has a missing Subject_ID.")
    print(visits[visits["Subject_ID"].isna()].to_string(index=False))
    raise ValueError("Cannot continue with missing visit subject IDs.")

if visits["Visit_Name"].isna().any():
    print("\nERROR: A visit has a missing Visit_Name.")
    print(visits[visits["Visit_Name"].isna()].to_string(index=False))
    raise ValueError("Cannot continue with missing visit names.")

unexpected_visit_names = visits[
    ~visits["Visit_Name"].isin(valid_visit_names)
]

if not unexpected_visit_names.empty:
    print("\nERROR: Unexpected visit names in Visits sheet:")
    print(unexpected_visit_names.to_string(index=False))
    raise ValueError("Cannot continue with invalid visit names.")

# Each subject should have at most one record per visit.
duplicate_visits = visits.duplicated(
    subset=["Subject_ID", "Visit_Name"],
    keep=False,
)

if duplicate_visits.any():
    print("\nERROR: Duplicate subject-visit records:")
    print(visits[duplicate_visits].to_string(index=False))
    raise ValueError("Cannot safely match duplicate visits.")

# Every visit must belong to a known participant.
unknown_visit_subjects = visits[
    ~visits["Subject_ID"].isin(known_subjects)
]

if not unknown_visit_subjects.empty:
    print("\nERROR: Visits contain unknown subjects:")
    print(unknown_visit_subjects.to_string(index=False))
    raise ValueError("Cannot continue with unknown visit subjects.")

# --------------------------------------------------
# 7. Calculate corrected visit dates
# --------------------------------------------------

day_offsets = {
    "Baseline": 0,
    "Week 4": 28,
    "Week 8": 56,
}

consent_dates = participants.set_index("Subject_ID")["Consent_Date"]

visits["Visit_Date"] = visits.apply(
    lambda row: (
        consent_dates.loc[row["Subject_ID"]]
        + pd.Timedelta(days=day_offsets[row["Visit_Name"]])
    ),
    axis=1,
)

# --------------------------------------------------
# 8. Match laboratory collection dates to visits
# --------------------------------------------------

visit_date_lookup = visits.set_index(
    ["Subject_ID", "Visit_Name"]
)["Visit_Date"]

lab["Collection_Date"] = lab.apply(
    lambda row: visit_date_lookup.loc[
        (row["Subject_ID"], row["Visit_Name"])
    ],
    axis=1,
)

# --------------------------------------------------
# 9. Show corrected date previews
# --------------------------------------------------

for subject in ["CT001", "CT030"]:
    print(f"\nCorrected visit dates for {subject}:")
    print(
        visits[visits["Subject_ID"] == subject][
            ["Subject_ID", "Visit_Name", "Visit_Date"]
        ].to_string(index=False)
    )

    print(f"\nCorrected laboratory dates for {subject}:")
    print(
        lab[lab["Subject_ID"] == subject][
            ["Lab_ID", "Visit_Name", "Collection_Date"]
        ].to_string(index=False)
    )

# --------------------------------------------------
# 10. Save to a separate workbook
# --------------------------------------------------

sheets["Participants"] = participants
sheets["Visits"] = visits
sheets["Laboratory"] = lab

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    for sheet_name, df in sheets.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print("\nSUCCESS: Corrected workbook created.")
print("Output:", output_file)
print("\nFinal record counts:")

for sheet_name, df in sheets.items():
    print(f"{sheet_name}: {len(df)} records")

print("\nThe original clinical_study_raw.xlsx was not modified.")
