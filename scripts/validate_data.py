
from pathlib import Path
import sys
import pandas as pd

# ============================================================
# 1. CONFIGURATION
# ============================================================

PROJECT_DIR = Path(__file__).resolve().parent.parent

# Choose mode from the command line.
# clean = corrected workbook and clean report
# test  = defective workbook and test report
MODE = sys.argv[1].lower() if len(sys.argv) > 1 else "clean"

if MODE not in {"clean", "test"}:
    raise SystemExit(
        "Invalid mode. Use: python validate_data.py clean"
        " or python validate_data.py test"
    )

if MODE == "clean":
    INPUT_FILENAME = "clinical_study_corrected.xlsx"
    REPORT_FILENAME = "validation_report.xlsx"
else:
    INPUT_FILENAME = "clinical_study_test.xlsx"
    REPORT_FILENAME = "test_validation_report.xlsx"

FILE_PATH = PROJECT_DIR / "data" / INPUT_FILENAME

EXPECTED_SHEETS = [
    "Participants",
    "Visits",
    "Laboratory",
    "Adverse_Events",
]

VALID_SEX = {"Female", "Male", "Other", "Not Reported"}
VALID_CONSENT = {"Yes", "No"}
VALID_ARMS = {"Treatment A", "Control B"}
VALID_VISITS = {"Baseline", "Week 4", "Week 8"}
VALID_VISIT_STATUS = {"Completed", "Missed", "Not Done"}
VALID_TESTS = {"Glucose"}
VALID_UNITS = {"mg/dL", "mmol/L"}
VALID_SERIOUSNESS = {"Serious", "Non-serious", "Pending assessment"}

VALID_OUTCOMES = {
    "Recovered",
    "Recovering",
    "Not Recovered",
    "Unknown",
    "Other",
}

DAY_OFFSETS = {
    "Baseline": 0,
    "Week 4": 28,
    "Week 8": 56,
}

ERRORS = []
REVIEWS = []


# ============================================================
# 2. HELPER FUNCTIONS
# ============================================================

def add_error(rule_id, message):
    ERRORS.append({
        "Severity": "ERROR",
        "Rule_ID": rule_id,
        "Issue": message,
    })


def add_review(rule_id, message):
    REVIEWS.append({
        "Severity": "REVIEW",
        "Rule_ID": rule_id,
        "Issue": message,
    })


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


def check_required_columns(sheet_name, df, required_columns):
    missing_columns = [
        column
        for column in required_columns
        if column not in df.columns
    ]

    if missing_columns:
        add_error(
            "STRUCT-001",
            f"{sheet_name}: missing columns {missing_columns}",
        )
        return False

    return True


def print_section(title):
    print("\n" + "=" * 70)
    print(title)
    print("=" * 70)


# ============================================================
# 3. READ WORKBOOK
# ============================================================

print("Clinical Data Management Validation")
print("Mode:", MODE.upper())
print("Workbook:", FILE_PATH)

if not FILE_PATH.exists():
    raise FileNotFoundError(
        f"Workbook not found: {FILE_PATH}"
    )

sheets = pd.read_excel(
    FILE_PATH,
    sheet_name=None,
)

missing_sheets = [
    sheet
    for sheet in EXPECTED_SHEETS
    if sheet not in sheets
]

if missing_sheets:
    raise ValueError(
        f"Workbook is missing sheets: {missing_sheets}"
    )

participants = sheets["Participants"].copy()
visits = sheets["Visits"].copy()
lab = sheets["Laboratory"].copy()
ae = sheets["Adverse_Events"].copy()

print("\nRecord counts:")
for name, df in sheets.items():
    print(f"{name}: {len(df)} records")


# ============================================================
# 4. REQUIRED COLUMN CHECKS
# ============================================================

print_section("STRUCTURAL VALIDATION")

structure_ok = True

required_columns = {
    "Participants": [
        "Subject_ID",
        "Age",
        "Sex",
        "Consent_Status",
        "Consent_Date",
        "Treatment_Arm",
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
    "Adverse_Events": [
        "AE_ID",
        "Subject_ID",
        "Event_Description",
        "Onset_Date",
        "Seriousness",
        "Outcome",
    ],
}

for sheet_name, df in [
    ("Participants", participants),
    ("Visits", visits),
    ("Laboratory", lab),
    ("Adverse_Events", ae),
]:
    result = check_required_columns(
        sheet_name,
        df,
        required_columns[sheet_name],
    )

    print(
        f"{sheet_name}: "
        f"{'PASS' if result else 'FAIL'} - required columns"
    )

    if not result:
        structure_ok = False

if not structure_ok:
    raise ValueError(
        "Structural validation failed. Fix the missing columns "
        "before continuing."
    )

parse_dates(participants, "Consent_Date")
parse_dates(visits, "Visit_Date")
parse_dates(lab, "Collection_Date")
parse_dates(ae, "Onset_Date")


# ============================================================
# 5. PARTICIPANT VALIDATION
# ============================================================

print_section("PARTICIPANT VALIDATION")

# VAL-001: Subject_ID required and unique
if participants["Subject_ID"].isna().any():
    add_error(
        "VAL-001",
        "Participants: missing Subject_ID.",
    )

duplicate_subjects = participants[
    participants["Subject_ID"].notna()
    & participants["Subject_ID"].duplicated(keep=False)
]

for subject_id in duplicate_subjects["Subject_ID"].unique():
    add_error(
        "VAL-001",
        f"Participants: duplicate Subject_ID {subject_id}.",
    )

# VAL-002: Age must be present, integer, and 18-80
age_numeric = pd.to_numeric(
    participants["Age"],
    errors="coerce",
)

for index, row in participants.iterrows():
    age = age_numeric.loc[index]
    subject = row["Subject_ID"]

    if pd.isna(age) or age != int(age):
        add_error(
            "VAL-002",
            f"Participant {subject}: Age is missing or not an integer.",
        )
    elif age < 18 or age > 80:
        add_review(
            "VAL-002",
            f"Participant {subject}: Age {age} is outside 18-80.",
        )

# VAL-003: Sex allowed values
for _, row in participants.iterrows():
    if missing(row["Sex"]) or row["Sex"] not in VALID_SEX:
        add_error(
            "VAL-003",
            f"Participant {row['Subject_ID']}: invalid Sex value "
            f"{row['Sex']!r}.",
        )

# VAL-004: Consent status
for _, row in participants.iterrows():
    if (
        missing(row["Consent_Status"])
        or row["Consent_Status"] not in VALID_CONSENT
    ):
        add_error(
            "VAL-004",
            f"Participant {row['Subject_ID']}: invalid Consent_Status.",
        )

# VAL-005: Consent date required and valid
for _, row in participants.iterrows():
    if pd.isna(row["Consent_Date"]):
        add_error(
            "VAL-005",
            f"Participant {row['Subject_ID']}: missing or invalid Consent_Date.",
        )

# VAL-006: Treatment arm
for _, row in participants.iterrows():
    if (
        missing(row["Treatment_Arm"])
        or row["Treatment_Arm"] not in VALID_ARMS
    ):
        add_error(
            "VAL-006",
            f"Participant {row['Subject_ID']}: invalid Treatment_Arm.",
        )

# Build participant lookups from nonmissing subject IDs.
valid_participants = participants[
    participants["Subject_ID"].notna()
].copy()

consent_lookup = (
    valid_participants
    .drop_duplicates("Subject_ID")
    .set_index("Subject_ID")["Consent_Date"]
    .to_dict()
)

participant_ids = set(consent_lookup.keys())

arm_counts = participants["Treatment_Arm"].value_counts()

print("Participant records:", len(participants))
print("Treatment arm counts:", arm_counts.to_dict())


# ============================================================
# 6. VISIT VALIDATION
# ============================================================

print_section("VISIT VALIDATION")

# VAL-007: Visit subject must exist
for _, row in visits.iterrows():
    subject = row["Subject_ID"]

    if missing(subject) or subject not in participant_ids:
        add_error(
            "VAL-007",
            f"Visit row {row.name + 2}: subject {subject!r} "
            "does not exist in Participants.",
        )

# VAL-008: Visit name allowed
for _, row in visits.iterrows():
    if missing(row["Visit_Name"]) or row["Visit_Name"] not in VALID_VISITS:
        add_error(
            "VAL-008",
            f"Visit row {row.name + 2}: invalid Visit_Name "
            f"{row['Visit_Name']!r}.",
        )

# VAL-009: Visit date valid
for _, row in visits.iterrows():
    if pd.isna(row["Visit_Date"]):
        add_error(
            "VAL-009",
            f"Visit row {row.name + 2}: missing or invalid Visit_Date.",
        )

# VAL-010: Visit status allowed
for _, row in visits.iterrows():
    if (
        missing(row["Visit_Status"])
        or row["Visit_Status"] not in VALID_VISIT_STATUS
    ):
        add_error(
            "VAL-010",
            f"Visit row {row.name + 2}: invalid Visit_Status "
            f"{row['Visit_Status']!r}.",
        )

# VAL-011: Subject_ID + Visit_Name unique
duplicate_visit_mask = visits.duplicated(
    subset=["Subject_ID", "Visit_Name"],
    keep=False,
)

for _, row in visits[duplicate_visit_mask].iterrows():
    add_error(
        "VAL-011",
        f"Duplicate visit for {row['Subject_ID']} / {row['Visit_Name']}.",
    )

# VAL-012: Visit windows relative to consent date
for _, row in visits.iterrows():
    subject = row["Subject_ID"]
    visit_name = row["Visit_Name"]

    if (
        subject not in consent_lookup
        or visit_name not in DAY_OFFSETS
        or pd.isna(row["Visit_Date"])
    ):
        continue

    consent_date = consent_lookup[subject]
    actual_date = row["Visit_Date"]

    if pd.isna(consent_date):
        continue

    day_number = (actual_date - consent_date).days

    if visit_name == "Baseline":
        if day_number != 0:
            add_review(
                "VAL-012",
                f"{subject} Baseline is Day {day_number}; expected Day 0.",
            )

    elif visit_name == "Week 4":
        if not 21 <= day_number <= 35:
            add_review(
                "VAL-012",
                f"{subject} Week 4 is Day {day_number}; "
                "expected Days 21-35.",
            )

    elif visit_name == "Week 8":
        if not 49 <= day_number <= 63:
            add_review(
                "VAL-012",
                f"{subject} Week 8 is Day {day_number}; "
                "expected Days 49-63.",
            )


# ============================================================
# 7. LABORATORY VALIDATION
# ============================================================

print_section("LABORATORY VALIDATION")

# VAL-013: Subject must exist
for _, row in lab.iterrows():
    subject = row["Subject_ID"]

    if missing(subject) or subject not in participant_ids:
        add_error(
            "VAL-013",
            f"Lab row {row.name + 2}: unknown or missing subject {subject!r}.",
        )

# VAL-014: Visit name allowed
for _, row in lab.iterrows():
    if missing(row["Visit_Name"]) or row["Visit_Name"] not in VALID_VISITS:
        add_error(
            "VAL-014",
            f"Lab row {row.name + 2}: invalid Visit_Name.",
        )

# VAL-015: Test name allowed
for _, row in lab.iterrows():
    if missing(row["Test_Name"]) or row["Test_Name"] not in VALID_TESTS:
        add_error(
            "VAL-015",
            f"Lab row {row.name + 2}: invalid Test_Name "
            f"{row['Test_Name']!r}.",
        )

# VAL-016: Result must be numeric and positive
lab_numeric = pd.to_numeric(
    lab["Result_Value"],
    errors="coerce",
)

for index, row in lab.iterrows():
    value = lab_numeric.loc[index]

    if pd.isna(value) or value <= 0:
        add_error(
            "VAL-016",
            f"Lab row {row.name + 2}: result is missing, non-numeric, "
            "or not positive.",
        )

# VAL-017: Unit allowed
for _, row in lab.iterrows():
    if missing(row["Unit"]) or row["Unit"] not in VALID_UNITS:
        add_error(
            "VAL-017",
            f"Lab row {row.name + 2}: invalid Unit {row['Unit']!r}.",
        )

# VAL-018: Mixed units require review
valid_lab_pairs = lab[
    lab["Subject_ID"].isin(participant_ids)
    & lab["Test_Name"].isin(VALID_TESTS)
    & lab["Unit"].isin(VALID_UNITS)
]

for subject, group in valid_lab_pairs.groupby("Subject_ID"):
    units = group["Unit"].dropna().unique()

    if len(units) > 1:
        add_review(
            "VAL-018",
            f"{subject}: mixed glucose units detected: {list(units)}. "
            "Confirm unit consistency or convert values before comparison.",
        )

# VAL-019: Collection date valid
for _, row in lab.iterrows():
    if pd.isna(row["Collection_Date"]):
        add_error(
            "VAL-019",
            f"Lab row {row.name + 2}: missing or invalid Collection_Date.",
        )

# VAL-020 and VAL-021: Match laboratory records to visits
visit_lookup = {}

for _, row in visits.iterrows():
    subject = row["Subject_ID"]
    visit_name = row["Visit_Name"]

    if (
        missing(subject)
        or missing(visit_name)
        or pd.isna(row["Visit_Date"])
    ):
        continue

    key = (subject, visit_name)

    if key not in visit_lookup:
        visit_lookup[key] = row

for _, row in lab.iterrows():
    subject = row["Subject_ID"]
    visit_name = row["Visit_Name"]

    if missing(subject) or missing(visit_name):
        continue

    key = (subject, visit_name)

    if key not in visit_lookup:
        add_error(
            "VAL-020",
            f"{row['Lab_ID']}: no matching visit for "
            f"{subject} / {visit_name}.",
        )
        continue

    matching_visit = visit_lookup[key]

    if matching_visit["Visit_Status"] != "Completed":
        add_error(
            "VAL-020",
            f"{row['Lab_ID']}: laboratory record is linked to a visit "
            f"with status {matching_visit['Visit_Status']!r}.",
        )

    if (
        pd.notna(row["Collection_Date"])
        and row["Collection_Date"] != matching_visit["Visit_Date"]
    ):
        add_review(
            "VAL-021",
            f"{row['Lab_ID']}: Collection_Date differs from "
            "the matching visit date.",
        )

# VAL-028: Lab_ID required and unique
if lab["Lab_ID"].isna().any():
    add_error(
        "VAL-028",
        "Laboratory: one or more Lab_ID values are missing.",
    )

duplicate_lab_ids = lab[
    lab["Lab_ID"].notna()
    & lab["Lab_ID"].duplicated(keep=False)
]

for lab_id in duplicate_lab_ids["Lab_ID"].unique():
    add_error(
        "VAL-028",
        f"Laboratory: duplicate Lab_ID {lab_id}.",
    )

# Also check duplicate subject + visit + test combinations
duplicate_measurements = lab.duplicated(
    subset=["Subject_ID", "Visit_Name", "Test_Name"],
    keep=False,
)

for _, row in lab[duplicate_measurements].iterrows():
    add_error(
        "VAL-028",
        f"Duplicate measurement: {row['Subject_ID']} / "
        f"{row['Visit_Name']} / {row['Test_Name']}.",
    )


# ============================================================
# 8. ADVERSE EVENT VALIDATION
# ============================================================

print_section("ADVERSE EVENT VALIDATION")

# VAL-022: Subject must exist
for _, row in ae.iterrows():
    subject = row["Subject_ID"]

    if missing(subject) or subject not in participant_ids:
        add_error(
            "VAL-022",
            f"AE row {row.name + 2}: unknown or missing subject {subject!r}.",
        )

# VAL-023: AE_ID required and unique
if ae["AE_ID"].isna().any():
    add_error(
        "VAL-023",
        "Adverse_Events: one or more AE_ID values are missing.",
    )

duplicate_ae_ids = ae[
    ae["AE_ID"].notna()
    & ae["AE_ID"].duplicated(keep=False)
]

for ae_id in duplicate_ae_ids["AE_ID"].unique():
    add_error(
        "VAL-023",
        f"Adverse_Events: duplicate AE_ID {ae_id}.",
    )

# VAL-024: Event description required
for _, row in ae.iterrows():
    if missing(row["Event_Description"]):
        add_error(
            "VAL-024",
            f"AE {row['AE_ID']}: Event_Description is missing.",
        )

# VAL-025: Onset date and study chronology
for _, row in ae.iterrows():
    subject = row["Subject_ID"]
    onset = row["Onset_Date"]

    if pd.isna(onset):
        add_error(
            "VAL-025",
            f"AE {row['AE_ID']}: missing or invalid Onset_Date.",
        )
        continue

    if subject in consent_lookup:
        consent_date = consent_lookup[subject]

        if pd.notna(consent_date) and onset < consent_date:
            add_review(
                "VAL-025",
                f"AE {row['AE_ID']}: onset date is before consent date.",
            )

# VAL-026: Seriousness allowed
for _, row in ae.iterrows():
    if (
        missing(row["Seriousness"])
        or row["Seriousness"] not in VALID_SERIOUSNESS
    ):
        add_error(
            "VAL-026",
            f"AE {row['AE_ID']}: invalid Seriousness value.",
        )

# VAL-027: Outcome optional, but if provided it must be allowed
for _, row in ae.iterrows():
    outcome = row["Outcome"]

    if not missing(outcome) and outcome not in VALID_OUTCOMES:
        add_error(
            "VAL-027",
            f"AE {row['AE_ID']}: invalid Outcome value {outcome!r}.",
        )


# ============================================================
# 9. DISPLAY RESULTS
# ============================================================

print_section("VALIDATION RESULTS")

print(f"Total errors: {len(ERRORS)}")
print(f"Total reviews: {len(REVIEWS)}")

if ERRORS:
    print("\nERROR DETAILS")

    for item in ERRORS:
        print(
            f"[{item['Severity']}] "
            f"[{item['Rule_ID']}] "
            f"{item['Issue']}"
        )
else:
    print("\nNo errors detected.")

if REVIEWS:
    print("\nREVIEW DETAILS")

    for item in REVIEWS:
        print(
            f"[{item['Severity']}] "
            f"[{item['Rule_ID']}] "
            f"{item['Issue']}"
        )
else:
    print("\nNo review items detected.")


# ============================================================
# 10. SAVE VALIDATION REPORT
# ============================================================

report_dir = PROJECT_DIR / "reports"
report_dir.mkdir(parents=True, exist_ok=True)

# Output filename is selected automatically by MODE.
report_file = report_dir / REPORT_FILENAME

error_df = pd.DataFrame(
    ERRORS,
    columns=["Severity", "Rule_ID", "Issue"],
)

review_df = pd.DataFrame(
    REVIEWS,
    columns=["Severity", "Rule_ID", "Issue"],
)

summary_df = pd.DataFrame([
    {"Metric": "Participants", "Value": len(participants)},
    {"Metric": "Visits", "Value": len(visits)},
    {"Metric": "Laboratory", "Value": len(lab)},
    {"Metric": "Adverse Events", "Value": len(ae)},
    {"Metric": "Total Errors", "Value": len(ERRORS)},
    {"Metric": "Total Reviews", "Value": len(REVIEWS)},
])

with pd.ExcelWriter(report_file, engine="openpyxl") as writer:
    summary_df.to_excel(
        writer,
        sheet_name="Summary",
        index=False,
    )

    error_df.to_excel(
        writer,
        sheet_name="Errors",
        index=False,
    )

    review_df.to_excel(
        writer,
        sheet_name="Reviews",
        index=False,
    )

print("\nValidation report saved to:")
print(report_file)

print("\nValidation complete.")
