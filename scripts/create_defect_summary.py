
from pathlib import Path
import pandas as pd


# ---------------------------------------------------------
# File paths
# ---------------------------------------------------------

PROJECT_DIR = Path(__file__).resolve().parent.parent

report_file = PROJECT_DIR / "reports" / "test_validation_report.xlsx"
output_file = PROJECT_DIR / "reports" / "defect_detection_summary.xlsx"


if not report_file.exists():
    raise FileNotFoundError(
        f"Validation report not found: {report_file.resolve()}"
    )


# ---------------------------------------------------------
# Read detected errors and reviews
# ---------------------------------------------------------

sheets = pd.read_excel(report_file, sheet_name=None)

findings = []

for sheet_name in ["Errors", "Reviews"]:
    if sheet_name not in sheets:
        continue

    df = sheets[sheet_name]

    required_columns = {"Rule_ID", "Issue"}
    if not required_columns.issubset(df.columns):
        raise ValueError(
            f"Sheet '{sheet_name}' must contain columns: "
            f"{sorted(required_columns)}"
        )

    for _, row in df.iterrows():
        rule_id = str(row["Rule_ID"]).strip()
        issue = str(row["Issue"]).strip()

        if rule_id and issue and rule_id.lower() != "nan":
            findings.append({
                "Finding_Type": sheet_name,
                "Rule_ID": rule_id,
                "Finding_Text": issue
            })


# ---------------------------------------------------------
# Define expected defect-to-finding matches
# ---------------------------------------------------------
# Each defect requires:
# 1. An exact validation rule ID
# 2. Record-specific text identifying the intended finding
#
# These checks test detection of seeded synthetic defects.
# They do not establish clinical source-data verification.


defects = [
    {
        "Defect_ID": "DEF-001",
        "Description": "CT005 has missing Age",
        "Expected_Rule": "VAL-002",
        "Keywords": ["CT005", "Age is missing"]
    },
    {
        "Defect_ID": "DEF-002",
        "Description": "Duplicate participant ID CT005",
        "Expected_Rule": "VAL-001",
        "Keywords": ["duplicate Subject_ID CT005"]
    },
    {
        "Defect_ID": "DEF-003",
        "Description": "CT004 Week 4 has invalid visit date",
        "Expected_Rule": "VAL-009",
        "Keywords": ["Visit row 12", "Visit_Date"]
    },
    {
        "Defect_ID": "DEF-004",
        "Description": "CT007 Week 4 outside visit window",
        "Expected_Rule": "VAL-012",
        "Keywords": ["CT007 Week 4", "Day 40"]
    },
    {
        "Defect_ID": "DEF-005",
        "Description": "LAB005 references unknown subject CT999",
        "Expected_Rule": "VAL-013",
        "Keywords": ["Lab row 6", "CT999"]
    },
    {
        "Defect_ID": "DEF-006",
        "Description": "Duplicate laboratory record LAB006",
        "Expected_Rule": "VAL-028",
        "Keywords": ["duplicate Lab_ID LAB006"]
    },
    {
        "Defect_ID": "DEF-007",
        "Description": "LAB007 has unsupported unit mg/L",
        "Expected_Rule": "VAL-017",
        "Keywords": ["Lab row 8", "mg/L"]
    },
    {
        "Defect_ID": "DEF-008",
        "Description": "LAB008 has a nonnumeric result",
        "Expected_Rule": "VAL-016",
        "Keywords": ["Lab row 9", "non-numeric"]
    },
    {
        "Defect_ID": "DEF-009",
        "Description": "CT008 Week 4 marked Missed despite lab data",
        "Expected_Rule": "VAL-020",
        "Keywords": ["LAB023", "Missed"]
    },
    {
        "Defect_ID": "DEF-010",
        "Description": "AE001 onset before consent",
        "Expected_Rule": "VAL-025",
        "Keywords": ["AE001", "before consent"]
    },
    {
        "Defect_ID": "DEF-011",
        "Description": "CT010 Baseline before expected Day 0",
        "Expected_Rule": "VAL-012",
        "Keywords": ["CT010 Baseline", "Day -1"]
    },
    {
        "Defect_ID": "DEF-012",
        "Description": "AE012 references unknown subject CT999",
        "Expected_Rule": "VAL-022",
        "Keywords": ["AE row 13", "CT999"]
    }
]


# ---------------------------------------------------------
# Match each expected defect to validation findings
# ---------------------------------------------------------

summary = []

for defect in defects:
    matched = []

    for finding in findings:
        # Match the rule ID exactly, not as a substring.
        rule_found = (
            finding["Rule_ID"] == defect["Expected_Rule"]
        )

        text = finding["Finding_Text"]

        # All record-specific keywords must appear in the same
        # finding. Case-insensitive comparison is intentional.
        keywords_found = all(
            keyword.casefold() in text.casefold()
            for keyword in defect["Keywords"]
        )

        if rule_found and keywords_found:
            matched.append(finding)

    if len(matched) == 1:
        status = "PASS"
        matched_text = matched[0]["Finding_Text"]
        match_count = 1
    elif len(matched) == 0:
        status = "NOT DETECTED"
        matched_text = "No exact matching finding in validation report"
        match_count = 0
    else:
        status = "REVIEW - MULTIPLE MATCHES"
        matched_text = " || ".join(
            finding["Finding_Text"] for finding in matched
        )
        match_count = len(matched)

    summary.append({
        "Defect_ID": defect["Defect_ID"],
        "Description": defect["Description"],
        "Expected_Rule": defect["Expected_Rule"],
        "Expected_Keywords": " | ".join(defect["Keywords"]),
        "Match_Count": match_count,
        "Status": status,
        "Matched_Finding": matched_text
    })


# ---------------------------------------------------------
# Create summary counts
# ---------------------------------------------------------

summary_df = pd.DataFrame(summary)

status_counts = (
    summary_df["Status"]
    .value_counts()
    .rename_axis("Status")
    .reset_index(name="Count")
)

detected_count = int((summary_df["Status"] == "PASS").sum())
not_detected_count = int(
    (summary_df["Status"] == "NOT DETECTED").sum()
)
review_count = int(
    (summary_df["Status"] == "REVIEW - MULTIPLE MATCHES").sum()
)


# ---------------------------------------------------------
# Save Excel report
# ---------------------------------------------------------

output_file.parent.mkdir(parents=True, exist_ok=True)

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    summary_df.to_excel(
        writer,
        sheet_name="Defect_Summary",
        index=False
    )

    status_counts.to_excel(
        writer,
        sheet_name="Summary",
        index=False
    )

    # Preserve the validation evidence used by this audit.
    findings_df = pd.DataFrame(findings)
    findings_df.to_excel(
        writer,
        sheet_name="Validation_Findings",
        index=False
    )


# ---------------------------------------------------------
# Display results
# ---------------------------------------------------------

print("=" * 70)
print("DEFECT DETECTION SUMMARY CREATED")
print("=" * 70)

print(
    summary_df[
        ["Defect_ID", "Expected_Rule", "Match_Count", "Status"]
    ].to_string(index=False)
)

print()
print(f"Total seeded defects: {len(summary_df)}")
print(f"Uniquely matched: {detected_count}")
print(f"Not detected: {not_detected_count}")
print(f"Multiple matches requiring review: {review_count}")
print(f"Saved to: {output_file.resolve()}")

if not_detected_count > 0 or review_count > 0:
    print(
        "\nAUDIT WARNING: Inspect unmatched or ambiguous defects "
        "before reporting the detection rate."
    )
else:
    print(
        "\nAll seeded defects have one matching validation finding "
        "under the configured rule and keyword checks."
    )
