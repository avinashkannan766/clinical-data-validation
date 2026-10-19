
from pathlib import Path
import pandas as pd


# ---------------------------------------------------------
# File paths
# ---------------------------------------------------------

report_file = Path("reports/test_validation_report.xlsx")
output_file = Path("reports/defect_detection_summary.xlsx")


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

    for _, row in sheets[sheet_name].iterrows():
        text = " | ".join(
            str(value)
            for value in row.tolist()
            if pd.notna(value)
        )

        if text.strip():
            findings.append({
                "Finding_Type": sheet_name,
                "Finding_Text": text
            })


# ---------------------------------------------------------
# Define expected defect-to-finding matches
# ---------------------------------------------------------

defects = [
    {
        "Defect_ID": "DEF-001",
        "Description": "CT005 has missing Age",
        "Expected_Rule": "VAL-002",
        "Keywords": ["CT005", "Age"]
    },
    {
        "Defect_ID": "DEF-002",
        "Description": "Duplicate participant ID CT005",
        "Expected_Rule": "VAL-001",
        "Keywords": ["CT005", "duplicate"]
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
        "Keywords": ["CT007", "Week 4"]
    },
    {
        "Defect_ID": "DEF-005",
        "Description": "LAB005 references unknown subject CT999",
        "Expected_Rule": "VAL-013",
        "Keywords": ["CT999"]
    },
    {
        "Defect_ID": "DEF-006",
        "Description": "Duplicate laboratory record LAB006",
        "Expected_Rule": "VAL-028",
        "Keywords": ["LAB006"]
    },
    {
        "Defect_ID": "DEF-007",
        "Description": "LAB007 has unsupported unit mg/L",
        "Expected_Rule": "VAL-017",
        "Keywords": ["mg/L"]
    },
    {
        "Defect_ID": "DEF-008",
        "Description": "LAB008 has a nonnumeric result",
        "Expected_Rule": "VAL-016",
        "Keywords": ["Lab row 9"]
    },
    {
        "Defect_ID": "DEF-009",
        "Description": "CT008 Week 4 marked Missed despite lab data",
        "Expected_Rule": "VAL-020",
        "Keywords": ["Missed"]
    },
    {
        "Defect_ID": "DEF-010",
        "Description": "AE001 onset before consent",
        "Expected_Rule": "VAL-025",
        "Keywords": ["AE001"]
    },
    {
        "Defect_ID": "DEF-011",
        "Description": "CT010 Baseline before expected Day 0",
        "Expected_Rule": "VAL-012",
        "Keywords": ["CT010", "Baseline"]
    },
    {
        "Defect_ID": "DEF-012",
        "Description": "AE012 references unknown subject CT999",
        "Expected_Rule": "VAL-022",
        "Keywords": ["CT999"]
    }
]


# ---------------------------------------------------------
# Compare expected defects against actual findings
# ---------------------------------------------------------

summary = []

for defect in defects:
    matched = []

    for finding in findings:
        text = finding["Finding_Text"]

        rule_found = defect["Expected_Rule"] in text
        keywords_found = all(
            keyword.lower() in text.lower()
            for keyword in defect["Keywords"]
        )

        if rule_found and keywords_found:
            matched.append(text)

    detected = len(matched) > 0

    summary.append({
        "Defect_ID": defect["Defect_ID"],
        "Description": defect["Description"],
        "Expected_Rule": defect["Expected_Rule"],
        "Status": "PASS" if detected else "NOT DETECTED",
        "Matched_Finding": (
            " || ".join(matched)
            if matched
            else "No matching finding in validation report"
        )
    })


# ---------------------------------------------------------
# Save summary workbook
# ---------------------------------------------------------

summary_df = pd.DataFrame(summary)

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    summary_df.to_excel(
        writer,
        sheet_name="Defect_Summary",
        index=False
    )

    counts = (
        summary_df["Status"]
        .value_counts()
        .rename_axis("Status")
        .reset_index(name="Count")
    )

    counts.to_excel(
        writer,
        sheet_name="Summary",
        index=False
    )

print("=" * 60)
print("DEFECT DETECTION SUMMARY CREATED")
print("=" * 60)

print(summary_df[["Defect_ID", "Expected_Rule", "Status"]].to_string(
    index=False
))

print()
print(f"Total seeded defects: {len(summary_df)}")
print(f"Detected: {(summary_df['Status'] == 'PASS').sum()}")
print(f"Not detected: {(summary_df['Status'] != 'PASS').sum()}")
print(f"Saved to: {output_file.resolve()}")
