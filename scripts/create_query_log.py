
from pathlib import Path
import re
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment

PROJECT_DIR = Path(__file__).resolve().parent.parent

REPORT_FILE = PROJECT_DIR / "reports" / "test_validation_report.xlsx"
TEST_FILE = PROJECT_DIR / "data" / "clinical_study_test.xlsx"
CLEAN_FILE = PROJECT_DIR / "data" / "clinical_study_corrected.xlsx"
OUTPUT_FILE = PROJECT_DIR / "reports" / "query_log.xlsx"

# These are illustrative statuses for this synthetic training project.
# They do not represent real site communications or clinical decisions.

if not REPORT_FILE.exists():
    raise FileNotFoundError(f"Validation report not found: {REPORT_FILE}")

if not TEST_FILE.exists() or not CLEAN_FILE.exists():
    raise FileNotFoundError("Test or corrected workbook is missing.")

errors = pd.read_excel(REPORT_FILE, sheet_name="Errors")
reviews = pd.read_excel(REPORT_FILE, sheet_name="Reviews")

test_sheets = pd.read_excel(TEST_FILE, sheet_name=None)
clean_sheets = pd.read_excel(CLEAN_FILE, sheet_name=None)

findings = pd.concat(
    [
        errors.assign(Source_Sheet="Errors"),
        reviews.assign(Source_Sheet="Reviews"),
    ],
    ignore_index=True,
)

test_participants = test_sheets["Participants"]
clean_participants = clean_sheets["Participants"]
test_visits = test_sheets["Visits"]
clean_visits = clean_sheets["Visits"]
test_lab = test_sheets["Laboratory"]
clean_lab = clean_sheets["Laboratory"]

DOMAIN_BY_RULE = {
    "VAL-001": "Participants",
    "VAL-002": "Participants",
    "VAL-003": "Participants",
    "VAL-004": "Participants",
    "VAL-005": "Participants",
    "VAL-006": "Participants",
    "VAL-007": "Visits",
    "VAL-008": "Visits",
    "VAL-009": "Visits",
    "VAL-010": "Visits",
    "VAL-011": "Visits",
    "VAL-012": "Visits",
    "VAL-013": "Laboratory",
    "VAL-014": "Laboratory",
    "VAL-015": "Laboratory",
    "VAL-016": "Laboratory",
    "VAL-017": "Laboratory",
    "VAL-018": "Laboratory",
    "VAL-019": "Laboratory",
    "VAL-020": "Laboratory",
    "VAL-021": "Laboratory",
    "VAL-022": "Adverse_Events",
    "VAL-023": "Adverse_Events",
    "VAL-024": "Adverse_Events",
    "VAL-025": "Adverse_Events",
    "VAL-026": "Adverse_Events",
    "VAL-027": "Adverse_Events",
    "VAL-028": "Laboratory",
}


def is_missing(value):
    return pd.isna(value) or (
        isinstance(value, str) and value.strip() == ""
    )


def display_value(value):
    if is_missing(value):
        return "<MISSING>"
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    return str(value)


def row_number_from_issue(issue):
    match = re.search(r"\brow (\d+)\b", str(issue), re.IGNORECASE)
    return int(match.group(1)) if match else None


def reference_change(rule_id, issue):
    """Return a documented synthetic test-to-reference difference."""
    issue = str(issue)

    # DEF-001: missing age for CT005
    if rule_id == "VAL-002" and "CT005" in issue:
        subject = "CT005"
        old = test_participants.loc[
            test_participants["Subject_ID"] == subject, "Age"
        ]
        new = clean_participants.loc[
            clean_participants["Subject_ID"] == subject, "Age"
        ]
        if len(old) and len(new):
            return {
                "domain": "Participants",
                "record_id": subject,
                "field": "Age",
                "old": old.iloc[0],
                "new": new.iloc[0],
                "reason": "Synthetic missing-value example compared with the clean reference workbook.",
            }

    # Row-based laboratory findings: map the reported Excel row to Lab_ID.
    if rule_id in {"VAL-016", "VAL-017"}:
        excel_row = row_number_from_issue(issue)
        if excel_row is not None:
            index = excel_row - 2
            if 0 <= index < len(test_lab):
                lab_id = test_lab.iloc[index]["Lab_ID"]
                test_row = test_lab.loc[test_lab["Lab_ID"] == lab_id]
                clean_row = clean_lab.loc[clean_lab["Lab_ID"] == lab_id]

                field = "Result_Value" if rule_id == "VAL-016" else "Unit"

                if len(test_row) and len(clean_row):
                    old = test_row.iloc[0][field]
                    new = clean_row.iloc[0][field]
                    if display_value(old) != display_value(new):
                        return {
                            "domain": "Laboratory",
                            "record_id": lab_id,
                            "field": field,
                            "old": old,
                            "new": new,
                            "reason": "Synthetic test value compared with the corresponding clean reference record.",
                        }

    # DEF-011: CT010 Baseline visit date
    if rule_id == "VAL-012" and "CT010 Baseline" in issue:
        subject = "CT010"
        visit = "Baseline"

        test_row = test_visits.loc[
            (test_visits["Subject_ID"] == subject)
            & (test_visits["Visit_Name"] == visit)
        ]
        clean_row = clean_visits.loc[
            (clean_visits["Subject_ID"] == subject)
            & (clean_visits["Visit_Name"] == visit)
        ]

        if len(test_row) and len(clean_row):
            return {
                "domain": "Visits",
                "record_id": f"{subject} / {visit}",
                "field": "Visit_Date",
                "old": test_row.iloc[0]["Visit_Date"],
                "new": clean_row.iloc[0]["Visit_Date"],
                "reason": "Synthetic visit-date example compared with the clean reference workbook.",
            }

    # DEF-003: invalid visit date; map by the Excel row number.
    if rule_id == "VAL-009":
        excel_row = row_number_from_issue(issue)
        if excel_row is not None:
            index = excel_row - 2
            if 0 <= index < len(test_visits):
                row = test_visits.iloc[index]
                subject = row["Subject_ID"]
                visit = row["Visit_Name"]

                clean_row = clean_visits.loc[
                    (clean_visits["Subject_ID"] == subject)
                    & (clean_visits["Visit_Name"] == visit)
                ]

                if len(clean_row):
                    return {
                        "domain": "Visits",
                        "record_id": f"{subject} / {visit}",
                        "field": "Visit_Date",
                        "old": row["Visit_Date"],
                        "new": clean_row.iloc[0]["Visit_Date"],
                        "reason": "Synthetic invalid-date example compared with the clean reference workbook.",
                    }

    return None


def scenario_for(rule_id, issue):
    """Create a small, explicitly simulated query lifecycle."""
    issue = str(issue)

    if rule_id == "VAL-002" and "CT005" in issue:
        return (
            "Closed",
            "Simulated response: the missing value was reviewed against the clean synthetic reference.",
            "Closed for demonstration after comparison with the reference workbook.",
        )

    if rule_id == "VAL-017" and "mg/L" in issue:
        return (
            "Resolved",
            "Simulated response: the unit entry was reviewed against the clean synthetic reference.",
            "Resolved for demonstration; see the linked change log.",
        )

    if rule_id == "VAL-016":
        return (
            "Answered",
            "Simulated response: a reference value is available for review.",
            "Awaiting independent verification in this training scenario.",
        )

    if rule_id == "VAL-009":
        return (
            "Open",
            "",
            "Date requires review before a resolution can be recorded.",
        )

    if rule_id == "VAL-012" and "CT010 Baseline" in issue:
        return (
            "Resolved",
            "Simulated response: the baseline date was compared with the clean reference.",
            "Resolved for demonstration; see the linked change log.",
        )

    if rule_id == "VAL-025" and "AE001" in issue:
        return (
            "Open",
            "",
            "Chronology requires review; do not automatically change adverse-event data.",
        )

    return (
        "Open",
        "",
        "Review the finding and document the response before changing status.",
    )


query_rows = []
change_rows = []
query_number = 1

for _, finding in findings.iterrows():
    rule_id = str(finding["Rule_ID"])
    issue = str(finding["Issue"])
    severity = str(finding["Severity"])

    query_id = f"QRY-{query_number:04d}"
    query_number += 1

    status, response, resolution = scenario_for(rule_id, issue)
    change = reference_change(rule_id, issue)

    # Only create a change-log row when the query is resolved/closed
    # and an actual test-to-reference difference is available.
    linked_change_id = ""

    if status in {"Resolved", "Closed"} and change is not None:
        linked_change_id = f"CHG-{len(change_rows) + 1:04d}"

        change_rows.append({
            "Change_ID": linked_change_id,
            "Query_ID": query_id,
            "Domain": change["domain"],
            "Record_ID": change["record_id"],
            "Field_Name": change["field"],
            "Original_Test_Value": display_value(change["old"]),
            "Reference_Value": display_value(change["new"]),
            "Reason": change["reason"],
            "Evidence_Source": "clinical_study_corrected.xlsx",
            "Change_Type": "Simulated comparison; not a live clinical correction",
        })

    query_rows.append({
        "Query_ID": query_id,
        "Source_Report": "test_validation_report.xlsx",
        "Domain": DOMAIN_BY_RULE.get(rule_id, "Review required"),
        "Severity": severity,
        "Rule_ID": rule_id,
        "Finding": issue,
        "Query_Text": f"Synthetic query: please review the finding under {rule_id}.",
        "Status": status,
        "Simulated_Response": response,
        "Resolution_Notes": resolution,
        "Linked_Change_ID": linked_change_id,
        "Training_Note": "Illustrative synthetic scenario; not a real site query.",
    })


query_df = pd.DataFrame(query_rows)
change_df = pd.DataFrame(change_rows)

status_order = ["Open", "Answered", "Resolved", "Closed"]
summary_df = pd.DataFrame([
    {"Metric": "Total findings used", "Value": len(query_df)},
    {"Metric": "Open", "Value": int((query_df["Status"] == "Open").sum())},
    {"Metric": "Answered", "Value": int((query_df["Status"] == "Answered").sum())},
    {"Metric": "Resolved", "Value": int((query_df["Status"] == "Resolved").sum())},
    {"Metric": "Closed", "Value": int((query_df["Status"] == "Closed").sum())},
    {"Metric": "Linked change-log entries", "Value": len(change_df)},
])

OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    query_df.to_excel(writer, sheet_name="Query_Log", index=False)
    change_df.to_excel(writer, sheet_name="Change_Log", index=False)
    summary_df.to_excel(writer, sheet_name="Summary", index=False)

    for sheet in writer.sheets.values():
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions

        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.alignment = Alignment(wrap_text=True, vertical="top")

        for column_cells in sheet.columns:
            max_length = max(
                len(str(cell.value or "")) for cell in column_cells
            )
            sheet.column_dimensions[
                column_cells[0].column_letter
            ].width = min(max(max_length + 2, 12), 48)

print("=" * 60)
print("QUERY LOG CREATED")
print("=" * 60)
print(f"Findings processed: {len(query_df)}")
print(f"Linked change entries: {len(change_df)}")
print(f"Output: {OUTPUT_FILE}")
print("\nStatus counts:")
print(query_df["Status"].value_counts().to_string())
