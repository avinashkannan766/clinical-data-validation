# Clinical Data Management Validation Project

## Overview

This project simulates an entry-level Clinical Data Management (CDM) workflow using a fictional clinical study and entirely synthetic participant data.

The objective is to demonstrate practical skills in clinical-style data review, data-quality validation, issue identification, defect testing, query tracking, and Excel-based reporting using Python.

**Study ID:** CTDM-001

| Dataset domain | Expected records in the clean dataset |
|---|---:|
| Participants | 30 |
| Visits | 90 |
| Laboratory | 90 |
| Adverse Events | 12 |

## Tools Used

- Python
- pandas
- openpyxl
- Microsoft Excel
- Git and GitHub

## Project Workflow

1. Define the fictional study structure and data dictionary.
2. Organize the data into Participants, Visits, Laboratory, and Adverse Events domains.
3. Preserve the raw workbook separately from the corrected and test workbooks.
4. Implement rule-based validation checks in Python.
5. Create a test dataset containing deliberately seeded data-quality defects.
6. Run validation against the clean and test datasets.
7. Compare detected findings with the seeded defect register.
8. Generate Excel reports for validation results and defect detection.
9. Create a simulated query log and change log to demonstrate issue tracking and documentation.

## Validation Coverage

The validation script checks for issues including:

- Missing and duplicate participant identifiers
- Missing or invalid participant values
- Invalid visit dates and visit-window deviations
- Laboratory records linked to unknown participants or visits
- Invalid laboratory results and units
- Duplicate laboratory records
- Laboratory records linked to visits marked as missed
- Adverse events linked to unknown participants
- Adverse-event date and chronology concerns
- Review-level findings requiring human assessment

Findings are classified as errors or review items. Review items are not automatically treated as confirmed errors.

## Results

### Clean workbook

The corrected workbook passed the implemented validation checks:

- Structural validation: PASS for all four domains
- Validation errors: 0
- Review items: 0

### Test workbook

A separate test workbook contains 12 deliberately seeded defects.

The defect-detection summary reported that all 12 seeded defects were detected by their expected validation rules. Additional findings may arise when one data-quality problem triggers multiple checks.

The latest test validation produced 19 errors and 5 review items, for 24 findings in total.

### Query and change logs

The query-management workbook contains:

- 24 logged validation findings
- Query IDs and status tracking
- Resolution notes and training notes
- Three simulated change-log entries linked to selected queries

The change log compares selected test values with values in the clean reference workbook. These are educational examples, not live clinical corrections or evidence of a real query-resolution process.

## Setup and Installation

### Prerequisites

- Python installed on your computer
- Git, if cloning the repository
- Microsoft Excel or compatible spreadsheet software to inspect the reports

### 1. Clone the repository

```powershell
git clone https://github.com/avinashkannan766/clinical-data-validation.git
cd clinical-data-validation
```

Alternatively, download the repository as a ZIP file from GitHub and extract it.

### 2. Install dependencies

```powershell
python -m pip install -r requirements.txt
```

The `requirements.txt` file lists the external Python packages needed by the project.

## Running the Validation

Run these commands from the project root directory.

### 1. Validate the clean dataset

```powershell
python .\scripts\validate_data.py clean
```

This runs the validation checks against the corrected workbook and generates `reports/validation_report.xlsx`.

### 2. Prepare the test dataset

Use the project's test-data generation script:

```powershell
python .\scripts\create_test_defects.py
```

This step is intended to prepare the synthetic test workbook with seeded defects.

### 3. Validate the test dataset

```powershell
python .\scripts\validate_data.py test
```

This generates `reports/test_validation_report.xlsx`.

### 4. Generate the defect-detection summary

```powershell
python .\scripts\create_defect_summary.py
```

This compares detected findings with the seeded defect register and generates `reports/defect_detection_summary.xlsx`.

### 5. Generate the query log

```powershell
python .\scripts\create_query_log.py
```

This generates `reports/query_log.xlsx`, including the query log, change log, and summary sheets.

**Note:** The query-log script uses the existing test-validation report and the test and corrected workbooks. Run the preceding steps first so those inputs are available and up to date.

## Repository Structure

```text
clinical-data-validation/
├── data/
│   ├── clinical_study_raw.xlsx
│   ├── clinical_study_corrected.xlsx
│   └── clinical_study_test.xlsx
├── docs/
│   ├── data_dictionary.xlsx
│   ├── seeded_defect_register.xlsx
│   └── study_specification.md
├── reports/
│   ├── validation_report.xlsx
│   ├── test_validation_report.xlsx
│   ├── defect_detection_summary.xlsx
│   └── query_log.xlsx
├── scripts/
│   ├── validate_data.py
│   ├── inspect_data.py
│   ├── fix_visit_dates.py
│   ├── create_test_defects.py
│   ├── create_defect_summary.py
│   └── create_query_log.py
├── requirements.txt
└── README.md
```

Generated reports may need to be recreated after the underlying datasets or scripts change. The repository may also contain additional local files that are intentionally excluded from version control.

## Important Limitations

This is an educational portfolio project using entirely synthetic data. It is not a real clinical trial and is not a validated clinical data management system.

The project does not claim production EDC experience, regulatory compliance, formal GCP/CDISC certification, or experience managing actual clinical-trial participants. Simulated query statuses and reference comparisons demonstrate documentation concepts only.

## Skills Demonstrated

- Structured clinical-style data handling
- Data cleaning and rule-based validation
- Python scripting with pandas
- Test-data preparation and seeded-defect detection
- Error and review-level finding classification
- Query tracking and simulated change documentation
- Excel-based reporting
- Git/GitHub version control
- Technical documentation and reproducible workflows
