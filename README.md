# Clinical Data Management Validation

A Python-based clinical data validation project that checks a synthetic clinical study dataset for missing values, invalid records, inconsistent dates, duplicate records, and other data quality issues. The project generates Excel validation reports and uses automated tests to verify expected behavior.

**Project type:** Synthetic training and portfolio project
**Language:** Python
**Tools:** pandas, openpyxl, pytest, Excel, Git, GitHub

---

## Project Overview

Clinical Data Management involves maintaining the quality, consistency, and reliability of clinical study data.

This project demonstrates a basic data validation workflow using a synthetic clinical study dataset. It checks data against predefined validation rules, separates errors from review findings, and produces Excel reports for further investigation.

The project includes both a clean dataset and an intentionally defective test dataset to demonstrate how the validation script identifies data quality issues.

**Important:** All study data used in this project are synthetic and created for learning and demonstration. This project does not represent real clinical trial data or actual clinical research experience.

---

## Objectives

- Validate the structure and contents of clinical study datasets.
- Identify missing or invalid participant information.
- Check visit dates and expected visit windows.
- Detect laboratory data inconsistencies and duplicate records.
- Identify adverse event data quality issues.
- Separate errors requiring correction from findings requiring review.
- Generate Excel reports containing validation results.
- Verify expected validation behavior through automated tests.

---

## Tools and Technologies

| Tool | Purpose |
|---|---|
| Python | Validation logic and workflow automation |
| pandas | Data handling and validation |
| openpyxl | Reading and writing Excel workbooks |
| pytest | Automated testing |
| Microsoft Excel | Reviewing synthetic datasets and reports |
| Git | Version control |
| GitHub | Project hosting and portfolio presentation |

---

## Dataset Description

The synthetic study uses the study identifier `CTDM-001`.

The dataset contains four main data tables.

| Dataset | Clean records | Purpose |
|---|---:|---|
| Participants | 30 | Participant identifiers, demographics, consent, and treatment assignment |
| Visits | 90 | Baseline, Week 4, and Week 8 visit records |
| Laboratory | 90 | Synthetic glucose measurements and collection dates |
| Adverse Events | 12 | Synthetic adverse event information |

The intentionally defective laboratory dataset contains 91 records because it includes a seeded duplicate record.

The two treatment arms contain 15 participants each:

- Treatment A: 15 participants
- Control B: 15 participants

---

## Validation Workflow

The validation process follows these steps:

1. Select the clean or test execution mode.
2. Load the corresponding synthetic Excel workbook.
3. Check that the expected worksheets and required columns are present.
4. Validate participant, visit, laboratory, and adverse event records.
5. Record detected issues with a rule identifier and descriptive message.
6. Separate errors from review findings.
7. Generate an Excel validation report.
8. Run automated tests to check the expected outputs.

### Execution Modes

**Clean mode**

Uses the corrected synthetic workbook and generates `reports/validation_report.xlsx`.

**Test mode**

Uses an intentionally defective synthetic workbook and generates `reports/test_validation_report.xlsx`.

---

## Validation Rules

The validation script includes checks covering the following areas.

### 1. Structural Validation

- Required worksheets
- Required columns

### 2. Participant Validation

- Duplicate participant identifiers
- Missing or invalid age values
- Participant identifiers referenced by other datasets

### 3. Visit Validation

- Valid visit dates
- Expected visit windows
- Participant references
- Visit status consistency

### 4. Laboratory Validation

- Valid laboratory results
- Allowed measurement units
- Valid participant references
- Matching visit records
- Collection date consistency
- Duplicate laboratory identifiers
- Duplicate measurements
- Laboratory records linked to missed visits

### 5. Adverse Event Validation

- Valid participant references
- Event onset date consistency with consent
- Valid seriousness values
- Valid outcome values when provided

The validation rules use rule identifiers such as `VAL-001`, `VAL-009`, and `VAL-028` to make findings easier to trace.

---

## Error vs Review Findings

The project distinguishes between errors and review findings.

**Errors** identify records that violate a defined validation rule and require investigation or correction.

**Review findings** identify records that may be inconsistent with the study specification and require assessment rather than automatic rejection.

For example:

- A missing participant age is recorded as an error.
- A visit outside the expected date window is recorded as a review finding.
- An adverse event onset date before consent is flagged for review.

These classifications are based on the rules implemented in this synthetic project. They should not be interpreted as universal clinical data management requirements.

---

## Validation Results

The current synthetic datasets produce the following results.

| Metric | Clean dataset | Test dataset |
|---|---:|---:|
| Participants | 30 | 30 |
| Visits | 90 | 90 |
| Laboratory records | 90 | 91 |
| Adverse events | 12 | 12 |
| Errors | 0 | 19 |
| Review findings | 0 | 5 |

The clean dataset passes the implemented validation checks without errors or review findings.

The intentionally defective dataset produces 19 errors and 5 review findings. Some findings are secondary effects of seeded defects. For example, changing a participant identifier can cause related visit and laboratory records to be flagged as referencing an unknown participant.

These results demonstrate the current rules against the prepared synthetic data; they do not prove that the validator detects every possible data quality issue.

---

## Seeded Defect Detection

The project includes a synthetic test workbook containing 12 intentionally seeded defects.

The defect detection summary compares each seeded defect with its expected validation rule and checks whether it matches exactly one finding.

Current summary:

- Seeded defects: 12
- Uniquely matched defects: 12
- Not detected: 0
- Multiple matches in the summary: 0

This summary verifies matching against the project's predefined defect register. It is not independent proof of complete defect detection, and secondary findings may occur when one defect affects related records.

---

## Automated Testing

The project includes automated tests using `pytest`.

The test suite checks:

- Missing-value detection
- Valid and invalid date parsing
- Behavior when a date column is absent
- Required-column validation
- Clean validation report output
- Expected findings in the defective-data report
- Unique matching of the 12 seeded defects

### Test Results

**Last recorded test result: 9 passed.**

The suite contains six helper-function tests and three report-output tests.

Run the tests from the project root:

```powershell
python -m pytest .\tests -v
```

The tests verify selected helper functions and generated report outputs. They do not yet test every validation rule directly.

---

## Repository Structure

```text
clinical-data-validation/
â”‚
â”œâ”€â”€ data/
â”‚   â”œâ”€â”€ clinical_study_raw.xlsx
â”‚   â”œâ”€â”€ clinical_study_corrected.xlsx
â”‚   â””â”€â”€ clinical_study_test.xlsx
â”‚
â”œâ”€â”€ docs/
â”‚   â”œâ”€â”€ data_dictionary.xlsx
â”‚   â”œâ”€â”€ seeded_defect_register.xlsx
â”‚   â””â”€â”€ study_specification.md
â”‚
â”œâ”€â”€ reports/
â”‚   â”œâ”€â”€ validation_report.xlsx
â”‚   â”œâ”€â”€ test_validation_report.xlsx
â”‚   â”œâ”€â”€ defect_detection_summary.xlsx
â”‚   â””â”€â”€ query_log.xlsx
â”‚
â”œâ”€â”€ scripts/
â”‚   â”œâ”€â”€ validate_data.py
â”‚   â”œâ”€â”€ validation_helpers.py
â”‚   â”œâ”€â”€ create_test_defects.py
â”‚   â”œâ”€â”€ create_defect_summary.py
â”‚   â””â”€â”€ create_query_log.py
â”‚
â”œâ”€â”€ tests/
â”‚   â”œâ”€â”€ test_validation_helpers.py
â”‚   â””â”€â”€ test_validation_outputs.py
â”‚
â”œâ”€â”€ .gitignore
â”œâ”€â”€ requirements.txt
â””â”€â”€ README.md
```

This structure highlights the main project files. Other scripts may also exist in the `scripts/` directory.

---

## Installation and Usage

### Prerequisites

- Python installed
- Git installed
- A terminal such as PowerShell

### 1. Clone the repository

```powershell
git clone https://github.com/avinashkannan766/clinical-data-validation.git
cd clinical-data-validation
```

### 2. Install dependencies

```powershell
python -m pip install -r requirements.txt
```

### 3. Run clean validation

```powershell
python .\scripts\validate_data.py clean
```

This validates the corrected synthetic dataset and generates:

```text
reports/validation_report.xlsx
```

### 4. Run test validation

```powershell
python .\scripts\validate_data.py test
```

This validates the intentionally defective dataset and generates:

```text
reports/test_validation_report.xlsx
```

### 5. Run automated tests

```powershell
python -m pytest .\tests -v
```

---

## Limitations

- All study data are synthetic.
- The project is intended for learning and portfolio demonstration.
- The validation rules are illustrative and based on the project's study specification.
- The project does not implement a production electronic data capture system.
- It does not claim CDISC compliance, regulatory validation, or GCP certification.
- The current automated tests cover selected helpers and report outputs rather than every validation rule.
- Seeded defect matching depends on predefined identifiers, expected rule IDs, and matching criteria.
- A zero-error report means no issues were detected by the implemented rules; it does not guarantee that the dataset is free of all errors.

---

## Skills Demonstrated

- Python-based data validation
- Excel dataset handling
- Data quality checks
- Missing-value and duplicate detection
- Date validation and consistency checks
- Error and review classification
- Validation report generation
- Automated testing with pytest
- Version control with Git
- Project documentation

---

## Future Improvements

Potential future enhancements include:

- Integrating shared helper functions directly into the main validator
- Adding direct tests for individual validation rules
- Improving the traceability between defects, findings, and query resolution
- Adding more test cases for unexpected or malformed input data
- Improving report formatting and summary visualizations

---

## Author

Created as a self-directed portfolio project to demonstrate foundational Python, data validation, and automated testing skills relevant to entry-level Clinical Data Management roles.

**GitHub repository:**
https://github.com/avinashkannan766/clinical-data-validation
