# Clinical Data Management Validation Project

## Overview

This project simulates an entry-level Clinical Data Management (CDM) workflow using a fictional clinical study and synthetic participant data.

The objective is to demonstrate practical skills in clinical data review, data-quality validation, issue identification, and validation reporting using Python and Microsoft Excel.

**Study ID:** CTDM-001  
**Participants:** 30  
**Visits:** 90  
**Laboratory records:** 90  
**Adverse-event records:** 12

## Tools Used

- Python
- pandas
- openpyxl
- Microsoft Excel

## Project Workflow

1. Define the fictional study structure and data dictionary.
2. Organize data into Participants, Visits, Laboratory, and Adverse Events domains.
3. Preserve the original raw workbook and prepare a separate corrected workbook.
4. Implement validation rules in Python.
5. Generate a test workbook containing deliberately introduced data-quality defects.
6. Run validation and compare detected findings with the seeded defect register.
7. Export validation results to Excel for review.

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
- Review-level findings that require human assessment

Findings are classified as errors or review items. Review items are not automatically treated as confirmed errors.

## Results

### Clean workbook

The corrected workbook passed the implemented validation checks:

- Structural validation: PASS for all four domains
- Validation errors: 0
- Review items: 0

### Test workbook

A separate test workbook was created with 12 deliberately seeded defects. The defect-detection summary reported that all 12 seeded defects were detected by their expected validation rules.

Some seeded defects also generated additional findings because a single data problem can affect related records or validation checks.

## Repository Structure

```text
Clinical/
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
│   └── defect_detection_summary.xlsx
├── scripts/
│   ├── validate_data.py
│   ├── inspect_data.py
│   ├── fix_visit_dates.py
│   ├── create_test_defects.py
│   └── create_defect_summary.py
└── README.md
```

## Important Limitations

This is an educational portfolio project using entirely synthetic data. It is not a real clinical trial and does not represent a validated clinical data management system. The project does not claim production EDC experience, regulatory compliance, or formal GCP/CDISC certification.

## Skills Demonstrated

- Structured clinical-style data handling
- Data cleaning and validation
- Python-based quality checks
- Defect tracking and rule-based testing
- Excel reporting
- Documentation and reproducible workflows