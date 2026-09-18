# SIOP Data Health Validator

## Overview

SIOP Data Health Validator is an automated data quality validation solution built using Python and Excel VBA.

The tool validates large volumes of SIOP planning workbooks and identifies structural, formatting, and data quality issues before downstream planning and reporting activities.

The solution provides business users with a simple Excel-based interface while leveraging Python for high-volume validation and report generation.

---

## Key Features

### File Validation

- Validates file naming conventions
- Detects duplicate files
- Verifies workbook format
- Confirms file availability

### Sheet Validation

- Validates mandatory worksheet presence
- Verifies expected column count
- Checks fixed column structures
- Validates monthly column formats

### Data Validation

- Data type validation
- Missing value detection
- Manual NA/N-A detection
- Excel formula error detection
- Business rule validation

### Reporting

- Automated Excel validation reports
- Conditional formatting for issues
- Structured issue categorization
- Detailed validation summary

---

## Solution Architecture

User
→ Excel VBA Frontend
→ Python Validation Engine
→ Validation Modules
→ Validation Report Generator
→ Excel Output Report

---

## Technologies Used

- Python
- Pandas
- OpenPyXL
- Xlwings
- Excel VBA
- Logging Framework

---

## Business Value

This solution significantly reduces manual data quality review effort by automatically identifying inconsistencies across SIOP planning workbooks before planning cycles begin.

Benefits include:

- Faster validation cycles
- Improved data quality
- Reduced manual review effort
- Standardized governance checks
- Improved planning data reliability

---

## Sample Validation Checks

- File naming compliance
- Missing worksheets
- Unexpected columns
- Invalid data types
- Blank mandatory fields
- Excel formula errors
- Invalid month formatting
- Duplicate file detection

---

## Screenshots

### User Interface

screenshots/dashboard.png

### Validation Report

screenshots/validation-report.png

---

## Author

Gayathri R

Analyst - Application Development

Portfolio Project
