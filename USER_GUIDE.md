# Ledger to Purchase CSV Converter User Guide

## Introduction
This application automates the process of converting financial Ledger Reports into the specific CSV format required for Logisys upload. It specifically handles data merging between your **Ledger Report** (Excel) and **Job Register** (CSV/Excel) to correctly map Job Numbers to Bill of Entry (BOE) numbers, while formatting tax codes and charges according to business rules.

## How to Use

### 1. Launching the App
1. Locate the application folder.
2. Double-click the `Ledger_to_CSV.exe` file (the application icon should be visible).
3. The application will launch in full-screen mode with the Nagarkot branding.

### 2. The Workflow (Step-by-Step)
1.  **Select Job Register**: 
    *   Click the **"Select Job Register"** button.
    *   Navigate and select your Job Register file.
    *   *Supported Formats*: `.csv` or `.xlsx`.
    *   *Note: The file must contain columns for "BOE No" (or similar) and "Job No" (or similar).*
2.  **Select Ledger Report**:
    *   Click the **"Select Ledger Report"** button.
    *   Select your Ledger Dump file.
    *   *Supported Format*: `.xlsx` (Excel).
    *   *Note: This usually requires the Job Register to be selected first.*
3.  **Process Files**:
    *   Click the blue **"Process & Generate CSV"** button.
    *   The application will read both files, match records, and apply formatting rules.
    *   *Note: Rows without a Receipt No. or BOE No. will be skipped and logged.*
4.  **Save Output**:
    *   The processed CSV file is automatically saved in a folder named `Kale Output` in the same directory as the application.
    *   The filename will include the current timestamp (e.g., `purchase_14-02-26 12-30.csv`).
    *   You will see a "Success" popup message upon completion.

## Interface Reference

| Control / Input | Description | Expected Format |
| :--- | :--- | :--- |
| **Select Job Register** | Selects the reference file linking BOE to Job Nos. | `.csv` or `.xlsx` |
| **Select Ledger Report** | Selects the raw financial data dump. | `.xlsx` |
| **Process & Generate CSV** | Triggers the conversion logic. | Button Action |
| **Processing Log** | Displays real-time status, errors, and skipped rows. | Live Text Output |

## Troubleshooting & Validations

If you see an error or the output isn't what you expect, check this table:

| Message / Issue | What it means | Solution |
| :--- | :--- | :--- |
| **"No Job Register file selected"** | You tried to select the Ledger Report or Process without picking the Job Register first. | Click "Select Job Register" first. |
| **"BOE column not found..."** | The Job Register file doesn't have a recognizable column header for Bill of Entry. | Ensure your CSV/Excel has a column named "BOE No", "BE No", or "Bill of Entry No". |
| **"Job No column not found..."** | The Job Register file is missing the Job Number column. | Ensure your file has "Job No", "Job Number", or "Ref No". |
| **"Skipping row ... missing Receipt No"** | A row in the Ledger Report has no Receipt Number. | Check the source Excel file for incomplete rows. |
| **"Skipping row ... missing BOE No"** | A row in the Ledger Report has no Bill of Entry Number. | Ensure BOE data is present in the source file. |
| **"Failed to create CSV: ..."** | A system error occurred during file writing. | Ensure the `Kale Output` folder isn't open or set to Read-Only. |
