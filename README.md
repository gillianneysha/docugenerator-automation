# HR DocuGenerator Automation

Automated HR document generation using Google Workspace.

## Project Summary

**DocuGenerator Automation** is a Google Workspace document automation solution built in Google Apps Script, designed to turn structured spreadsheet data into contract and agreement documents with minimal manual effort.

This project was co-developed by **Gillianne Ysha Lim** and **Ludrein Reimar Salvador**.

## What We Built

- A Google Sheets-driven document automation engine.
- A template registry for Google Docs templates with placeholder scanning.
- Batch document generation from row-based spreadsheet data.
- Output as Google Docs and PDF, with support for both formats.
- A configurable clause library for optional contract sections.

## Key Accomplishments

- Enabled dynamic contract generation from spreadsheet rows.
- Built validation that detects missing or mismatched template placeholders.
- Created clause-driven conditional content for HMO, leaves, probation, and other legal options.
- Added flexible row selection and filter-based generation modes.
- Implemented per-row success/failure logging with clear error reporting.
- Supported multi-line values, formatted sheet data, and derived runtime placeholders.

## Core Features

- **Template Management**
  - Register templates using Google Docs.
  - Store a required file-name pattern for each registered template.
  - Auto-scan `{{placeholder}}` tags in document templates.
  - Validate templates against spreadsheet headers and shared settings.
  - Re-validate templates on demand to catch issues before generation.

- **Document Generation**
  - Generate one document per spreadsheet row.
  - Support bulk generation for large batches.
  - Produce output as Google Doc, PDF, or both.
  - Customizable file naming using placeholder values.

- **Clause Library & Conditional Content**
  - Configure legal clause text from spreadsheet data instead of hardcoding it.
  - Support dynamic clauses for HMO, leave policies, business tools, and more.
  - Allow HR to update wording directly in the spreadsheet.

- **Data-driven Validation**
  - Detect missing required fields and invalid placeholder references.
  - Prevent broken template registration.
  - Provide user-friendly error messages for missing folders, invalid templates, and permission issues.

### Template Registry columns

The `Template Registry` sheet uses this column order:

`Template Name | File Name Pattern | Google Doc ID | Source Sheet | Output Folder ID | Status`

Run **Document Automation > Initialize Sheets (first-time setup)** after updating an existing spreadsheet. The setup step upgrades the registry schema and repairs rows created by older versions where the status and IDs were shifted after the file-name pattern column was introduced.

## Architecture Overview

- **Google Sheets** is the primary data source.
- **Google Docs** serves as the template source.
- **Google Apps Script** powers automation and generation.
- Key script files:
  - `Code.js`
  - `TemplateManager.js`
  - `DocumentGenerator.js`
  - `DriveService.js`
  - `SpreadsheetService.js`
  - `UI.js`
  - `Utilities.js`
- UI components are located in the `HTML/` folder.

## Why This Matters

- Eliminates manual contract assembly and repetitive copy/paste work.
- Keeps templates flexible for frequent client-driven changes.
- Lets HR manage templates and clause text without modifying code.
- Reduces generation time and minimizes human error.

## What Was Improved

- Added support for multiple contract templates and source tabs.
- Built a reusable clause registry so new conditional clauses can be added without code changes.
- Improved sheet handling so long text and formatted values render correctly in documents.
- Added friendly validation and template revalidation workflows.
- Enhanced output handling and Drive folder management.

## Result

A complete, scalable document automation solution for HR and contract generation workflows, including:

- configurable templates;
- Excel/Google Sheets data-driven generation;
- clause-based conditional content;
- PDF support;
- user-friendly validation and logging.

## Credits

- **Gillianne Ysha P. Lim**
- **Ludrein Reimar R. Salvador**
- **Project team**

## License & Security

- This repository is licensed under the **MIT License**.
- The code and documentation are intended for use as a document automation solution built on Google Workspace.
- No sensitive credentials or private keys should be stored in this repository.
- This project does not include `.env` files, but if any environment-specific configuration is needed, keep it outside source control and never commit secrets or script IDs.
- Do not store Apps Script `scriptId` values in the repository; keep them in a separate secure place if needed.
- Use `.gitignore` to exclude local or temporary files that should not be shared.

## Technologies Used

- Google Apps Script
- Google Sheets
- Google Docs
- Google Drive
- Apps Script HTML service

