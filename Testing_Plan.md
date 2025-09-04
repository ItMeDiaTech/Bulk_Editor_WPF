# Bulk Editor - Comprehensive Testing Plan (v1.2.0)

## 1. Introduction

This document outlines the test cases for validating the recent enhancements and bug fixes implemented in the Bulk Editor application. The goal is to ensure all new and modified features are working as expected, the application is stable, and no regressions have been introduced.

## 2. Test Environment

- **Application Version:** 1.2.0
- **Operating System:** Windows 10/11
- **Prerequisites:**
  - A valid `appsettings.json` with a working API endpoint for lookups.
  - Sample `.docx` files for testing.
  - A test `lookup.csv` file for hyperlink replacement rules.
  - A test `replacements.csv` file for text replacement rules.

## 3. Test Cases

### 3.1. Feature: Startup Update Check & Notifications

| Test ID | Scenario                   | Test Steps                                                                        | Expected Result                                                                                                                    |
| ------- | -------------------------- | --------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------- |
| UC-01   | **Update Available**       | 1. Ensure a newer version is available on GitHub. <br> 2. Launch the application. | A notification appears at the bottom of the window indicating "A new version is available." The "Update" button should be visible. |
| UC-02   | **No New Update**          | 1. Ensure the running version is the latest. <br> 2. Launch the application.      | No update notification appears.                                                                                                    |
| UC-03   | **No Internet Connection** | 1. Disconnect the machine from the internet. <br> 2. Launch the application.      | The application starts without crashing. No update notification appears. A log entry may indicate the update check failed.         |
| UC-04   | **Perform Update**         | 1. From the state in UC-01, click the "Update" button.                            | A confirmation dialog appears. On 'Yes', the installer for the new version is downloaded and launched.                             |

---

### 3.2. Feature: Settings & API Connection Test

| Test ID | Scenario                   | Test Steps                                                                                                                    | Expected Result                                                                                   |
| ------- | -------------------------- | ----------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------- |
| AC-01   | **Valid API Connection**   | 1. Navigate to Settings. <br> 2. Ensure the `ApiLookupUrl` is correct in `appsettings.json`. <br> 3. Click "Test Connection". | A success message ("Connection successful!") is displayed.                                        |
| AC-02   | **Invalid API Connection** | 1. Navigate to Settings. <br> 2. Set the `ApiLookupUrl` to a non-existent endpoint. <br> 3. Click "Test Connection".          | An error message ("Connection failed.") is displayed.                                             |
| AC-03   | **UI Verification**        | 1. Navigate to Settings.                                                                                                      | The UI should **not** contain fields for "API Key" or "GitHub Owner". The layout should be clean. |

---

### 3.3. Feature: Revert Last Process

| Test ID | Scenario                              | Test Steps                                                                                                                                                                                      | Expected Result                                                                                                                                                                             |
| ------- | ------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| RV-01   | **Successful Revert**                 | 1. Load one or more `.docx` files. <br> 2. Load replacement/hyperlink rules. <br> 3. Click "Process Documents". <br> 4. Verify the documents are modified. <br> 5. Click "Revert Last Process". | The documents are restored to their original state (before processing). A success message is shown. The "Revert" button becomes disabled.                                                   |
| RV-02   | **Revert Without Prior Action**       | 1. Launch the application. <br> 2. Click "Revert Last Process".                                                                                                                                 | The button should be disabled, so this action is not possible.                                                                                                                              |
| RV-03   | **Multiple Processes, Single Revert** | 1. Process a set of documents (Session 1). <br> 2. Process another set of documents (Session 2). <br> 3. Click "Revert Last Process".                                                           | Only the modifications from Session 2 are reverted. The "Revert" button becomes disabled.                                                                                                   |
| RV-04   | **Backup Folder Verification**        | 1. Before processing, check the `%LOCALAPPDATA%\BulkEditor\backups` directory. <br> 2. After processing, check the directory again. <br> 3. After reverting, check the directory a final time.  | 1. The directory is empty or doesn't exist. <br> 2. The directory contains a new folder with backups of the processed files. <br> 3. The backup folder for the reverted session is deleted. |

---

### 3.4. Feature: Hardened Replacement Services

| Test ID | Scenario                                       | Test Steps                                                                                                                                           | Expected Result                                                                                                                                                                                        |
| ------- | ---------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| HR-01   | **Text Replacement (Capitalization)**          | 1. Create a rule to replace "word" with "replacement". <br> 2. Process a document containing "word", "Word", and "WORD".                             | The document is updated with "replacement", "Replacement", and "REPLACEMENT" respectively.                                                                                                             |
| HR-02   | **Hyperlink Replacement (Invalid Content ID)** | 1. Create a hyperlink rule with a `Content ID` that is not in the `lookup.csv`. <br> 2. Process a document containing a matching hyperlink.          | The hyperlink is not replaced. The document's status shows a warning, and inspecting the log/errors should show a `ProcessingError` with a message like "Could not lookup document for Content ID...". |
| HR-03   | **Hyperlink Replacement (Missing Title)**      | 1. Modify the `lookup.csv` so a specific `Content ID` has an empty `Title`. <br> 2. Create a rule for that `Content ID`. <br> 3. Process a document. | The hyperlink is not replaced. A `ProcessingError` is generated indicating the title was missing.                                                                                                      |
| HR-04   | **Structured Error Logging**                   | 1. Run any scenario that is expected to fail (like HR-02 or HR-03). <br> 2. Observe the `Document` object in the UI/debugger.                        | The `ProcessingErrors` list for the affected document contains a `ProcessingError` object with the correct `RuleId`, `Message`, and `Severity`.                                                        |
