# 📄Automated Rostering System

This project automates staff rostering and generates detailed analysis reports directly within Excel using VBA. It is designed to simplify shift assignment, track actual vs. planned duties, and generate summary reports for effective duty management.

---

## ✨ Features
- **Automated Roster Duplication**
  - Create timestamped copies of the main roster for record-keeping.
  - Automatically removes all buttons and form controls in the duplicated sheet.
  - Protects the duplicated sheet for view-only access.

- **Shift Analysis Report**
  - Generate slot-wise analysis for:
    - Loan Mail Box
    - Morning
    - Afternoon
    - AOH
    - Saturday AOH
  - Compare **system-generated duty counts** against **actual roster counts**.
  - Highlight discrepancies with clear differences and % difference.

- **Total Summary Table**
  - Combine data from all shift analysis tables.
  - Sum up all counters across slots into one consolidated view.
  - Automatically handle empty tables (no staff) gracefully.

- **Sheet Protection & Password Authentication**
  - Protect roster sheets with a password (`rostering2025`).
  - Prompt users to enter the password before running key operations.

- **Customizable Workflow**
  - Prompt users to select the `ActualRoster_*` sheet for analysis.
  - Automatically names the analysis sheet as `AnalysisReport_YYYYMMDD_HHMM`.

---

## 🛠 Requirements
- **Excel Version**:
  - Microsoft Excel 2016 or later  
  - Microsoft 365 (Recommended for best compatibility)
- **Macros**:
  - Macros must be enabled.
- **VBA Reference Libraries**:
  - No external libraries required (uses built-in VBA objects like `Scripting.Dictionary`).

---

## 📂 Project Structure
| Module/Procedure               | Description                                                                 |
|---------------------------------|-----------------------------------------------------------------------------|
| `DuplicateActualRoster`         | Duplicates the roster sheet, removes buttons, and protects it.             |
| `MasterGenerateAllAnalyses`     | Master procedure to generate all shift analyses and summary reports.       |
| `GenerateShiftAnalysisBlock`    | Generates an individual shift analysis table (slot-wise).                  |
| `GenerateTotalSummaryTable`     | Combines all shift analysis tables into one total summary.                 |
| `ProtectRosterSheet`            | Protects the roster sheet with a predefined password.                     |
| `UnprotectRosterSheet`          | Prompts the user for a password to unprotect the roster sheet.             |

---

## 🚀 Usage Guide
1. **Open the Workbook**
   - Enable macros when prompted.
   
2. **Generate a Roster Copy**
   - Run `DuplicateActualRoster` to create a new `ActualRoster_*` sheet.

3. **Populate the Roster**
   - Assign shifts using the main `Roster` sheet.

4. **Generate Analysis**
   - Run `MasterGenerateAllAnalyses`.
   - Select an `ActualRoster_*` sheet when prompted.
   - Enter the password if required.
   - View the generated report in `AnalysisReport_YYYYMMDD_HHMM`.

5. **Review Total Summary**
   - The summary table will be generated to the right of the individual analyses.

---

## 🧩 Example Workflow
1. **Start with a base `Roster`** → Duplicate using `DuplicateActualRoster`.  
2. **Edit the new roster** (`ActualRoster_20250728_2328`) with actual data.  
3. **Generate Analysis Report** → A new `AnalysisReport_20250728_2328` will be created.  
4. **Check Summary Table** → Validate staff duty distribution across all shifts.

---

## ⚠ Notes
- Ensure staff lists are correctly filled; blank personnel lists will generate empty analysis tables.
- The macro skips missing tables but still creates placeholders in the analysis report.
- Protecting sheets prevents accidental edits but allows sorting and filtering.
