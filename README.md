# Excel VBA – PDF Data Extraction & Automated Reconciliation Management

An Excel VBA-based automation system designed to process multiple PDF documents, extract structured financial information, identify customer details, and maintain organized records through rule-based validation and automated data management.

The system supports individual file selection and bulk folder imports, extracts information using filename parsing and regular expressions, and automatically inserts or updates records based on predefined matching conditions.

---

## 1. Demo Link

- **Demo:** Not publicly hosted.
- **Screenshots:** Available in the `Screenshots` folder.
- **Source Code:** Available in this repository.

## 2. Table of Contents

1. Business Understanding
2. Data Understanding
3. Screenshots of Results
4. Technologies
5. Setup
6. Approach
7. Business Rules & Conditions
8. Key VBA Functions
9. Challenges & Solutions
10. Project Status
11. Credits

---

## 3. Business Understanding

Processing financial confirmation documents manually can be time-consuming, particularly when dealing with multiple PDF files and maintaining corresponding customer records.

Manual processing also increases the possibility of duplicate entries, inconsistent data, incorrect financial values, and difficulties tracking previously processed documents.

The objective of this project is to automate PDF document processing and financial record management using Excel VBA.

### Project Objectives

- Automate the import of multiple PDF documents.
- Extract customer identification details from PDF filenames.
- Retrieve customer information from a master worksheet.
- Extract reference numbers and associated financial amounts.
- Identify existing records using multiple matching conditions.
- Automatically insert new records or update existing records.
- Maintain structured and traceable document records.

The main technical challenge was extracting relevant information from PDF binary content without using an external PDF extraction application.

---

## 4. Data Understanding

The project uses two structured Excel worksheets and externally stored PDF documents.

### Workbook Structure

| Worksheet | Purpose |
|---|---|
| Balance_Confirmation | Stores imported document details, extracted financial information, and document links. |
| Customer_Master | Maintains customer codes and corresponding customer names. |

### PDF File Structure

The system expects PDF filenames in the following format:

`CustomerCode_Quarter_Year.pdf`

Example:

`640432D035_Q4_2025.pdf`

The filename provides three important attributes:

- Customer Code
- Reporting Quarter
- Year

These attributes are used to identify customers, calculate financial years, and match existing records.

### Data Processing Fields

| Field | Source |
|---|---|
| Customer Code | PDF filename |
| Customer Name | Customer Master |
| Quarter | PDF filename |
| Year | PDF filename |
| Financial Year | Calculated |
| PDF Filename | Original PDF |
| Confirmation Status | Automated status assignment |
| File Date | PDF file metadata |
| Financial Amount | PDF binary content |
| Reference Number | PDF binary content |

The project does not require a separate external dataset. It uses the customer master worksheet and user-selected PDF documents.

---

## 5. Screenshots of Visualizations/Results

The following screenshots demonstrate the PDF import workflow and automated record management.

## Screenshots of Results

| PDF File Selection | Folder-Based Bulk Import |
|:---:|:---:|
| ![PDF File Selection](Screenshots/PDF%20File%20Selection.png) | ![Folder-Based Bulk Import](Screenshots/Folder-Based%20Bulk%20Import.png) |
| **Automated Record Insertion** | **Import Completion Summary** |
| ![Automated Record Insertion](Screenshots/Automated%20Record%20Insertion.png) | ![Import Completion Summary](Screenshots/Import%20Completion%20Summary.png) |

Refer to the `Screenshots` folder for visual demonstrations.

---

## 6. Technologies

- **Microsoft Excel:** Structured data storage and record management.
- **Excel VBA:** Automation and business rule implementation.
- **VBScript.RegExp:** Pattern-based extraction of reference numbers and financial amounts.
- **FileDialog:** Interactive PDF file and folder selection.
- **Binary File Handling:** Reading PDF file contents.
- **String Manipulation:** Filename parsing and extracted text cleaning.
- **Worksheet Automation:** Automated insertion, updating, and formatting.
- **Conditional Logic:** Rule-based record classification and duplicate detection.

---

## 7. Setup

### Prerequisites

- Microsoft Excel with VBA support.
- Excel workbook saved in `.xlsm` format.
- PDF documents following the required filename convention.
- Customer master data maintained in the workbook.

### Installation

1. Clone or download this repository.
2. Open the Excel workbook in Microsoft Excel.
3. Enable macros if the workbook is from a trusted source.
4. Ensure the following worksheets exist:
   - `Balance_Confirmation`
   - `Customer_Master`
5. Populate the customer master worksheet with customer codes and names.
6. Import the VBA module into the workbook if it is not already included.
7. Save the workbook in Excel Macro-Enabled Workbook (`.xlsm`) format.

### Usage Instructions

1. Run the `ImportBalanceConfirmationPDFs43` macro.
2. Select the preferred import method:
   - Yes: Select individual or multiple PDF files.
   - No: Select an entire folder containing PDF documents.
   - Cancel: Exit the import operation.
3. The system processes the selected PDF documents.
4. Extracted information is validated and matched against existing records.
5. New records are inserted and existing records are updated.
6. Review the final import summary.

**Note:** The original PDF files remain in their selected locations. The workbook stores clickable hyperlinks to those documents.

---

## 8. Approach

The project follows a structured PDF processing and record management workflow.

### Step 1: PDF Selection

Provides two import options:

- Individual or multiple PDF selection using File Picker.
- Bulk PDF collection from a selected folder using Folder Picker.

### Step 2: Filename Parsing

Uses the `Split()` function to separate filename components using underscores.

Extracts:
- Customer Code
- Quarter
- Year

### Step 3: Customer Master Lookup

Searches the customer master worksheet using the extracted customer code and retrieves the corresponding customer name.

### Step 4: Financial Year Calculation

Calculates the financial year using the extracted year.

Example:

`2025 → 2025-2026`

### Step 5: PDF Binary Processing

Reads the selected PDF file using VBA binary file operations.

The extracted binary content is cleaned by removing null characters before pattern matching.

### Step 6: Reference Number Extraction

Uses regular expressions to identify reference numbers matching predefined patterns.

### Step 7: Financial Amount Extraction

Searches the text following the identified reference number and extracts the associated numeric amount.

Handles:
- Positive amounts.
- Negative amounts with a leading minus sign.
- Negative amounts with a trailing minus sign.
- Comma-separated financial values.
- Decimal values.

### Step 8: Duplicate Detection

Checks whether a record already exists using a combination of:

- Customer Code
- Quarter
- Year

### Step 9: Record Insertion or Update

If no matching record exists:
- Inserts a new row.
- Copies the required formatting.
- Stores customer and document details.
- Adds a clickable PDF hyperlink.
- Stores extracted financial information when available.

If a matching record exists:
- Updates the customer name.
- Replaces the PDF hyperlink.
- Updates confirmation status and file date.
- Updates the financial amount and reference number when a reference number is found.

### Step 10: Final Processing

- Renumbers serial numbers.
- Automatically adjusts column widths.
- Restores Excel application settings.
- Displays an import summary.

---

## 9. Business Rules & Conditions

The system implements predefined conditions to maintain consistency and prevent duplicate records.

### A. File Selection Rules

- Users can select individual PDF files or an entire folder.
- Multiple PDF selection is supported.
- Folder selection collects PDF files from the selected folder.
- Processing stops if no PDF files are selected or found.

### B. Filename Validation Rules

- Filenames must contain exactly three underscore-separated components.
- The first component represents the customer code.
- The second component represents the reporting quarter.
- The third component represents a numeric year.
- Files that do not satisfy these conditions are skipped.

### C. Customer Lookup Rules

- Customer codes are matched against the master worksheet.
- Leading and trailing spaces are removed.
- Non-breaking spaces are removed from master customer codes.
- Text comparison is case-insensitive.
- If no customer match is found, the customer name remains blank.

### D. Duplicate Detection Rules

A record is considered existing when all three conditions match:

1. Customer Code
2. Reporting Quarter
3. Year

If a match is found, the existing record is updated rather than inserting another record.

### E. PDF Extraction Rules

- Reference numbers are identified using predefined regular expression patterns.
- Amount extraction is performed only when a reference number is identified.
- The amount is searched within the text following the reference number.
- Both leading and trailing negative signs are supported.
- Commas and unnecessary spaces are removed before numeric conversion.

### F. Record Management Rules

**For New Records:**
- Insert the record at the top of the worksheet.
- Copy formatting from the existing template row.
- Generate a clickable PDF hyperlink.
- Assign confirmation status.
- Store file modification date.
- Store extracted amount and reference number when available.

**For Existing Records:**
- Update customer name.
- Update PDF hyperlink.
- Update confirmation status.
- Update file date.
- Update amount and reference number only when a reference number is found.

### G. Final Validation

- Serial numbers are regenerated after processing.
- Column widths are adjusted automatically.
- Excel screen updating and events are restored.
- Import totals are displayed through a completion message.

---

## 10. Key VBA Functions

| Function / Procedure | Responsibility |
|---|---|
| `ImportBalanceConfirmationPDFs43()` | Main procedure for selecting, processing, and importing PDF files. |
| `GetFinancialYear()` | Generates financial-year information. |
| `GetCustomerName()` | Retrieves customer names from the master worksheet. |
| `FindExistingRecord()` | Identifies existing records using customer code, quarter, and year. |
| `RenumberSerialNumbers()` | Regenerates sequential record numbers. |
| `ReadPDFBinary()` | Reads PDF file contents using binary access. |
| `ExtractReferenceAndAmount()` | Extracts reference numbers and associated financial amounts using regular expressions. |

### Important VBA Concepts Demonstrated

- File and Folder Dialog Operations
- Collection Objects
- Binary File Reading
- Regular Expression Pattern Matching
- String Manipulation
- Composite-Key Matching
- Conditional Record Insertion and Updating
- Dynamic Worksheet Operations
- Error Handling
- Automated Data Formatting

---

## 11. Challenges & Solutions

| Challenge | Solution |
|---|---|
| Processing multiple PDF files | Implemented individual file and folder selection modes. |
| Extracting customer information | Used structured filename parsing and master data lookup. |
| Reading PDF content | Implemented binary file reading. |
| Identifying reference numbers | Used regular expression pattern matching. |
| Extracting negative financial values | Added support for leading and trailing minus signs. |
| Preventing duplicate records | Implemented composite-key matching using customer code, quarter, and year. |
| Maintaining updated records | Implemented conditional insertion and update operations. |
| Improving processing experience | Disabled screen updating and events during processing. |

---

## 12. Project Status

**Status:** Core functionality implemented.

The current implementation supports:
- Individual and bulk PDF imports.
- Filename-based information extraction.
- Customer master lookup.
- Reference number and financial amount extraction.
- Duplicate record identification.
- Automated record insertion and updating.
- Import summary generation.

---

## 13. Credits

Developed independently as a practical Excel VBA automation project.

The project demonstrates the application of VBA programming, regular expressions, file handling, and rule-based data processing to automate document management workflows.

