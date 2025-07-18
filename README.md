# Automated Invoice Data Extraction with Google Apps Script

This repository contains a Google Apps Script to automate the entire invoice processing workflow, from document scanning to recording structured data in Google Sheets and archiving documents.

## Key Features

- **Multi-Format Support:** Automatically processes invoice files in PDF, image (JPG, PNG), and `.docx` formats.
- **2 Mechanisms of Data Extraction:**
    - For PDFs and images, it uses the **OCR API** to extract text.
    - For `.docx` files, it uses a conversion to a Google Doc to read the text directly.
- **Dual-Parser System:** Uses separate, optimized parsing logic for each data source type (i.e., in common occurence, OCR results are often unstructured vs. cleaner direct `.docx` reads of which both are often subject to different mechanisms of data extraction).
- **Automated Archiving:**
    - Creates a dynamic folder structure for payment documentation with the format `Year/Recipient Name/Month`.
    - Saves a raw text copy of each processed document with the filename format `DDMMYYYY-HHmmss-OriginalFileName.txt` in a separate folder.
    - Moves the processed invoice file to an archive folder.
- **Spreadsheet Recording:** Automatically adds a new row to a Google Sheet containing structured data, processing status, and a direct link to the documentation folder.
- **Alternative Mechanism:** Equipped with a retry mechanism for failed external API connections and a stable file conversion method.

## Project Structure

```
.
├── Code.gs         # Main file containing all the Google Apps Script logic.
└── appsscript.json # Manifest file (auto-generated by Google).
└── README.md       # This documentation.
```

## Setup and Configuration

Before running the script, ensure you have completed the following setup in your Google environment.

### 1. Google Drive Configuration
Create the following 4 folders in your Google Drive and note their IDs (the ID can be found in the folder's URL):
- **Incoming Invoices Folder:** Where you will place new invoice files to be processed.
- **Processed Invoices Folder:** Where the script will move files after they have been processed.
- **Payment Documentation Folder:** The main folder where the `Year/Recipient Name/Month` structure will be created.
- **Raw Text Folder:** A dedicated folder to store all the raw text extraction `.txt` files.

### 2. Google Sheet Configuration
- Create a new Google Sheet file.
- Rename "Sheet1" to your desired name (e.g., "Invoices").
- Create headers in the first row in this exact order:
  `Invoice Source`, `API Source`, `Processing Time`, `Invoice Number`, `Invoice Date`, `Recipient`, `Total Payment`, `Status`, `Payment Document`

### 3. Get an API Key
- This script uses OCR API (in this case **OCR.space**) to process PDFs/images.
- If you prefered not to modify this script further for different API, register on the [OCR.space website](https://ocr.space/ocrapi/free) to get your free `apikey`.

### 4. Script Configuration (`Code.gs`)
Open the `Code.gs` file and fill in all the variables in the `USER CONFIGURATION` section with the Folder IDs and API Key you have prepared.

```javascript
// ===============================================================
// USER CONFIGURATION (Universal)
// ===============================================================

// -- API Configuration (Currently == OCR.space) --
const OCR_API_KEY = 'YOUR_API_KEY_HERE'; 
const OCR_API_ENDPOINT = '[https://api.ocr.space/parse/image](https://api.ocr.space/parse/image)';

// -- Google Drive & Sheets Configuration --
const SOURCE_FOLDER_ID = 'YOUR_SOURCE_FOLDER_ID_HERE';
const PROCESSED_FOLDER_ID = 'YOUR_PROCESSED_FOLDER_ID_HERE';
const PAYMENT_DOCS_ROOT_FOLDER_ID = 'YOUR_PAYMENT_DOCS_FOLDER_ID_HERE';
const RAW_TEXT_FOLDER_ID = 'YOUR_RAW_TEXT_FOLDER_ID_HERE';
const SHEET_NAME = 'Invoices'; // Adjust to your sheet name
```

### 5. Enable Advanced Services
This script requires the **Drive API Service** to convert `.docx` files.
- In the Apps Script Editor, click on the **Services +** menu.
- Find and select **Drive API**.
- Click **Add**. Ensure the Identifier is `Drive`.

## Installation and Execution

1.  Open the Google Sheet you prepared.
2.  Go to **Extensions > Apps Script**.
3.  Copy the entire content of the `Code.gs` file and paste it into the editor, replacing any existing code.
4.  Perform the **Script Configuration** and **Enable Advanced Services** steps as described above.
5.  Save the project.
6.  **Initial Authorization:** Run the `processInvoicesUniversal` function manually from the editor for the first time. Google will prompt you with a series of permission requests. Review and allow all of them.
7.  **Set Up the Trigger:**
    - In the editor, click on the **Triggers** menu (clock icon).
    - Click **Add Trigger**.
    - Choose the `processInvoicesUniversal` function to run.
    - Select the event source as `Time-driven`.
    - Set the interval (e.g., `Hour timer` to run every hour).
    - Save the trigger.

## How to Use

Simply place an invoice file (PDF, JPG, PNG, or DOCX) into the "Incoming Invoices Folder" in your Google Drive. The script will automatically process it at the next interval based on the trigger you set.

---

Contributions and suggestions for improvement are welcome. Please create an *issue* or a *pull request*.
