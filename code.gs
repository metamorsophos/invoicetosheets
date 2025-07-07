// ===============================================================
// USER CONFIGURATION (Universal)
// ===============================================================

// -- API Configuration (Current == OCR.space) --
const OCR_API_KEY = 'YOUR_API_KEY'; 
const OCR_API_ENDPOINT = 'https://api.ocr.space/parse/image';

// -- Google Drive & Sheets Configuration --
const SOURCE_FOLDER_ID = '1mBw9Xo45G5VMGuWk1bBp4BiMkNfK41r1'; // Folder for new invoices to be processed
const PROCESSED_FOLDER_ID = '1Ga1ualQP2NJHTS3JUtK5A0hHnSAWoB6I'; // Archive folder for processed invoices
const PAYMENT_DOCS_ROOT_FOLDER_ID = '1kRFCLXAUZ1GC0PAc3h9W8PIU56wdlEEd'; // Root folder for generated payment proofs
const RAW_TEXT_FOLDER_ID = '12CYwot9UNDoflol80NDh8RVQgWNOwDqj'; // Folder for raw text logs from OCR
const SHEET_NAME = 'Invoice_Processing_Log'; // Sheet name for logging results

// ===============================================================
// MAIN FUNCTION & PROCESSING
// ===============================================================

/**
 * The main function to be run by a trigger.
 * FINAL VERSION: Separates parsing logic for OCR and DOCX.
 */
function processUniversalInvoices() {
  const sourceFolder = DriveApp.getFolderById(SOURCE_FOLDER_ID);
  const processedFolder = DriveApp.getFolderById(PROCESSED_FOLDER_ID);
  const rawTextFolder = DriveApp.getFolderById(RAW_TEXT_FOLDER_ID);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  if (!sheet) {
    Logger.log(`Error: Sheet "${SHEET_NAME}" not found.`);
    return;
  }

  let apiSource = 'Failed to Get Source';
  const match = OCR_API_ENDPOINT.match(/:\/\/(.[^/]+)/);
  if (match != null) {
    apiSource = match[1];
  }

  const files = sourceFolder.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    let status = 'Error';
    let paymentFolderUrl = 'Failed to Create';
    let data = {};
    let fullText = null;
    let dataSource = '';

    try {
      const mimeType = file.getMimeType();

      if (mimeType === MimeType.PDF || mimeType.startsWith('image/')) {
        dataSource = apiSource;
        const ocrApiResponse = callOcrAPI(file.getBlob()); // This function would need to be adapted for the new API
        const parsedOcr = parseOcrApiResponse(ocrApiResponse); // This function would need to be adapted for the new API
        if (parsedOcr.text) {
          fullText = parsedOcr.text;

          // --- DEBUGGING SPECIFICALLY FOR OCR PROCESS ---
          Logger.log(`--- Raw Text from OCR API for file: ${fileName} ---`);
          Logger.log(fullText);
          Logger.log("----------------------------------------------------");
          // --- END DEBUG ---

          data = parseOcrText(fullText); // Using the OCR parser
        } else {
          data = { invoiceNumber: parsedOcr.error };
        }

      } else if (mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
        dataSource = 'Direct Read (.docx)';
        let tempDocxId = null;
        let tempGdocId = null;
        try {
          // Create a temporary DOCX file in a designated folder
          const tempDocxFile = rawTextFolder.createFile(file.getBlob());
          tempDocxId = tempDocxFile.getId();
          
          // Convert the temporary DOCX to a Google Doc
          const resource = { title: `temp_gdoc_conversion_${fileName}`, mimeType: 'application/vnd.google-apps.document' };
          const convertedGdoc = Drive.Files.copy(resource, tempDocxId);
          tempGdocId = convertedGdoc.id;
          
          // Open the Google Doc and extract text
          const doc = DocumentApp.openById(tempGdocId);
          fullText = doc.getBody().getText();
          data = parseDocxText(fullText); // Using the DOCX parser
        } finally {
          // Clean up: remove the temporary Google Doc and DOCX files
          if (tempGdocId) { Drive.Files.remove(tempGdocId); }
          if (tempDocxId) { Drive.Files.remove(tempDocxId); }
        }

      } else {
        Logger.log(`Unsupported file type: ${fileName} (${mimeType}). Skipping.`);
        continue;
      }

      if (fullText && data.invoiceNumber) {
        const now = new Date();
        const timestamp = `${now.getDate().toString().padStart(2, '0')}${(now.getMonth() + 1).toString().padStart(2, '0')}${now.getFullYear()}-${now.getHours().toString().padStart(2, '0')}${now.getMinutes().toString().padStart(2, '0')}${now.getSeconds().toString().padStart(2, '0')}`;
        const textFileName = `${timestamp}-${fileName}.txt`;
        rawTextFolder.createFile(textFileName, fullText);

        if (data.recipient && data.recipient !== 'Not Found' && data.date && data.date !== 'Not Found') {
          paymentFolderUrl = createPaymentFolder(data.recipient, data.date);
          status = 'Success';
        } else {
          status = 'Incomplete Data';
        }
      } else {
        if (!data.invoiceNumber) {
           data = {invoiceNumber: 'Extraction Failed', date: 'Extraction Failed', recipient: 'Extraction Failed', total: 'Extraction Failed'};
        }
      }

    } catch (e) {
      status = 'Critical Error';
      data.invoiceNumber = e.toString();
    }

    sheet.appendRow([
      fileName, dataSource, new Date(), data.invoiceNumber, data.date, data.recipient, data.total, status, paymentFolderUrl
    ]);
    file.moveTo(processedFolder);
  }
}

// ===============================================================
// OCR PROVIDER SPECIFIC FUNCTIONS
// NOTE: These functions are placeholders and would need to be rewritten to work with the new API endpoint (e.g., Google Cloud Vision).
// ===============================================================
function callOcrAPI(fileBlob) {
  // This function would need to be completely changed for Google Vision API.
  // It would involve creating a JSON request body with the image content (base64 encoded)
  // and feature type (e.g., 'TEXT_DETECTION').
  // The endpoint requires a ?key= parameter in the URL.
  return { error: 'callOcrAPI function not adapted for the new API.' };
}

function parseOcrApiResponse(ocrResponse) {
  // This function would also need to be changed. Google Vision returns a JSON
  // with a 'fullTextAnnotation' object containing the text.
  return { error: 'parseOcrApiResponse function not adapted for the new API.' };
}


// ===============================================================
// INTERNAL LOGIC FUNCTIONS
// ===============================================================

/**
 * [PDF PARSER] Parses text from the OCR API result.
 * This regex is configured for "Quantum Dynamics Inc." invoices.
 * Configure this line of script so that the parts of your invoice can be recognized.
 */
function parseOcrText(text) {
  const data = {};

  // Invoice Number Pattern: INV/QD/{YEAR}/{ID}
  let matchInvoiceNo = text.match(/INV\/QD\/\d{4}\/[A-Z\d]+[\/']\d+/i);
  data.invoiceNumber = matchInvoiceNo ? matchInvoiceNo[0] : 'Not Found';

  // Invoice Date Pattern
  let matchDate = text.match(/Invoice Date\s*:?\s*(\d{1,2}\s+[A-Z]+\s+\d{4})/i);
  data.date = matchDate ? matchDate[1].trim() : 'Not Found';

  // Recipient Pattern
  let matchRecipient = text.match(/Bill To\s*:\s*([^\n\r]+)/i); // Common alternative "Bill To:"
  data.recipient = matchRecipient ? matchRecipient[1].trim() : 'Not Found';

  // Total Amount Pattern
  let matchTotal = text.match(/Total Due\s*Rp\.?\s*([\d.,O-]+)/i); // Common alternative "Total Due"
  if (matchTotal) {
    let capturedString = matchTotal[1];
    let stringWithZeros = capturedString.replace(/O/g, '0');
    let cleanTotal = stringWithZeros.replace(/[.,-]/g, '');
    data.total = parseFloat(cleanTotal);
  } else {
    data.total = 'Not Found';
  }
  return data;
}


/**
 * [DOCX PARSER] Parses text from a .docx file.
 * This parsing method looks for unique data patterns, not necessarily adjacent labels.
 * Configure this line of script so that the parts of your invoice can be recognized.
 */
function parseDocxText(text) {
  const data = {};

  // Invoice Number Pattern: Looks for INV/QD/... format anywhere in the text.
  let matchInvoiceNo = text.match(/(INV\/QD\/\d{4}\/[A-Z\d\/]+\/\d+)/i);
  data.invoiceNumber = matchInvoiceNo ? matchInvoiceNo[1] : 'Not Found';

  // Invoice Date Pattern: Looks for DD Month YYYY format anywhere in the text.
  let matchDate = text.match(/(\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4})/i);
  data.date = matchDate ? matchDate[1].trim() : 'Not Found';

  // Recipient Pattern: Looks for text after "Bill To:"
  let matchRecipient = text.match(/Bill To:?\s*([^\n\r]+)/i);
  data.recipient = matchRecipient ? matchRecipient[1].trim() : 'Not Found';

  // Total Amount Pattern: Looks for text after "Total Due"
  let matchTotal = text.match(/Total Due\s*Rp\.\s*([\d.,-]+)/i);
  if (matchTotal) {
    let cleanTotal = matchTotal[1].replace(/\./g, '').replace(/,-$/, '').replace(/,/, '.');
    data.total = parseFloat(cleanTotal);
  } else {
    data.total = 'Not Found';
  }

  return data;
}

function createPaymentFolder(recipientName, invoiceDateString) {
  try {
    let currentFolder = DriveApp.getFolderById(PAYMENT_DOCS_ROOT_FOLDER_ID);
    const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    
    const parts = invoiceDateString.replace(/,/g, '').split(' ');
    const monthIndex = months.findIndex(m => m.toUpperCase() === parts[1].toUpperCase());
    const dateObj = new Date(parts[2], monthIndex, parts[0]);

    if (isNaN(dateObj.getTime())) { return 'Date Format Error'; }
    
    const year = dateObj.getFullYear().toString();
    
    // Get or create the Year folder
    const yearFolders = currentFolder.getFoldersByName(year);
    if (yearFolders.hasNext()) {
      currentFolder = yearFolders.next();
    } else {
      currentFolder = currentFolder.createFolder(year);
    }

    // Get or create the Recipient folder
    const recipientFolders = currentFolder.getFoldersByName(recipientName);
    if (recipientFolders.hasNext()) {
      currentFolder = recipientFolders.next();
    } else {
      currentFolder = currentFolder.createFolder(recipientName);
    }
    
    // Format the month folder name as "MM - MonthName"
    const monthNumber = (dateObj.getMonth() + 1).toString().padStart(2, '0');
    const monthName = months[dateObj.getMonth()];
    const monthFolderName = `${monthNumber} - ${monthName}`;

    // Get or create the Month folder
    const monthFolders = currentFolder.getFoldersByName(monthFolderName);
    if (monthFolders.hasNext()) {
      currentFolder = monthFolders.next();
    } else {
      currentFolder = currentFolder.createFolder(monthFolderName);
    }
    
    return currentFolder.getUrl();
  } catch (e) {
    return 'Failed to Create Folder';
  }
}
