// ===============================================================
// USER CONFIGURATION (Universal)
// ===============================================================

// -- API Configuration (Currently == OCR.space) --
const OCR_API_KEY = 'YOUR_API_KEY'; 
const OCR_API_ENDPOINT = 'https://api.ocr.space/parse/image';

// -- Google Drive & Sheets Configuration --
const SOURCE_FOLDER_ID = '1fRVqYo40G4VMGuWk1bBp4BiMkNfK30r0'; // PDF folder to be processed
const PROCESSED_FOLDER_ID = '1Ea0ualQP2NJHTS3JUtK5A0hHnSAWoA5H'; // Folder for processed PDFs
const PAYMENT_DOCS_ROOT_FOLDER_ID = '1jRFCLXAUZ1GC0PAc3h9W8PIU56wdlEEc'; // Root folder for payment proofs
const RAW_TEXT_FOLDER_ID = '11CYwot9UNDoflol80NDh8RVQgWNOwCqi'; // Folder for raw text files
const SHEET_NAME = 'Example'; // Change to your sheet name if different

// ===============================================================
// MAIN FUNCTION & PROCESSING
// ===============================================================

/**
 * The main function to be executed by a trigger.
 * Separates parsing logic for OCR and DOCX.
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
    let paymentFolderUrl = 'Creation Failed';
    let data = {};
    let fullText = null; 
    let dataSource = '';

    try {
      const mimeType = file.getMimeType();

      if (mimeType === MimeType.PDF || mimeType.startsWith('image/')) {
        dataSource = apiSource;
        const ocrApiResponse = callOcrApi(file.getBlob());
        const parsedOcr = parseOcrApiResponse(ocrApiResponse);
        if (parsedOcr.text) {
          fullText = parsedOcr.text;
          
          // --- DEBUG SPECIFICALLY FOR OCR PROCESS ---
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
          const tempDocxFile = rawTextFolder.createFile(file.getBlob());
          tempDocxId = tempDocxFile.getId();
          const resource = { title: `temp_gdoc_conversion_${fileName}`, mimeType: 'application/vnd.google-apps.document' };
          const convertedGdoc = Drive.Files.copy(resource, tempDocxId);
          tempGdocId = convertedGdoc.id;
          const doc = DocumentApp.openById(tempGdocId);
          fullText = doc.getBody().getText();
          data = parseDocxText(fullText); // Using the DOCX parser
        } finally {
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
// OCR PROVIDER-SPECIFIC FUNCTIONS
// ===============================================================

function callOcrApi(fileBlob) {
  const payload = { 'file': fileBlob, 'language': 'eng', 'isTable': 'true', 'OCREngine': '1' };
  const options = { 'method': 'post', 'headers': { 'apikey': OCR_API_KEY }, 'payload': payload, 'muteHttpExceptions': true };
  for (let i = 0; i < 3; i++) {
    try {
      const response = UrlFetchApp.fetch(OCR_API_ENDPOINT, options);
      if (response.getResponseCode() == 200) { return JSON.parse(response.getContentText()); }
    } catch (e) {
      Logger.log(`Attempt ${i + 1} failed: ${e.toString()}`);
    }
    if (i < 2) Utilities.sleep(2000);
  }
  return { error: 'Failed to connect to the API after 3 attempts.' };
}

function parseOcrApiResponse(ocrResponse) {
  if (ocrResponse && !ocrResponse.IsErroredOnProcessing && ocrResponse.ParsedResults && ocrResponse.ParsedResults.length > 0) {
    return { text: ocrResponse.ParsedResults[0].ParsedText };
  } else if (ocrResponse && ocrResponse.ErrorMessage) {
    return { error: ocrResponse.ErrorMessage.join('; ') };
  } else {
    return { error: 'Invalid or empty API response.' };
  }
}

// ===============================================================
// INTERNAL LOGIC FUNCTIONS
// ===============================================================

/**
 * [PDF PARSER] Parses text from a structured OCR API result.
 * Last edit: ?.
 */
function parseOcrText(text) {
  const data = {};
  
  let matchInvoiceNumber = text.match(/INV\/TAP\/\d{4}\/[A-Z\d]+[\/']\d+/i);
  data.invoiceNumber = matchInvoiceNumber ? matchInvoiceNumber[0] : 'Not Found';

  let matchDate = text.match(/Tanggal Faktur\s*:?\s*(\d{1,2}\s+[A-Z]+\s+\d{4})/i);
  data.date = matchDate ? matchDate[1].trim() : 'Not Found';

  let matchRecipient = text.match(/Kepa\s*da\s*Yth,?\s*([^\n\r]+)/i);
  data.recipient = matchRecipient ? matchRecipient[1].trim() : 'Not Found';

  let matchTotal = text.match(/Total Tag\s*Ihan\s*Rp\.?\s*([\d.,O-]+)/i);
  if (matchTotal) {
    let capturedString = matchTotal[1];
    let stringWithZeros = capturedString.replace(/O/g, '0'); // Replace OCR error 'O' with '0'
    let cleanTotal = stringWithZeros.replace(/[.,-]/g, ''); 
    data.total = parseFloat(cleanTotal);
  } else {
    data.total = 'Not Found';
  }
  return data;
}

/**
 * [DOCX PARSER] Parses text from a .docx file.
 * The parsing method looks for unique data patterns, not adjacent labels.
 */
function parseDocxText(text) {
  const data = {};
  
  // Invoice Number Pattern: Searches for INV/TAP/... format anywhere in the text.
  let matchInvoiceNumber = text.match(/(INV\/TAP\/\d{4}\/[A-Z\d\/]+\/\d+)/i);
  data.invoiceNumber = matchInvoiceNumber ? matchInvoiceNumber[1] : 'Not Found';

  // Invoice Date Pattern: Searches for DD Month YYYY format anywhere in the text.
  let matchDate = text.match(/(\d{1,2}\s+(?:Januari|Februari|Maret|April|Mei|Juni|Juli|Agustus|September|Oktober|November|Desember)\s+\d{4})/i);
  data.date = matchDate ? matchDate[1].trim() : 'Not Found';

  // Recipient Pattern: Searches for text after "Kepada Yth,"
  let matchRecipient = text.match(/Kepada Yth,?\s*([^\n\r]+)/i);
  data.recipient = matchRecipient ? matchRecipient[1].trim() : 'Not Found';

  // Total Invoice Pattern: Searches for text after "Total Tagihan"
  let matchTotal = text.match(/Total Tagihan\s*Rp\.\s*([\d.,-]+)/i);
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
    const months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    const parts = invoiceDateString.replace(/,/g, '').split(' ');
    const monthIndex = months.findIndex(m => m.toUpperCase() === parts[1].toUpperCase());
    const dateObj = new Date(parts[2], monthIndex, parts[0]);

    if (isNaN(dateObj.getTime())) {
      return 'Date Format Error';
    }
    const year = dateObj.getFullYear().toString();
    
    // Get or create Year folder
    const yearFolders = currentFolder.getFoldersByName(year);
    if (yearFolders.hasNext()) {
      currentFolder = yearFolders.next();
    } else {
      currentFolder = currentFolder.createFolder(year);
    }

    // Get or create Recipient folder
    const recipientFolders = currentFolder.getFoldersByName(recipientName);
    if (recipientFolders.hasNext()) {
      currentFolder = recipientFolders.next();
    } else {
      currentFolder = currentFolder.createFolder(recipientName);
    }
    
    const monthNumber = (dateObj.getMonth() + 1).toString().padStart(2, '0');
    const monthName = months[dateObj.getMonth()];
    const monthFolderName = `${monthNumber} - ${monthName}`;

    // Get or create Month folder
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
