const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const builder = require('xmlbuilder');

// Get the input file path from the command-line arguments
const inputFilePath = process.argv[2];
if (!inputFilePath) {
  console.error("No input file specified. Please provide a valid .xls file path.");
  process.exit(1);
}

// Function to parse Excel and generate OFX
function convertExcelToOFX(inputExcelFile) {
  // Determine the output path with the same file name but different extension
  const inputDir = path.dirname(inputExcelFile);
  const inputBaseName = path.basename(inputExcelFile, path.extname(inputExcelFile)); // Base name without extension
  const outputOFXFile = path.join(inputDir, `${inputBaseName}.ofx`); // Create the output path with .ofx extension

  // Read the Excel file
  const workbook = xlsx.readFile(inputExcelFile);

  // Assuming the first sheet contains the transactions
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  // Parse the sheet into JSON format
  const transactions = xlsx.utils.sheet_to_json(sheet);

  // Create OFX structure
  const ofx = builder.create('OFX')
    .ele('SIGNONMSGSRSV1')
      .ele('SONRS')
        .ele('STATUS')
          .ele('CODE', 0).up()
          .ele('SEVERITY', 'INFO').up()
        .up()
        .ele('DTSERVER', new Date().toISOString()).up()
        .ele('LANGUAGE', 'ENG').up()
      .up()
    .up()
    .ele('BANKMSGSRSV1')
      .ele('STMTTRNRS')
        .ele('TRNUID', '1001').up()
        .ele('STATUS')
          .ele('CODE', 0).up()
          .ele('SEVERITY', 'INFO').up()
        .up()
        .ele('STMTRS')
          .ele('CURDEF', 'EUR').up()
          .ele('BANKTRANLIST');

  // Map the Excel columns to OFX fields
  transactions.forEach((transaction, index) => {
    const { Transactiebedrag, Omschrijving, Transactiedatum } = transaction;

    // Extract payee and description from "Omschrijving" field using the dynamic extraction function
    const payee = extractPayee(Omschrijving);
    const description = Omschrijving || 'No description available';

    // Use the date directly from the Excel file and convert it to OFX format
    const date = formatDate(Transactiedatum);

    ofx.ele('STMTTRN')
      .ele('TRNTYPE', Transactiebedrag < 0 ? 'DEBIT' : 'CREDIT').up() // DEBIT for negative amounts, CREDIT for positive
      .ele('DTPOSTED', date).up() // Use formatted date from Transactiedatum
      .ele('TRNAMT', Transactiebedrag).up() // Assuming 'Transactiebedrag' is the amount
      .ele('FITID', 'T' + index).up() // Unique ID for each transaction
      .ele('NAME', payee).up() // Use extracted payee
      .ele('MEMO', description).up() // Full description
    .up();
  });

  // Close the OFX structure
  ofx.up().up().up().up();

  // Write the OFX file to the same directory as the input Excel file, using the same base name
  const ofxString = ofx.end({ pretty: true });
  fs.writeFileSync(outputOFXFile, ofxString);
  console.log(`OFX file created successfully at ${outputOFXFile}`);
}

// Utility function to convert Excel's YYYYMMDD to OFX's required format (YYYYMMDDHHMMSS)
function formatDate(dateValue) {
  return `${dateValue}000000`; // Append 000000 to match the OFX format
}

// Extracts the payee based on the description field
function extractPayee(description) {
  const staticFormats = { "ABN AMRO Bank N.V.": "ABN AMRO Bank N.V.", "RENTEAFSLUITING": "Renteafsluiting" };
  for (const key in staticFormats) {
    if (description.startsWith(key)) return staticFormats[key];
  }
  const dynamicFormats = ["BEA,", "GEA,", "APP", "eCom"];
  for (const format of dynamicFormats) {
    if (description.startsWith(format)) {
      const start = 33;
      const end = description.indexOf(",", start) !== -1 ? description.indexOf(",", start) : start + 32;
      return description.substring(start, end).trim();
    }
  }
  if (description.startsWith("SEPA Overboeking")) {
    const match = description.match(/Naam:\s*([^\s].*?)\s*(?=IBAN|BIC|Omschrijving|$)/);
    return match ? match[1].trim() : "Unknown Payee";
  }
  if (description.startsWith("/TRTP/")) {
    const match = description.match(/\/NAME\/([^/]+)\//);
    return match ? match[1].trim() : "Unknown Payee";
  }
  return "Unknown Format";
}

// Run the converter with the input file path provided as an argument
convertExcelToOFX(inputFilePath);