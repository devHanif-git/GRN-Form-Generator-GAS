// Release on 22/12/2022 by devHanif AKA Hanif Firdaus, made with love <3
function onOpen() {
  functionMenuV2()
}

function functionMenuV2() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('GRN');
  //menu.addItem('Generate GRN Form 2022', 'createGRNFormV2')
  menu.addItem('Generate GRN Form 2023', 'createGRNFormV22023')
  menu.addToUi();
}

function createGRNFormV22023() {
  // Declare variables to hold the Google Doc templates
  const googleDocTemplate = DriveApp.getFileById('15WIbaGXKBlk0Utl6COInkBCT8muBxm5pLBgAhzze9rk');
  const googleDocTemplateMoreThan13 = DriveApp.getFileById('1Jyo7Nxq1fNOpMP4NpT47TVFe_ci1MJD8MxNRYfnVNDc');
  const googleDocTemplateMoreThan26 = DriveApp.getFileById('1XYM9_b-516FLkAf5opyRXbYZ7qboVb1mitKSUU_r27M');
  const googleDocTemplateMoreThan39 = DriveApp.getFileById('1MzS9-YA1I02kBiELbMtdaFIBGZYlCAf7cA7gCX5IF-g');
  const googleDocTemplateMoreThan52 = DriveApp.getFileById('1CRv22SCMHWMIyFkpuaFY_oPvuDoslaLKYw3ax0EpBMI');
  const googleDocTemplateMoreThan65 = DriveApp.getFileById('1_sUN2xmbvEY9Ou1ifKEbGSTfai1rzzDlWqRw4U9Nbx4');
  const googleDocTemplateMoreThan78 = DriveApp.getFileById('15RWcuy0rwHBy6Mun6HvDWQHcJzfVIpQY9SxFA3ZFAoA');
  const googleDocTemplateMoreThan91 = DriveApp.getFileById('1MqKvoeLhTMwb7CxWp8zhuw2UrhQFRVnHV1QfRtvNgKw');
  const googleDocTemplateMoreThan104 = DriveApp.getFileById('1Y3hdfibMSKwg17ww5Iu4v73tu8QiFAgOvclKprqxZ00');
  const googleDocTemplateMoreThan117 = DriveApp.getFileById('1kUeaW84hiUVQuQ98oOCHziwchi0gafs6rs0UEImDugY');

  // Declare a variable to hold the destination folder ID
  const destinationFolder = DriveApp.getFolderById('11RMjq0sW2th-fTiWRQyoUBlqN-T2ZSMq')

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('GRN2023')

  // Setup Column
  const noCol = 1, dateCol = 2, supplierCol = 3, customerCol = 4, donoCol = 5, grnCol = 6, poCol = 7;
  const modelCol = 8, itemnoCol = 9, descCol = 10, pnCol = 11, qtyCol = 12, rmkCol = 13;
  const linkCol= 14, statusCol = 15;
  
  // Get rows from DataRange
  const rows = sheet.getDataRange().getValues();

  // Setup Variable
  let currentDate = "", currentGRN = "", currentSupplier = "", currentDONO = "", currentPO = "";
  let currentCust = "", currentModel = "";
  let currentDoc, currentDocCopy, body;

  // Setup flag
  let flag = 0, flagtrig13 = false, flagtrig26 = false, flagtrig39 = false, flagtrig52 = false;
  let flagtrig65 = false, flagtrig78 = false, flagtrig91 = false, flagtrig104 = false, flagtrig117 = false;

  // Setup flag for check if any is grn is generate
  let flaggrn = 0;

  // Loop to get all data in rows 
  for (let i = 0; i < rows.length; i++) {  
    // retrive data from each row.
    const row = rows[i],  no = row[noCol-1], date = new Date(row[dateCol-1]), supplier = row[supplierCol-1];
    const customer = row[customerCol-1], dono = row[donoCol-1], grn = row[grnCol-1], po = row[poCol-1];
    const model = row[modelCol-1], itemno = row[itemnoCol-1], desc = row[descCol-1], pn = row[pnCol-1]; 
    const qty = row[qtyCol-1], rmk = row[rmkCol-1];

    // Setup for DATE STYLING
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    const formattedDate = day + "/" + month + "/" + year;

    // Skip with condition
    if (row[1] === "GOOD RECEIVING NOTE (GRN)" || row[1] === "Year : " || row[1] === "DATE" || 
        row[14] != "UNFINISHED") {
      continue;
    }
    
    // If the current row is for a new supplier, create a new doc
    if (supplier != currentSupplier && supplier != "") {
      if (currentDocCopy) {
        if(flag > 0) {
          // Clear placeholders from {{NO2}} to {{REMARK13}}
          for (let j = 2; j <= 52; j++) {
            body.replaceText(`{{NO${j}}}`, "");
            body.replaceText(`{{DESC${j}}}`, "");
            body.replaceText(`{{PN${j}}}`, "");
            body.replaceText(`{{QTYDO${j}}}`, "");
            body.replaceText(`{{REMARK${j}}}`, "");
          }
        }
        // Save the previous doc
        currentDocCopy.saveAndClose();
        flaggrn = 1;
      }
      currentDate = formattedDate;
      currentGRN = grn;
      currentSupplier = supplier;
      currentDONO = dono;
      currentPO = po;
      currentCust = customer;
      currentModel = model;
      flag = 0;

      currentDoc = googleDocTemplate.makeCopy(`[${currentGRN}] ${currentSupplier} - GRN Form (1)` ,  
                                               destinationFolder)
      currentDocCopy = DocumentApp.openById(currentDoc.getId())
      body = currentDocCopy.getBody();
      //get the URL of the doc
      const docUrl = currentDoc.getUrl();
      sheet.getRange(i+1, linkCol).setValue(docUrl);
    } 
     // If the current row is for the same customer, append the values to the current doc
    else if (supplier == currentSupplier && supplier != "") {
      if (currentDocCopy) {
        if(flag > 0) {
          // Clear placeholders from {{NO2}} to {{REMARK13}}
          for (let j = 2; j <= 130; j++) {
            body.replaceText(`{{NO${j}}}`, "");
            body.replaceText(`{{DESC${j}}}`, "");
            body.replaceText(`{{PN${j}}}`, "");
            body.replaceText(`{{QTYDO${j}}}`, "");
            body.replaceText(`{{REMARK${j}}}`, "");
          }
        }
        // Save the previous doc
        currentDocCopy.saveAndClose();
        flaggrn = 1;
      }
      currentDate = formattedDate;
      currentGRN = grn;
      currentSupplier = supplier;
      currentDONO = dono;
      currentPO = po;
      currentCust = customer;
      currentModel = model;
      flag = 0; 

      currentDoc = googleDocTemplate.makeCopy(`[${currentGRN}] ${currentSupplier} - GRN Form (1)` , 
                                               destinationFolder)
      currentDocCopy = DocumentApp.openById(currentDoc.getId())
      body = currentDocCopy.getBody();
      //get the URL of the doc
      const docUrl = currentDoc.getUrl();
      sheet.getRange(i+1, linkCol).setValue(docUrl);
    }

    // If flag is more than 13 then save and create new docs for item 14 and More.
    if (flag >= 13 && flag <= 26 && !flagtrig13) {
      currentDocCopy.saveAndClose();
      flaggrn = 1;

      currentDoc = googleDocTemplateMoreThan13.makeCopy(`[${currentGRN}] ${currentSupplier} - GRN Form (2)` , 
                   destinationFolder);
      currentDocCopy = DocumentApp.openById(currentDoc.getId())
      body = currentDocCopy.getBody();
      //get the URL of the doc
      const docUrl = currentDoc.getUrl();
      sheet.getRange(i+1, linkCol).setValue(docUrl);
      // Set the flag to true
      flagtrig13 = true;
    } 

    // If flag is more than 26 then save and create new docs for item 27 and More.
    else if (flag >=26 && flag <= 39 && !flagtrig26) {
      currentDocCopy.saveAndClose();
      flaggrn = 1;

      currentDoc = googleDocTemplateMoreThan26.makeCopy(`[${currentGRN}] ${currentSupplier} - GRN Form (3)` , 
                   destinationFolder);
      currentDocCopy = DocumentApp.openById(currentDoc.getId())
      body = currentDocCopy.getBody();
      //get the URL of the doc
      const docUrl = currentDoc.getUrl();
      sheet.getRange(i+1, linkCol).setValue(docUrl);
      // Set the flag to true
      flagtrig26 = true;            
    }

    // 39 > < 52
    else if (flag >= 39 && flag <= 52 && !flagtrig39) {
      currentDocCopy.saveAndClose();
      flaggrn = 1;

      currentDoc = googleDocTemplateMoreThan39.makeCopy(`[${currentGRN}] ${currentSupplier} - GRN Form (4)` , 
                   destinationFolder);
      currentDocCopy = DocumentApp.openById(currentDoc.getId())
      body = currentDocCopy.getBody();
      //get the URL of the doc
      const docUrl = currentDoc.getUrl();
      sheet.getRange(i+1, linkCol).setValue(docUrl);
      // Set the flag to true
      flagtrig39 = true;        
    }

    // 52 > < 65
    else if (flag >= 52 && flag <= 65 && !flagtrig52) {
      currentDocCopy.saveAndClose();
      flaggrn = 1;

      currentDoc = googleDocTemplateMoreThan52.makeCopy(`[${currentGRN}] ${currentSupplier} - GRN Form (5)` , 
                   destinationFolder);
      currentDocCopy = DocumentApp.openById(currentDoc.getId())
      body = currentDocCopy.getBody();
      //get the URL of the doc
      const docUrl = currentDoc.getUrl();
      sheet.getRange(i+1, linkCol).setValue(docUrl);
      // Set the flag to true
      flagtrig52 = true; 
    }

    // 65 > < 78
    else if (flag >= 65 && flag <= 78 && !flagtrig65) {
      currentDocCopy.saveAndClose();
      flaggrn = 1;

      currentDoc = googleDocTemplateMoreThan65.makeCopy(`[${currentGRN}] ${currentSupplier} - GRN Form (6)` , 
                   destinationFolder);
      currentDocCopy = DocumentApp.openById(currentDoc.getId())
      body = currentDocCopy.getBody();
      //get the URL of the doc
      const docUrl = currentDoc.getUrl();
      sheet.getRange(i+1, linkCol).setValue(docUrl);
      // Set the flag to true
      flagtrig65 = true; 
    }

    // 78 > < 91
    else if (flag >= 78 && flag <= 91 && !flagtrig78) {
      currentDocCopy.saveAndClose();
      flaggrn = 1;

      currentDoc = googleDocTemplateMoreThan78.makeCopy(`[${currentGRN}] ${currentSupplier} - GRN Form (7)` , 
                   destinationFolder);
      currentDocCopy = DocumentApp.openById(currentDoc.getId())
      body = currentDocCopy.getBody();
      //get the URL of the doc
      const docUrl = currentDoc.getUrl();
      sheet.getRange(i+1, linkCol).setValue(docUrl);
      // Set the flag to true
      flagtrig78 = true; 
    }

    // 91 > < 104
    else if (flag >= 91 && flag <= 104 && !flagtrig91) {
      currentDocCopy.saveAndClose();
      flaggrn = 1;

      currentDoc = googleDocTemplateMoreThan91.makeCopy(`[${currentGRN}] ${currentSupplier} - GRN Form (8)` , 
                   destinationFolder);
      currentDocCopy = DocumentApp.openById(currentDoc.getId())
      body = currentDocCopy.getBody();
      //get the URL of the doc
      const docUrl = currentDoc.getUrl();
      sheet.getRange(i+1, linkCol).setValue(docUrl);
      // Set the flag to true
      flagtrig91 = true; 
    }

    // 104 > < 117
    else if (flag >= 104 && flag <= 117 && !flagtrig104) {
      currentDocCopy.saveAndClose();
      flaggrn = 1;

      currentDoc = googleDocTemplateMoreThan104.makeCopy(`[${currentGRN}] ${currentSupplier} - GRN Form (9)` , 
                   destinationFolder);
      currentDocCopy = DocumentApp.openById(currentDoc.getId())
      body = currentDocCopy.getBody();
      //get the URL of the doc
      const docUrl = currentDoc.getUrl();
      sheet.getRange(i+1, linkCol).setValue(docUrl);
      // Set the flag to true
      flagtrig104 = true; 
    }

    // 117 > < 130
    else if (flag >= 117 && flag <= 130 && !flagtrig117) {
      currentDocCopy.saveAndClose();
      flaggrn = 1;

      currentDoc = googleDocTemplateMoreThan117.makeCopy(`[${currentGRN}] ${currentSupplier} - GRN Form (10)` , 
                   destinationFolder);
      currentDocCopy = DocumentApp.openById(currentDoc.getId())
      body = currentDocCopy.getBody();
      //get the URL of the doc
      const docUrl = currentDoc.getUrl();
      sheet.getRange(i+1, linkCol).setValue(docUrl);
      // Set the flag to true
      flagtrig117 = true; 
    }

    // Replace the placeholders with the values from the Spreadsheet
    //HEADER
    body.replaceText("{{DATE}}", currentDate);
    body.replaceText('{{GRNNO}}', currentGRN);
    body.replaceText('{{SUPPLIER}}', currentSupplier);
    body.replaceText('{{SDONO}}', currentDONO);
    body.replaceText('{{CGPO}}', currentPO);
    body.replaceText('{{CUS}}', currentCust);
    body.replaceText('{{MODEL}}', currentModel);
    //BODY
    body.replaceText("{{NO" + itemno + "}}", itemno);
    body.replaceText("{{DESC" + itemno + "}}", desc);
    body.replaceText("{{PN" + itemno + "}}", pn);
    body.replaceText("{{QTYDO" + itemno + "}}", qty);
    body.replaceText("{{REMARK" + itemno + "}}", rmk);

    // Update "Status" column in current row to "FINISHED"
    sheet.getRange(i+1, statusCol).setValue("FINISHED");

    //FOR FLAG
    flag++;
  }
  
  // Save and close the last document
  if (currentDocCopy) {
    if(flag > 0) {
      // Clear placeholders from {{NO2}} to {{REMARK130}}
      for (var i = 2; i <= 130; i++) {
      body.replaceText(`{{NO${i}}}`, "");
      body.replaceText(`{{DESC${i}}}`, "");
      body.replaceText(`{{PN${i}}}`, "");
      body.replaceText(`{{QTYDO${i}}}`, "");
      body.replaceText(`{{REMARK${i}}}`, "");
      }
    } 
    currentDocCopy.saveAndClose();
    flaggrn = 1;
    //RESET FLAG
    flag = 0;
    flagtrig13 = false; flagtrig26 = false; flagtrig39 = false; flagtrig52 = false;
    flagtrig65 = false; flagtrig78 = false; flagtrig91 = false; flagtrig104 = false;
    flagtrig117 = false;
  }
  
  if (flaggrn == 1){
    Browser.msgBox("Completed: A new GRN Form has been generated.");
  } else {
    Browser.msgBox('NO GRN Form was generated.\\n\\nReason: The status is already "FINISHED" OR\\n there is no new  line for me to generate.');
  } 
}





function createGRNFormV2() {
  // Declare variables to hold the Google Doc templates
  var googleDocTemplate = DriveApp.getFileById('15WIbaGXKBlk0Utl6COInkBCT8muBxm5pLBgAhzze9rk');
  var googleDocTemplateMoreThan13 = DriveApp.getFileById('1Jyo7Nxq1fNOpMP4NpT47TVFe_ci1MJD8MxNRYfnVNDc');
  var googleDocTemplateMoreThan26 = DriveApp.getFileById('1XYM9_b-516FLkAf5opyRXbYZ7qboVb1mitKSUU_r27M');
  var googleDocTemplateMoreThan39 = DriveApp.getFileById('1MzS9-YA1I02kBiELbMtdaFIBGZYlCAf7cA7gCX5IF-g');
  var googleDocTemplateMoreThan52 = DriveApp.getFileById('1CRv22SCMHWMIyFkpuaFY_oPvuDoslaLKYw3ax0EpBMI');
  var googleDocTemplateMoreThan65 = DriveApp.getFileById('1_sUN2xmbvEY9Ou1ifKEbGSTfai1rzzDlWqRw4U9Nbx4');
  var googleDocTemplateMoreThan78 = DriveApp.getFileById('15RWcuy0rwHBy6Mun6HvDWQHcJzfVIpQY9SxFA3ZFAoA');
  var googleDocTemplateMoreThan91 = DriveApp.getFileById('1MqKvoeLhTMwb7CxWp8zhuw2UrhQFRVnHV1QfRtvNgKw');
  var googleDocTemplateMoreThan104 = DriveApp.getFileById('1Y3hdfibMSKwg17ww5Iu4v73tu8QiFAgOvclKprqxZ00');
  var googleDocTemplateMoreThan117 = DriveApp.getFileById('1kUeaW84hiUVQuQ98oOCHziwchi0gafs6rs0UEImDugY');

  // Declare a variable to hold the destination folder ID
  var destinationFolder = DriveApp.getFolderById('11RMjq0sW2th-fTiWRQyoUBlqN-T2ZSMq')

  var sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('GRN2022')

  // Setup Column
  var noCol = 1, dateCol = 2, supplierCol = 3, customerCol = 4, donoCol = 5, grnCol = 6, poCol = 7;
  var modelCol = 8, itemnoCol = 9, descCol = 10, pnCol = 11, qtyCol = 12, rmkCol = 13;
  var linkCol= 14, statusCol = 15;
  
  // Get rows from DataRange
  var rows = sheet.getDataRange().getValues();

  // Setup Variable
  var currentDate = "", currentGRN = "", currentSupplier = "", currentDONO = "", currentPO = "";
  var currentCust = "", currentModel = "";
  var currentDoc, currentDocCopy, body;

  // Setup flag
  var flag = 0, flagtrig13 = false, flagtrig26 = false, flagtrig39 = false, flagtrig52 = false;
  var flagtrig65 = false, flagtrig78 = false, flagtrig91 = false, flagtrig104 = false, flagtrig117 = false;

  // Setup flag for check if any is grn is generate
  var flaggrn = 0;

  // Loop to get all data in rows 
  for (var i = 0; i < rows.length; i++) {  
    // retrive data from each row.
    var row = rows[i],  no = row[noCol-1], date = new Date(row[dateCol-1]), supplier = row[supplierCol-1];
    var customer = row[customerCol-1], dono = row[donoCol-1], grn = row[grnCol-1], po = row[poCol-1];
    var model = row[modelCol-1], itemno = row[itemnoCol-1], desc = row[descCol-1], pn = row[pnCol-1]; 
    var qty = row[qtyCol-1], rmk = row[rmkCol-1];

    // Setup for DATE STYLING
    var day = String(date.getDate()).padStart(2, '0');
    var month = String(date.getMonth() + 1).padStart(2, '0');
    var year = date.getFullYear();
    var formattedDate = day + "/" + month + "/" + year;

    // Skip with condition
    if (row[1] === "GOOD RECEIVING NOTE (GRN)" || row[1] === "Year : " || row[1] === "DATE" || 
        row[14] != "UNFINISHED") {
      continue;
    }
    
    // If the current row is for a new supplier, create a new doc
    if (supplier != currentSupplier && supplier != "") {
      if (currentDocCopy) {
        if(flag > 0) {
          // Clear placeholders from {{NO2}} to {{REMARK13}}
          for (var j = 2; j <= 52; j++) {
            body.replaceText(`{{NO${j}}}`, "");
            body.replaceText(`{{DESC${j}}}`, "");
            body.replaceText(`{{PN${j}}}`, "");
            body.replaceText(`{{QTYDO${j}}}`, "");
            body.replaceText(`{{REMARK${j}}}`, "");
          }
        }
        // Save the previous doc
        currentDocCopy.saveAndClose();
        flaggrn = 1;
      }
      currentDate = formattedDate;
      currentGRN = grn;
      currentSupplier = supplier;
      currentDONO = dono;
      currentPO = po;
      currentCust = customer;
      currentModel = model;
      flag = 0;

      currentDoc = googleDocTemplate.makeCopy(`[${currentGRN}] ${currentSupplier} - GRN Form (1)` ,  
                                               destinationFolder)
      currentDocCopy = DocumentApp.openById(currentDoc.getId())
      body = currentDocCopy.getBody();
      //get the URL of the doc
      var docUrl = currentDoc.getUrl();
      sheet.getRange(i+1, linkCol).setValue(docUrl);
    } 
     // If the current row is for the same customer, append the values to the current doc
    else if (supplier == currentSupplier && supplier != "") {
      if (currentDocCopy) {
        if(flag > 0) {
          // Clear placeholders from {{NO2}} to {{REMARK13}}
          for (var j = 2; j <= 130; j++) {
            body.replaceText(`{{NO${j}}}`, "");
            body.replaceText(`{{DESC${j}}}`, "");
            body.replaceText(`{{PN${j}}}`, "");
            body.replaceText(`{{QTYDO${j}}}`, "");
            body.replaceText(`{{REMARK${j}}}`, "");
          }
        }
        // Save the previous doc
        currentDocCopy.saveAndClose();
        flaggrn = 1;
      }
      currentDate = formattedDate;
      currentGRN = grn;
      currentSupplier = supplier;
      currentDONO = dono;
      currentPO = po;
      currentCust = customer;
      currentModel = model;
      flag = 0; 

      currentDoc = googleDocTemplate.makeCopy(`[${currentGRN}] ${currentSupplier} - GRN Form (1)` , 
                                               destinationFolder)
      currentDocCopy = DocumentApp.openById(currentDoc.getId())
      body = currentDocCopy.getBody();
      //get the URL of the doc
      var docUrl = currentDoc.getUrl();
      sheet.getRange(i+1, linkCol).setValue(docUrl);
    }

    // If flag is more than 13 then save and create new docs for item 14 and More.
    if (flag >= 13 && flag <= 26 && !flagtrig13) {
      currentDocCopy.saveAndClose();
      flaggrn = 1;

      currentDoc = googleDocTemplateMoreThan13.makeCopy(`[${currentGRN}] ${currentSupplier} - GRN Form (2)` , 
                   destinationFolder);
      currentDocCopy = DocumentApp.openById(currentDoc.getId())
      body = currentDocCopy.getBody();
      //get the URL of the doc
      var docUrl = currentDoc.getUrl();
      sheet.getRange(i+1, linkCol).setValue(docUrl);
      // Set the flag to true
      flagtrig13 = true;
    } 

    // If flag is more than 26 then save and create new docs for item 27 and More.
    else if (flag >=26 && flag <= 39 && !flagtrig26) {
      currentDocCopy.saveAndClose();
      flaggrn = 1;

      currentDoc = googleDocTemplateMoreThan26.makeCopy(`[${currentGRN}] ${currentSupplier} - GRN Form (3)` , 
                   destinationFolder);
      currentDocCopy = DocumentApp.openById(currentDoc.getId())
      body = currentDocCopy.getBody();
      //get the URL of the doc
      var docUrl = currentDoc.getUrl();
      sheet.getRange(i+1, linkCol).setValue(docUrl);
      // Set the flag to true
      flagtrig26 = true;            
    }

    // 39 > < 52
    else if (flag >= 39 && flag <= 52 && !flagtrig39) {
      currentDocCopy.saveAndClose();
      flaggrn = 1;

      currentDoc = googleDocTemplateMoreThan39.makeCopy(`[${currentGRN}] ${currentSupplier} - GRN Form (4)` , 
                   destinationFolder);
      currentDocCopy = DocumentApp.openById(currentDoc.getId())
      body = currentDocCopy.getBody();
      //get the URL of the doc
      var docUrl = currentDoc.getUrl();
      sheet.getRange(i+1, linkCol).setValue(docUrl);
      // Set the flag to true
      flagtrig39 = true;        
    }

    // 52 > < 65
    else if (flag >= 52 && flag <= 65 && !flagtrig52) {
      currentDocCopy.saveAndClose();
      flaggrn = 1;

      currentDoc = googleDocTemplateMoreThan52.makeCopy(`[${currentGRN}] ${currentSupplier} - GRN Form (5)` , 
                   destinationFolder);
      currentDocCopy = DocumentApp.openById(currentDoc.getId())
      body = currentDocCopy.getBody();
      //get the URL of the doc
      var docUrl = currentDoc.getUrl();
      sheet.getRange(i+1, linkCol).setValue(docUrl);
      // Set the flag to true
      flagtrig52 = true; 
    }

    // 65 > < 78
    else if (flag >= 65 && flag <= 78 && !flagtrig65) {
      currentDocCopy.saveAndClose();
      flaggrn = 1;

      currentDoc = googleDocTemplateMoreThan65.makeCopy(`[${currentGRN}] ${currentSupplier} - GRN Form (6)` , 
                   destinationFolder);
      currentDocCopy = DocumentApp.openById(currentDoc.getId())
      body = currentDocCopy.getBody();
      //get the URL of the doc
      var docUrl = currentDoc.getUrl();
      sheet.getRange(i+1, linkCol).setValue(docUrl);
      // Set the flag to true
      flagtrig65 = true; 
    }

    // 78 > < 91
    else if (flag >= 78 && flag <= 91 && !flagtrig78) {
      currentDocCopy.saveAndClose();
      flaggrn = 1;

      currentDoc = googleDocTemplateMoreThan78.makeCopy(`[${currentGRN}] ${currentSupplier} - GRN Form (7)` , 
                   destinationFolder);
      currentDocCopy = DocumentApp.openById(currentDoc.getId())
      body = currentDocCopy.getBody();
      //get the URL of the doc
      var docUrl = currentDoc.getUrl();
      sheet.getRange(i+1, linkCol).setValue(docUrl);
      // Set the flag to true
      flagtrig78 = true; 
    }

    // 91 > < 104
    else if (flag >= 91 && flag <= 104 && !flagtrig91) {
      currentDocCopy.saveAndClose();
      flaggrn = 1;

      currentDoc = googleDocTemplateMoreThan91.makeCopy(`[${currentGRN}] ${currentSupplier} - GRN Form (8)` , 
                   destinationFolder);
      currentDocCopy = DocumentApp.openById(currentDoc.getId())
      body = currentDocCopy.getBody();
      //get the URL of the doc
      var docUrl = currentDoc.getUrl();
      sheet.getRange(i+1, linkCol).setValue(docUrl);
      // Set the flag to true
      flagtrig91 = true; 
    }

    // 104 > < 117
    else if (flag >= 104 && flag <= 117 && !flagtrig104) {
      currentDocCopy.saveAndClose();
      flaggrn = 1;

      currentDoc = googleDocTemplateMoreThan104.makeCopy(`[${currentGRN}] ${currentSupplier} - GRN Form (9)` , 
                   destinationFolder);
      currentDocCopy = DocumentApp.openById(currentDoc.getId())
      body = currentDocCopy.getBody();
      //get the URL of the doc
      var docUrl = currentDoc.getUrl();
      sheet.getRange(i+1, linkCol).setValue(docUrl);
      // Set the flag to true
      flagtrig104 = true; 
    }

    // 117 > < 130
    else if (flag >= 117 && flag <= 130 && !flagtrig117) {
      currentDocCopy.saveAndClose();
      flaggrn = 1;

      currentDoc = googleDocTemplateMoreThan117.makeCopy(`[${currentGRN}] ${currentSupplier} - GRN Form (10)` , 
                   destinationFolder);
      currentDocCopy = DocumentApp.openById(currentDoc.getId())
      body = currentDocCopy.getBody();
      //get the URL of the doc
      var docUrl = currentDoc.getUrl();
      sheet.getRange(i+1, linkCol).setValue(docUrl);
      // Set the flag to true
      flagtrig117 = true; 
    }

    // Replace the placeholders with the values from the Spreadsheet
    //HEADER
    body.replaceText("{{DATE}}", currentDate);
    body.replaceText('{{GRNNO}}', currentGRN);
    body.replaceText('{{SUPPLIER}}', currentSupplier);
    body.replaceText('{{SDONO}}', currentDONO);
    body.replaceText('{{CGPO}}', currentPO);
    body.replaceText('{{CUS}}', currentCust);
    body.replaceText('{{MODEL}}', currentModel);
    //BODY
    body.replaceText("{{NO" + itemno + "}}", itemno);
    body.replaceText("{{DESC" + itemno + "}}", desc);
    body.replaceText("{{PN" + itemno + "}}", pn);
    body.replaceText("{{QTYDO" + itemno + "}}", qty);
    body.replaceText("{{REMARK" + itemno + "}}", rmk);

    // Update "Status" column in current row to "FINISHED"
    sheet.getRange(i+1, statusCol).setValue("FINISHED");

    //FOR FLAG
    flag++;
  }
  
  // Save and close the last document
  if (currentDocCopy) {
    if(flag > 0) {
      // Clear placeholders from {{NO2}} to {{REMARK26}}
      for (var i = 2; i <= 130; i++) {
      body.replaceText(`{{NO${i}}}`, "");
      body.replaceText(`{{DESC${i}}}`, "");
      body.replaceText(`{{PN${i}}}`, "");
      body.replaceText(`{{QTYDO${i}}}`, "");
      body.replaceText(`{{REMARK${i}}}`, "");
      }
    } 
    currentDocCopy.saveAndClose();
    flaggrn = 1;
    //RESET FLAG
    flag = 0;
    flagtrig13 = false; flagtrig26 = false; flagtrig39 = false; flagtrig52 = false;
    flagtrig65 = false; flagtrig78 = false; flagtrig91 = false; flagtrig104 = false;
    flagtrig117 = false;
  }
  
  if (flaggrn == 1){
    Browser.msgBox("Completed: A new GRN Form has been generated.");
  } else {
    Browser.msgBox('NO GRN Form was generated.\\n\\nReason: The status is already "FINISHED" OR\\n there is no new  line for me to generate.');
  } 
}