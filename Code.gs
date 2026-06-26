function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const infoSheet = ss.getSheetByName('info');
  const lastRow = infoSheet.getLastRow();
  let flavors = [];
  
  if (lastRow >= 2) {
    flavors = infoSheet.getRange("A2:A" + lastRow).getValues().flat().filter(String);
  }
  
  // Directly wraps content to guarantee a safe evaluation context for the browser
  return ContentService.createTextOutput(JSON.stringify({ flavors: flavors }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let ordersSheet = ss.getSheetByName('orders');
    
    if (!ordersSheet) {
      ordersSheet = ss.insertSheet('orders');
      ordersSheet.appendRow(["Timestamp", "Full Name", "Flavour", "20pcs Qty", "3pcs Qty", "1pcs Qty", "Other Qty", "Total Price"]);
    }
    
    // Accept standard form post parameter keys directly
    const p = e.parameter;
    
    ordersSheet.appendRow([
      new Date(),
      p.fullName,
      p.flavour,
      p.multipack20,
      p.multipack3,
      p.multipack1,
      p.otherQty,
      p.totalPrice
    ]);
    
    // Return plain text success response
    return ContentService.createTextOutput("Success").setMimeType(ContentService.MimeType.TEXT);
  } catch (error) {
    return ContentService.createTextOutput("Error: " + error.toString()).setMimeType(ContentService.MimeType.TEXT);
  }
}
