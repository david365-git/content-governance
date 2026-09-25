/**
 * STANDALONE HTML CLEANUP
 * Removes specific Ninja, Figure, and Amazon blocks from the active cell.
 * Writes the cleaned content to Column T of the same row.
 */
function cleanAndTransferHtmlStandalone() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const activeCell = sheet.getActiveCell();
  const rowIndex = activeCell.getRowIndex();
  
  // Basic Validation
  if (rowIndex < 1) {
    Browser.msgBox("Selection Error", "Please select a cell containing HTML content.", Browser.Buttons.OK);
    return;
  }

  let content = activeCell.getValue();
  
  // Ensure we are working with a string
  if (typeof content !== 'string' || content === "") {
    ss.toast("The active cell is empty or doesn't contain text.", "⚠️ Error", 5);
    return;
  }

  // Regex Patterns for removal
  const patterns = [
    // 1. Image Ninja Pro Wrapper
    /<div class="image-ninja-pro-wrapper aligncenter"[\s\S]*?>/g,
    
    // 2. Lazy Load Images
    /<img loading="lazy"[\s\S]*?>/g,
    
    // 3. Figure tags with IDs
    /<figure id=[\s\S]*?>/g,
    
    // 4. Amazon Blocks
    /<p>[\s\S]*?<\/table>/g,
    
    // 5. Authority Ninja Box
    /[\s\S]*?<\/div><\/div>/g
  ];

  // Process the text
  let cleanedContent = content;
  patterns.forEach(regex => {
    cleanedContent = cleanedContent.replace(regex, '');
  });

  // Write results to Column T (Column 20)
  sheet.getRange(rowIndex, 20).setValue(cleanedContent);
  
  // Provide feedback via Toast (bottom-right notification)
  ss.toast("Cleaned HTML has been written to Column T.", "✅ Success", 3);
}