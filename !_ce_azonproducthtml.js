/**
 * ce_AzonProductHtml.gs - GENERATE AMAZON PRODUCT HTML FOR ACTIVE ROW
 * Operates entirely within the "azon" sheet. Reads prod-1/prod-2/prod-3
 * (columns B, C, D) and intro-text (column E) from the active row,
 * looks up each product name in column H (prod-name) on the same sheet,
 * and builds an HTML block using column O ("html image + button") for
 * each matched product. Writes the result to cell B5.
 */

function generateAzonProductHTML() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const azonSheet = ss.getSheetByName("azon");
  if (!azonSheet) return "ERROR: 'azon' sheet not found.";

  const row = 2; // Always reads from row 2 — no need to select a cell first

  const prod1 = String(azonSheet.getRange(row, 2).getValue() || "").trim(); // B
  const prod2 = String(azonSheet.getRange(row, 3).getValue() || "").trim(); // C
  const prod3 = String(azonSheet.getRange(row, 4).getValue() || "").trim(); // D
  const introText = String(azonSheet.getRange(row, 5).getValue() || "").trim(); // E

  const azonData = azonSheet.getDataRange().getValues();

  function escapeHTML(text) {
    return String(text || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function findProductRow(prodName) {
    if (!prodName) return null;
    const target = prodName.trim().toLowerCase();
    for (let i = 1; i < azonData.length; i++) {
      const name = String(azonData[i][7] || "").trim(); // H = prod-name
      if (name.toLowerCase() === target) {
        return azonData[i];
      }
    }
    return null;
  }

  const products = [prod1, prod2, prod3];
  let cellsHtml = "";
  let notFound = [];

  products.forEach(function(prodName) {
    if (!prodName) return;
    const productRow = findProductRow(prodName);
    if (!productRow) { notFound.push(prodName); return; }

    const name = escapeHTML(productRow[7]);   // H = prod-name
    const imgLink = escapeHTML(productRow[8]); // I = img-link
    const affLink = escapeHTML(productRow[9]); // J = aff-link

    cellsHtml +=
      '<td style="border: 1px solid #ddd; border-radius: 8px; padding: 10px; text-align: center; vertical-align: top;">' +
      '<img style="max-height: 100px; width: auto; display: block; margin: 0 auto;" src="' + imgLink + '" alt="' + name + '" width="100" />\n' +
      '<div style="font-weight: bold; margin: 10px 0;">' + name + '</div>\n' +
      '<a style="display: inline-block; background: #28a745; color: #fff; padding: 8px 8px; border-radius: 10px; font-weight: bold; text-decoration: none;" href="' + affLink + '" target="_blank" rel="nofollow noopener">Shop Now</a></td>\n';
  });

  if (!cellsHtml) {
    return "ERROR: No matching products found in 'azon' sheet for: " + products.filter(String).join(", ");
  }

  let template =
    '<!-- New Azon -->\n' +
    '<h2 style="background: #f8f9fa; border-left: 5px solid #28a745; padding: 10px 15px; font-weight: bold; font-size: 1.1em; margin-bottom: 20px; border-radius: 4px;">' + escapeHTML(introText) + '</h2>\n' +
    '<table style="width: 100%; border-spacing: 10px;">\n' +
    '<tbody>\n' +
    '<tr>\n' +
    cellsHtml +
    '</tr>\n' +
    '</tbody>\n' +
    '</table>';

  if (notFound.length > 0) {
    template += '\n<!-- NOT FOUND IN AZON SHEET: ' + notFound.join(', ') + ' -->';
  }

  azonSheet.getRange('B5').setValue(template);

  return template;
}
function getB5Content() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const azonSheet = ss.getSheetByName("azon");
  return String(azonSheet.getRange('B5').getValue() || "");
}

function openCopyB5Dialog() {
  const html = HtmlService.createHtmlOutputFromFile('ce_Dialog_CopyB5')
    .setWidth(250)
    .setHeight(80);
  SpreadsheetApp.getUi().showModalDialog(html, 'Copying…');
}