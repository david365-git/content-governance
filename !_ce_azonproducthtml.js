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
  let cardsHtml = "";
  let notFound = [];

  products.forEach(function(prodName) {
    if (!prodName) return;
    const productRow = findProductRow(prodName);
    if (!productRow) { notFound.push(prodName); return; }

    const name = escapeHTML(productRow[7]);    // H = prod-name
    const imgLink = escapeHTML(productRow[8]); // I = img-link
    const affLink = escapeHTML(productRow[9]); // J = aff-link

    cardsHtml +=
      '<div style="border: 1px solid #ddd; border-radius: 8px; padding: 16px; text-align: center; display: flex; flex-direction: column; justify-content: space-between; gap: 12px; background: #fff;">\n' +
      '<div>\n' +
      '<img style="height: 100px; width: auto; max-width: 100%; display: block; margin: 0 auto;" src="' + imgLink + '" alt="' + name + ' product image" width="100" height="100" loading="lazy" />\n' +
      '<div style="font-weight: bold; margin: 12px 0;">' + name + '</div>\n' +
      '</div>\n' +
      '<div><a style="display: inline-block; background: #28a745; color: #fff; padding: 8px 16px; border-radius: 8px; font-weight: bold; text-decoration: none;" href="' + affLink + '" target="_blank" rel="sponsored nofollow noopener noreferrer">Check Price on Amazon</a></div>\n' +
      '</div>\n';
  });

  if (!cardsHtml) {
    return "ERROR: No matching products found in 'azon' sheet for: " + products.filter(String).join(", ");
  }

  let template =
    '<!-- New Azon -->\n' +
    '<div class="abbey-affiliate-box" style="margin: 24px 0;">\n' +
    '<p style="background: #f8f9fa; border-left: 5px solid #28a745; padding: 10px 15px; font-weight: bold; font-size: 1.1em; margin: 0 0 8px 0; border-radius: 4px;">' + escapeHTML(introText) + '</p>\n' +
    '<p style="font-size: 0.85em; color: #555; margin: 0 0 16px 0; font-style: italic;">Disclosure: As an Amazon Associate, Abbey Floor Care earns from qualifying purchases at no extra cost to you.</p>\n' +
    '<div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(min(100%, 220px), 1fr)); gap: 16px;">\n' +
    cardsHtml +
    '</div>\n' +
    '</div>';
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