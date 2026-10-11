/**
 * AZON CSS product-table alternative.
 * B5 remains the original generator's HTML.
 * B40 receives an independent copy with CSS background-image tiles.
 *
 * Run azonCssInstallMenu() ONCE after clasp push.
 * It installs a separate "AZON CSS" menu on spreadsheet open.
 */
function generateAzonCssProductHtml() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('azon');
  if (!sh) throw new Error('Worksheet "azon" was not found.');

  var html = String(sh.getRange('B5').getValue() || '');
  if (!html.trim()) throw new Error('B5 is empty. Generate AZON Product HTML first.');

  var changed = 0;
  var output = html.replace(/<img\b[^>]*>/gi, function(tag) {
    var src = tag.match(/\bsrc\s*=\s*(["'])(.*?)\1/i);
    if (!src || !src[2]) return tag;
    var alt = tag.match(/\balt\s*=\s*(["'])(.*?)\1/i);
    var width = tag.match(/\bwidth\s*=\s*(["']?)(\d+)\1/i);
    var px = width ? Math.min(150, Math.max(60, Number(width[2]))) : 100;
    var safeSrc = src[2].replace(/&amp;/g, '&').replace(/'/g, '%27')
      .replace(/"/g, '%22').replace(/</g, '%3C').replace(/>/g, '%3E');
    var safeAlt = (alt ? alt[2] : 'Product image')
      .replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
    changed++;
    return '<div role="img" aria-label="' + safeAlt + '" ' +
      'style="width:' + px + 'px;height:100px;display:block;margin:0 auto;' +
      'background-image:url(\'' + safeSrc + '\');background-size:contain;' +
      'background-position:center;background-repeat:no-repeat;"></div>';
  });

  if (!changed) throw new Error('No image tags were found in B5; B40 was not changed.');
  sh.getRange('B40').setValue(output);
  SpreadsheetApp.getUi().alert(
    'AZON CSS HTML created in B40. ' + changed +
    ' image(s) converted. Original B5 remains unchanged.'
  );
}

/** Creates the separate custom menu. Called by the installable open trigger. */
function azonCssShowMenu() {
  SpreadsheetApp.getUi()
    .createMenu('AZON CSS')
    .addItem('Generate AZON CSS Product HTML (B40)', 'generateAzonCssProductHtml')
    .addToUi();
}

/** Run once to install the menu on future spreadsheet opens. */
function azonCssInstallMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var exists = ScriptApp.getProjectTriggers().some(function(t) {
    return t.getHandlerFunction() === 'azonCssShowMenu' &&
      t.getEventType() === ScriptApp.EventType.ON_OPEN;
  });
  if (!exists) {
    ScriptApp.newTrigger('azonCssShowMenu').forSpreadsheet(ss).onOpen().create();
  }
  azonCssShowMenu();
}
