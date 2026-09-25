function migrateAllQueryData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var postsSheet = ss.getSheetByName('posts');
  if (!postsSheet) {
    SpreadsheetApp.getUi().alert('posts sheet not found.');
    return;
  }

  var periods = [
    { colName: 'Queries',        sheetName: 'GSC Queries'   },
    { colName: 'Basel Queries',  sheetName: 'GSC Baseline'  },
    { colName: 'Benmrk Queries', sheetName: 'GSC Benchmark' },
    { colName: 'Montr Queries',  sheetName: 'GSC Monitor'   }
  ];

  var headers = postsSheet.getRange(1, 1, 1, postsSheet.getLastColumn())
    .getValues()[0]
    .map(function(h) { return String(h).trim(); });

  var postIdIdx = headers.indexOf('Post ID');
  if (postIdIdx === -1) {
    SpreadsheetApp.getUi().alert('Post ID column not found.');
    return;
  }

  var allData = postsSheet.getDataRange().getValues();
  var totalRows = 0;

  periods.forEach(function(period) {
    var colIdx = headers.indexOf(period.colName);
    if (colIdx === -1) {
      Logger.log('Column not found: ' + period.colName);
      return;
    }

    var targetSheet = ss.getSheetByName(period.sheetName);
    if (!targetSheet) {
      targetSheet = ss.insertSheet(period.sheetName);
    } else {
      targetSheet.clearContents();
    }

    targetSheet.getRange(1, 1, 1, 4).setValues([
      ['Post ID', 'Query', 'Clicks', 'Impressions']
    ]);

    var outputRows = [];

    for (var i = 1; i < allData.length; i++) {
      var postId = String(allData[i][postIdIdx] || '').trim();
      if (!postId) continue;

      var rawQueries = String(allData[i][colIdx] || '').trim();
      if (!rawQueries) continue;

      // Normalise all line ending variants
      rawQueries = rawQueries.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

      var lines = rawQueries.split('\n');

      lines.forEach(function(line) {
        line = line.trim();
        if (!line) return;

        // Skip header lines
        if (/^(top queries|clicks|impressions|ctr|position)$/i.test(line)) return;

        // Try tab separator first, then multiple spaces, then single space
        var parts;
        if (line.indexOf('\t') !== -1) {
          parts = line.split('\t');
        } else if (/\s{2,}/.test(line)) {
          parts = line.split(/\s{2,}/);
        } else {
          parts = line.split(/\s+/);
        }

        // Clean all parts
        parts = parts.map(function(p) { return p.trim(); }).filter(function(p) { return p.length > 0; });

        if (parts.length < 2) return;

        // First part is always the query
        // Find the first numeric part after the query — that is clicks
        // Find the second numeric part — that is impressions
        var query = '';
        var clicks = 0;
        var impressions = 0;
        var numericParts = [];
        var queryParts = [];

        for (var p = 0; p < parts.length; p++) {
          // A part is numeric if it is a number or a percentage
          if (/^\d+(\.\d+)?%?$/.test(parts[p])) {
            numericParts.push(parts[p]);
          } else if (numericParts.length === 0) {
            // Still in the query portion
            queryParts.push(parts[p]);
          }
        }

        query = queryParts.join(' ').toLowerCase().trim();
        if (!query) return;

        // Skip if query looks like a header
        if (/^(top queries|clicks|impressions|ctr|position)$/i.test(query)) return;

        clicks = numericParts.length >= 1 ? parseInt(numericParts[0]) || 0 : 0;
        impressions = numericParts.length >= 2 ? parseInt(numericParts[1]) || 0 : 0;

        if (impressions < 1) return;

        outputRows.push([postId, query, clicks, impressions]);
      });
    }

    if (outputRows.length > 0) {
      targetSheet.getRange(2, 1, outputRows.length, 4).setValues(outputRows);
      totalRows += outputRows.length;
      Logger.log(period.sheetName + ': ' + outputRows.length + ' rows written');
    } else {
      Logger.log(period.sheetName + ': no rows found');
    }
  });

  SpreadsheetApp.getUi().alert('Migration complete. ' + totalRows + ' total query rows written across 4 sheets.');
}

function verifyQuerySheetMigration() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var postsSheet = ss.getSheetByName('posts');
  if (!postsSheet) {
    SpreadsheetApp.getUi().alert('posts sheet not found.');
    return;
  }

  var periods = [
    { colName: 'Queries',        sheetName: 'GSC Queries'   },
    { colName: 'Basel Queries',  sheetName: 'GSC Baseline'  },
    { colName: 'Benmrk Queries', sheetName: 'GSC Benchmark' },
    { colName: 'Montr Queries',  sheetName: 'GSC Monitor'   }
  ];

  var headers = postsSheet.getRange(1, 1, 1, postsSheet.getLastColumn())
    .getValues()[0]
    .map(function(h) { return String(h).trim(); });

  var postIdIdx = headers.indexOf('Post ID');
  var allData = postsSheet.getDataRange().getValues();

  var report = 'QUERY SHEET VERIFICATION REPORT\n\n';

  periods.forEach(function(period) {
    var colIdx = headers.indexOf(period.colName);
    var targetSheet = ss.getSheetByName(period.sheetName);

    report += '--- ' + period.sheetName + ' ---\n';

    if (colIdx === -1) {
      report += 'SOURCE COLUMN NOT FOUND: ' + period.colName + '\n\n';
      return;
    }

    if (!targetSheet) {
      report += 'TARGET SHEET NOT FOUND\n\n';
      return;
    }

    // Count posts sheet rows with data in this column
    var postsWithData = 0;
    var totalLinesInPosts = 0;
    for (var i = 1; i < allData.length; i++) {
      var postId = String(allData[i][postIdIdx] || '').trim();
      if (!postId) continue;
      var raw = String(allData[i][colIdx] || '').trim();
      if (!raw) continue;
      postsWithData++;
      var lines = raw.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
      lines.forEach(function(line) {
        line = line.trim();
        if (!line) return;
        if (/^(top queries|clicks|impressions|ctr|position)$/i.test(line)) return;
        totalLinesInPosts++;
      });
    }

    // Count rows in target sheet
    var sheetRows = targetSheet.getLastRow() - 1; // minus header
    var sheetPostIds = new Set();
    if (sheetRows > 0) {
      var sheetData = targetSheet.getRange(2, 1, sheetRows, 1).getValues();
      sheetData.forEach(function(row) {
        if (row[0]) sheetPostIds.add(String(row[0]).trim());
      });
    }

    report += 'Posts with data in source column: ' + postsWithData + '\n';
    report += 'Total query lines in source column: ' + totalLinesInPosts + '\n';
    report += 'Rows written to sheet (excl. header): ' + sheetRows + '\n';
    report += 'Unique Post IDs in sheet: ' + sheetPostIds.size + '\n';

    var diff = totalLinesInPosts - sheetRows;
    if (diff === 0) {
      report += 'STATUS: MATCH\n\n';
    } else {
      report += 'STATUS: DIFFERENCE OF ' + diff + ' rows (lines with 0 impressions are excluded)\n\n';
    }
  });

  Logger.log(report);
  SpreadsheetApp.getUi().alert(report);
}