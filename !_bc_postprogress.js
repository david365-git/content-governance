function bc_getPostProgressStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var postsSheet = ss.getSheetByName('posts');
  var progressSheet = ss.getSheetByName('posts_progress');

  // Get active Post ID from posts sheet
  var activeRow = postsSheet.getActiveCell().getRow();
  var postId = postsSheet.getRange(activeRow, 4).getValue();

  if (!postId) {
    return {
      exists: false,
      stage: 0,
      label: 'PRODUCTION — no Post ID found in active row',
      bgColor: '#e37400',
      textColor: 'white'
    };
  }

  // Find Post ID in posts_progress
  var progressData = progressSheet.getDataRange().getValues();
  var headers = progressData[0];
  var colIndex = {};
  for (var i = 0; i < headers.length; i++) {
    colIndex[headers[i]] = i;
  }

  var postRow = null;
  for (var r = 1; r < progressData.length; r++) {
    if (String(progressData[r][colIndex['Post ID']]) === String(postId)) {
      postRow = progressData[r];
      break;
    }
  }

  // New post — not found in posts_progress
  if (!postRow) {
    progressSheet.appendRow([postId]);
    return {
      exists: false,
      stage: 1,
      label: 'PRODUCTION — STAGE 1 — New post added to progress sheet',
      bgColor: '#1a73e8',
      textColor: 'white'
    };
  }

  var ce1date = postRow[colIndex['CE-1-date']];
  var ce2date = postRow[colIndex['CE-2-date']];
  var wpDate  = postRow[colIndex['WP-date']];

  // Stage determination
  if (wpDate) {
    return {
      exists: true,
      stage: 4,
      label: 'PRODUCTION — PUBLISHED — Published on ' + wpDate,
      bgColor: '#1e8e3e',
      textColor: 'white'
    };
  }

  if (ce2date) {
    return {
      exists: true,
      stage: 3,
      label: 'PRODUCTION — CONTENT COMPLETE — CE-2 done on ' + ce2date,
      bgColor: '#0f6e56',
      textColor: 'white'
    };
  }

  var elg2date = postRow[colIndex['ELG-2-date']];

  if (ce1date && elg2date) {
    return {
      exists: true,
      stage: 2,
      label: 'PRODUCTION — STAGE 2 — Run: W2A → W2B → W3 → W4 → W5 → W6 → W7',
      bgColor: '#1a73e8',
      textColor: 'white'
    };
  }

  if (ce1date && !elg2date) {
    return {
      exists: true,
      stage: 1,
      label: 'PRODUCTION — STAGE 1 — CE-1 complete, awaiting ELG-2',
      bgColor: '#1a73e8',
      textColor: 'white'
    };
  }

  return {
    exists: true,
    stage: 1,
    label: 'PRODUCTION — STAGE 1',
    bgColor: '#1a73e8',
    textColor: 'white'
  };
}
function bc_writeELGToProgress(stage) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var postsSheet = ss.getSheetByName('posts');
  var progressSheet = ss.getSheetByName('posts_progress');

  var activeRow = postsSheet.getActiveCell().getRow();
  var postId = postsSheet.getRange(activeRow, 4).getValue();

  var progressData = progressSheet.getDataRange().getValues();
  var headers = progressData[0];
  var colIndex = {};
  for (var i = 0; i < headers.length; i++) {
    colIndex[headers[i]] = i;
  }

  for (var r = 1; r < progressData.length; r++) {
    if (String(progressData[r][colIndex['Post ID']]) === String(postId)) {
      var now = new Date();
      if (stage === 1 || stage === undefined) {
        progressSheet.getRange(r + 1, colIndex['ELG-1'] + 1).setValue('Done');
        progressSheet.getRange(r + 1, colIndex['ELG-1-date'] + 1).setValue(now);
      } else if (stage === 2) {
        progressSheet.getRange(r + 1, colIndex['ELG-2'] + 1).setValue('Done');
        progressSheet.getRange(r + 1, colIndex['ELG-2-date'] + 1).setValue(now);
      }
      return 'ELG Stage ' + stage + ' written to progress sheet.';
    }
  }

  return 'ELG write failed — Post ID not found in posts_progress.';
}
function bc_writeCEToProgress(stage) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var postsSheet = ss.getSheetByName('posts');
  var progressSheet = ss.getSheetByName('posts_progress');

  var activeRange = postsSheet.getActiveRange();
  if (!activeRange) return 'Error: No active selection found on posts sheet.';

  var activeRow = activeRange.getRow();
  var postId = postsSheet.getRange(activeRow, 4).getValue();

  if (!postId) {
    return 'Error: No Post ID found in column D of row ' + activeRow;
  }

  var progressData = progressSheet.getDataRange().getValues();
  var headers = progressData[0];
  var colIndex = {};
  for (var i = 0; i < headers.length; i++) {
    colIndex[headers[i]] = i;
  }

  var stageNum = parseInt(stage) || 1;

  for (var r = 1; r < progressData.length; r++) {
    if (String(progressData[r][colIndex['Post ID']]) === String(postId)) {
      var now = new Date();
      if (stageNum === 1) {
        progressSheet.getRange(r + 1, colIndex['CE-1'] + 1).setValue('Done');
        progressSheet.getRange(r + 1, colIndex['CE-1-date'] + 1).setValue(now);
      } else if (stageNum === 2) {
        progressSheet.getRange(r + 1, colIndex['CE-2'] + 1).setValue('Done');
        progressSheet.getRange(r + 1, colIndex['CE-2-date'] + 1).setValue(now);
      }
      return 'CE Stage ' + stageNum + ' updated for Post ID: ' + postId;
    }
  }

  return 'CE write failed — Post ID (' + postId + ') not found in posts_progress.';
}

function bc_writeWPToProgress() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var postsSheet = ss.getSheetByName('posts');
  var progressSheet = ss.getSheetByName('posts_progress');

  var activeRange = postsSheet.getActiveRange();
  if (!activeRange) return 'Error: No active selection found on posts sheet.';

  var activeRow = activeRange.getRow();
  var postId = postsSheet.getRange(activeRow, 4).getValue();

  if (!postId) {
    return 'Error: No Post ID found in column D of row ' + activeRow;
  }

  var progressData = progressSheet.getDataRange().getValues();
  var headers = progressData[0];
  var colIndex = {};
  for (var i = 0; i < headers.length; i++) {
    colIndex[headers[i]] = i;
  }

  for (var r = 1; r < progressData.length; r++) {
    if (String(progressData[r][colIndex['Post ID']]) === String(postId)) {
      var now = new Date();
      progressSheet.getRange(r + 1, colIndex['WP'] + 1).setValue('Done');
      progressSheet.getRange(r + 1, colIndex['WP-date'] + 1).setValue(now);
      postsSheet.getRange(activeRow, 87).setValue(now);
      var flagCell = postsSheet.getRange(activeRow, 1);
      var existingFlag = String(flagCell.getValue() || '').trim();
      var newFlag = existingFlag.indexOf('H') > -1 ? 'HP' : 'P';
      flagCell.setValue(newFlag);
      return 'WP published — timestamp written to progress sheet and posts sheet for Post ID: ' + postId;
    }
  }

  return 'WP write failed — Post ID (' + postId + ') not found in posts_progress.';
}

function bc_cycleRunColour() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var postsSheet = ss.getSheetByName('posts');
    if (!postsSheet) return { success: false, message: "Posts sheet not found." };

    var activeRow = postsSheet.getActiveRange().getRow();
    if (activeRow < 2) return { success: false, message: "Select a data row first." };

    var cell = postsSheet.getRange(activeRow, 4); // Post ID column
    var currentBg = String(cell.getBackground() || '').toLowerCase();

    var WHITE_VALUES = ['', '#ffffff', '#fff', 'white'];
    var YELLOW = '#ffe600';
    var BURGUNDY = '#800020';

    if (WHITE_VALUES.indexOf(currentBg) > -1) {
      cell.setBackground(YELLOW).setFontColor('#000000');
      return { success: true, message: "Post ID marked yellow (first Content Intelligence run)." };
    } else if (currentBg === YELLOW) {
      cell.setBackground(BURGUNDY).setFontColor('#ffffff');
      return { success: true, message: "Post ID marked burgundy (second Content Intelligence run)." };
    } else if (currentBg === BURGUNDY) {
      return { success: true, message: "Post ID already burgundy — no change." };
    } else {
      // Unexpected existing colour — treat as first run to be safe
      cell.setBackground(YELLOW).setFontColor('#000000');
      return { success: true, message: "Post ID marked yellow (unrecognised prior colour reset)." };
    }

  } catch (e) {
    return { success: false, message: "bc_cycleRunColour error: " + e.message };
  }
}