/* ============================================================
   ce_ProblemRegister
   PROBLEM REGISTER — Dynamic index of problems addressed
   across the entire posts estate.

   Built from:
     Primary Query Cluster Owned  — the problem this page owns
     Observed Query Cluster       — related problems it touches
     Problem Angle                — how this page addresses it
     Article Type                 — what type of page it is
     Entity Role                  — governed function of the page
     Primary Entity               — Material + Intent
     Page Governance Summary      — structured governance summary
     Stone Type                   — which material
     URL                          — page identity
     Title                        — human-readable label

   Sheet: "Problem Register"
   Columns:
     A — Problem (from Primary Query Cluster Owned)
     B — Stone Type
     C — Article Type
     D — Angle Code
     E — Angle Statement
     F — URL
     G — Title
     H — Last Updated

   Used by W-1 to:
     1. Check if the current page's problem is already addressed
     2. Show which articles address it and from which angle
     3. Confirm the active article's angle is genuinely different
     4. Auto-assign problem angle from governance data

   Angle auto-assignment uses:
     Article Type + Entity Role + Primary Entity (intent component)
     to deterministically select the correct angle without
     requiring manual input.

   Conflict classification:
     Geo Service Page vs Geo Service Page (same angle) = Location Variant
     All other same-problem conflicts = Potential Cannibalisation

   Matching strategy: keyword-based intent overlap
   Two intents match if they share 2+ significant keywords.
============================================================ */

// ============================================================
// ANGLE DEFINITIONS
// ============================================================
var ANGLE_DEFINITIONS = {
  'WHY': {
    code: 'WHY',
    label: 'Material Explanation',
    arrow: '→ Why is this happening to my floor?',
    template: 'explains the material-level factors behind {problem}, so you understand what you are actually dealing with'
  },
  'ID': {
    code: 'ID',
    label: 'Diagnosis',
    arrow: '→ Which specific condition do I have?',
    template: 'helps you identify the specific condition behind {problem}, so you can choose the right solution'
  },
  'FIX': {
    code: 'FIX',
    label: 'Method / Process',
    arrow: '→ How is this professionally corrected?',
    template: 'explains the professional process used to address {problem}, stage by stage'
  },
  'PREVENT': {
    code: 'PREVENT',
    label: 'Maintenance',
    arrow: '→ How do I stop this happening again?',
    template: 'explains the maintenance approach needed for {problem}, helping reduce avoidable deterioration and repeat issues'
  },
  'CHOOSE': {
    code: 'CHOOSE',
    label: 'Contractor Selection',
    arrow: '→ Who should I hire and how do I evaluate them?',
    template: 'helps you find and evaluate the right specialist for {problem}'
  },
  'LOCAL': {
    code: 'LOCAL',
    label: 'Geo Service',
    arrow: '→ Can someone fix this near me?',
    template: 'explains the professional service available for {problem} in {locality}'
  },
  'PROVE': {
    code: 'PROVE',
    label: 'Case Study / Proof',
    arrow: '→ Has this actually been fixed for someone like me?',
    template: 'shows how {problem} was addressed on a real floor in {locality}, with clear before-and-after evidence'
  },
  'DECIDE': {

      code: 'DECIDE',

      label: 'Value / Comparison',

      arrow: '→ Which option is right for my situation?',

      template: 'helps you decide which option is the right fit for {problem}'

    }
};

// ============================================================
// STOP WORDS
// ============================================================
var STOP_WORDS = [
  'a','an','the','and','or','but','in','on','at','to','for','of','with',
  'is','are','was','were','be','been','being','have','has','had','do','does',
  'did','will','would','could','should','may','might','shall','not','no',
  'that','this','these','those','it','its','they','them','their','there',
  'what','why','how','when','where','which','who','whom',
  'after','before','during','from','into','through','about','above','below',
  'my','your','our','their','his','her','its','me','you','us',
  'despite','still','even','just','also','only','very','so','too',
  'marble','floor','floors'
];

// ============================================================
// KEYWORD EXTRACTION
// ============================================================
function extractKeywords(phrase) {
  if (!phrase) return [];
  var words = phrase.toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .split(/\s+/)
    .filter(function(w) {
      return w.length > 2 && STOP_WORDS.indexOf(w) === -1;
    });
  var seen = {};
  return words.filter(function(w) {
    if (seen[w]) return false;
    seen[w] = true;
    return true;
  });
}

// ============================================================
// INTENT OVERLAP
// ============================================================
function intentOverlap(phraseA, phraseB) {
  var keywordsA = extractKeywords(phraseA);
  var keywordsB = extractKeywords(phraseB);
  if (keywordsA.length === 0 || keywordsB.length === 0) return 0;
  var shared = keywordsA.filter(function(k) {
    return keywordsB.indexOf(k) > -1;
  });
  return shared.length;
}

// ============================================================
// AUTO ANGLE ASSIGNMENT
// Deterministic mapping from Article Type + Entity Role +
// Primary Entity intent component.
// No user input required.
// ============================================================
function resolveAngleCode(articleType, entityRole, primaryEntity, primaryProblem) {
  var at = String(articleType || '').trim();
  var er = String(entityRole  || '').trim();
  var pe = String(primaryEntity || '').toLowerCase().trim();
  var pp = String(primaryProblem || '').toLowerCase().trim();

  // Fully deterministic by Article Type first
  if (at === 'Hub Page')         return 'WHY';
  if (at === 'Geo Service Page') return 'LOCAL';
  if (at === 'Service Page') return 'FIX';
  if (at === 'Case Study')       return 'PROVE';
  if (at === 'Method Guide')     return 'FIX';
  if (at === 'Diagnostic Guide') return 'ID';
  if (at === 'Buyer Guide') {

    // If the reader is deciding whether a material/product/option is suitable,
    // this is a decision/comparison angle.
    if (
      pp.indexOf('deciding whether') > -1 ||
      pp.indexOf('suits') > -1 ||
      pp.indexOf('suitable') > -1 ||
      pp.indexOf('comparison') > -1 ||
      pp.indexOf('worth') > -1
    ) {
      return 'DECIDE';
    }

    // Otherwise contractor/local-advice Buyer Guides use CHOOSE.
    if (pe.indexOf('contractor selection') > -1) return 'CHOOSE';
    if (pe.indexOf('local advice') > -1) return 'CHOOSE';

  return 'DECIDE';
}

  // Educational Guide — resolve by Entity Role + Primary Entity intent
  if (at === 'Educational Guide') {
    if (er === 'Maintenance Entity')           return 'PREVENT';
    if (er === 'Intervention Method Entity')   return 'FIX';
    if (er === 'Diagnostic Entity') {
      if (pe.indexOf('local advice')          > -1) return 'CHOOSE';
      if (pe.indexOf('contractor selection')  > -1) return 'CHOOSE';
      if (pe.indexOf('damage cause')          > -1) return 'WHY';
      if (pe.indexOf('diagnosis')             > -1) return 'WHY';
      if (pe.indexOf('method safety')         > -1) return 'WHY';
      if (pe.indexOf('cost')                  > -1) return 'DECIDE';
      if (pe.indexOf('comparison')            > -1) return 'DECIDE';
      if (pp.indexOf('restoration specialist selection') > -1) return 'WHY';
      return 'ID'; // default for Educational Guide Diagnostic Entity
    }
    return 'ID'; // fallback for Educational Guide
  }

  // Fallback for any unrecognised Article Type
  return 'ID';
}

function buildAngleString(angleCode, problem, locality) {
  if (angleCode === 'EXEMPT') return '';
  var def = ANGLE_DEFINITIONS[angleCode];
  if (!def) return angleCode;

  var stmt = def.template
    .replace('{problem}', problem)
    .replace('{locality}', locality || 'your area');
  return def.code + ' — ' + def.label + '\n' +
         def.arrow + '\n' +
         'This article ' + stmt + '.';
}

// ============================================================
// AUTO-ASSIGN ANGLE TO ACTIVE ROW
// Called from sidebar. Reads governance data from active row,
// determines angle deterministically, writes to Problem Angle column.
// ============================================================
function autoAssignProblemAngle(targetRow) {
  try {
    var ss         = SpreadsheetApp.getActiveSpreadsheet();
    var postsSheet = ss.getSheetByName("posts");
    if (!postsSheet) return { success: false, message: "Posts sheet not found." };

    var activeRow = targetRow || postsSheet.getActiveRange().getRow();
    if (activeRow < 2) return { success: false, message: "Select a data row first." };

    var lastCol = postsSheet.getLastColumn();
    var headers = postsSheet.getRange(1, 1, 1, lastCol).getValues()[0]
                            .map(function(h) { return String(h).trim(); });
    var row     = postsSheet.getRange(activeRow, 1, 1, lastCol).getValues()[0];

    function val(name) {
      var idx = headers.indexOf(name);
      return idx > -1 ? String(row[idx] || '').trim() : '';
    }

    var articleType    = val('Article Type');
    var entityRole     = val('Entity Role');
    var primaryEntity  = val('Primary Entity');
    var primaryProblem = val('Primary Query Cluster Owned');

    if (!articleType) {
      return { success: false, message: "Article Type not set for this row — run governance prompts first." };
    }

    var angleCode = resolveAngleCode(
      articleType,
      entityRole,
      primaryEntity,
      primaryProblem
    );
    var locality = val('Locality');
    var angleString = buildAngleString(
      angleCode,
      primaryProblem,
      locality
    );

    

    // Write to Problem Angle column
    var angleIdx = headers.indexOf('Problem Angle');
    if (angleIdx === -1) {
      return { success: false, message: "Problem Angle column not found in Posts sheet." };
    }

    postsSheet.getRange(activeRow, angleIdx + 1).setValue(angleString);

    return {
      success: true,
      message: "Angle assigned: " + angleCode + " — " + ANGLE_DEFINITIONS[angleCode].label,
      angleCode: angleCode,
      angleString: angleString
    };

  } catch(e) {
    return { success: false, message: "Auto-assign error: " + e.toString() };
  }
}

// ============================================================
// BUILD PROBLEM REGISTER
// ============================================================
function buildProblemRegister() {
  try {
    var ss         = SpreadsheetApp.getActiveSpreadsheet();
    var postsSheet = ss.getSheetByName("posts");
    if (!postsSheet) return { success: false, message: "Posts sheet not found." };

    var lastCol  = postsSheet.getLastColumn();
    var lastRow  = postsSheet.getLastRow();
    if (lastRow < 2) return { success: false, message: "No data rows in posts sheet." };

    var headers  = postsSheet.getRange(1, 1, 1, lastCol).getValues()[0]
                             .map(function(h) { return String(h).trim(); });
    var allData  = postsSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    function colIdx(name) { return headers.indexOf(name); }

    var iTitle       = colIdx("Title");
    var iURL         = colIdx("URL");
    var iStone       = colIdx("Stone Type");
    var iType        = colIdx("Article Type");
    var iEntityRole  = colIdx("Entity Role");
    var iPrimEntity  = colIdx("Primary Entity");
    var iGovSummary  = colIdx("Page Governance Summary");
    var iPrimary     = colIdx("Primary Query Cluster Owned");
    var iObserved    = colIdx("Observed Query Cluster");
    var iAngle       = colIdx("Problem Angle");
    var iDrift       = colIdx("Drift Status");
    var iRewrite     = colIdx("Rewrite Status");
    var iLocality    = colIdx("Locality");

    var registerRows = [];
    var today        = new Date().toLocaleDateString("en-GB");

    allData.forEach(function(row) {
      var title      = String(row[iTitle]      || "").trim();
      var url        = String(row[iURL]        || "").trim();
      var stone      = String(row[iStone]      || "").trim();
      var artType    = String(row[iType]       || "").trim();
      var entityRole = String(row[iEntityRole] || "").trim();
      var primEntity = String(row[iPrimEntity] || "").trim();
      var govSummary = String(row[iGovSummary] || "").trim();
      var primary    = String(row[iPrimary]    || "").trim();
      var observed   = String(row[iObserved]   || "").trim();
      var angle      = String(row[iAngle]      || "").trim();
      var drift      = String(row[iDrift]      || "").trim();
      var rewrite    = String(row[iRewrite]    || "").trim();
      var locality = String(row[iLocality] || "").trim();

      if (!url || !primary) return;

      // Auto-assign angle if not already set
      if (!angle && artType) {
        var autoCode = resolveAngleCode(
          artType,
          entityRole,
          primEntity,
          primary
        );
        if (autoCode !== 'EXEMPT') {
          angle = buildAngleString(
          autoCode,
          primary,
          locality
          );
        }
      }

      var angleCode = "";
      var angleStmt = "";
      if (angle) {
        var dashIdx = angle.indexOf(" — ");
        if (dashIdx > -1) {
          angleCode = angle.substring(0, dashIdx).trim();
          angleStmt = angle.substring(dashIdx + 3).trim();
        } else {
          angleCode = angle;
        }
      }

      registerRows.push([
        primary, stone, artType, angleCode, angleStmt,
        url, title, today, entityRole, primEntity, drift, rewrite
      ]);

      if (observed && observed !== primary) {
        var clusters = observed.split(/[;,]/).map(function(c) { return c.trim(); });
        clusters.forEach(function(cluster) {
          if (cluster && cluster !== primary) {
            registerRows.push([
              cluster, stone, artType,
              angleCode + " (secondary)", angleStmt,
              url, title, today, entityRole, primEntity, drift, rewrite
            ]);
          }
        });
      }
    });

    if (registerRows.length === 0) {
      return { success: false, message: "No register entries built — check Primary Query Cluster Owned column is populated." };
    }

    var regSheet = ss.getSheetByName("Problem Register");
    if (!regSheet) {
      regSheet = ss.insertSheet("Problem Register");
    }

    regSheet.clearContents();
    var regHeaders = [
      "Problem", "Stone Type", "Article Type", "Angle Code",
      "Angle Statement", "URL", "Title", "Last Updated",
      "Entity Role", "Primary Entity", "Drift Status", "Rewrite Status"
    ];
    regSheet.getRange(1, 1, 1, regHeaders.length).setValues([regHeaders]);
    regSheet.getRange(2, 1, registerRows.length, regHeaders.length).setValues(registerRows);
    regSheet.getRange(1, 1, 1, regHeaders.length)
            .setFontWeight("bold").setBackground("#1a237e").setFontColor("#ffffff");
    regSheet.setFrozenRows(1);

    return {
      success: true,
      message: "Problem Register built — " + registerRows.length + " entries from " + allData.length + " posts rows.",
      count: registerRows.length
    };

  } catch(e) {
    return { success: false, message: "Build error: " + e.toString() };
  }
}

// ============================================================
// CHECK PROBLEM REGISTER
// ============================================================
function checkProblemRegister(problem, stoneType, activeUrl, activeArticleType, activeAngleCode) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (!stoneType) {
      var posts = ss.getSheetByName("posts");
      if (posts) {
        var activeRow = posts.getActiveRange() ? posts.getActiveRange().getRow() : 0;
        if (activeRow >= 2) {
          var headers = posts.getRange(1, 1, 1, posts.getLastColumn()).getValues()[0]
                             .map(function(h) { return String(h).trim(); });
          var stIdx = headers.indexOf("Stone Type");
          if (stIdx > -1) {
            stoneType = String(posts.getRange(activeRow, stIdx + 1).getValue() || "").trim();
          }
        }
      }
    }

    var regSheet = ss.getSheetByName("Problem Register");
    if (!regSheet) {
      return {
        success: true, found: false,
        message: "Problem Register not yet built — run Build Register first.",
        matches: []
      };
    }

    var lastRow = regSheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, found: false, message: "Problem Register is empty.", matches: [] };
    }

    var data    = regSheet.getRange(2, 1, lastRow - 1, 12).getValues();
    var matches = [];

    data.forEach(function(row) {
      var regProblem     = String(row[0]  || "").trim();
      var regStone       = String(row[1]  || "").trim();
      var regArticleType = String(row[2]  || "").trim();
      var regAngleCode   = String(row[3]  || "").trim();
      var regAngleStmt   = String(row[4]  || "").trim();
      var regUrl         = String(row[5]  || "").trim();
      var regTitle       = String(row[6]  || "").trim();
      var regEntityRole  = String(row[8]  || "").trim();
      var regDrift       = String(row[10] || "").trim();
      var regRewrite     = String(row[11] || "").trim();

      // Skip active page
      var normRegUrl    = regUrl.replace(/\/+$/, '').toLowerCase().trim();
      var normActiveUrl = String(activeUrl || '').replace(/\/+$/, '').toLowerCase().trim();
      if (normActiveUrl && normRegUrl === normActiveUrl) return;

      // Skip different stone types
      var normRegStone  = regStone.toLowerCase().trim();
      var normStone     = String(stoneType || '').toLowerCase().trim();
      if (normStone && normRegStone && normRegStone !== normStone) return;

      // Keyword-based intent matching
      var overlap = intentOverlap(problem, regProblem);
      if (overlap < 2) return;

      // Classify conflict type
      var conflictType = 'POTENTIAL_CANNIBALISATION';

      // Geo vs Geo same angle = location variant, not cannibalisation
      if (
        activeArticleType === 'Geo Service Page' &&
        regArticleType    === 'Geo Service Page' &&
        activeAngleCode   === 'LOCAL' &&
        regAngleCode.replace(' (secondary)', '') === 'LOCAL'
      ) {
        conflictType = 'LOCATION_VARIANT';
      }

      matches.push({
        problem:         regProblem,
        stoneType:       regStone,
        articleType:     regArticleType,
        angleCode:       regAngleCode,
        angleStmt:       regAngleStmt,
        url:             regUrl,
        title:           regTitle,
        entityRole:      regEntityRole,
        drift:           regDrift,
        rewrite:         regRewrite,
        overlap:         overlap,
        conflictType:    conflictType
      });
    });

    // Sort: cannibalisation conflicts first, then location variants, then by overlap
    matches.sort(function(a, b) {
      if (a.conflictType !== b.conflictType) {
        return a.conflictType === 'POTENTIAL_CANNIBALISATION' ? -1 : 1;
      }
      return b.overlap - a.overlap;
    });

    var cannibalisationCount = matches.filter(function(m) {
      return m.conflictType === 'POTENTIAL_CANNIBALISATION';
    }).length;

    var locationVariantCount = matches.filter(function(m) {
      return m.conflictType === 'LOCATION_VARIANT';
    }).length;

    var message = "";
    if (matches.length === 0) {
      message = "No other articles address this problem — this is a new angle";
    } else {
      var parts = [];
      if (cannibalisationCount > 0) {
        parts.push(cannibalisationCount + " potential cannibalisation conflict(s)");
      }
      if (locationVariantCount > 0) {
        parts.push(locationVariantCount + " location variant(s) — same problem, different area");
      }
      message = parts.join(" | ");
    }

    return {
      success:                true,
      found:                  matches.length > 0,
      count:                  matches.length,
      cannibalisationCount:   cannibalisationCount,
      locationVariantCount:   locationVariantCount,
      matches:                matches,
      message:                message
    };

  } catch(e) {
    return { success: false, message: "Check error: " + e.toString(), matches: [] };
  }
}