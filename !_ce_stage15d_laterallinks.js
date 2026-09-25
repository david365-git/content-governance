/**
 * ================================================================================
 * ce_Stage15D_LateralLinks.gs - W1.5D LATERAL LINKS
 * ================================================================================
 *
 * Lateral link assignment to section plan
 *
 * Reads:  Col 101 (CW) — W1.5C Enriched Plan
 * Saves:  Col 102 (CX) — W1.5D Final Plan
 * ================================================================================
 */

function getW15DData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('posts');
  const row = sheet.getActiveRange().getRow();

  if (row < 2) {
    return { error: 'Select a data row first.' };
  }

  const d = getActiveRowDataMap();

  const material = d["Stone Type"] || "UNKNOWN";
  const articleType = d["Article Type"] || "General";
  const primaryEntity = d["Primary Entity"] || material;
  const currentUrl = String(d["URL"] || "").trim();
  const hubUrl = String(d["Feeds Hub"] || "").trim();

  const maxTotalLinks = getMaxLateralLinks(articleType);

  if (maxTotalLinks === null) {
    return {
      error:
        'Could not determine max lateral links for article type: ' +
        articleType
    };
  }

  let sectionPlan =
    String(sheet.getRange(row, 101).getValue() || "").trim();

  if (!sectionPlan) {
    return {
      error:
        'No enriched section plan found in Column CW. Run W1.5C first.'
    };
  }

  sectionPlan = sectionPlan.replace(
    /\[(https?:\/\/[^\]]+)\]\([^)]+\)/g,
    '$1'
  );

  const hubAlreadyAssigned =
    !!hubUrl &&
    sectionPlan.indexOf('Internal link: ' + hubUrl) !== -1;

  const assignHubLink =
    maxTotalLinks > 0 &&
    hubUrl.length > 0;

  const lateralLinksToAssign =
    maxTotalLinks === 0
      ? 0
      : assignHubLink
        ? Math.max(0, maxTotalLinks - 1)
        : maxTotalLinks;

  const siloMap =
    buildEnhancedSiloMap(
      material,
      row,
      currentUrl
    );

  return {
    material: material,
    articleType: articleType,
    primaryEntity: primaryEntity,
    currentUrl: currentUrl,
    maxTotalLinks: maxTotalLinks,
    lateralLinksToAssign: lateralLinksToAssign,
    assignHubLink: assignHubLink,
    hubAlreadyAssigned: hubAlreadyAssigned,
    hubUrl: hubUrl,
    sectionPlan: sectionPlan,
    siloMap: siloMap
  };
}
/**
 * Get max lateral links allowed for an article type from Article Type Control Sheet
 */
function getMaxLateralLinks(articleType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Article Type Control Sheet");
  if (!sheet) return null;
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const typeCol = headers.indexOf("Canonical Article Type Label");
  const maxLinksCol = headers.indexOf("Max Lateral Links");
  
  if (typeCol === -1 || maxLinksCol === -1) {
    Logger.log("ERROR: Article Type Control Sheet missing required columns");
    return null;
  }
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][typeCol]).trim() === articleType) {
      const value = data[i][maxLinksCol];
      const parsed = parseInt(value);
      return isNaN(parsed) ? 0 : parsed;
    }
  }
  
  Logger.log("WARNING: Article type '" + articleType + "' not found in Article Type Control Sheet");
  return 2; // Default fallback
}

function buildEnhancedSiloMap(material, activeRow, currentUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('posts');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const matIndex = headers.indexOf("Stone Type");
  const urlIndex = headers.indexOf("URL");
  const titleIndex = headers.indexOf("Title");
  const articleTypeIdx = headers.indexOf("Article Type");
  const primaryEntityIdx = headers.indexOf("Primary Entity");
  const entityRoleIdx = headers.indexOf("Entity Role");
  const supportingEntitiesIdx = headers.indexOf("Supporting Entities Core");
  const metaTitleIdx = headers.indexOf("Current Meta Title");
  const headingsIdx = headers.indexOf("Current Headings");
  const safeHandoffIdx = headers.indexOf("Safe Handoff Pages");
  const cannibalGuardIdx = headers.indexOf("Cannibalisation Guardrail");
  
  let siloPages = [];
  
  for (let i = 1; i < data.length; i++) {
    if (i === activeRow) continue; // Skip current article by row
    if (data[i][matIndex] !== material) continue; // Only same material
    
    const url = String(data[i][urlIndex] || "").trim();
    if (!url) continue;
    
    // CRITICAL: Skip current article by URL to prevent self-referencing
    if (url === currentUrl) continue;
    
    const articleType = String(data[i][articleTypeIdx] || "").trim();
    
    // Filter out unwanted page types for lateral linking
    const excludedTypes = ["Hub Page", "Service Page", "Geo Service Page"];
    if (excludedTypes.indexOf(articleType) !== -1) continue;
    
    const title = String(data[i][titleIndex] || "").trim();
    const primaryEntity = String(data[i][primaryEntityIdx] || "").trim();
    const entityRole = String(data[i][entityRoleIdx] || "").trim();
    const supportingEntities = String(data[i][supportingEntitiesIdx] || "").trim();
    const metaTitle = String(data[i][metaTitleIdx] || "").trim();
    const headings = String(data[i][headingsIdx] || "").trim();
    const safeHandoff = String(data[i][safeHandoffIdx] || "").trim();
    const cannibalGuard = String(data[i][cannibalGuardIdx] || "").trim();
    
    // Extract first 2-3 H2s
    let h2Sample = "";
    if (headings) {
      const h2Match = headings.match(/<h2>[^<]+/gi);
      if (h2Match && h2Match.length > 0) {
        const firstThree = h2Match.slice(0, 3).map(h => h.replace(/<h2>\s*/i, '').trim());
        h2Sample = firstThree.join(" | ");
      }
    }
    
    siloPages.push({
      url: url,
      title: title,
      articleType: articleType,
      primaryEntity: primaryEntity,
      entityRole: entityRole,
      supportingEntities: supportingEntities,
      metaTitle: metaTitle,
      h2s: h2Sample,
      safeHandoff: safeHandoff,
      cannibalGuard: cannibalGuard
    });
  }
  
  return siloPages;
}

function buildW15DPrompt() {
  try {

    const data = getW15DData();

    if (data.error) {
      return 'ERROR: ' + data.error;
    }

    let siloMapText = '';

    data.siloMap.forEach(function(page, index) {

      siloMapText +=
        'PAGE ' + (index + 1) + '\n' +
        'URL: ' + page.url + '\n' +
        'TITLE: ' + page.title + '\n' +
        'TYPE: ' + page.articleType + '\n' +
        'PRIMARY ENTITY: ' + page.primaryEntity + '\n';

      if (page.supportingEntities) {
        siloMapText +=
          'COVERS: ' +
          page.supportingEntities +
          '\n';
      }

      if (page.h2s) {
        siloMapText +=
          'H2s: ' +
          page.h2s +
          '\n';
      }

      if (page.cannibalGuard) {
        siloMapText +=
          'AVOID: ' +
          page.cannibalGuard +
          '\n';
      }

      siloMapText += '\n';
    });

    return `
W1.5D — LATERAL LINK SELECTION

ROLE:
Choose the most appropriate lateral-link destinations.

IMPORTANT:
You are NOT rewriting the section plan.
You are NOT inserting URLs.
You are only selecting which approved silo PAGE should be linked from which SECTION.

LATERAL LINKS REQUIRED:
${data.lateralLinksToAssign}

--- CURRENT GOVERNED SECTION PLAN ---

${data.sectionPlan}

--- AVAILABLE SILO PAGES ---

${siloMapText}

--- RULES ---

1. Return exactly ${data.lateralLinksToAssign} assignment${data.lateralLinksToAssign === 1 ? '' : 's'}.

2. Only choose PAGE numbers from the AVAILABLE SILO PAGES above.

3. Only choose SECTION numbers that currently contain:

Internal link: None

4. Do NOT select a section that already contains an Internal link.

5. Do NOT select the same SECTION more than once.

6. Do NOT select the same PAGE more than once.

7. Choose the page whose subject is most relevant to the section.

8. Do not reproduce or rewrite any part of the section plan.

--- OUTPUT FORMAT ---

Return JSON only in exactly this structure:

{
  "assignments": [
    {
      "section": 2,
      "page": 1
    }
  ]
}

Return exactly ${data.lateralLinksToAssign} assignment${data.lateralLinksToAssign === 1 ? '' : 's'} inside "assignments".

No explanation.
No markdown fences.
`.trim();

  } catch (e) {

    return (
      'ERROR in buildW15DPrompt: ' +
      e.message
    );
  }
}

function saveW15DUpdatedPlan(updatedPlan) {
  try {

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('posts');
    const row = sheet.getActiveRange().getRow();

    if (row < 2) {
      return {
        success: false,
        message: 'Select a data row first.'
      };
    }

    const data = getW15DData();

    if (data.error) {
      return {
        success: false,
        message: data.error
      };
    }

    const sourcePlan =
      String(
        sheet.getRange(row, 101).getValue() || ''
      ).trim();

    if (!sourcePlan) {
      return {
        success: false,
        message:
          'Column CW is empty — W1.5C must complete before W1.5D.'
      };
    }

    let cleaned =
      String(updatedPlan || '')
        .trim()
        .replace(/^```json\s*/i, '')
        .replace(/^```\s*/i, '')
        .replace(/\s*```$/i, '')
        .trim();

    let parsed;

    try {
      parsed = JSON.parse(cleaned);
    } catch (e) {
      return {
        success: false,
        message:
          'W1.5D returned invalid JSON — CX not changed.'
      };
    }

    if (
      !parsed ||
      !Array.isArray(parsed.assignments)
    ) {
      return {
        success: false,
        message:
          'W1.5D JSON does not contain an assignments array — CX not changed.'
      };
    }

    if (
      parsed.assignments.length !==
      data.lateralLinksToAssign
    ) {
      return {
        success: false,
        message:
          'Expected exactly ' +
          data.lateralLinksToAssign +
          ' lateral link assignment(s), but received ' +
          parsed.assignments.length +
          '. CX not changed.'
      };
    }

    const usedSections = {};
    const usedPages = {};
    const assignments = [];

    for (
      let i = 0;
      i < parsed.assignments.length;
      i++
    ) {

      const assignment =
        parsed.assignments[i] || {};

      const section =
        Number(assignment.section);

      const page =
        Number(assignment.page);

      if (
        !Number.isInteger(section) ||
        section < 1
      ) {
        return {
          success: false,
          message:
            'Invalid section number in W1.5D assignment — CX not changed.'
        };
      }

      if (
        !Number.isInteger(page) ||
        page < 1 ||
        page > data.siloMap.length
      ) {
        return {
          success: false,
          message:
            'Invalid silo page number in W1.5D assignment — CX not changed.'
        };
      }

      if (usedSections[section]) {
        return {
          success: false,
          message:
            'W1.5D selected SECTION ' +
            section +
            ' more than once — CX not changed.'
        };
      }

      if (usedPages[page]) {
        return {
          success: false,
          message:
            'W1.5D selected PAGE ' +
            page +
            ' more than once — CX not changed.'
        };
      }

      usedSections[section] = true;
      usedPages[page] = true;

      assignments.push({
        section: section,
        url: data.siloMap[page - 1].url
      });
    }

    const lines =
      sourcePlan.split('\n');

    assignments.forEach(function(assignment) {

      let sectionStart = -1;
      let sectionEnd = lines.length;

      for (
        let i = 0;
        i < lines.length;
        i++
      ) {

        const match =
          String(lines[i]).match(
            /^SECTION\s+(\d+):/i
          );

        if (!match) {
          continue;
        }

        const number =
          Number(match[1]);

        if (
          number === assignment.section
        ) {
          sectionStart = i;
          continue;
        }

        if (
          sectionStart !== -1 &&
          number !== assignment.section
        ) {
          sectionEnd = i;
          break;
        }
      }

      if (sectionStart === -1) {
        throw new Error(
          'SECTION ' +
          assignment.section +
          ' does not exist in the governed W1.5C plan.'
        );
      }

      let linkLine = -1;

      for (
        let i = sectionStart;
        i < sectionEnd;
        i++
      ) {

        if (
          /^\s*Internal link:\s*None\s*$/i.test(
            lines[i]
          )
        ) {
          linkLine = i;
          break;
        }
      }

      if (linkLine === -1) {
        throw new Error(
          'SECTION ' +
          assignment.section +
          ' does not have "Internal link: None".'
        );
      }

      const indent =
        String(lines[linkLine]).match(/^\s*/)[0];

      lines[linkLine] =
        indent +
        'Internal link: ' +
        assignment.url;
    });

    let finalPlan =
      lines.join('\n');

    finalPlan =
      finalPlan.replace(
        /Stage 1\.5C complete\. Waiting for Stage 1\.5D\.\s*$/,
        'Stage 1.5D complete.'
      );

    if (
      !finalPlan.endsWith(
        'Stage 1.5D complete.'
      )
    ) {
      return {
        success: false,
        message:
          'Could not create the W1.5D completion line — CX not changed.'
      };
    }

    sheet
      .getRange(row, 102)
      .setValue(finalPlan);

    logPipelineResume(
      'W1.5D — Lateral Links Final Plan',
      ''
    );

    return {
      success: true,
      message:
        'W1.5D lateral links inserted safely into the governed W1.5C plan and saved to Column CX.'
    };

  } catch (e) {

    return {
      success: false,
      message:
        'Error saving W1.5D final plan: ' +
        e.message
    };
  }
}

function saveW15DFallbackPlan() {
  try {

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('posts');
    const row = sheet.getActiveRange().getRow();

    if (row < 2) {
      return {
        success: false,
        message: 'Select a data row first.'
      };
    }

    const sourcePlan =
      String(
        sheet.getRange(row, 101).getValue() || ''
      ).trim();

    if (!sourcePlan) {
      return {
        success: false,
        message:
          'Column CW is empty — cannot create W1.5D fallback.'
      };
    }

    let fallbackPlan =
      sourcePlan.replace(
        /Stage 1\.5C complete\. Waiting for Stage 1\.5D\.\s*$/,
        'Stage 1.5D complete.'
      );

    if (
      !fallbackPlan.endsWith(
        'Stage 1.5D complete.'
      )
    ) {
      return {
        success: false,
        message:
          'Could not create W1.5D fallback completion line.'
      };
    }

    sheet
      .getRange(row, 102)
      .setValue(fallbackPlan);

    return {
      success: true,
      message:
        'W1.5D fallback saved — governed W1.5C plan preserved unchanged.'
    };

  } catch (e) {

    return {
      success: false,
      message:
        'W1.5D fallback error: ' +
        e.message
    };
  }
}

function buildW15DRepairPrompt(failedOutput, failureMessage) {

  const data = getW15DData();

  if (data.error) {
    return 'ERROR: ' + data.error;
  }

  return `
W1.5D — LATERAL LINK CORRECTION

The previous W1.5D output failed validation.

VALIDATION FAILURE:
${failureMessage}

--- GOVERNED W1.5C SOURCE PLAN ---
${data.sectionPlan}

--- FAILED W1.5D OUTPUT ---
${failedOutput}

--- CORRECTION RULES ---

Correct ONLY the Internal link fields necessary to resolve the validation failure.

Do NOT change:
- SECTION headings
- TSM Requirement
- Word Budget
- Content Brief
- Visual Pattern
- List eligible
- Images
- Entities
- section order
- any other governed content

The governed W1.5C hub-link placement is LOCKED.

Hub URL:
${data.hubUrl || 'None'}

${
  data.hubAlreadyAssigned
    ? `The hub URL is already present in the W1.5C source plan.
It MUST remain in exactly the same section.
Do NOT move it.`
    : ''
}

Assign exactly ${data.lateralLinksToAssign} lateral link${data.lateralLinksToAssign === 1 ? '' : 's'} using only approved URLs already supplied by W1.5D.

The only permitted changes are values after:

Internal link:

Return the COMPLETE corrected section plan.

First line must begin:
SECTION 1:

Last line must be:
Stage 1.5D complete.

No explanation.
No markdown fences.
`.trim();
}