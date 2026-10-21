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


// W1.5D owns the single material-hub assignment. It never invents a URL.
function bc_w15dAssignHub_(plan, hubUrl, articleType) {
  var text = String(plan || '');
  if (articleType === 'Hub Page' || !/^https?:\/\//i.test(String(hubUrl || '')) ||
      /hub-page-temp/i.test(hubUrl)) return text;
  var lines = text.split('\n');
  // Remove legacy W1.5C hub assignment if still present, avoiding duplicates.
  lines = lines.map(function(line) {
    return /^\s*Internal link:\s*/i.test(line) &&
      line.indexOf(hubUrl) >= 0 ? line.replace(/Internal link:\s*.+$/i, 'Internal link: None') : line;
  });
  var heading = -1, chosen = -1;
  for (var i = 0; i < lines.length; i++) {
    if (/^SECTION\s+\d+:/i.test(lines[i])) heading = i;
    if (heading >= 0 && /^\s*Internal link:\s*None\s*$/i.test(lines[i]) && chosen < 0) chosen = i;
  }
  if (chosen >= 0) {
    lines[chosen] = lines[chosen].replace(/Internal link:\s*None/i, 'Internal link: ' + hubUrl);
  }
  return lines.join('\n');
}

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
  const isHub = articleType === 'Hub Page';
  // Missing or placeholder hubs are not emitted as links.

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

  sectionPlan = bc_w15dAssignHub_(sectionPlan, hubUrl, articleType);
  const hubAlreadyAssigned = !!hubUrl && sectionPlan.indexOf('Internal link: ' + hubUrl) !== -1;
  const assignHubLink = !isHub && hubAlreadyAssigned;

  // The mandatory hub link does not consume the optional lateral-link allowance.
  const lateralLinksToAssign = Math.max(0, maxTotalLinks);

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

  const localityIdx = headers.indexOf("Locality");
  const evidenceNames = ['Project Location & Property','Observed Conditions','Client Requirements',
    'Materials & Construction','Inspection & Tests','Restoration Methods',
    'Products & Equipment','Technical Findings','Project Results','Maintenance Handover'];
  const statusIdx = headers.indexOf('Evidence Status');
  const postIdIdx = headers.indexOf('Post ID');
  const exportSheet = ss.getSheetByName('site-export');
  const originalById = {};
  if (exportSheet && exportSheet.getLastRow()>1) {
    const sources=exportSheet.getRange(2,1,exportSheet.getLastRow()-1,15).getValues();
    sources.forEach(function(src) {
      const id=String(src[0]||'').trim();
      if(id && !(id in originalById)) originalById[id]=String(src[14]||'');
      else if(id) originalById[id]=null;
    });
  }

  const currentLocality =
    localityIdx > -1
      ? String(
          sheet.getRange(activeRow, localityIdx + 1).getValue() || ""
        ).trim()
      : "";

  // National service-area boundaries: a Scottish service/geo-service page
  // may only use confirmed Scottish projects, not English projects.
  const currentArticleType = String(data[activeRow-1][articleTypeIdx]||'').trim();
  function confirmedRegion(locality,title,url) {
    const val=[locality,title,url].join(' ').toLowerCase();
    if (/\b(scotland|scottish|edinburgh|glasgow|stirling|dundee|aberdeen|inverness|perth|fife|falkirk|paisley|kinross|livingston|musselburgh|dalkeith|dunfermline|east lothian|west lothian|midlothian|borders)\b/.test(val))return 'Scotland';
    if (/\b(england|english|london|surrey|burton|stoke|staffordshire|nottingham|derby|birmingham|somerset|northampton|walsall|tutbury|penkhull|new malden|bristol|oxford|york|leicester|manchester|liverpool)\b/.test(val))return 'England';
    return '';
  }
  const currentRegion=confirmedRegion(currentLocality,
    String(data[activeRow-1][titleIndex]||''),currentUrl);
  // Original WordPress anchors are review candidates, never automatically approved.
  const currentId = postIdIdx >= 0 ? String(data[activeRow - 1][postIdIdx] || '').trim() : '';
  const originalCurrent = String(originalById[currentId] || '');
  const sourceLinks = [];
  const originalHrefRe = /<a\b[^>]*\bhref\s*=\s*["']([^"']+)["']/gi;
  let hrefMatch;
  while ((hrefMatch = originalHrefRe.exec(originalCurrent)) !== null) {
    const href = String(hrefMatch[1] || '').trim();
    if (/^https?:\/\/(?:www\.)?abbeyfloorcare\.co\.uk(?:\/|$)/i.test(href) ||
        (/^\//.test(href) && !/^\/\//.test(href))) {
      if (sourceLinks.indexOf(href) < 0) sourceLinks.push(href);
    }
  }
  function linkPath_(url) {
    return String(url || '').replace(/^https?:\/\/[^/]+/i, '').split(/[?#]/)[0]
      .replace(/\/+$/, '').toLowerCase() || '/';
  }
  let siloPages = [];

  for (let i = 1; i < data.length; i++) {

    // Skip current article.
    if (i + 1 === activeRow) continue;

    // Same material only.
    if (data[i][matIndex] !== material) continue;

    const url =
      String(data[i][urlIndex] || "").trim();

    if (!url) continue;

    // Prevent self-linking.
    if (url === currentUrl) continue;

    const articleType =
      String(data[i][articleTypeIdx] || "").trim();
    if (currentArticleType === 'Service Page' || currentArticleType === 'Geo Service Page') {
      const candidateRegion = confirmedRegion(
        localityIdx >= 0 ? String(data[i][localityIdx]||'') : '',
        String(data[i][titleIndex]||''),url);
      // Fail closed for unclassified local evidence; never cross national boundaries.
      if (articleType === 'Case Study' || articleType === 'Geo Service Page' ||
          articleType === 'Service Page') {
        if (!currentRegion || !candidateRegion || currentRegion !== candidateRegion) continue;
      }
    }

    const candidateLocality =
      localityIdx > -1
        ? String(data[i][localityIdx] || "").trim()
        : "";

    /*
     * LOCATION GOVERNANCE
     *
     * If the current article has a locality:
     *
     * 1. A page with the SAME locality is allowed.
     * 2. A page with a DIFFERENT locality is excluded.
     * 3. A page with NO locality may still be considered as
     *    a general non-geographic supporting page.
     */
    if (
      currentLocality &&
      candidateLocality &&
      articleType !== 'Case Study' &&
      candidateLocality.toLowerCase() !==
        currentLocality.toLowerCase()
    ) {
      continue;
    }

    /*
     * Hub is handled separately.
     * Generic Service Pages remain excluded.
     *
     * Geo Service Pages are allowed only when their locality
     * matches the current article. A Geo Service Page with no
     * locality is not safe to use.
     */
    if (
      articleType === "Hub Page" ||
      articleType === "Service Page"
    ) {
      continue;
    }

    if (
      articleType === "Geo Service Page" &&
      (
        !currentLocality ||
        !candidateLocality ||
        candidateLocality.toLowerCase() !==
          currentLocality.toLowerCase()
      )
    ) {
      continue;
    }

    // Case Studies of the same material can illustrate technically relevant
    // work outside the current locality. The prompt forbids local claims.

    const title =
      String(data[i][titleIndex] || "").trim();

    const primaryEntity =
      String(data[i][primaryEntityIdx] || "").trim();

    const entityRole =
      String(data[i][entityRoleIdx] || "").trim();

    const supportingEntities =
      String(data[i][supportingEntitiesIdx] || "").trim();

    const metaTitle =
      String(data[i][metaTitleIdx] || "").trim();

    const headings =
      String(data[i][headingsIdx] || "").trim();

    const safeHandoff =
      String(data[i][safeHandoffIdx] || "").trim();

    const cannibalGuard =
      String(data[i][cannibalGuardIdx] || "").trim();
    let caseEvidence = '';
    if (articleType==='Case Study' && statusIdx>=0 && postIdIdx>=0 &&
        typeof bc_caseEvidenceDigest_==='function') {
      const original=originalById[String(data[i][postIdIdx]||'').trim()];
      if (original && String(data[i][statusIdx]||'').trim()===
          'SOURCE-QUOTED | sha256:'+bc_caseEvidenceDigest_(original)) {
        const findings=[];
        evidenceNames.forEach(function(label) {
          const col=headers.indexOf(label);
          if(col<0)return;
          try {
            const items=JSON.parse(String(data[i][col]||'[]'));
            if(!Array.isArray(items))return;
            const facts=items.filter(x=>x && x.fact && x.quote)
              .slice(0,2).map(x=>String(x.fact).slice(0,140));
            if(facts.length)findings.push(label+': '+facts.join('; '));
          } catch(e) {}
        });
        caseEvidence=findings.join(' | ').slice(0,2500);
      }
    }

    let h2Sample = "";

    if (headings) {

      const h2Match =
        headings.match(/<h2>[^<]+/gi);

      if (
        h2Match &&
        h2Match.length > 0
      ) {

        const firstThree =
          h2Match
            .slice(0, 3)
            .map(function(h) {
              return h
                .replace(/<h2>\s*/i, '')
                .trim();
            });

        h2Sample =
          firstThree.join(" | ");
      }
    }

    siloPages.push({
      url: url,
      originallyLinked: sourceLinks.some(function(link) { return linkPath_(link) === linkPath_(url); }),
      title: title,
      articleType: articleType,
      locality: candidateLocality,
      sameLocality:
        !!currentLocality &&
        !!candidateLocality &&
        candidateLocality.toLowerCase() ===
          currentLocality.toLowerCase(),
      primaryEntity: primaryEntity,
      entityRole: entityRole,
      supportingEntities: supportingEntities,
      metaTitle: metaTitle,
      h2s: h2Sample,
      safeHandoff: safeHandoff,
      cannibalGuard: cannibalGuard,
      caseEvidence: caseEvidence
    });
  }

  /*
   * Put exact-locality candidates first.
   * General non-geographic pages follow.
   */
  siloPages.sort(function(a, b) {

    if (
      a.sameLocality &&
      !b.sameLocality
    ) {
      return -1;
    }

    if (
      !a.sameLocality &&
      b.sameLocality
    ) {
      return 1;
    }

    return 0;
  });

  siloPages.originalLinks = sourceLinks;
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
        'LINK ALREADY IN ORIGINAL ARTICLE (REVIEW ONLY): ' + (page.originallyLinked ? 'YES' : 'NO') + '\n' +
        'LOCALITY: ' + (page.locality || 'None') + '\n' +
        'SAME LOCALITY AS CURRENT ARTICLE: ' +
          (page.sameLocality ? 'YES' : 'NO') + '\n' +
        'PRIMARY ENTITY: ' + page.primaryEntity + '\n';

      if (page.supportingEntities) {
        siloMapText +=
          'COVERS: ' +
          page.supportingEntities +
          '\n';
      }

      if (page.caseEvidence) {
        siloMapText += 'SOURCE-QUOTED CASE STUDY FACTS: ' + page.caseEvidence + '\n';
      }
      if (page.safeHandoff) {
        siloMapText += 'SAFE HANDOFF: ' + page.safeHandoff + '\n';
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

--- ORIGINAL ARTICLE LINKS — REVIEW, DO NOT RETAIN AUTOMATICALLY ---
${(data.siloMap.originalLinks || []).length ? data.siloMap.originalLinks.join('\\n') : 'No original internal links found.'}
An original link survives only if explicitly selected as an AVAILABLE SILO PAGE below.
Reject original links without relevance, geographic eligibility or a matching approved candidate.

--- AVAILABLE SILO PAGES ---

${siloMapText}

--- RULES ---

1. Return up to ${data.lateralLinksToAssign} assignment${data.lateralLinksToAssign === 1 ? '' : 's'}.

2. Only choose PAGE numbers from the AVAILABLE SILO PAGES above.

3. Only choose SECTION numbers that currently contain:

Internal link: None

4. Do NOT select a section that already contains an Internal link.

5. Do NOT select the same SECTION more than once.

6. Do NOT select the same PAGE more than once.

7. LOCATION PRIORITY:
   - If one or more relevant pages have SAME LOCALITY AS CURRENT ARTICLE: YES, choose the most relevant of those first.
   - Only use a page with LOCALITY: None when no suitable same-locality page exists.
   - Never prefer a non-geographic page over an equally relevant same-locality Case Study or Geo Service Page.
   - An out-of-area Case Study may support a technical point, but must NOT imply the project happened in the current article's locality.

8. Choose links only where they help the reader with the specific section. A zero-assignment result is valid when nothing is relevant. Match Case Studies using verified conditions, methods, findings and outcomes; do not infer undocumented project work. For other article types use primary entity, headings, supporting entities, safe handoff and cannibalisation guardrails.

9. Mandatory hub link in the W1.5C source plan is LOCKED and must remain in its section. The hub is not counted in the optional lateral-link limit. Do not reproduce or rewrite any part of the section plan.

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

Return up to ${data.lateralLinksToAssign} assignment${data.lateralLinksToAssign === 1 ? '' : 's'} inside "assignments".

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

    const sourcePlan = bc_w15dAssignHub_(
      String(sheet.getRange(row, 101).getValue() || '').trim(),
      data.hubUrl, data.articleType);

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
      parsed.assignments.length >
      data.lateralLinksToAssign
    ) {
      return {
        success: false,
        message:
          'Allowed at most ' +
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

    // Contact is a W1.5D-governed destination, not a W2B decision.
    // Keep it separate from the material hub and discretionary lateral links.
    if (!lines.some(function(line){return /^Contact link:\s*\/contact\/?\s*$/i.test(line);})) {
      const finalSection = lines.map(function(line,i){return /^SECTION\s+\d+:/i.test(line)?i:-1;})
        .filter(function(i){return i>=0;}).pop();
      if (finalSection === undefined) throw new Error('Cannot place governed Contact link without a section.');
      let end=lines.length;
      for(let i=finalSection+1;i<lines.length;i++){
        if (/^SECTION\s+\d+:/i.test(lines[i]) || /^Stage 1\.5C complete\./i.test(lines[i])) {end=i;break;}
      }
      lines.splice(end,0,'Contact link: /contact');
    }
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

    const sourcePlan = bc_w15dAssignHub_(
      String(sheet.getRange(row, 101).getValue() || '').trim(),
      getW15DData().hubUrl, getW15DData().articleType);

    if (!sourcePlan) {
      return {
        success: false,
        message:
          'Column CW is empty — cannot create W1.5D fallback.'
      };
    }

    // A fallback must never allow an article without its required material hub.
    const governance = getW15DData();
    if (governance.error) return {success:false,message:governance.error};
    const fallbackSource = bc_w15dAssignHub_(sourcePlan, governance.hubUrl, governance.articleType).replace(/(Stage 1\.5C complete\. Waiting for Stage 1\.5D\.)\s*$/, 'Contact link: /contact\n\n$1');
    let fallbackPlan =
      fallbackSource.replace(
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
- Depth (Light / Standard / Deep), where supplied; do not introduce numerical word limits
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