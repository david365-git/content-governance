/**
 * ce_FactCheck.gs - W2B.05 FACT CHECK AGAINST SOURCE ARTICLE
 * Compares W2B/New HTML output against the original site-export article
 * to catch fabricated facts and unsupported embellishments introduced
 * during the AI rewrite pipeline.
 */

function buildFactCheckPrompt(sourceColumn) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();

  const requestedCol = (sourceColumn === 'CT') ? 'CT' : 'EU';

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(h) { return String(h).trim(); });
  const colPostId = headers.indexOf('Post ID');
  const colTitle = 1; // column B

  if (colPostId === -1) {
    return { success: false, message: "ERROR: Post ID column not found." };
  }

  const rowData = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  const postId = String(rowData[colPostId] || '').trim().replace(/\.0$/, '');
  const title = String(rowData[colTitle] || '').trim();

  const rowMap = {};
  headers.forEach(function(h, i) {
    rowMap[h] = rowData[i];
  });

  const governedMaterial = String(
    rowMap["Stone Type"] ||
    rowMap["Material"] ||
    ""
  ).trim();

  const governedLocality = String(
    rowMap["Locality"] ||
    rowMap["Location"] ||
    ""
  ).trim();

  const articleType = String(
    rowMap["Article Type"] ||
    ""
  ).trim();

  const rewriteBrief = String(
    rowMap["Page Rewrite Brief"] ||
    ""
  ).trim();

  const finalSectionPlan = String(
    rowMap["W1.5D Final Plan"] ||
    rowMap["Final Plan"] ||
    ""
  ).trim();

  const authorityBrief = String(
    rowMap["Authority Brief"] ||
    ""
  ).trim();

  const feedsHub = String(
    rowMap["Feeds Hub"] ||
    ""
  ).trim();

  const safeHandoffPages = String(
    rowMap["Safe Handoff Pages"] ||
    ""
  ).trim();

  // Same-row extracted Case Study evidence (GS:HB). Blank cells are not proof
  // an event did not occur. These entries may contain extracted claims, not
  // independently confirmed facts; weigh each against original and verified notes.
  const caseStudyEvidenceFields = [
    "Project Location & Property",
    "Observed Conditions",
    "Client Requirements & Restoration Objectives",
    "Materials & Construction",
    "Inspection & Tests",
    "Restoration Methods",
    "Products & Equipment",
    "Technical Findings",
    "Project Results",
    "Maintenance Handover"
  ];
  const caseStudyEvidence = caseStudyEvidenceFields
    .map(function(field) {
      return { field: field, value: String(rowMap[field] || '').trim() };
    })
    .filter(function(entry) { return entry.value.length > 0; });
  const caseStudyEvidenceBlock = caseStudyEvidence.length
    ? caseStudyEvidence.map(function(entry) {
        return entry.field + ": " + entry.value;
      }).join("\n\n")
    : "No extracted Case Study evidence entries supplied for this row.";

  // Automatically recheck a previously saved fix if one exists — GE holds
  // the W2B.05 Fixed HTML from an earlier fact-check fix pass. Checking GE
  // first (regardless of which source column was originally audited) means
  // a second "Generate Fact-Check Prompt" click always verifies the actual
  // fix, not the untouched original — no manual column-switching needed.
  const geHtml = String(sh.getRange(row, 187).getValue() || '').trim();
  const targetCol = geHtml ? 'GE' : requestedCol;
  const generatedHtml = geHtml || String(sh.getRange(requestedCol + row).getValue() || '').trim();

  if (!postId) {
    return { success: false, message: "ERROR: Post ID is empty for this row." };
  }
  if (!generatedHtml) {
    return { success: false, message: "ERROR: Column " + targetCol + " is empty for this row." };
  }

  // Look up original source HTML from site-export
  let originalHtml = '';
  try {
    const exportSheet = ss.getSheetByName('site-export');
    if (!exportSheet) {
      return { success: false, message: "ERROR: 'site-export' sheet not found." };
    }
    const exportData = exportSheet.getDataRange().getValues();
    const exportHeaders = exportData[0].map(function(h) { return String(h).trim(); });
    const exportIdCol = exportHeaders.indexOf('ID');
    const exportHtmlCol = exportHeaders.indexOf('Full Post HTML');
    if (exportIdCol === -1 || exportHtmlCol === -1) {
      return { success: false, message: "ERROR: Could not find ID or Full Post HTML columns in site-export." };
    }
    for (let i = 1; i < exportData.length; i++) {
      if (String(exportData[i][exportIdCol]).trim().replace(/\.0$/, '') === postId) {
        originalHtml = String(exportData[i][exportHtmlCol] || '').trim();
        break;
      }
    }
  } catch (e) {
    return { success: false, message: "ERROR reading site-export: " + e.message };
  }

  if (!originalHtml) {
    return { success: false, message: "ERROR: No original article found in site-export for Post ID: " + postId };
  }

  const prompt =
`You are performing W2B.05 PROJECT FACT CHECK.

PURPOSE:
Check whether the rewritten case study has invented, contradicted or materially exaggerated facts about the ACTUAL documented project.

This stage is NOT a general technical fact-check, SEO check, internal-link check, heading check or governance-compliance check. Other stages handle those jobs.

PAGE CONTEXT:
MATERIAL: ${governedMaterial || "Not supplied"}
ARTICLE TYPE: ${articleType || "Not supplied"}
LOCALITY: ${governedLocality || "Not supplied"}

SHORT REWRITE SCOPE:
${rewriteBrief || "No active Rewrite Brief supplied."}

CASE STUDY EVIDENCE — CURRENT POST ROW (EXTRACTED RECORD, NOT AUTOMATICALLY VERIFIED):
${caseStudyEvidenceBlock}

EVIDENCE HANDLING:
Read the ten Case Study evidence fields alongside the source article. Treat extracted entries as supporting project context, not as independent proof of every detail: compare them with the original and any explicitly confirmed owner information. Do not promote an AI-inferred or contradictory entry into a verified project fact. An empty field means no evidence supplied, not that an action did not occur.
Before applying the five-condition unsupported-project-claim test, check whether an assertion is supported by an actual evidence-field statement or an explicitly confirmed job fact. Preserve supported observations and methods, including first-person experience. Where evidence conflicts, do not resolve the conflict by invention; use PASS_WITH_NOTES if no safe determination is possible.

ORIGINAL SOURCE ARTICLE — HISTORICAL REFERENCE, NOT AN EXHAUSTIVE TECHNICAL RECORD:
${originalHtml}

REWRITTEN ARTICLE — CHECK PROJECT CLAIMS IN THIS VERSION:
${generatedHtml}

EVIDENCE PRIORITY — HARD LOCK:
The original article is a historical reference but may omit accurate details. Absence of a statement is NOT evidence it is false. A FAIL requires a demonstrable contradiction or an objectively significant error, not merely missing corroboration.
Protect accurate professional background, mechanisms and ordinary technical detail. Do not use "not documented" alone as the sourceEvidence for FAIL.
If a detail cannot be confirmed or denied, use PASS_WITH_NOTES with no fixes, unless all five conditions of the MATERIAL UNSUPPORTED PROJECT CLAIM rule are met and a safe neutral correction exists.
Do not assume that a colour-enhancing sealer is a surface coating merely because the original omitted "impregnating". Conversely, confirm actual product behaviour when material to classification.
For the Northampton terrazzo example, the owner has confirmed: a cement matrix; colour-enhancing impregnating sealer and its pore-penetration/protective behaviour; the client removed all wooden boards; and the work area was a busy hallway. These known facts must not be labelled unsupported on that project. Do not transfer these Northampton facts to other projects.
Distinguish generic repair limitations (such as repairs not being perfectly invisible) from a specific observed defect.
Only propose an automatic edit where the corrected wording resolves an unambiguous, material error. Otherwise PASS_WITH_NOTES with no changes.
Do not flag generic, accurate technical explanations simply because they add useful detail.

KNOWLEDGE-BOUNDARY FACT-CHECK RULE — ALL ARTICLE TYPES:
EDITORIAL KNOWLEDGE BOUNDARY — ALL ARTICLE TYPES AND MATERIALS:
A. CONFIRMED SOURCE FACTS: Preserve the original account's actual events, observations, people, dates, dimensions, treatments, order, client preferences and results. Later confirmed notes can clarify ambiguous shorthand. Apparent soundness and worn appearance can coexist; distinguish general condition from local defects.
B. PROFESSIONAL TRADE KNOWLEDGE: Explain the purpose and typical mechanism of a documented operation in clear UK homeowner English. Do not claim an illustrative technique, grit, product, diagnosis or equipment was actually used unless the source establishes it. Do not force grinding, honing, polishing or sealing on materials or jobs where these are unsuitable.
C. UNCERTAIN DIAGNOSTICS: Explain plausible causes conditionally, not as confirmed project findings. Shorthand such as "movement" is context-dependent: it may describe cracks, delamination, loose material or other issues. Do not automatically infer substrate or structural causes.
D. ARTICLE-TYPE VOICE: Preserve I/we where the original establishes practitioner involvement; never change we to I, invent a named technician, or mandate first-person in guides. Case Studies tell the job's story; Diagnostic Guides distinguish symptoms; Method Guides explain safe steps; Buyer Guides support choices; Educational Guides explain principles; Hub Pages organise topics; service pages explain suitability; correction pieces address misconceptions.
E. EDITORIAL QUALITY: Avoid "near me" in reader-facing narrative, unsupported local-property stereotypes, repeated headings, fabricated client reactions and empty technical padding. Preserve useful source detail and avoid boilerplate.
Do not reject a technically sound generic explanation solely because it is absent from the original. Flag a definite job-site claim about unrecorded specific tools, grit, measurements, client reactions, personnel or outcomes when materially misleading. A qualified diagnostic possibility is not an assertion that the cause was found on site. Review image captions and headings too. Preserve the established PASS_WITH_NOTES mechanism for uncertainties that cannot be safely repaired atomically.

PROJECT EVIDENCE PROTECTION — TARGETED CHECKS:
Perform a dedicated comparison of new project-specific statements against the ORIGINAL SOURCE ARTICLE, including header, body, photographs, alt text, video titles and captions.
1. Separate (a) documented events, (b) ordinary professional explanation or conditional advice, and (c) claims that a specific inspection, test, decision criterion, tool, treatment or observation actually occurred on this job.
2. Scrutinise new assertions such as a controlled/consistent lighting inspection, decisions based on the quantity released during each cleaning pass, a test area being performed, or a media item depicting a particular treatment stage. These are examples of categories to verify, NOT claims to flag automatically on every article.
3. Do not silently convert generic good practice into an assertion that the team did it on this project. If the source supports only general practice, keep the useful advice conditional and remove any misleading claim that it occurred.
4. A new, positively asserted and materially specific on-site procedure or decision criterion that lacks source support may be classified MATERIAL UNSUPPORTED PROJECT CLAIM only when it would meaningfully misrepresent what was done or observed. Absence alone does NOT prove a false claim; genuine uncertainty belongs in PASS_WITH_NOTES.
5. For an unsupported photo/video treatment-stage identification, compare against the source's existing caption/title. Where the rewrite newly asserts a different stage without supporting evidence, use a minimal source-grounded caption/title replacement if the claim materially misrepresents the media. Do not invent a visual description from an unseen image or video; preserve src, video URL, iframe and figure structure.
6. Where a materially unsupported project-specific claim has an unambiguous safe, neutral rewrite, provide an atomic visible-text fix. Retain verified project facts and accurate professional explanations. Do not erase real restoration detail simply to achieve PASS.
7. Do not turn stylistic issues, optional extra links, minor caption phrasing, or harmless general test-area advice into blocking errors. CX link governance and editorial audits remain separate.
8. In sourceEvidence, quote the pertinent original wording and explain the precise unsupported project-specific implication; never write only 'not documented' as the justification.

MATERIAL UNSUPPORTED PROJECT CLAIM — FIVE-CONDITION DECISION RULE:
A new assertion qualifies for FAIL with issueType MATERIAL UNSUPPORTED PROJECT CLAIM ONLY if ALL five conditions hold:
1. NEW: The specific job-site detail was introduced by the rewrite and is not already stated or reasonably entailed by the original project account or other confirmed project evidence actually supplied here.
2. JOB-SPECIFIC: It positively claims an inspection, test, measurement, named tool/product, particular treatment stage, observed result, client reaction or decision criterion OCCURRED on this job; it is not conditional advice or an explanation of ordinary trade practice.
3. UNSUPPORTED: The supplied project evidence does not substantiate that particular assertion. Check meaning and reasonable paraphrase, not exact-word matching. Do not presume the historical original is an exhaustive record.
4. MATERIAL: Its asserted specificity could meaningfully mislead a homeowner about what was actually done, seen or decided. Do not escalate harmless descriptions, professional explanation, style or minor ambiguity.
5. SAFELY CORRECTABLE: A short, unambiguous, atomically matchable visible-text replacement can remove only the unsupported job attribution while preserving genuine experience, useful trade explanation, grammar, section purpose and all links. If not, record uncertainty as PASS_WITH_NOTES instead of inventing a repair.
When ALL five hold, return FAIL with exact atomic fixes, not PASS_WITH_NOTES. For new stage-specific photo or video claims, prefer the original source-grounded description rather than inferring what unseen media depicts. Do NOT mechanically change every first-person observation into generic advice: authentic source-supported experience must remain first-person.
For example, 'Looking at the floor under consistent lighting helped...' and 'The amount of material released determined whether another pass was safer...' warrant this test if they are newly attributed to this job without project support. They are NOT automatic FAIL keywords.
General guidance such as 'A test area can help confirm the likely appearance' is not a project claim. A statement may be true despite being absent from an old article, so use PASS_WITH_NOTES if the five conditions cannot be established confidently.
Only use authenticated case-study evidence if it is actually included in the supplied prompt context; do not assume access to unseen ledgers.
For sourceEvidence, quote the nearest original project evidence, explain the unsupported newly asserted attribution and why it is materially misleading; missing documentation alone does not establish a factual contradiction.

FACT-CHECK SCOPE — MATERIAL ERRORS ONLY:
This is a focused check of the documented project, not a line-by-line challenge to technical explanation.
Check:
1. Correct material identity and any material distinctions.
2. Correct room, floor area and indoor/outdoor setting.
3. Correct town/locality and project location.
4. Correct original defect or reason for the work.
5. Correct treatments and steps actually performed, with no contradictory invented action.
6. Correct material/product compatibility.
7. Correct sealer classification: impregnating versus surface coating, and colour-enhancing versus non-enhancing. Do not infer a sealer's type solely from its described effect.
8. Equipment and methods: flag only a clear contradiction with the documented work; allow general equipment explanations, typical abrasives and illustrative grit sequences.
9. Work sequence: flag material contradictions, not harmless narrative rearrangement.
10. Existing damage: flag invented serious defects or falsely asserted causes.
11. Results: flag demonstrably false or materially misleading outcomes, not reasonable general benefits clearly distinguished from observed results.
12. Client actions, quotations, approvals, reactions or specific advice: flag fabricated project events, not general homeowner guidance.
13. Safety: flag only a clear and significant material-damage risk or unsafe instruction, not reasonable professional explanations.
14. Internal consistency: flag contradictory material, room, location, sealer, treatment, quantity or outcome claims.

MATERIALITY RULE — HARD LOCK:
FAIL only for a clear material factual contradiction, materially misleading project-specific claim, or significant unsafe recommendation. The finding must identify the concrete error and the supporting source evidence or clear technical incompatibility.
Do not require the source article to substantiate general trade knowledge, material science, typical equipment behaviour, illustrative grit numbers, general maintenance advice, explanatory reasons for using a method, or contextual links.
Do not turn general education into a purported historical event. Distinguish "this is how a planetary grinder works" from "we used 400-grit diamonds on this job".
Never fail just because a technically sound explanatory detail is absent from the original case study.

ESTABLISHED STONE AND TILE RESTORATION KNOWLEDGE — HARD LOCK:
Apply this rule to ALL materials and article types: marble, limestone, travertine, terrazzo, granite, slate, sandstone, terracotta, Victorian and quarry tiles, porcelain, glazed and unglazed ceramic tiles, grout, and other stone or tile surfaces.
Preserve technically sound general explanations of each material's properties, cleaning chemistry, surface finishes, porosity, equipment, methods, maintenance, and inherent repair or restoration limitations, even if those principles are not described in the original article.
For example, repairs to stone, terrazzo and tiles may be blended carefully but cannot be assumed completely invisible; differences in aggregate, colour, grain, texture, glaze or finish can remain detectable. This is general technical education, not a claim that a particular defect was observed on the job.
Distinguish GENERAL PRINCIPLES from PROJECT-SPECIFIC FACTS. General explanations of how grinding, honing, polishing, sealing, grout cleaning or tile repairs work are permissible if accurate and applicable. Claims that a particular product, machine, grit, repair material or treatment was ACTUALLY USED, or that a specific limitation was OBSERVED on this project, must be supported by the original project evidence.
Do not infer that a specific sealer was impregnating or film-forming, that a particular floor had a given binder or substrate, or that a client observed a repair difference unless supported by evidence.
When a sentence combines an undocumented project action with valid general technical explanation, correct only the unsupported action and retain the explanation.
Do not reject technical information merely because it adds useful detail beyond the original case-study record.
Do not remove useful technical detail solely to make the article less specific.
Do not invent certainty where product formulation, substrate condition or client actions are unknown.
Judge meaning, not literal wording. Minor word choice, style, SEO, headings, CTA, hyperlinks and non-material ambiguity are outside this check.
When uncertain or when no demonstrable material error exists, PASS or PASS WITH NOTES rather than FAIL.
For every FAIL fix, choose the smallest uniquely matchable visible-text fragment, and preserve existing links. Never use a fix to introduce unrelated editorial changes.

COMPLETE FIRST-PASS REVIEW — HARD LOCK:
Before returning the first response, inspect the ENTIRE rewritten article from the opening to the final paragraph, including introduction, each section, captions, conclusions and aftercare.
Cross-check every material project-specific assertion against the original source and the 14 criteria above. In particular, review who performed actions, client reactions or expectations, room and locality, terrazzo binder or substrate type, equipment actually used, sealer classification and documented outcomes.
Make a second complete internal sweep for material errors you may have overlooked. Consolidate ALL demonstrable material errors into the FIRST FAIL response; do not stop after finding the first few.
For each discovered factual error, search the entire rewritten HTML for the same claim in H1/H2/H3 headings, captions, body text, callouts and conclusion. Supply a separate exact-text fix for each occurrence where needed; never correct only the paragraph if its heading still makes the same false claim.
Do not ration findings, invent issues to fill a quota, or broaden the materiality threshold. Supplementary technical education remains permitted.
Each fix must address a distinct material error, and must be independently justified by precise evidence. Only after completing both sweeps return the JSON.
On verification, apply exactly the same materiality threshold; do not introduce stricter editorial standards or reclassify harmless background explanations as project facts.

CLASSIFICATION:
PASS: No material errors and no noteworthy uncertainty.
PASS_WITH_NOTES: No material errors or safely repairable material unsupported project assertions; minor genuine uncertainty or editorial observation worth recording, with no fixes.
FAIL: Demonstrable factual contradictions, material safety errors, OR newly introduced material project-specific assertions meeting ALL five conditions above, each with unambiguous safe atomic fixes. Missing documentation alone is not FAIL.

OUTPUT FORMAT — HARD LOCK:
Return ONLY valid JSON, no fences.
For PASS:
{
  "status": "PASS",
  "summary": "No material factual errors found.",
  "fixes": []
}
For PASS WITH NOTES:
{
  "status": "PASS_WITH_NOTES",
  "summary": "Brief non-blocking observation.",
  "fixes": []
}
For FAIL:
{
  "status": "FAIL",
  "summary": "Brief description of the material factual errors.",
  "fixes": [
    {
      "fixLabel": "short label",
      "issueType": "FACTUAL CONTRADICTION or MATERIAL UNSUPPORTED PROJECT CLAIM or MATERIAL SAFETY ERROR",
      "targetText": "unique exact visible-text fragment of the error",
      "contextText": "exact nearby visible text only if target repeats, otherwise empty string",
      "newText": "minimal replacement; blank only if safe to delete",
      "sourceEvidence": "precise source contradiction, technical incompatibility, or for a five-condition unsupported claim: source quotation plus explanation of newly introduced material job-specific attribution"
    }
  ]
}
Only include genuinely material FAIL findings. Return all such findings together in the initial review.

ATOMIC FIX RULES:
1. targetText and contextText are visible text, not HTML.
2. Prefer a short, unique targetText.
3. If targetText occurs more than once, provide contextText from the same local sentence/paragraph.
4. newText must make the minimum correction necessary.
5. If deletion is safest, use an empty newText.
6. Do not replace one unsupported project fact with another.
7. Preserve general explanatory material that is not itself a false project claim.

LINK PRESERVATION:
8. targetText may span visible anchor text when necessary.
9. If targetText crosses an existing link, put [[LINK_1]], [[LINK_2]], etc. in newText wherever the original link must remain.
10. Do not reproduce URLs or HTML. The application code restores the original exact anchor HTML.
11. Every crossed link must have a matching placeholder in the same order.

Return the JSON only.`;
  return { success: true, prompt: prompt, postId: postId, title: title };
}
function buildFactCheckFixPrompt(findingsText, sourceColumn) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();

  const requestedCol = (sourceColumn === 'CT') ? 'CT' : 'EU';

  // Same auto-detect as buildFactCheckPrompt — always build the fix against
  // the latest saved state (GE if a previous fix exists), never the
  // untouched original, so consecutive fix passes compound correctly.
  const geHtml = String(sh.getRange(row, 187).getValue() || '').trim();
  const targetCol = geHtml ? 'GE' : requestedCol;
  const html = geHtml || String(sh.getRange(requestedCol + row).getValue() || '').trim();

  if (!html) {
    return { success: false, message: "ERROR: Column " + targetCol + " is empty for this row." };
  }
  if (!findingsText || !findingsText.trim()) {
    return { success: false, message: "ERROR: No fact-check findings provided." };
  }

  const prompt =
`You are correcting specific fabricated or unsupported claims in an article, identified by a fact-check against the original source project.

FINDINGS TO ADDRESS (only fix items marked FACTUAL CONTRADICTION or UNSUPPORTED EMBELLISHMENT):
${findingsText}

RULES:
1. Only rewrite the specific fragments flagged above. Leave everything else in the article completely unchanged.
2. Remove or correct the fabricated/embellished claim so it accurately reflects only what the source evidence supports. Do not replace one fabrication with another — if the source says nothing on a point, remove the claim rather than inventing a different one.
3. Keep the surrounding sentence natural and grammatically correct after the fix.
4. Do NOT rewrite, wrap, or alter any <a href="..."> link or its surrounding anchor tag unless the finding explicitly flags that link's text. Never convert HTML links into markdown-style syntax.
5. Return ONLY a JSON array, no markdown fences, no preamble, in this exact format:
[
  {"fixLabel": "short label for this fix", "old": "exact original fragment from the article below", "new": "exact replacement fragment"}
]
6. Every double quote character inside the old/new string values MUST be escaped as \\" so the result is valid JSON. Use straight quotes only.
7. The "old" fragment must be copied verbatim from the article HTML below — exact characters — so it can be located by an automated find-and-replace. Include enough surrounding text to make each "old" fragment unique within the article.
8. REDUNDANCY CHECK (Hard Lock): Before finalising each replacement, check the surrounding paragraphs in the ARTICLE HTML below (not just the flagged fragment) for another sentence that already describes the same technical step or claim. If the article elsewhere already covers the same point, do not reinsert it — tighten or shorten your replacement so the corrected article does not end up stating the same fact twice in nearby paragraphs.
9. ${ce_getConcreteLanguageStandard()}
10. ${ce_getFixBatchIntegrityCheck()}
11. ${ce_getContentPreservationRules()}
12. ${ce_getSourceEvidencePriorityLock()}
13. ${ce_getCustomerFacingCopyOnlyLock()}
14. ${ce_getLocationVarietyWithoutFabricationLock()}

ARTICLE HTML:
${html}`;

  return { success: true, prompt: prompt };
}

function applyFactCheckFix(rawJson, sourceColumn) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();

  const requestedCol = (sourceColumn === 'CT') ? 'CT' : 'EU';

  // Work on the latest GE if a previous fix exists.
  const geHtmlCheck =
    String(sh.getRange(row, 187).getValue() || '').trim();

  const targetCol =
    geHtmlCheck ? 'GE' : requestedCol;

  if (!rawJson || !rawJson.trim()) {
    return {
      success: false,
      message: "No response provided."
    };
  }

  let fixes;

  let cleaned = rawJson
    .trim()
    .replace(/^```json/i, '')
    .replace(/^```/, '')
    .replace(/```$/, '')
    .trim();

  try {
    fixes = JSON.parse(cleaned);

  } catch (e) {

    try {
      const repaired = cleaned.replace(
        /"(old|new|fixLabel)":\s*"([\s\S]*?)"(?=\s*[,}])/g,
        function(match, key, value) {

          const fixedValue =
            value.replace(/(?<!\\)"/g, '\\"');

          return '"' + key + '": "' + fixedValue + '"';
        }
      );

      fixes = JSON.parse(repaired);

    } catch (e2) {

      return {
        success: false,
        message:
          "Could not parse response as JSON: " +
          e.message +
          " (repair attempt also failed: " +
          e2.message +
          ")"
      };
    }
  }

  if (!Array.isArray(fixes) || fixes.length === 0) {
    return {
      success: false,
      message: "No fixes found in parsed response."
    };
  }

  const html =
    (targetCol === 'GE')
      ? String(sh.getRange(row, 187).getValue() || '')
      : String(sh.getRange(targetCol + row).getValue() || '');

  if (!html) {
    return {
      success: false,
      message:
        "Column " +
        targetCol +
        " is empty for this row."
    };
  }

  const brokenLinkPattern =
    /<a\s+href="\[https?:\/\/[^\]]+\]\(/i;

  const validationErrors = [];
  const seenOld = {};

  // ---------------------------------------------------------
  // VALIDATE EVERY FIX BEFORE CHANGING GE
  // ---------------------------------------------------------

  fixes.forEach(function(fix, index) {

    const label =
      String(
        fix.fixLabel ||
        ('Fix ' + (index + 1))
      );

    if (
      typeof fix.old !== 'string' ||
      fix.old.length === 0
    ) {
      validationErrors.push(
        label + ' — OLD text is missing.'
      );
      return;
    }

    if (typeof fix.new !== 'string') {
      validationErrors.push(
        label +
        ' — NEW must be a string, including "" for deletion.'
      );
      return;
    }

    if (
      brokenLinkPattern.test(fix.old) ||
      brokenLinkPattern.test(fix.new)
    ) {
      validationErrors.push(
        label +
        ' — contains malformed markdown-in-HTML link syntax.'
      );
      return;
    }

    if (seenOld[fix.old]) {
      validationErrors.push(
        label +
        ' — duplicate OLD target in same batch.'
      );
      return;
    }

    seenOld[fix.old] = true;

    const occurrences =
      html.split(fix.old).length - 1;

    if (occurrences === 0) {
      validationErrors.push(
        label +
        ' — OLD fragment not found in article.'
      );
      return;
    }

    if (occurrences > 1) {
      validationErrors.push(
        label +
        ' — OLD fragment matches ' +
        occurrences +
        ' places.'
      );
      return;
    }
  });

  // Reject overlapping OLD targets.
  for (let i = 0; i < fixes.length; i++) {
    for (let j = i + 1; j < fixes.length; j++) {

      const oldA =
        String(fixes[i].old || '');

      const oldB =
        String(fixes[j].old || '');

      if (
        oldA &&
        oldB &&
        (
          oldA.indexOf(oldB) !== -1 ||
          oldB.indexOf(oldA) !== -1
        )
      ) {
        validationErrors.push(
          'Overlapping OLD targets detected between "' +
          (fixes[i].fixLabel || ('Fix ' + (i + 1))) +
          '" and "' +
          (fixes[j].fixLabel || ('Fix ' + (j + 1))) +
          '".'
        );
      }
    }
  }

  // Atomic rule: one bad fix means NOTHING changes.
  if (validationErrors.length > 0) {
    return {
      success: false,
      message:
        'Fact-check fix batch rejected — no changes applied. ' +
        validationErrors.join(' | ')
    };
  }

  // ---------------------------------------------------------
  // APPLY ONLY AFTER THE WHOLE BATCH PASSES
  // ---------------------------------------------------------

  let updatedHtml = html;
  const applied = [];

  fixes.forEach(function(fix, index) {

    updatedHtml =
      updatedHtml.replace(
        fix.old,
        fix.new
      );

    applied.push(
      String(
        fix.fixLabel ||
        ('Fix ' + (index + 1))
      )
    );
  });

  // Fact-check final/working HTML is GE.
  sh.getRange(row, 187).setValue(bc_governLinksAtStage_(updatedHtml, sh, row));

  if (typeof logPipelineResume === 'function') {
    logPipelineResume(
      "W2B.05 — Fact-Check Fix Applied Atomically (GE)",
      ""
    );
  }

  return {
    success: true,
    message:
      fixes.length +
      '/' +
      fixes.length +
      ' fact-check fix(es) applied atomically to GE. Applied: ' +
      applied.join(', '),
    appliedCount: fixes.length
  };
}

function passThroughFactCheckToGE() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();

  const html = String(sh.getRange(row, 151).getValue() || '').trim(); // EU = column 151
  if (!html) {
    return { success: false, message: "ERROR: Column EU (W2B Raw HTML) is empty for this row." };
  }

  sh.getRange(row, 187).setValue(bc_governLinksAtStage_(html, sh, row)); // GE = column 187
  if (typeof logPipelineResume === 'function') {
    logPipelineResume("W2B.05 — Fact-Check PASS, copied EU to GE", "");
  }

  return { success: true, message: "No issues found. EU copied unchanged to column GE (W2B.05 Fixed HTML)." };
}

function ce_checkMaterialCategoryIntegrity_(html) {

  var d = getActiveRowDataMap();

  var material = String(
    d["Stone Type"] ||
    d["Material"] ||
    ""
  ).trim();

  if (!material || !html) {
    return { passed: true };
  }

  var materialLower = material.toLowerCase();

  // These governed materials are NOT natural stone.
  // They may legitimately be described as tiles, but not as stone.
  var nonStoneMaterialTests = [
    { test: /ceramic/i, articleTest: /\bceramic\b/i },
    { test: /porcelain/i, articleTest: /\bporcelain\b/i },
    { test: /terracotta/i, articleTest: /\bterracotta\b/i },
    { test: /quarry/i, articleTest: /\bquarry\b/i },
    { test: /victorian/i, articleTest: /\bvictorian\b/i },
    { test: /edwardian/i, articleTest: /\bedwardian\b/i },
    { test: /encaustic/i, articleTest: /\bencaustic\b/i }
  ];

  var governedTest = null;

  for (var i = 0; i < nonStoneMaterialTests.length; i++) {
    if (nonStoneMaterialTests[i].test.test(materialLower)) {
      governedTest = nonStoneMaterialTests[i].articleTest;
      break;
    }
  }

  // Natural stone and other materials are not governed by this hard check.
  if (!governedTest) {
    return { passed: true };
  }

  var plainText = String(html)
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  var sentences =
    plainText.match(/[^.!?]+[.!?]+|[^.!?]+$/g) || [];

  var directStoneReference =
    /\b(each stone|the stone|this stone|that stone|stone behaves|stone surface|stone itself|natural stone)\b/i;

  var comparisonLanguage =
    /\b(differs? from|different from|unlike|compared with|compared to|rather than|as opposed to|is not natural stone|isn't natural stone|not a natural stone)\b/i;

  for (var s = 0; s < sentences.length; s++) {

    var sentence = String(sentences[s]).trim();

    if (
      governedTest.test(sentence) &&
      directStoneReference.test(sentence) &&
      !comparisonLanguage.test(sentence)
    ) {

      return {
        passed: false,
        finding:
          'ISSUE 1\n' +
          'TYPE: FACTUAL CONTRADICTION\n' +
          'CLAIM: "' + sentence + '"\n' +
          'WHY IT FAILS: The governed material is "' +
          material +
          '", but this sentence directly describes that material as stone. ' +
          'The governed material may legitimately be described as tile, but it must not be reassigned to the natural-stone material category.\n' +
          'SOURCE EVIDENCE: Governed Material = "' +
          material +
          '".'
      };
    }
  }

  return { passed: true };
}