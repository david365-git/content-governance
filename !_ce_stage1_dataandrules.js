/**
 * ================================================================================
 * ce_stage1_dataandrules.gs - STAGE1 DATAANDRULES
 * ================================================================================
 * 
 * Data loading and rule injection for content generation
 * 
 * Part of Abbey Floor Care Content Pipeline v77+
 * Reorganised: March 2026
 * ================================================================================
 */
/* ============================================================
   PRIMARY QUERY CLUSTER SANITISER
   Cleans up "Primary Query Cluster Owned" before it gets baked
   into a Hard Lock instruction (Problem Angle opening sentence).
   Trims whitespace and flags likely malformed values (e.g.
   duplicated/overlapping phrases) for manual review rather than
   silently rewriting meaning.
============================================================ */
function sanitisePrimaryClusterPhrase(raw) {
  var text = String(raw || "").trim().replace(/\s+/g, " ");
  if (!text) return text;

  // Detect duplicate or near-duplicate segments joined by "and"
  // e.g. "slate floor cleaner and clean slate in Nottingham"
  var parts = text.split(/\s+and\s+/i);
  if (parts.length === 2) {
    var wordsA = parts[0].toLowerCase().split(/\s+/).filter(function(w) { return w.length > 3; });
    var wordsB = parts[1].toLowerCase().split(/\s+/).filter(function(w) { return w.length > 3; });
    var overlap = wordsA.filter(function(w) { return wordsB.indexOf(w) > -1; });
    if (overlap.length >= 2) {
      return text + " [REVIEW: Primary Query Cluster Owned may be malformed — overlapping phrases detected. Check this value in the posts sheet.]";
    }
  }

  return text;
}


function getSiloRoleMap(material) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("posts");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const activeRowData = getActiveRowDataMap();
  const activePostID = String(activeRowData["Post ID"]);
  
  const matIndex   = headers.indexOf("Stone Type");
  const roleIndex  = headers.indexOf("Entity Role");
  const titleIndex = headers.indexOf("Title");
  const idIndex    = headers.indexOf("Post ID");
  const urlIndex   = headers.indexOf("URL");
  
  let siloMap = [];
  for (let i = 1; i < data.length; i++) {
    const postID = String(data[i][idIndex]);
    if (postID === activePostID) continue;
    if (data[i][matIndex] === material) {
      const role  = data[i][roleIndex]  || "Role Not Yet Assigned";
      const title = data[i][titleIndex] || "No Title";
      const url   = data[i][urlIndex]   || "";
      siloMap.push(`- URL: ${url} | ROLE: ${role} | TITLE: ${title}`);
    }
  }
  return siloMap.length > 0
    ? `--- SEMANTIC SILO MAP (${material}) ---\n${siloMap.join('\n')}`
    : `--- SEMANTIC SILO MAP ---\nNo other ${material} posts found yet.`;
}



function buildStage1APrompt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const d = getActiveRowDataMap();

  const postID        = d["Post ID"];
  const material      = d["Stone Type"];
  const articleType   = d["Article Type"] || "General";
  const primaryEntity = d["Primary Entity"] || "Not Set";
  const govSummary    = d["Page Governance Summary"] || "No summary provided.";

  const techMatrix = getTechnicalMatrix(material);
  const siloMap    = getSiloRoleMap(material);
  const dna        = getTechDNA(material);

  const fragility      = dna ? dna["Fragility Calibration"]     || "" : "";
  const confidence     = dna ? dna["Transformation Confidence"] || "High" : "High";
  const anchorEntities = dna ? dna["Anchor Entities"]           || "" : "";
  const prohibited     = dna ? dna["Prohibited LLM Behaviours"] || "" : "";
  const ukLocalisation = dna ? dna["UK Localisation Notes"]     || "" : "";
  const supportingEntities = d["Supporting Entities Core"] || "None";

  // Read stored Authority Brief from col DN
  const authorityBrief = getAuthorityBrief();

  // Read stored Problem Angle from posts sheet
  const problemAngle = (function() {
    var headers = SpreadsheetApp.getActiveSpreadsheet()
                    .getSheetByName("posts")
                    .getRange(1, 1, 1, SpreadsheetApp.getActiveSpreadsheet().getSheetByName("posts").getLastColumn())
                    .getValues()[0].map(function(h){ return String(h).trim(); });
    var idx = headers.indexOf("Problem Angle");
    if (idx === -1) return "";
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("posts");
    var row = sheet.getActiveRange().getRow();
    return row >= 2 ? String(sheet.getRange(row, idx + 1).getValue() || "").trim() : "";
  })();

  if (!postID) return "ERROR: No Post ID found.";

  const exportSheet  = ss.getSheetByName("site-export");
  const exportData   = exportSheet.getDataRange().getValues();
  const exportHeaders = exportData[0];
  const idIndex      = exportHeaders.indexOf("ID");
  const htmlIndex    = exportHeaders.indexOf("Full Post HTML");

  let rawHTML      = "";
  let originalHTML = "";
  for (let i = 1; i < exportData.length; i++) {
    if (String(exportData[i][idIndex]) === String(postID)) {
      rawHTML      = String(exportData[i][htmlIndex] || "");
      originalHTML = cleanHtmlForLLM(rawHTML);
      break;
    }
  }

  return `
STAGE 1A — DATA LOADING
ROLE: Senior UK SEO & Stone Restoration Strategist
ARTICLE TYPE: ${articleType}

--- SYSTEM INSTRUCTION: HOLD ---
This is Stage 1A of the pipeline. Your only job is to ingest and retain all data below.
Do not begin any rewrite. Do not produce any HTML. Do not summarise or analyse.
Do not apply any rules — rules arrive in Stages 1B, 1C, and 1D.

Your entire response must be exactly: "Stage 1A data received. Waiting for Stage 1B."
No other text. No analysis. No commentary. No image inventory.

${authorityBrief ? "--- AUTHORITY BRIEF (Pre-Pipeline Governance) ---\nThis brief was generated from entity analysis of the original page.\nIt governs the entire pipeline — all stages must comply with the constraints below.\n" + authorityBrief + "\n--- END AUTHORITY BRIEF ---\n\n" : ""}${(function() {
    var problemAngle = d["Problem Angle"] || "";
    var primaryCluster = sanitisePrimaryClusterPhrase(d["Primary Query Cluster Owned"] || "");
    var angleCode = "";
    var firstLine = problemAngle.split("\n")[0].trim();
    var dashIndex = firstLine.indexOf(" — ");
    if (dashIndex > -1) angleCode = firstLine.substring(0, dashIndex).trim();
    if (!problemAngle) return "";
    return "--- PROBLEM ANGLE (Opening Sentence Governance) ---\n" +
           "This article addresses: " + primaryCluster + "\n" +
           "Approach: " + problemAngle.split("\n")[0] + "\n" +
           "OPENING SENTENCE LAW: The opening sentence of the article header and the first paragraph of Section 1\n" +
           "must directly reflect this problem angle. The reader must recognise their specific problem\n" +
           "from the very first sentence. Do not open with a general statement about the material.\n" +
           "--- END PROBLEM ANGLE ---\n\n";
  })()}--- PAGE GOVERNANCE DATA ---
POST ID: ${postID}
PRIMARY ENTITY: ${primaryEntity}
ARTICLE TYPE: ${articleType}
GOVERNANCE SUMMARY: ${govSummary}
MATERIAL: ${material}
TRANSFORMATION CONFIDENCE: ${confidence}

--- GOVERNANCE FIELDS (from sheet — read-only context) ---
NEW H1: ${d["New H1"] || "Not set"}
NEW META TITLE: ${d["New Meta Title"] || "Not set"}
NEW META DESCRIPTION: ${d["New Meta Description"] || "Not set"}
SCHEMA (JSON-LD): ${d["Schema (JSON-LD)"] || "Not set"}
PRIMARY SEARCH TERM: ${d["Primary Search Term"] || "Not set"}
STRATEGIC REASONING: ${d["Strategic Reasoning"] || "Not set"}
TOPICAL INTENT: ${d["Topical Intent"] || "Not set"}
MATRIX ROLE: ${d["Matrix Role"] || "Not set"}
KEY DECISIONS EXPLAINED: ${d["Key Decisions Explained"] || "Not set"}

--- MATERIAL BEHAVIOUR CALIBRATION ---
${fragility}

--- PROHIBITED LLM BEHAVIOURS FOR THIS MATERIAL ---
${prohibited}

--- TECHNICAL SPECIFICATION MATRIX (HARD-LOCK) ---
${techMatrix}

--- SEMANTIC REQUIREMENTS ---
CORE SUPPORTING ENTITIES: ${supportingEntities}
ANCHOR ENTITIES (CO-OCCURRENCE TARGETS): ${anchorEntities}

--- SILO CONTEXT (INTERNAL LINKING GOVERNANCE) ---
${siloMap}

--- INTERNAL LINKING DATA ---
1. Use the full URL from the Silo Map for all href values — no placeholders.
2. Integrate 5-8 relevant links from the Silo Map naturally.
3. Anchor text MUST match the intent of the target Role.

--- UK LOCALISATION DATA ---
1. Use British English exclusively (e.g., Specialist, Metres, Taps, Splashback).
2. ${ukLocalisation}

--- SECTION CONTENT GOVERNANCE ---
1. Primary Entity Focus: 80% of the word count must serve the PRIMARY ENTITY: ${primaryEntity}.
2. Peripheral Summarisation: Sections that do not strengthen the Primary Entity are limited to
   maximum 300 words. Summarise principles and use the full URL from the Silo Map to link to the designated authority page.

--- ORIGINAL HTML ---
${originalHTML}

ALL CHUNKS NOW LOADED.
Stage 1A data received. Waiting for Stage 1B.
`.trim();
}

/* ============================================================
   STAGE 1B — STRUCTURAL RULES
   H2 naming, ID mapping, hub-intro architecture, HTML structure.
   Planning rules only — applied before any content is written.
============================================================ */



function buildStage1BPrompt() {
  return `
STAGE 1B — STRUCTURAL RULES
ROLE: Senior UK SEO & Stone Restoration Strategist

--- SYSTEM INSTRUCTION: HOLD ---
This is Stage 1B. Stage 1A data is already held in your context window.
Your job is to ingest these structural rules and retain them for execution in Stage 2B.
Do not produce any HTML. Do not write any content.
Your entire response must be exactly: "Stage 1B structural rules received."
Nothing else.

--- RULE 1: H2 SUBHEADINGS (Hard Lock) ---
Every H2 must name the primary entity or mechanism explicitly.
Framing constructions that open with "Where…", "When…", "How to…" without the material
or process name present are prohibited. The stone type, defect name, or intervention
method must appear in the heading itself.
❌ WRONG: "Where DIY Filling Has Limits"
✅ CORRECT: "Cementitious Grout: Appropriate Use Cases and Limits"

--- RULE 2: H2 AND ID MAPPING (Hard Lock) ---
Every section id value, H2 heading text, and section content subject must all describe
the same topic. A section whose id says "sealing" must have an H2 about sealing and
content about sealing — not content about polishing.
The quick-links block at the top of the article must reference the correct id value
for each link. Every anchor must point to the section that covers that topic —
not to the section that happens to appear in that position in the document.
Before writing any HTML, list every planned section in this format and confirm alignment:
  SECTION [N]: id="[id]" | H2="[heading]" | CONTENT="[one sentence on what this covers]"
If any row shows a mismatch — correct it before writing.

--- RULE 3: HTML STRUCTURE (Hard Lock) ---
Every opening tag must have a matching closing tag.
No content may appear outside its intended parent element.
The <header> element must contain only the intro paragraph — a single <p> tag.
Quick-links navigation must never appear inside <header>.
For Hub Pages, quick-links must appear in <section id="hub-intro"> immediately after </header>.
No section content belongs inside <header>.
Every <section> must have an id attribute that matches its quick-links anchor.
No <p> tags may appear as direct siblings of <header> outside the header element.
All <figure> elements must appear inside their parent <section> — not between sections.

--- RULE 4: HUB-INTRO NAVIGATION-ONLY (Hard Lock — hub pages only) ---
The hub-intro section has one job: routing the homeowner to the correct section.
It must contain:
  - One orienting sentence only — maximum 15 words
  - The navigation list with symptom-led anchor labels
It must NOT contain:
  - Any restatement of header content
  - Any explanation of material behaviour or mechanisms
  - Any paragraph that would also work as a header paragraph
If more than one sentence appears before the navigation list — delete the excess.

--- HOLD REMINDER ---
Retain all four structural rules. Respond only with: "Stage 1B structural rules received."
`.trim();
}

/* ============================================================
   STAGE 1C — PARAGRAPH GENERATION RULES
   Topic sentence, framing verbs, bold, noun repetition,
   paragraph length. Sentence-level generation rules.
============================================================ */



function buildStage1CPrompt() {
  return `
STAGE 1C — PARAGRAPH GENERATION RULES
ROLE: Senior UK SEO & Stone Restoration Strategist

--- SYSTEM INSTRUCTION: HOLD ---
This is Stage 1C. Stages 1A and 1B are already held in your context window.
Your job is to ingest these paragraph-level generation rules and retain them for Stage 2B.
Do not produce any HTML. Do not write any content.
Your entire response must be exactly: "Stage 1C paragraph rules received."
Nothing else.

--- RULE 5: PARAGRAPH ARCHITECTURE (Hard Lock) ---
Generate each paragraph by writing the topic sentence first. The semantic payload —
the named entity or mechanism being discussed — must appear in the first clause of the
first sentence. Each paragraph's opening sentence must be able to stand alone as a
complete, indexable statement of the point that follows.

ACTIVE VOICE REQUIREMENT (Hard Lock):
The subject of each sentence must perform the action, not receive it. Do not use passive
constructions ("was needed", "was applied", "was removed", "is treated", "was chosen").
Rewrite passive sentences so a named or implied active agent performs the verb.
❌ WRONG: "Controlled alkaline cleaning was needed because residue had built up."
✅ CORRECT: "Technicians applied controlled alkaline cleaning because residue had built up."
❌ WRONG: "The old sealer was stripped using a solvent-based remover."
✅ CORRECT: "A solvent-based remover stripped the old sealer."
This rule applies within the third-person voice already required by the Article Type
Control Lock — it does not authorise first-person "we" language.

Prohibited opening constructions:
"In…", "For…", "If…", "When…", "There are…", "It is…", "One of…", "Many…", "This is…"

If the topic sentence opens with any prohibited construction — discard it and rewrite
before continuing to the supporting sentences.

FRAMING VERB PROHIBITION: The first verb must state what the entity does, causes,
prevents, or produces — not introduce it. Prohibited first verbs:
explains, explains why, are explained, describes, refers to, presents, discusses,
covers, addresses, outlines, details, shows, illustrates, highlights, demonstrates,
demonstrate, appear in, appear across, can be seen in, documents, reveals.

❌ WRONG: "Travertine floor behaviour explains why this defect appears frequently."
✅ CORRECT: "Travertine's natural cavity structure spreads residue from tile pits into grout lines."
❌ WRONG: "Domestic steam cleaning presents additional risks for travertine grout."
✅ CORRECT: "Domestic steam cleaning forces liquified soil deeper into grout pores because it heats without extracting."

INTERNAL LINK SIGNPOSTING PROHIBITION:
Passive signposting constructions are prohibited as paragraph openers:
"…are explained in", "…are examined in", "…are covered in", "…can be seen in",
"…appear in", "…is discussed in", "…details are available in"
A paragraph must never exist solely to point to another page.
Signposting sentences must be placed as the final sentence of a substantive paragraph.

STANDALONE SIGNPOSTING PARAGRAPH PROHIBITION:
A single-sentence paragraph whose sole purpose is to point to another page is prohibited.
The link must always be attached to a substantive paragraph as its final sentence.
❌ WRONG: "<p>Stone-safe cleaning principles are explained in detail in [link].</p>"
✅ CORRECT: "[Substantive content.] Practical maintenance routines are explained in [link]."

--- RULE 6: BOLD TEXT (Hard Lock) ---
Do not bold the opening entity or mechanism name in any paragraph.
Entity prominence is achieved through sentence position alone — not typographic emphasis.
Maximum 3–4 bold instances per page total.
Do not bold anchor entities, co-occurrence terms, or SEO target phrases.
MARKDOWN BOLD PROHIBITION: Never use **text** syntax. Use <strong>text</strong> only.
Any instance of **text** in the output is an error.

--- RULE 7: PARAGRAPH LENGTH VARIATION (Hard Lock) ---
Do not produce three or more consecutive single-sentence paragraphs.
Where three single-sentence paragraphs appear in sequence — consolidate at least two
into a single paragraph of 2–3 sentences before proceeding.

--- RULE 8: CROSS-PARAGRAPH NOUN REPETITION (Hard Lock) ---
Do not open more than two consecutive paragraphs with the same noun or noun phrase.
If three or more consecutive paragraphs would open with the same subject noun —
restructure at least one to lead with a related entity, consequence, or mechanism.
❌ WRONG: Four consecutive paragraphs opening with "Travertine grout blackening…"
✅ CORRECT: Vary with mechanism, consequence, or affected zone as subject.

--- HOLD REMINDER ---
Retain all four paragraph generation rules. Respond only with: "Stage 1C paragraph rules received."
`.trim();
}

/* ============================================================
   STAGE 1D — OUTPUT QUALITY RULES
   Signposting audit, image preservation, COMM flag,
   defect completion, process section depth, cluster tone,
   footer CTA, schema flag, DOM flag.
   Review rules — applied before the HTML is finalised.
============================================================ */



function buildStage1DPrompt() {
  const d = getActiveRowDataMap();
  const govSummary = d["Page Governance Summary"] || "";
  const articleType = d["Article Type"] || "General";

  let commInstruction = "";
  if (govSummary.includes("[COMM:No]")) {
    commInstruction = `RULE 10: COMMERCIAL CONTENT PROHIBITION [COMM:No] (Hard Lock):
Do not include affiliate tables, Amazon links, Shop Now buttons, product recommendation
blocks, or any commercial call-to-action. Remove any such elements from the original HTML.
Do not create new commercial sections. Commercial elements must not appear anywhere in output.`;
  } else if (govSummary.includes("[COMM:BelowFold]")) {
    commInstruction = `RULE 10: COMMERCIAL CONTENT GOVERNANCE [COMM:BelowFold]:
Commercial product references permitted but must be repositioned to a single dedicated
section at the end of all editorial content, immediately before the author box.
LINK PRESERVATION RULE (ALL OUTBOUND COMMERCIAL LINKS — Hard Lock):
Every outbound product, supplier, and equipment link present in the original HTML must
appear in the commercial section output. This includes supplier sites, retailer links,
and any other third-party product or equipment URL — not just Amazon affiliate links.
Extract the exact href from the original anchor. Do not expand, reconstruct, or substitute URLs.
The href must be byte-for-byte identical to the original.
Do not create new product recommendations — preserve only links already in the original HTML.
Convert Amazon affiliate links to: <a href="https://amzn.to/XXXXXXX">Product Name</a>
All other outbound commercial links: preserve href and anchor text exactly as supplied.`;
  } else if (govSummary.includes("[COMM:Yes]")) {
    commInstruction = `RULE 10: COMMERCIAL CONTENT [COMM:Yes]:
Commercial content including affiliate links may remain in their current position.
MARKUP VALIDITY (Hard Lock): While preserving the commercial content's position and product data exactly, correct any malformed HTML in the block before output — every opening tag must have a matching closing tag, and no <p> tag may appear orphaned or unclosed inside a <td> or <div>. Do not alter product names, prices, URLs, or image sources — only correct structurally invalid tag nesting.

SCOPE BOUNDARY OVERRIDE (Hard Lock — takes priority over the retention rule above):
Before retaining any commercial block, check its topic against the Rewrite Brief's SCOPE
BOUNDARY above. If the Rewrite Brief explicitly lists an intent for removal (e.g. "remove
Maintenance / Aftercare entirely", "remove Cleaning entirely") and this commercial block's
heading or product focus matches that removed intent, you MUST remove the block entirely —
do not retain it, do not reposition it, do not exempt its heading from the H2 lock.
The Rewrite Brief's Scope Boundary is the more specific, more recently governed instruction
and always overrides general commercial-content preservation rules.
This applies regardless of the block's original heading, position, or formatting.
If the commercial block's topic is NOT covered by any removed intent in the Scope Boundary,
retain it per the rule above as normal.`;
  } else {
    commInstruction = `RULE 10: COMMERCIAL CONTENT:
No COMM flag detected. Do not add commercial content. Preserve existing commercial
elements only if present in the original HTML.`;
  }

  const schemaFlag = `RULE 11: SCHEMA FLAG: HowTo schema is not used on this site. Do not insert any schema marker comment.`;

  return `
STAGE 1D — OUTPUT QUALITY RULES
ROLE: Senior UK SEO & Stone Restoration Strategist

--- SYSTEM INSTRUCTION: HOLD ---
This is Stage 1D. Stages 1A, 1B, and 1C are already held in your context window.
Your job is to ingest these output quality rules and retain them for Stage 2B.
These rules are applied as a review pass before the HTML is finalised — not during writing.
Do not produce any HTML. Do not write any content.
Your entire response must be exactly: "Stage 1D output quality rules received."
Nothing else.

--- RULE 9: PRE-OUTPUT SIGNPOSTING AUDIT (Hard Lock) ---
Before finalising HTML output, scan every <p> tag containing an <a href> link.
Apply both checks to every such paragraph:

CHECK A — STANDALONE PARAGRAPH:
If the <p> contains only one sentence and that sentence contains a link — FAIL.
Merge it as the final sentence of the preceding <p> and delete the standalone paragraph.

CHECK B — ISOLATED CLOSING SENTENCE:
If the link appears only in the final sentence and that sentence is the only sentence
in the paragraph — FAIL. The signposting sentence must share a <p> with at least one
substantive sentence delivering technical or editorial content.

CLOSING SENTENCE RULE: A signposting sentence must not be the only sentence after
the last substantive point. Either place it between two substantive sentences, or add
a further substantive sentence after it.

❌ FAIL: "<p>Repair methods are discussed in [link].</p>"
✅ PASS: "<p>[Substantive.] [Substantive.] Repair methods are in [link]. Correct technique prevents further damage.</p>"

${commInstruction}

${schemaFlag}

--- RULE 12: DOM FLAG ---
If the original HTML contains an element with class 'authority-ninja-box' or 'author-bio',
insert this comment immediately before that element — not at the start or end of the output,
but at the exact position where the element appears:
<!-- [DOM: Verify bio div renders inside <article> tags in WordPress theme output] -->

--- RULE 13: IMAGE PRESERVATION (Hard Lock) ---
All figure, img, and figcaption elements from the original HTML must appear in the output.
This is a count-verified requirement — not a best-effort instruction.

STEP 1 — COUNT BEFORE WRITING:
Before writing any HTML, count every <figure> and standalone <img> element in the original HTML.
This includes minimal legacy images — a bare <img> inside a <figure> with only inline float
styles and no srcset, sizes, or width counts as one image.
Record this number internally as ORIGINAL_IMAGE_COUNT.

STEP 2 — IDENTIFY EACH IMAGE:
For each image record:
  - src attribute value (the URL — this is the permanent identifier)
  - alt attribute value
  - Whether it is full WordPress format (has srcset, sizes, class) or legacy format (src and alt only)

STEP 3 — PLACE DURING WRITING:
Place each image at the most contextually relevant point in the new structure.
If a section that contained an image no longer exists — place the image in the nearest
thematically related section. Do not discard any image regardless of its format.

LEGACY IMAGE UPGRADE RULE:
If an image is legacy format (src and alt only, no srcset, sizes, width, height, class) —
upgrade it to current WordPress figure format using this template:
<figure class="wp-caption aligncenter" style="width: 100%">
<img loading="lazy" decoding="async" src="[ORIGINAL SRC]" alt="[ORIGINAL ALT]" />
<figcaption class="wp-caption-text">[Write one descriptive sentence for the caption]</figcaption>
</figure>
Preserve the original src and alt exactly. Write a new caption based on the image context.
Remove the original float inline styles — use aligncenter instead.

STEP 4 — VERIFY BEFORE FINALISING:
Before producing final output, count every <figure> and <img> in your output.
If output count is less than ORIGINAL_IMAGE_COUNT — STOP.
Find the missing image(s) by src URL and insert them before continuing.

FULL FORMAT IMAGES:
Images that already have srcset, sizes, width, height, and class must be copied
character-for-character with all attributes unchanged.

ZERO TOLERANCE: A missing image — regardless of format — is a hard failure. No exceptions.

--- RULE 13B: SECTION 1 SCOPE BOUNDARY FOR CASE STUDIES (Hard Lock) ---
This rule applies only when Article Type is Case Study.
Section 1 must cover entry condition and problem identification ONLY — what the floor looked like, what the homeowner noticed, and why normal cleaning had stopped working.
Section 1 must NOT preview, name, or rule out any intervention method, technique, or approach.
Prohibited in Section 1: any sentence stating what will NOT be done (e.g. "this was not a job for flattening, grinding or changing the surface"), any sentence naming the chosen remedy category, any sentence explaining why a technique is appropriate.
That reasoning belongs in Section 2 (Problem & Intervention), where the technique choice and its justification are documented.
SELF-CHECK BEFORE OUTPUT: Before finalising Section 1, scan every sentence for words describing what was or was not done to correct the floor (flatten, grind, strip, hone, seal, apply, remove, treat). Any such sentence must be deleted from Section 1 or moved to Section 2.
WORD COUNT SELF-CHECK: Before finalising Section 1, count the total words in Section 1 including all H3 sub-blocks but excluding any commercial product table. If the count exceeds the WORD BUDGET maximum specified in the approved section plan, cut content — do not carry it forward as extra length. Prioritise cutting supporting detail over cutting the core problem description.

--- RULE 14: NAMED DEFECT COMPLETION (Hard Lock) ---
Every technical defect named in the article must be explained to completion.
A defect is complete when all three elements are present in the same section:
  ELEMENT 1: What it is — one sentence defining the defect.
  ELEMENT 2: What the homeowner sees — one sentence on visible symptom.
  ELEMENT 3: What is done about it — one sentence naming the correction.
Applies to: etching, micro-scratching, lippage, picture-framing, efflorescence,
filler collapse, delamination, sealer failure, surface fines loss, residue lock-in,
grout haze, colour loss, fading — and any other named defect.
Do not name a defect without completing all three elements before the next paragraph.

PICTURE-FRAMING EXAMPLE:
✅ Element 1: "Lippage is the height variation between adjacent tiles at their shared edge."
✅ Element 2: "When polishing runs across tiles with lippage, the machine rides higher edges and polishes them less, leaving a darker border — this is picture-framing."
✅ Element 3: "Milling before honing levels the tiles and eliminates the condition."

--- RULE 15: MINIMUM PARAGRAPH DEPTH FOR PROCESS SECTIONS (Hard Lock) ---
Every process section — cleaning, honing, polishing, sealing, restoration, burnishing,
stripping, re-grouting, recolouring — must contain minimum three body paragraphs
before any figure element or closing internal link.
Each paragraph must contain minimum two sentences.
A single-sentence paragraph naming a process step without expanding it is prohibited.

PROCESS PARAGRAPH ROLE VALIDATION (Strict Enforcement):
Each of the three paragraphs must explicitly match one of these roles.
If a paragraph does not clearly match its role — rewrite it before continuing.

  PARAGRAPH 1 — SYMPTOM (Hard Lock):
  - MUST open with a homeowner-visible problem or condition using "If your..." framing
  - MUST NOT contain process explanation in the first sentence
  - MUST describe what the homeowner sees, feels, or notices
  ❌ FAIL: "Professional cleaning works by removing contamination from the surface."
  ✅ PASS: "If your marble keeps looking dirty after mopping, the soil has moved below the surface."

  PARAGRAPH 2 — MECHANISM (Hard Lock):
  - MUST explicitly name the process in the first sentence
  - MUST explain how the process works physically or chemically
  - Must include at least one specific technical detail — chemistry, grit, tool, or sequence
  ❌ FAIL: "The surface is treated to remove contamination."
  ✅ PASS: "Professional cleaning works through controlled alkaline chemistry and wet vacuum extraction."

  PARAGRAPH 3 — OUTCOME (Hard Lock):
  - MUST explicitly describe the result for the homeowner
  - MUST include a comparison between professional and incomplete work
  - MUST include one of: "significantly better", "easier to maintain", or "uniform finish"
  ❌ FAIL: "The treatment improves the appearance of the floor."
  ✅ PASS: "The floor will look significantly better once the contamination layer is removed — and easier to maintain than a worn surface."

SELF-CHECK BEFORE OUTPUT (Hard Lock):
Before writing final HTML, internally label each paragraph in every process section:
  [SYMPTOM] — [MECHANISM] — [OUTCOME]
If any label is missing or a paragraph does not clearly match its role — rewrite before output.
Do not include labels in the output HTML.
If all three roles are not present before any <figure> — the section is a hard failure.

FIGURE PLACEMENT RULE:
Figures must appear AFTER all three role paragraphs — not between them.
A figure placed before paragraph 3 is a structural failure.

SURFACE FINISH DISTINCTION RULE: Where a section discusses surface refinishing outcomes,
include the distinction sentence from the Stone-Specific Prompt Rules block (injected in Stage 2B part 2).
Use the exact framing provided for this stone type — do not substitute generic polish/hone language.

--- RULE 16: LIST GOVERNANCE (Hard Lock) ---
Maximum two list elements permitted per article. A list is only permitted when ALL three
criteria are met simultaneously:
  CRITERION 1: Items are genuinely discrete — each stands alone without needing explanation.
  CRITERION 2: Four or more items — fewer than four reads better as prose ("x, y, and z").
  CRITERION 3: The reader needs to scan and identify — not read and understand.

APPROVED list use cases:
  ✅ Discrete items a reader checks against their own situation (e.g. acidic liquids that etch marble)
  ✅ Sequential numbered steps in a Method Guide where order is critical
  ✅ A set of named products or materials with no explanatory dependency

PROHIBITED list use cases:
  ❌ Stain types or defect categories — these require diagnostic explanation, not labels
  ❌ Benefits of professional treatment — these require substantive prose
  ❌ Causes of a problem — prose with mechanism is required, not a label list
  ❌ Any list that replaces a paragraph — a list supplements prose, never substitutes it
  ❌ A list section immediately followed by another list section

ENFORCEMENT: Before finalising output, count all <ul> and <ol> elements.
If count exceeds 2 — convert excess lists to prose. Apply CRITERION checks to remaining lists.
If a list fails any criterion — convert to prose immediately.

--- RULE 17: H2 HEADING GOVERNANCE (Hard Lock) ---
Every H2 heading must pass all four tests before it is used:

TEST 1 — READER RECOGNITION (PRIORITY TEST):
The heading must name a condition, problem, or question the homeowner recognises from their
own experience. A homeowner scanning headings must find themselves in the heading.
  ❌ FAIL: "Marble Surface Cap and Absorption Behaviour"
  ✅ PASS: "Why stains spread and go deeper than the original spill"
  ❌ FAIL: "Marble Formation and Calcite Crystal Behaviour in Stain Removal"
  ✅ PASS: "Why marble stains form differently from other floors"

TEST 2 — NO TRAILING GENERIC SUFFIXES:
Remove: "Explained", "Overview", "Guide", "Information", "Details", "Considerations"
These add no meaning. If the heading requires a suffix to work — rewrite the heading.
  ❌ FAIL: "Marble Soil and Stain Behaviour Explained"
  ✅ PASS: "Why the floor looks dirty again after cleaning"

TEST 3 — NO STONE TYPE PREFIX REPETITION:
Do not begin every H2 with the stone type name. If the article is about marble, every
heading does not need to start with "Marble". Use it selectively where it aids clarity.
  ❌ FAIL: "Marble Stain Types and Identification" / "Marble Cleaning Mechanisms" /
           "Marble Sealing Behaviour" — three consecutive marble-prefixed headings
  ✅ PASS: Vary — "Stain types and how to identify them" / "Why cleaning has limits" /
           "How sealing changes the floor's response to spills"

TEST 4 — MATCHES SECTION CONTENT:
The heading must accurately describe what the section covers. A reader who reads only
the heading must know what they will find in that section.

PRIORITY RULE: Test 1 (Reader Recognition) takes precedence over Test 4 (Content Accuracy).
If a heading passes Test 1 but seems to fail Test 4 — favor Test 1.
Symptom-first and question-format headings are preferred even if they don't fully describe mechanism content.

ENFORCEMENT: Before finalising output, apply all four tests to every H2.
Prioritize reader recognition over content description accuracy.

--- RULE 18: CLUSTER TONE CONSISTENCY (Hard Lock) ---
The cluster label from Stage 0 must be maintained across every paragraph of the header
block — not just the opening sentence.

CLUSTER B HEADER — prohibited words and phrases anywhere in the header:
"reassuring", "reassurance", "don't worry", "the good news is", "fortunately",
"happily", "rest assured", "in most cases there is no need to panic"
Test: read the full header. If it sounds like comfort rather than confidence — rewrite.

CLUSTER A HEADER — prohibited: mechanism-first sentences, damage-first sentences,
any sentence naming the problem before naming the solution or normal expectation.

CLUSTER C HEADER — prohibited: sentences qualifying professional expertise, sentences
leading with homeowner uncertainty rather than specialist knowledge.

--- RULE 19: FOOTER CTA DIRECT NEXT STEP (Hard Lock) ---
The footer must end with a direct next step for the homeowner.
The final element must be one of:
  - A direct instruction: e.g. "Contact us to arrange a no-obligation floor assessment."
  - A button anchor: <a href="/contact" class="cta-button">Book a floor assessment</a>
  - An action sentence: e.g. "A floor assessment identifies what is needed before work is agreed."
A footer ending on an explanation sentence without directing the homeowner to act is incomplete.

--- HOLD REMINDER ---
Retain all output quality rules. Respond only with: "Stage 1D output quality rules received."
`.trim();
}

function testStage1B() {
  Logger.log(buildStage1BPrompt());
}

function saveW1AOutput() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('posts');
  var row = sheet.getActiveRange().getRow();
  if (row < 2) return { success: false, message: 'Select a data row first' };

  var d = getActiveRowDataMap();
  var postID = d["Post ID"];
  if (!postID) return { success: false, message: 'No Post ID found' };

  // Get original HTML separately
  var exportSheet = ss.getSheetByName("site-export");
  var exportData = exportSheet.getDataRange().getValues();
  var exportHeaders = exportData[0];
  var idIndex = exportHeaders.indexOf("ID");
  var htmlIndex = exportHeaders.indexOf("Full Post HTML");
  var originalHTML = "";
  for (var i = 1; i < exportData.length; i++) {
    if (String(exportData[i][idIndex]) === String(postID)) {
      originalHTML = cleanHtmlForLLM(String(exportData[i][htmlIndex] || ""));
      break;
    }
  }

  // Build each block separately
  var techMatrix   = getTechnicalMatrix(d["Stone Type"] || "");
  var siloMap      = getSiloRoleMap(d["Stone Type"] || "");
  var dna          = getTechDNA(d["Stone Type"] || "");
  var fragility    = dna ? dna["Fragility Calibration"]     || "" : "";
  var prohibited   = dna ? dna["Prohibited LLM Behaviours"] || "" : "";
  var ukLocal      = dna ? dna["UK Localisation Notes"]     || "" : "";
  var anchorEntities     = dna ? dna["Anchor Entities"]     || "" : "";
  var supportingEntities = d["Supporting Entities Core"]    || "None";
  var authorityBrief     = getAuthorityBrief();
  var primaryEntity      = d["Primary Entity"] || "Not Set";
  var articleType        = d["Article Type"]   || "General";
  var govSummary         = d["Page Governance Summary"] || "No summary provided.";
  var confidence         = dna ? dna["Transformation Confidence"] || "High" : "High";

  // Problem Angle
  var problemAngle = (function() {
    var headers = ss.getSheetByName("posts")
                    .getRange(1, 1, 1, ss.getSheetByName("posts").getLastColumn())
                    .getValues()[0].map(function(h){ return String(h).trim(); });
    var idx = headers.indexOf("Problem Angle");
    if (idx === -1) return "";
    return row >= 2 ? String(sheet.getRange(row, idx + 1).getValue() || "").trim() : "";
  })();

  // Column EV — Core prompt header + problem angle + UK localisation + section governance
  var evContent =
    "STAGE 1A — DATA LOADING\n" +
    "ROLE: Senior UK SEO & Stone Restoration Strategist\n" +
    "ARTICLE TYPE: " + articleType + "\n\n" +
    "--- SYSTEM INSTRUCTION: HOLD ---\n" +
    "This is Stage 1A of the pipeline. Your only job is to ingest and retain all data below.\n" +
    "Do not begin any rewrite. Do not produce any HTML. Do not summarise or analyse.\n" +
    "Do not apply any rules — rules arrive in Stages 1B, 1C, and 1D.\n\n" +
    "Your entire response must be exactly: \"Stage 1A data received. Waiting for Stage 1B.\"\n" +
    "No other text. No analysis. No commentary. No image inventory.\n\n" +
    (problemAngle ? "--- PROBLEM ANGLE (Opening Sentence Governance) ---\n" +
    "This article addresses: " + sanitisePrimaryClusterPhrase(d["Primary Query Cluster Owned"] || "") + "\n" +
    "Approach: " + problemAngle.split("\n")[0] + "\n" +
    "OPENING SENTENCE LAW: The opening sentence of the article header and the first paragraph of Section 1\n" +
    "must directly reflect this problem angle. The reader must recognise their specific problem\n" +
    "from the very first sentence. Do not open with a general statement about the material.\n" +
    "--- END PROBLEM ANGLE ---\n\n" : "") +
    "--- UK LOCALISATION DATA ---\n" +
    "1. Use British English exclusively (e.g., Specialist, Metres, Taps, Splashback).\n" +
    "2. " + ukLocal + "\n\n" +
    "--- SECTION CONTENT GOVERNANCE ---\n" +
    "1. Primary Entity Focus: 80% of the word count must serve the PRIMARY ENTITY: " + primaryEntity + ".\n" +
    "2. Peripheral Summarisation: Sections that do not strengthen the Primary Entity are limited to\n" +
    "   maximum 300 words. Summarise principles and use the full URL from the Silo Map to link to the designated authority page.\n\n" +
    "--- INTERNAL LINKING DATA ---\n" +
    "1. Use the full URL from the Silo Map for all href values — no placeholders.\n" +
    "2. Integrate 5-8 relevant links from the Silo Map naturally.\n" +
    "3. Anchor text MUST match the intent of the target Role.";

  // Column FA (157) — Authority Brief
  var faContent = authorityBrief
    ? "--- AUTHORITY BRIEF (Pre-Pipeline Governance) ---\n" +
      "This brief was generated from entity analysis of the original page.\n" +
      "It governs the entire pipeline — all stages must comply with the constraints below.\n" +
      authorityBrief + "\n--- END AUTHORITY BRIEF ---"
    : "";

  // Column FB (158) — Silo Map
  var fbContent = "--- SILO CONTEXT (INTERNAL LINKING GOVERNANCE) ---\n" + siloMap;

  // Column FC (159) — TSM
  var fcContent = "--- TECHNICAL SPECIFICATION MATRIX (HARD-LOCK) ---\n" + techMatrix;

  // Column FD (160) — Semantic Requirements
  var fdContent =
    "--- SEMANTIC REQUIREMENTS ---\n" +
    "CORE SUPPORTING ENTITIES: " + supportingEntities + "\n" +
    "ANCHOR ENTITIES (CO-OCCURRENCE TARGETS): " + anchorEntities;

  // Column FE (161) — Governance Fields
  var feContent =
    "--- PAGE GOVERNANCE DATA ---\n" +
    "POST ID: " + postID + "\n" +
    "PRIMARY ENTITY: " + primaryEntity + "\n" +
    "ARTICLE TYPE: " + articleType + "\n" +
    "GOVERNANCE SUMMARY: " + govSummary + "\n" +
    "MATERIAL: " + (d["Stone Type"] || "") + "\n" +
    "TRANSFORMATION CONFIDENCE: " + confidence + "\n\n" +
    "--- GOVERNANCE FIELDS (from sheet — read-only context) ---\n" +
    "NEW H1: " + (d["New H1"] || "Not set") + "\n" +
    "NEW META TITLE: " + (d["New Meta Title"] || "Not set") + "\n" +
    "NEW META DESCRIPTION: " + (d["New Meta Description"] || "Not set") + "\n" +
    "SCHEMA (JSON-LD): " + (d["Schema (JSON-LD)"] || "Not set") + "\n" +
    "PRIMARY SEARCH TERM: " + (d["Primary Search Term"] || "Not set") + "\n" +
    "STRATEGIC REASONING: " + (d["Strategic Reasoning"] || "Not set") + "\n" +
    "TOPICAL INTENT: " + (d["Topical Intent"] || "Not set") + "\n" +
    "MATRIX ROLE: " + (d["Matrix Role"] || "Not set") + "\n" +
    "KEY DECISIONS EXPLAINED: " + (d["Key Decisions Explained"] || "Not set");

  // Column FF (162) — Material Behaviour
  var ffContent =
    "--- MATERIAL BEHAVIOUR CALIBRATION ---\n" + fragility + "\n\n" +
    "--- PROHIBITED LLM BEHAVIOURS FOR THIS MATERIAL ---\n" + prohibited;

  // Save all columns — each wrapped so we can identify which one is oversized
  var blocks = [
    { col: 152, label: 'EV (core prompt header)', content: evContent },
    { col: 153, label: 'EW (original HTML)', content: originalHTML },
    { col: 157, label: 'FA (Authority Brief)', content: faContent },
    { col: 158, label: 'FB (Silo Map)', content: fbContent },
    { col: 159, label: 'FC (TSM)', content: fcContent },
    { col: 160, label: 'FD (Semantic Requirements)', content: fdContent },
    { col: 161, label: 'FE (Governance Fields)', content: feContent },
    { col: 162, label: 'FF (Material Behaviour)', content: ffContent }
  ];

  for (var i = 0; i < blocks.length; i++) {
    var b = blocks[i];
    var len = String(b.content || '').length;
    if (len > 50000) {
      return { success: false, message: 'Column ' + b.label + ' is ' + len + ' characters — exceeds the 50,000 cell limit. Trim this block before saving.' };
    }
  }

  sheet.getRange(row, 152).setValue(evContent);  // EV
  sheet.getRange(row, 153).setValue(originalHTML); // EW
  sheet.getRange(row, 157).setValue(faContent);  // FA
  sheet.getRange(row, 158).setValue(fbContent);  // FB
  sheet.getRange(row, 159).setValue(fcContent);  // FC
  sheet.getRange(row, 160).setValue(fdContent);  // FD
  sheet.getRange(row, 161).setValue(feContent);  // FE
  sheet.getRange(row, 162).setValue(ffContent);  // FF

  return { success: true, message: 'W1A saved across columns EV, EW, FA–FF' };
}
function saveW1BOutput() {
  var prompt = buildStage1BPrompt();
  if (!prompt) {
    return { success: false, message: 'Failed to build W1B prompt' };
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('posts');
  var row = sheet.getActiveRange().getRow();
  if (row < 2) return { success: false, message: 'Select a data row first' };
  sheet.getRange(row, 154).setValue(prompt); // Column EX
  return { success: true, message: 'W1B saved to column EX' };
}
function saveW1COutput() {
  var prompt = buildStage1CPrompt();
  if (!prompt) {
    return { success: false, message: 'Failed to build W1C prompt' };
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('posts');
  var row = sheet.getActiveRange().getRow();
  if (row < 2) return { success: false, message: 'Select a data row first' };
  sheet.getRange(row, 155).setValue(prompt); // Column EY
  return { success: true, message: 'W1C saved to column EY' };
}
function saveW1DOutput() {
  var prompt = buildStage1DPrompt();
  if (!prompt) {
    return { success: false, message: 'Failed to build W1D prompt' };
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('posts');
  var row = sheet.getActiveRange().getRow();
  if (row < 2) return { success: false, message: 'Select a data row first' };
  sheet.getRange(row, 156).setValue(prompt); // Column EZ
  return { success: true, message: 'W1D saved to column EZ' };
}
function saveAllW1Outputs() {
  try {
    var r1a = saveW1AOutput();
    if (!r1a.success) return { success: false, message: 'W1A failed: ' + r1a.message };
    var r1b = saveW1BOutput();
    if (!r1b.success) return { success: false, message: 'W1B failed: ' + r1b.message };
    var r1c = saveW1COutput();
    if (!r1c.success) return { success: false, message: 'W1C failed: ' + r1c.message };
    var r1d = saveW1DOutput();
    if (!r1d.success) return { success: false, message: 'W1D failed: ' + r1d.message };
    return {
      success: true,
      message: 'W1A–W1D saved to EV–EZ, with W1A supporting data in FA–FF'
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}