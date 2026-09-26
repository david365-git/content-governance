/**
 * ================================================================================
 * ce_Stage2_ContentGeneration.gs - STAGE2 CONTENTGENERATION
 * ================================================================================
 * ce_Stage2_ContentGeneration
 * Content execution - HTML generation and governance fields
 * 
 * Part of Abbey Floor Care Content Pipeline v77+
 * Reorganised: March 2026
 * ================================================================================
 */

function getStage2AData() {
  const d = getActiveRowDataMap();

  const primaryCluster = String(d["Primary Query Cluster Owned"] || "").toLowerCase();
  const clusterLabel = primaryCluster.includes("darken") ||
                       primaryCluster.includes("dirty") ||
                       primaryCluster.includes("won't clean") ||
                       primaryCluster.includes("resist") ||
                       primaryCluster.includes("hold dirt") ? "CLUSTER B" :
                       primaryCluster.includes("care") ||
                       primaryCluster.includes("maintenance") ||
                       primaryCluster.includes("how to clean") ? "CLUSTER A" :
                       primaryCluster.includes("service") ||
                       primaryCluster.includes("specialist") ||
                       primaryCluster.includes("near me") ? "CLUSTER C" : "CLUSTER B";

  return `
STAGE 2A — PRE-WRITE REMINDER
ROLE: Senior UK SEO & Stone Restoration Strategist

============================================================
OUTPUT CONTRACT — HARD RULES (READ FIRST — NON-NEGOTIABLE)
============================================================
These rules override all learned patterns. Commit to them before reading further.

RULE A — NO H1 TAG:
Do NOT include an <h1> tag anywhere in the output.
The H1 is managed by WordPress outside this HTML body.
INCORRECT: <header><h1>Title</h1><p>Intro</p></header>
CORRECT (no video):   <header><p>Intro text only</p></header>
CORRECT (with video): <header><p>[intro]</p><figure class="video-embed">...</figure><p>[bridging sentence]</p></header>

RULE B — HEADER STRUCTURE:
The <header> element contains the intro paragraph as its first <p> tag.
If a Header Video is assigned in the section plan, the header also contains a video figure block and a bridging <p> tag after it.
If no Header Video is assigned, the header contains only the single intro <p> tag.
No heading tags of any kind. No nav blocks.

RULE C — BIO_PARAGRAPH MARKER IS MANDATORY:
You MUST output a BIO_PARAGRAPH: line as the second-to-last element, immediately before <footer>.
This line does not look like HTML. Output it anyway. It is required.
Format: BIO_PARAGRAPH: [2-3 sentences of plain text — no HTML tags]
A missing BIO_PARAGRAPH line means the output is incomplete and will be rejected.

REQUIRED OUTPUT ORDER:
<header><p>...</p></header>
<section id="section-1">...</section>
[all sections]
BIO_PARAGRAPH: [plain text paragraph]
<footer><p>...</p></footer>

If ANY of these contract rules are broken, the output is invalid.
============================================================

--- SYSTEM INSTRUCTION ---
The Stage 1.5B section plan has been reviewed and approved.
You are about to write the HTML in Stage 2B.
Before execution begins, confirm you will apply these seven rules to every section.
Your entire response must be exactly: "Seven rules confirmed. Ready to execute."
Nothing else.

--- SEVEN RULES TO APPLY DURING HTML GENERATION ---

RULE 1 — CLUSTER TONE:
This article is classified as: ${clusterLabel}
The cluster label above must hold across every paragraph of the header block.
CLUSTER B prohibits comfort language anywhere in the header: "reassuring", "reassurance",
"the good news is", "fortunately", "happily", "rest assured".
CLUSTER A prohibits mechanism-first or damage-first header sentences.
CLUSTER C prohibits sentences qualifying expertise or leading with homeowner uncertainty.

RULE 2 — DEFECT COMPLETION:
Name no defect without all three elements present before the next paragraph:
(1) What it is. (2) What the homeowner sees. (3) What is done about it.

RULE 3 — PROCESS SECTION DEPTH:
This rule applies ONLY to sections explicitly covering a process intervention in 
Educational Guides, Method Guides, and Diagnostic Guides — NOT Case Studies.

Process sections require three body paragraphs minimum before any figure or link.
Para 1: homeowner symptom. Para 2: mechanism and correction. Para 3: outcome.
Surface finish section must use the distinction framing specified in the Stone-Specific Prompt Rules block below.

EXEMPT FROM RULE 3 — do NOT apply the three-paragraph minimum to:
— Case Studies (governed by H3 narrative structure rules from Content Governance instead)
— Hub routing sections (sections whose purpose is directing readers to other pages)
— Comparison sections (sections contrasting two conditions or finish types)
— Diagnostic threshold sections (sections giving the reader a self-check)
— Any section with a word budget of 120-180 words — the three-paragraph minimum 
   cannot be satisfied within this budget alongside figures and links
   
For exempt sections: write the minimum prose needed to serve the reader's purpose
within the governed word budget. Do not pad to satisfy a paragraph count.

FOR CASE STUDIES SPECIFICALLY: Use H3 subheadings to structure narrative sections 
longer than 300 words as specified in Content Governance H3 Usage Rules.
H3s are structural markers, not paragraph replacements — they organize the narrative 
into digestible phases (Initial Condition → Process → Outcome) without counting 
toward paragraph minimums.

RULE 4 — HUB-INTRO (hub pages only):
One orienting sentence maximum before the navigation list.
No restatement of header content.

RULE 5 — FOOTER CTA:
Footer must end with a button anchor, direct instruction, or action sentence.
A footer ending on an explanation sentence has failed this rule.

RULE 6 — PARAGRAPH OPENING VARIETY:
The conditional symptom-anchor opening — "If your [material] shows...", "If your floor...", "If your [material] has..." — is permitted ONCE per section as the orientation sentence only.
All subsequent paragraphs within that section must open with varied professional prose: a declarative statement, a mechanism explanation, a cause-and-effect observation, or a direct practical statement.
A section where every paragraph opens with a conditional "If..." construction has failed this rule.
The reader has been oriented by the first sentence. Subsequent paragraphs must develop the explanation in natural prose, not repeat the orientation device.
CORRECT: One conditional opener followed by 2-3 paragraphs of authoritative explanation.
WRONG: Four consecutive paragraphs each opening with "If your [material]..."

RULE 7 — SECTION OPENER ROTATION (Four-Way Pivot):
The opening sentence of each section must rotate through these four styles — never use the same style for two consecutive sections:
  Style A — CONDITIONAL: "If your [material] shows [symptom]..." — diagnostic, reader recognition
  Style B — FACT-FIRST: Opens with the physical reality — "Dull patches that refuse to shift are almost never dirt..."
  Style C — EXPERT WARNING: Opens with consequence of wrong action — "Repeatedly scrubbing a patchy floor often makes it worse..."
  Style D — NOUN-FIRST: Leads with the technical entity — "Capillary action is the reason spills appear to grow..."
The HEADER must never open with a Conditional — use Fact-First, Expert Warning, or a direct authority statement.
Section 1 may use any style. Sections 2 onward must vary from the previous section.
Two consecutive sections opening with "If your [material]..." is a failure of this rule.
ROUTING SECTION EXEMPTION: For sections whose sole purpose is routing the reader to other pages (e.g. Internal Estate Routing, hub navigation sections), Style B (Fact-First) or Style C (Expert Warning) are the only appropriate choices. Style A (Conditional) is prohibited on routing sections as the reader is not diagnosing a symptom.
`.trim();
}

/* ============================================================
   STAGE 2B — EXECUTE
   HTML generation only. No TDF row. No column counting.
   The nine governance fields are generated separately:
     Stage 0  → 5 analytical fields (pushed from W0)
     Stage 2C → 4 fields: H1, Meta Title, Meta Description, Schema
============================================================ */

function buildTotalWordCountTargetBlock(articleType) {
  try {
    var ceiling = getTierWordCountCeiling(articleType);
    if (!ceiling) return "";

    var total = ceiling.total;
    var minTarget = ceiling.minTarget;
    var maxTarget = ceiling.maxTarget;

    return "------------------------------------------------------------\n" +
      "TOTAL ARTICLE LENGTH TARGET (Hard Lock — this is a CEILING, not a floor)\n" +
      "------------------------------------------------------------\n" +
      "Article Type: " + articleType + "\n" +
      "The governed section word budgets from the Tier Structural Coverage Matrix sum to " + total + " words for this article type.\n" +
      "TARGET TOTAL LENGTH: " + minTarget + "-" + maxTarget + " words.\n" +
      "Writing beyond " + maxTarget + " words is a failure of this rule.\n" +
      "BEFORE YOU FINALISE OUTPUT: count your draft's approximate total length. If it exceeds " + maxTarget + " words, do not simply stop or cut the final section short. Instead, revisit every section and identify content that can be tightened — a warning or condition restated in more than one section, an elaboration that repeats an earlier point, a paragraph that could state the same fact more directly. Shorten those areas, then re-total before producing your final output.\n" +
      "Do not add extra explanatory paragraphs, restate a warning or condition already covered elsewhere in the article, or expand a section beyond its governed word budget to create a feeling of thoroughness.\n" +
      "If competitor pages for this search term average far fewer words than this target, that is not licence to go below the Tier Matrix per-section floors — it is a signal that this niche rewards concision, so do not treat the total target above as something to fill.\n" +
      "------------------------------------------------------------";
  } catch(e) {
    return "";
  }
}

function getStage2BData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const d = getActiveRowDataMap();
  const articleTypeBlock = buildArticleTypeControlBlockGPT(ss);

  // Read W1A–W1D pre-assembled outputs from sheet columns EV–FF
  const sheet = ss.getSheetByName('posts');
  const row = sheet.getActiveRange().getRow();
  const w1aOutput  = String(sheet.getRange(row, 152).getValue() || '').trim(); // EV — core
  const w1aHtml    = String(sheet.getRange(row, 153).getValue() || '').trim(); // EW — HTML
  const w1bOutput  = String(sheet.getRange(row, 154).getValue() || '').trim(); // EX — W1B
  const w1cOutput  = String(sheet.getRange(row, 155).getValue() || '').trim(); // EY — W1C
  const w1dOutput  = String(sheet.getRange(row, 156).getValue() || '').trim(); // EZ — W1D
  const w1aAuthBrief  = String(sheet.getRange(row, 157).getValue() || '').trim(); // FA — Authority Brief
  const w1aSiloMap    = String(sheet.getRange(row, 158).getValue() || '').trim(); // FB — Silo Map
  const w1aTsm        = String(sheet.getRange(row, 159).getValue() || '').trim(); // FC — TSM
  const w1aSemantic   = String(sheet.getRange(row, 160).getValue() || '').trim(); // FD — Semantic
  const w1aGovFields  = String(sheet.getRange(row, 161).getValue() || '').trim(); // FE — Governance Fields
  const w1aMatBehav   = String(sheet.getRange(row, 162).getValue() || '').trim(); // FF — Material Behaviour

  // ═══════════════════════════════════════════════════════════
  // Content Governance Integration
  // ═══════════════════════════════════════════════════════════
  const articleType = String(d["Article Type"] || "General").trim();
  const govSheet = ss.getSheetByName("Content Governance");
  let governanceBlock = "";

  if (govSheet) {
    const govData = govSheet.getDataRange().getValues();
    let govRow = null;
    
    for (let i = 1; i < govData.length; i++) {
      if (String(govData[i][0]).trim() === articleType) {
        govRow = govData[i];
        break;
      }
    }
    
    if (govRow) {
  governanceBlock = `
============================================================
CONTENT GOVERNANCE LOCKS — READ FIRST (NON-NEGOTIABLE)
============================================================
ARTICLE TYPE: ${articleType}

VOICE RULE:
${govRow[1]}

INTENT BOUNDARY:
${govRow[2]}

STRUCTURAL LOCK:
${govRow[3]}

FORBIDDEN PATTERNS:
${govRow[4]}

IMPROVEMENT DEFINITION:
${govRow[5]}

H2 PATTERN RULES:
${govRow[6]}

H3 USAGE RULES:
${govRow[7]}
============================================================

`;
}
  }
  // ═══════════════════════════════════════════════════════════

  // Extract image URLs directly from the original HTML for inline injection into Stage 2B
  // Reads site-export HTML for the active post — same source as Stage 1A
  var imageUrls = [];
  var videoUrls = [];
  try {
    var ss2b      = SpreadsheetApp.getActiveSpreadsheet();
    var postSheet = ss2b.getSheetByName("posts");
    var postRow   = postSheet.getActiveRange().getRow();
    var postId2b  = String(postSheet.getRange(postRow, 1).getValue() || "").trim();
    // Read Post ID by header name
    var hdr2b     = postSheet.getRange(1, 1, 1, postSheet.getLastColumn()).getValues()[0]
                             .map(function(h){ return String(h).trim(); });
    var pidIdx    = hdr2b.indexOf("Post ID");
    if (pidIdx > -1) postId2b = String(postSheet.getRange(postRow, pidIdx + 1).getValue() || "").trim();

    // Get original HTML from site-export
    var exportSh  = ss2b.getSheetByName("site-export");
    var rawHtml2b = "";
    if (exportSh && postId2b) {
      var expData = exportSh.getDataRange().getValues();
      var expHdr  = expData[0];
      var expIdX  = expHdr.indexOf("ID");
      var expHtX  = expHdr.indexOf("Full Post HTML");
      if (expIdX > -1 && expHtX > -1) {
        for (var ei = 1; ei < expData.length; ei++) {
          if (String(expData[ei][expIdX]).trim() === postId2b) {
            rawHtml2b = String(expData[ei][expHtX] || "");
            break;
          }
        }
      }
    }

    // Extract image src URLs — full size only (skip dimension-suffixed variants)
    // Also skip decorative patterns
    var IMAGE_EXT  = /\.(jpg|jpeg|png|webp|gif|avif)(\?[^"']*)?$/i;
    var DIM_SUFFIX = /\-\d+x\d+(@\d+x)?\.(jpg|jpeg|png|webp|gif|avif)$/i;
    var SKIP_PAT   = /logo|icon|avatar|sprite|pixel|tracking/i;
    var seenUrls   = [];

    // Match src= attributes
    var srcRe = /src=["']([^"']+)["']/gi;
    var sm;
    while ((sm = srcRe.exec(rawHtml2b)) !== null) {
      var url = sm[1].trim();
      if (!IMAGE_EXT.test(url)) continue;
      if (DIM_SUFFIX.test(url)) continue;
      if (SKIP_PAT.test(url)) continue;
      if (seenUrls.indexOf(url) > -1) continue;
      seenUrls.push(url);
      imageUrls.push(url);
    }

    // Extract video iframes separately
    var iframeRe = /src=["']([^"']*(?:youtube|vimeo|wistia|loom|embed)[^"']*)["']/gi;
    var vm;
    while ((vm = iframeRe.exec(rawHtml2b)) !== null) {
      var vurl = vm[1].trim();
      if (seenUrls.indexOf(vurl) > -1) continue;
      seenUrls.push(vurl);
      videoUrls.push(vurl);
      // Remove from imageUrls if accidentally captured
      imageUrls = imageUrls.filter(function(u) { return u !== vurl; });
    }
  } catch(imgEx) {
    // Extraction failed — proceed without list
  }

  // Build figBlock from extracted URLs
  var figBlock = "";
  var suggestionBlock = "";

  // VIDEO placement instruction
  if (videoUrls.length > 0) {
    figBlock += "------------------------------------------------------------\n" +
                "VIDEO PRESERVATION (Hard Lock):\n" +
                "The following video embed(s) existed in the original HTML and MUST appear in your output.\n" +
                "Use EXACTLY this iframe HTML for each video — copy character for character:\n\n";
    videoUrls.forEach(function(vurl, vi) {
      figBlock += (vi + 1) + ". <p><center><iframe loading=\"lazy\" src=\"" + vurl + "\" width=\"560\" height=\"315\" frameborder=\"0\" allowfullscreen=\"allowfullscreen\"></iframe></center></p>\n";
    });
    figBlock += "\nVIDEO PLACEMENT RULES:\n";
    if (videoUrls.length >= 2) {
      figBlock += "- Video 1: Place INSIDE the <header> block using the header video format specified in Section 3A.\n";
      figBlock += "- Video 2: Place at the END of Section 4, after all paragraphs in that section, before the closing </section> tag.\n";
    } else {
      figBlock += "- Place the video INSIDE the <header> block using the header video format specified in Section 3A.\n";
    }
    figBlock += "- Do NOT place any video inside a content section, footer, or bio box.\n" +
                "- A missing video is a CRITICAL audit failure.\n" +
                "------------------------------------------------------------\n";
  }

  // IMAGE URL list — inline injection
  // Image placement is now specified in the section plan (from W1.5C/W1.5D)
  // Global image list removed to prevent conflicting instructions
  var imagePlacementBlock = "IMAGE PLACEMENT:\n" +
             "Each section below specifies which image(s) to place in that section.\n" +
             "Insert each image exactly where specified in the section plan.\n" +
             "Wrap each image in: <figure class=\"wp-caption aligncenter\" style=\"width: 700px\"><img src=\"[url]\" alt=\"[descriptive alt]\" /><figcaption class=\"wp-caption-text\">[diagnostic caption max 15 words]</figcaption></figure>\n" +
             "------------------------------------------------------------";

  // Prepend image placement to figBlock — preserve any video preservation block already built
  figBlock = imagePlacementBlock + (figBlock ? "\n" + figBlock : "");

  const stage2AContent = getStage2AData();

  return {
    part1: `STAGE 2B — EXECUTE — READ ALL CONTEXT BELOW BEFORE OUTPUTTING
NOTE: This is a single complete prompt. Read everything before writing your first token.
Your first output token must be <header>.
OUTPUT: HTML body content only.
Do not output a tab-delimited row. Do not output column data. Do not number columns.
Your entire output must be the revised HTML body — nothing before it, nothing after it.

${w1aOutput ? '--- STAGE 1A CORE ---\n' + w1aOutput + '\n\n' : ''}${w1aAuthBrief ? w1aAuthBrief + '\n\n' : ''}${w1aGovFields ? w1aGovFields + '\n\n' : ''}${w1aMatBehav ? w1aMatBehav + '\n\n' : ''}${w1aTsm ? w1aTsm + '\n\n' : ''}${w1aSemantic ? w1aSemantic + '\n\n' : ''}${w1aSiloMap ? w1aSiloMap + '\n\n' : ''}${w1aHtml ? '--- ORIGINAL HTML ---\n' + w1aHtml + '\n\n' : ''}${w1bOutput ? '--- STAGE 1B STRUCTURAL RULES ---\n' + w1bOutput + '\n\n' : ''}${w1cOutput ? '--- STAGE 1C PARAGRAPH RULES ---\n' + w1cOutput + '\n\n' : ''}${w1dOutput ? '--- STAGE 1D OUTPUT QUALITY RULES ---\n' + w1dOutput + '\n\n' : ''}--- STAGE 2A PRE-WRITE RULES ---
${stage2AContent}

${governanceBlock}
------------------------------------------------------------
READER-FIRST EDITORIAL PRINCIPLE (Hard Lock — applies to every paragraph)
------------------------------------------------------------
The reader is a busy homeowner with a specific problem and limited time.
Every sentence must earn its place by answering one of two questions:
  1. Is this helping me understand or solve my problem?
  2. Should I keep reading or have I found what I need?

PARAGRAPH STRUCTURE LAW:
  First paragraph of every section: MUST open with the problem or symptom the reader is experiencing.
    The reader has just scanned to a new H2 heading and is deciding whether this section is relevant.
    The first sentence must answer "is this my problem?" before introducing any mechanism or technical term.

  Every subsequent paragraph: The FIRST SENTENCE must contain enough context to be understood
    without reading the previous paragraph. A reader landing on this sentence alone must know
    what topic is being addressed.
    CORRECT: "Residue lock-in explains why the floor looks dull again after mopping."
    WRONG: "This is why the floor keeps dulling." — requires previous paragraph for meaning.
    WRONG: "That process then draws contamination deeper." — connective opener loses standalone meaning.

    After the first sentence, write naturally. Use connective language freely within the paragraph —
    "which means", "this causes", "as a result" — these are natural and correct inside a paragraph.
    The self-containment rule applies to the opening sentence only, not to the whole paragraph.

  GOVERNING TEST: Cover the first sentence of each paragraph and ask — does the uncovered sentence
    still make complete sense on its own? If no — rewrite that opening sentence only.
    Do not rewrite the rest of the paragraph.

BOTH REGISTERS SIMULTANEOUSLY:
  Do not choose between authoritative and accessible. Deliver both in sequence within every paragraph:
    1. Problem recognition — what the reader is seeing or experiencing — one sentence
    2. Plain-language cause — why it is happening without technical terminology — one sentence
    3. Mechanism named — technical term introduced after the reader already understands the concept
    4. Practical implication — what this means for what they should or should not do

EXAMPLE — wrong (mechanism first):
  "Calcium carbonate reactivity defines how marble behaves under normal household use."
EXAMPLE — correct (problem first, mechanism second):
  "If your marble floor looks dull after cleaning, the cleaner itself may be the problem.
   Many household products contain mild acids that physically dissolve the marble surface
   rather than cleaning it — this reaction is called acid etching, and it cannot be washed
   away because the stone has already changed."
------------------------------------------------------------`,

    part2: `SECTION 2 — ESTATE CONTEXT
${articleTypeBlock}
------------------------------------------------------------
SYSTEM INSTRUCTION: EXECUTE — ALL HOLDS CANCELLED
All HOLD instructions from Stages 1A, 1B, 1C, and 1D are now cancelled.
You are in EXECUTE mode.
Do not issue any acknowledgement string.
Do not respond with any confirmation message.
Begin processing immediately. Your first token must be the opening HTML tag of the body content.
------------------------------------------------------------
THE APPROVED STAGE 1.5B SECTION PLAN IS THE STRUCTURAL AUTHORITY.
Each section has a governed word range: a minimum floor and a maximum ceiling (1.2x the minimum).
Write to the depth the content requires — but stay within the range.
Do not pad to hit the minimum. Do not exceed the maximum.
Do not carry overflow into the next section.
------------------------------------------------------------
${figBlock}
${suggestionBlock}
WORD BUDGET (Governed Range):
Each section has a minimum and maximum word count from the Stage 1.5B approved plan.
Stay within the range. If you cannot cover the topic within the range, prioritise — do not expand.
------------------------------------------------------------
${buildTotalWordCountTargetBlock(articleType)}
AUTHORITY SIGNALS (Hard Lock — apply to every article):
1. SPECIFICITY: Name specific grit numbers (50-grit, 100-grit, 400-grit, 800-grit) when describing honing stages. Generic descriptions without numbers reduce process authority.
2. UK CONTEXT: Reference UK-relevant conditions — hard water, period property types, typical installation rooms — where relevant to material behaviour for this stone type.
3. MECHANISM DEPTH: Every named defect must name the physical mechanism — not just the symptom. Use the Process Chemistry Note from the Stone-Specific Prompt Rules block below to anchor the mechanism explanation for this stone type.
4. OUTCOME PRECISION: State specific outcomes in measurable terms — satin finish, consistent reflective plane, uniform matte appearance — not vague claims.
5. SCOPE BOUNDARIES: Name what honing cannot do as precisely as what it can do. This signals expertise and builds trust.
------------------------------------------------------------
${buildStonePromptRulesBlock(d["Stone Type"] || "stone")}
${buildSemanticFingerprintBlock(d["Stone Type"] || "stone")}
${buildRecoveryBlueprintInjectionBlock(bc_getRecoveryBlueprintEntities())}
${buildSerpValidationChecklistBlock(bc_getSerpValidationEntities())}
${buildTfidfVocabularyBlock(bc_getTfidfTerms())}
${buildRecallGapBlock(bc_getRecallGapTerms())}
${buildCompetitorVocabBlock(bc_getCompetitorVocabGap(), bc_getCompetitorWordCount())}
DIAGNOSTIC IMAGE CAPTIONS (Hard Lock):
Every image caption must function as a diagnostic tool — not a description of what is shown.
The reader is scanning to identify whether the image matches their problem.
A caption that only names or describes the image fails. A caption that connects the image to the reader's situation passes.

DIAGNOSTIC CAPTION PATTERNS (use any of these — vary across the article):
  - Reader recognition: "If your floor looks like this..."
  - Named condition: "This is [condition] — [one sentence on what it means for the reader]"
  - Observation led: "Dark patches like these indicate..."
  - Stage based: "Floors at this stage need..."
  - Pattern recognition: "Hallways showing this pattern have..."

WRONG — purely descriptive (fails):
  "Acid etching leaves dull patches where the marble surface has been chemically damaged."
  "Old sealers and adhesives being removed from an original tile floor."
  "Coating removal in progress on the Minton tile floor."

CORRECT — diagnostic (passes):
  "If your floor looks like this after cleaning, the problem is etching — not dirt."
  "This is coating removal — softened residue must be extracted, not spread around."
  "Floors at this stage need controlled extraction before sealing can begin."

VARIATION RULE: No two captions in the same article may open with the same pattern.
Vary the diagnostic approach across the four or five images — do not repeat the same opener.
------------------------------------------------------------
OUTCOME STATEMENTS (Hard Lock — honest full range):
Professional restoration outcomes must be stated accurately — not defensively.
  CORRECT range: "better than new in many cases" through to "improved where structural conditions limit full correction"
  Do NOT default to "aims for a consistent visual plane rather than unrealistic perfection" — this undersells accurate outcomes.
  ALWAYS state: the floor will look significantly better than before intervention.
  ALWAYS state: the floor will be easier to clean and maintain after professional restoration.
  Limitation statements are reserved for specific named conditions ONLY:
    — Lippage exceeding 2mm without prior grinding
    — Deep iron oxidation staining
    — Structural fissures affecting surface integrity
  Do NOT apply limitation statements as general disclaimers on every outcome.
------------------------------------------------------------
UNIVERSAL BENEFIT STATEMENTS (Hard Lock — every article regardless of type):
Every article must contain ALL THREE of the following benefit statements,
placed at the most relevant point in the content — not gathered at the end:
  1. APPEARANCE: The floor will look significantly better — and in many cases better than when first installed.
     Where grinding, honing, and polishing are the primary intervention for this material and article type, reference them as the mechanism. Where they are not applicable or are prohibited for this material — omit them and state the appearance benefit in terms of the correct intervention instead.
  2. MAINTAINABILITY: A professionally restored and correctly sealed floor is significantly easier
     to clean and maintain than a worn or incorrectly treated floor.
  3. LONGEVITY: Correct ongoing maintenance — pH-neutral cleaning, grit removal before wet mopping,
     resealing at the right interval — is the single most important factor in extending the floor's life.
     This must be stated in the article and a contextual internal link to the maintenance article included.
------------------------------------------------------------
MAINTENANCE EDUCATION ELEMENT (Hard Lock — every article regardless of type):
Every article must include at minimum:
  — One sentence on what correct ongoing maintenance achieves for this specific stone type
  — One sentence naming one specific thing to avoid and why
  — A contextual internal link to the maintenance or cleaning tips article at the point where this appears
This is not a standalone maintenance section — it is integrated into the most relevant existing section.
------------------------------------------------------------
PROFESSIONAL DECISION REASONING (Hard Lock — every article regardless of type):
Whenever the article states that a specific material, product, method, technique,
or sealer type was chosen or used for this job — rather than an available
alternative — it must state the reason for that choice, grounded in the floor's
actual condition, the finish type, or the environment. A choice reported as a
bare fact, with no stated reason, has failed this rule.
  WRONG: "A surface sealer was used for this floor."
  CORRECT: "A surface sealer was chosen for this floor because the repairs
  needed a film that would bind the new filler to the surrounding stone,
  rather than the invisible, no-film protection an impregnating sealer gives."
This applies to any decision point named in the article — sealer type, filler
type, honing grit sequence, cleaning chemistry, extraction method, or repair
technique. The reason does not need its own sentence — it can be folded into
the same sentence that reports the choice, provided the causal link is explicit.
------------------------------------------------------------
SUMMARY BOX RENDERING (Tier 4 only):
If Stage 1.5B contains a SUMMARY_BOX, render it immediately before the opening paragraph:
<div class="abbey-summary-box">
<p class="abbey-summary-box__title">At a Glance</p>
<ul class="abbey-summary-box__list">
<li><strong>Problem:</strong> [Problem line]</li>
<li><strong>Solution:</strong> [Solution line]</li>
<li><strong>Result:</strong> [Result line]</li>
</ul>
</div>
If no SUMMARY_BOX present — omit entirely.
------------------------------------------------------------

============================================================
HARD STOP — DO NOT OUTPUT ANY HTML
============================================================
You have received context only. No execution instruction has been given yet.
Your only permitted response to this message is exactly:
"Context received. Waiting for section plan."
Any HTML output at this point will be discarded.
Wait for the next message before writing anything.
============================================================`,

    recoveryBlueprintEntities: bc_getRecoveryBlueprintEntities(),
    serpValidationEntities:    bc_getSerpValidationEntities(),
    tfidfTerms:                bc_getTfidfTerms(),
    recallGapTerms:            bc_getRecallGapTerms(),
    part3: buildStage2BPart3A(d) + '\n\n' + buildStage2BPart3B(d)
  };
}

function buildTfidfVocabularyBlock(terms) {
  if (!terms || terms.length === 0) return "";

  var block = "";
  block += "------------------------------------------------------------\n";
  block += "TF-IDF VOCABULARY TARGETS (Hard Lock — Distribute Naturally)\n";
  block += "------------------------------------------------------------\n";
  block += "The following terms are the dominant vocabulary Google associates with this page.\n";
  block += "They must appear naturally distributed across the article — not clustered in one section.\n";
  block += "Do not force terms into sentences where they do not belong naturally.\n\n";
  block += "RANKED VOCABULARY TARGETS:\n";

  terms.forEach(function(t, i) {
    block += (i + 1) + ". " + t.term + " (weight: " + t.score.toFixed(3) + ")\n";
  });

  block += "\nDISTRIBUTION RULE:\n";
  block += "- Top 3 terms: must appear in at least 3 separate sections\n";
  block += "- Terms 4-8: must appear in at least 2 separate sections\n";
  block += "- Remaining terms: at least one natural appearance across the article\n";
  block += "------------------------------------------------------------\n";
  return block;
}

function buildRecallGapBlock(gapTerms) {
  if (!gapTerms || gapTerms.length === 0) return "";

  var block = "";
  block += "------------------------------------------------------------\n";
  block += "RECALL GAP TERMS (Hard Lock — Missing Primary Query Cluster Vocabulary)\n";
  block += "------------------------------------------------------------\n";
  block += "The following terms appear in the primary query cluster for this page\n";
  block += "but are currently absent or underrepresented in the existing article.\n";
  block += "These are coverage gaps that must be addressed in the new article.\n\n";
  block += "GAP TERMS TO INCLUDE:\n";
  block += gapTerms.join(", ") + "\n\n";
  block += "COVERAGE RULE:\n";
  block += "Each gap term above must appear at least once in the article body.\n";
  block += "Integrate naturally — do not force into sentences where they do not belong.\n";
  block += "Priority: place gap terms in sections where the topic is most directly relevant.\n";
  block += "------------------------------------------------------------\n";
  return block;
}

function buildSerpValidationChecklistBlock(entities) {
  if (!entities || entities.length === 0) return "";

  var block = "";
  block += "------------------------------------------------------------\n";
  block += "SERP VALIDATION COVERAGE CHECKLIST (Hard Lock)\n";
  block += "------------------------------------------------------------\n";
  block += "The following entities are confirmed present on competitor pages in the live SERP.\n";
  block += "Every entity in this list MUST appear somewhere in your output before the response is complete.\n";
  block += "A missing SERP-validated entity is a coverage failure.\n\n";
  block += "CONFIRMED SERP ENTITIES:\n";
  block += entities.join(", ") + "\n\n";
  block += "COVERAGE RULE:\n";
  block += "Before outputting, verify each entity above appears at least once in your HTML.\n";
  block += "If any are absent — add them naturally to the most relevant section before completing.\n";
  block += "------------------------------------------------------------\n";
  return block;
}

function buildRecoveryBlueprintInjectionBlock(entities) {
  if (!entities) return "";

  var block = "";
  block += "------------------------------------------------------------\n";
  block += "RECOVERY BLUEPRINT ENTITY RANKING (Hard Lock — Priority Order)\n";
  block += "------------------------------------------------------------\n";
  block += "Entity coverage is governed by tier. Apply the rules below to every section.\n\n";

  if (entities.priority && entities.priority.length > 0) {
    block += "PRIORITY ENTITIES (must appear — anchor co-occurrence targets):\n";
    block += entities.priority.join(", ") + "\n";
    block += "Rule: Each priority entity MUST appear explicitly by name at least once.\n";
    block += "Surround each with its full semantic fingerprint vocabulary as specified above.\n";
    block += "PROXIMITY RULE: When the entity name appears, at least 3 of its fingerprint terms MUST appear within 2-3 sentences of the entity name.\n";
    block += "Do not place fingerprint terms in a different paragraph from the entity name.\n";
    block += "Distribute fingerprint terms naturally in the sentences immediately surrounding the entity name — not clustered after it.\n\n";
    block += "CAUSAL CHAIN RULE (Hard Lock — applies to all Priority Entities):\n";
    block += "When two or more Priority Entities appear in the same section, you MUST state the mechanical relationship between them.\n";
    block += "Do not treat entities as isolated topics. Show how one entity creates the condition for another to act.\n";
    block += "REQUIRED pattern: '[Entity A] allows/causes/enables [mechanism] which produces [Entity B symptom].'\n";
    block += "CORRECT: 'The open pore network in unglazed clay allows moisture to move upward through the tile body, carrying dissolved salts that crystallise at the surface as efflorescence.'\n";
    block += "WRONG: 'Unglazed clay porosity is a key characteristic. Efflorescence is a common problem on these floors.'\n";
    block += "The wrong version names both entities but gives no causal connection — an embedding model cannot map the relationship.\n\n";
  }

  if (entities.supporting && entities.supporting.length > 0) {
    block += "SUPPORTING ENTITIES (use where contextually relevant):\n";
    block += entities.supporting.join(", ") + "\n";
    block += "Rule: Include where the section topic naturally supports it. Do not force.\n";
    block += "PROXIMITY RULE: When a supporting entity name appears, at least 2 of its fingerprint terms MUST appear within 2-3 sentences of the entity name.\n\n";
  }

  if (entities.peripheral && entities.peripheral.length > 0) {
    block += "PERIPHERAL ENTITIES (one appearance sufficient):\n";
    block += entities.peripheral.join(", ") + "\n";
    block += "Rule: One natural mention across the entire article is sufficient.\n\n";
  }

  if (entities.gscGap && entities.gscGap.length > 0) {
    block += "GSC GAP ENTITIES (high impression, low rank — reinforce these):\n";
    block += entities.gscGap.join(", ") + "\n";
    block += "Rule: These terms have strong search demand but weak current ranking.\n";
    block += "Distribute naturally across at least two sections.\n\n";
  }

  block += "------------------------------------------------------------\n";
  return block;
}

function buildStage2BPart3A(d) {
  // PART 3A: Context and rules - paste first, wait for READY
  var st = String(d["Stone Type"] || "stone").trim();
  var stLower = st.toLowerCase();
  var articleType = String(d["Article Type"] || "General").trim();
  
  return "---\n\n" +
    "SECTION 3A — PAGE DATA & CONTEXT\n" +
    "Post ID: " + d["Post ID"] + "\n" +
    "Material: " + d["Stone Type"] + "\n" +
    "Article Type: " + d["Article Type"] + "\n\n" +
    "Read the following context and formatting rules carefully.\n" +
    "When you have finished reading, reply with exactly: READY\n" +
    "Do not begin writing yet. Just confirm you are ready.\n\n" +
    "------------------------------------------------------------\n" +
    "HEADER FORMAT (Hard Lock):\n" +
    "The <header> element must contain the intro paragraph and, if a header video is assigned, a video figure block.\n" +
    "Do NOT add an <h1> tag inside the header. The H1 is set in WordPress — never in the body HTML.\n" +
    "If a Header Video is assigned in the section plan, place it inside <header> using this EXACT format:\n" +
    "<figure class=\"video-embed\" style=\"margin:24px 0;text-align:center\"><iframe width=\"560\" height=\"315\" src=\"[video url]\" title=\"[article title]\" frameborder=\"0\" allow=\"accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture\" allowfullscreen></iframe><figcaption>[one sentence describing the video]</figcaption></figure>\n" +
    "Follow the figure immediately with a short bridging <p> tag introducing the article below.\n" +
    "CORRECT: <header><p>[intro]</p><figure class=\"video-embed\">...</figure><p>[bridging sentence]</p></header>\n" +
    "WRONG: <header><p>[intro]</p></header> — missing the header video that was assigned.\n" +
    "Do NOT add a <nav>, <ul>, quick-links block, or any navigation list inside <header>.\n" +
    (articleType === "Hub Page"
      ? "For Hub Page articles, quick-links navigation is mandatory but must appear in a separate structural block immediately after </header> using:\n" +
        "<section id=\"hub-intro\">\n" +
        "<p>[one orienting sentence only]</p>\n" +
        "<nav>\n" +
        "[quick-links list]\n" +
        "</nav>\n" +
        "</section>\n" +
        "NAV LINK COUNT (Hard Lock):\n" +
        "The quick-links nav list MUST include a link to EVERY numbered section in the approved section plan.\n" +
        "Do NOT select a subset of sections — all sections must be represented.\n" +
        "A nav list that omits any section is incomplete and will fail audit.\n" +
        "Count the sections in the approved plan below and verify your nav list matches before outputting.\n"
      : "This page is a " + articleType + " and must not include a nav block.\n") +
    "CORRECT: <header><p>[intro sentence]</p></header>\n" +
    "WRONG: <header><h1>[title]</h1><p>[intro]</p></header>\n" +
    "\n" +
    "CRITICAL: The header is NOT Section 1.\n" +
    (articleType === "Hub Page"
      ? "On Hub Pages, structure must be: <header>...</header> then <section id=\"hub-intro\">...</section> then <section id=\"section-1\"><h2>...</h2>...</section>\n"
      : "Section 1 begins AFTER the header and MUST have its own <section> tag and H2 heading.\n" +
        "Structure: <header>...</header> then <section id=\"section-1\"><h2>...</h2>...</section>\n") +
    "------------------------------------------------------------\n" +
    (String(d["Feeds Hub"] || "").trim() && 
 String(d["Feeds Hub"]).trim() !== String(d["URL"] || "").trim()
  ? "HUB LINK (Hard Lock — mandatory):\n" +
    "This article MUST contain at least one link to the hub page.\n" +
    "Hub URL: " + String(d["Feeds Hub"]).trim() + "\n" +
    "Use the full hub URL above directly in your href — do not use placeholders.\n" +
    "Place the hub link naturally within the body content — not in the footer.\n" +
    "A missing hub link is a Check 2 audit failure that will block push to sheet.\n" +
    "------------------------------------------------------------\n"
  : "") +
    "INTERNAL LINK POSITIONING (Hard Lock — Contextual Placement):\n" +
    "Internal links must appear at the point where the reader's need for the linked content is highest.\n" +
    "This is the moment the reader has recognised their specific problem and needs to know more.\n" +
    "Do NOT place internal links as standalone signposting sentences at the end of sections.\n" +
    "Place links WITHIN substantive paragraphs where the linked topic is being discussed.\n" +
    "CORRECT: Link embedded mid-paragraph where the concept is explained and reader need peaks.\n" +
    "WRONG: 'Further detail on this topic is available in [link].' — standalone signpost at section end.\n" +
    "------------------------------------------------------------\n" +
    "INTERNAL LINK COUNT REQUIREMENT (Hard Lock):\n" +
    "This article must contain 5-8 internal links from the silo map provided in Stage 1A.\n" +
    "This requirement applies regardless of the lateral link count set in Stage 1.5D.\n" +
    "Lateral links and internal links are distinct categories:\n" +
    "— Lateral links are silo-to-silo links assigned by Stage 1.5D.\n" +
    "— Internal links are contextual links embedded during generation from the Stage 1A silo map.\n" +
    "A lateral link count of zero does NOT mean zero internal links.\n" +
    "The section plan showing 'Internal link: None' reflects lateral link assignment only.\n" +
    "You MUST still embed 5-8 contextual internal links from the silo map during generation.\n" +
    "A missing internal link count is a Check 2 audit failure that will block push to sheet.\n" +
    "------------------------------------------------------------\n" +
    "ANCHOR TEXT (Hard Lock — No Bare URLs):\n" +
    "Every internal link MUST use descriptive anchor text — never a bare URL.\n" +
    "CORRECT: <a href=\"https://www.abbeyfloorcare.co.uk/home-garden/marble-care-cleaning-repair-and-restoration-explained/\">marble care, cleaning, repair and restoration explained</a>\n" +
    "WRONG: <a href=\"https://www.abbeyfloorcare.co.uk/home-garden/marble-care-cleaning-repair-and-restoration-explained/\">https://www.abbeyfloorcare.co.uk/home-garden/marble-care-cleaning-repair-and-restoration-explained/</a>\n" +
    "WRONG: <a href=\"https://www.abbeyfloorcare.co.uk/home-garden/marble-care-cleaning-repair-and-restoration-explained/\">click here</a>\n" +
    "WRONG: <a href=\"https://www.abbeyfloorcare.co.uk/home-garden/marble-care-cleaning-repair-and-restoration-explained/\">this article</a>\n" +
    "Anchor text must describe the content of the target page in plain language — typically 3–8 words.\n" +
    "A bare URL used as anchor text is a critical output error.\n" +
    "------------------------------------------------------------\n" +
    "SIGNPOSTING FORMAT (Hard Lock):\n" +
    "Every <p> containing an internal link MUST contain a minimum of 2 sentences. The link sentence must never be the only sentence.\n" +
    "REQUIRED PATTERN: <p>[Substantive sentence about the topic.] [Second substantive sentence.] [Link sentence.]</p>\n" +
    "COUNT CHECK: Before writing any <p> with a link, count the sentences. If count is 1 — add a substantive sentence. No exceptions.\n" +
    "A <p> containing only a link sentence is a critical output error.\n" +
    "------------------------------------------------------------\n" +
    (function() {
      var st = String(d["Stone Type"] || "stone").trim();
      var stLower = st.toLowerCase();
      var varietyMap = {
        "marble":      "polished, honed and tumbled marble",
        "travertine":  "honed and filled, tumbled, tumbled and filled, and polished travertine",
        "limestone":   "polished, honed and brushed limestone",
        "slate":       "riven, honed and tumbled slate",
        "terrazzo":    "poured and tile terrazzo",
        "victorian":   "geometric, encaustic and plain Victorian tiles",
        "sandstone":   "riven, sawn and tumbled sandstone",
        "quartzite":   "polished, honed and brushed quartzite",
        "porcelain":   "polished, matt and textured porcelain",
        "terracotta":  "sealed, waxed and unsealed terracotta"
      };
      var row = d["Post ID"] ? parseInt(String(d["Post ID"])) % 4 : 0;
      var varieties = varietyMap[stLower] || "a wide range of " + stLower + " types";
      var sentences = [
        "Abbey Floor Care has worked with " + varieties + " in homes, hotels and heritage buildings across the country for nearly 30 years.",
        "With nearly 30 years of experience across the country, Abbey Floor Care has restored " + varieties + " in residential, commercial and heritage settings.",
        "Abbey Floor Care brings nearly 30 years of practical experience with " + varieties + ", working across homes, hotels and period buildings throughout the UK.",
        "Having restored " + varieties + " for nearly 30 years in homes, hotels and heritage buildings nationwide, Abbey Floor Care understands how each stone behaves under real conditions."
      ];
      var eeatSentence = sentences[row] || sentences[0];
      var articleTypeCheck = String(d["Article Type"] || "").trim();
      if (articleTypeCheck !== "Service Page" && articleTypeCheck !== "Geo Service Page") {
        return "";
      }
      return "EEAT EXPERIENCE SIGNAL (Service Page and Geo Service Page only — Hard Lock):\n" +
        "The provider selection section MUST include one sentence establishing Abbey Floor Care experience.\n" +
        "Use this exact sentence for this article:\n" +
        eeatSentence + "\n" +
        "Place it as the opening sentence of the provider selection section. Do not omit it.\n" +
        "------------------------------------------------------------\n";
    })() +
    ([" Service Page","Geo Service Page","Case Study"].indexOf(d["Article Type"]) > -1 && String(d["Location Context"] || "").trim() ?
      "LOCATION CONTEXT (Hard Lock — " + d["Article Type"] + "):\n" +
      "The background below describes the general property types and housing stock found in this locality.\n" +
      "It was generated before this specific project's defects and mechanism were known — treat it as raw\n" +
      "material, not as text to insert unchanged.\n\n" +
      "LOCATION BACKGROUND (raw material — do not insert verbatim):\n" +
      String(d["Location Context"]).trim() + "\n\n" +
      "SYNTHESIS RULE (Hard Lock):\n" +
      "You MUST include one substantive paragraph, in either the opening section or the provider/\n" +
      "contractor selection section, that connects this locality's housing stock or layout (drawn\n" +
      "from the background above) directly to the SPECIFIC defect or mechanism found on THIS project\n" +
      "— the crack, void, filler failure, or wear pattern this article actually documents.\n" +
      "The paragraph must state a causal link between the property style/layout and the actual damage\n" +
      "found — not a general description of the area's architecture on its own.\n" +
      "  WRONG (generic housing history, no causal link to this project):\n" +
      "  \"New Malden has a broad mix of Edwardian houses, 1930s semi-detached and detached homes...\"\n" +
      "  CORRECT (locality detail tied to this project's actual defect):\n" +
      "  \"In New Malden's extended homes, travertine often bridges the original subfloor and a newer\n" +
      "  rear-extension base — exactly the kind of movement that opened the long crack on this floor.\"\n" +
      "Do not simply drop the LOCATION BACKGROUND text in unchanged. Rewrite it so it earns its place\n" +
      "by explaining something about THIS project, not the area in general.\n" +
      "------------------------------------------------------------\n"
    : "") +
    "DAVID ALLEN IMAGE HARD LOCK:\n" +
    "David_Allen.jpg must NEVER appear anywhere in your HTML output — not in sections, not in figures, not standalone.\n" +
    "The bio box is built by the script from a marker you provide — do NOT output any HTML for it.\n" +
    "------------------------------------------------------------\n" +
    "AUTHOR BIO MARKER (Hard Lock — required in every article):\n" +
    "After the last </section> and before <footer>, output EXACTLY this marker on its own line:\n" +
    "BIO_PARAGRAPH: [Your 2-3 sentence bio paragraph here]\n\n" +
    "Rules for the bio paragraph text:\n" +
    "- 2-3 sentences, specific to THIS article topic, stone type and article type\n" +
    "- Reference David Allen, 30+ years experience, Abbey Floor Care\n" +
    "- Case Study: reference the specific location and what was corrected\n" +
    "- Method Guide: reference the specific process covered\n" +
    "- Diagnostic Guide: reference the specific condition diagnosed\n" +
    "- Buyer Guide: reference the evaluation or selection focus\n" +
    "- Plain text only — no HTML tags inside the paragraph\n" +
    "- The script will wrap it in the correct bio box HTML automatically\n" +
    "EXAMPLE OUTPUT (do not copy — write fresh for this article):\n" +
    "BIO_PARAGRAPH: David Allen has restored marble floors across the UK for over 30 years, including this Stoke-on-Trent project where surface wear and uneven tiles were corrected through full mechanical refinement. His work focuses on accurate diagnosis and the correct restoration sequence to return clarity and performance.\n" +
    "------------------------------------------------------------\n" +
    "FOOTER FORMAT (Hard Lock — use exactly this structure):\n" +
    "<footer>\n" +
    "<p>[One sentence outcome statement.] <a href=\"/contact\">Contact us to arrange a no-obligation " + stLower + " floor assessment.</a></p>\n" +
    "</footer>\n" +
    "------------------------------------------------------------\n\n" +
    "CRITICAL: Do not begin writing until next input received.\n" +
    "Reply READY when you have read and understood these rules.";
    
}

function buildStage2BPart3B(d) {
  // PART 3B: Execution with H2 lock at the very top
  var sectionPlan = getSectionPlan();
  var hubUrl = String(d["Feeds Hub"] || "").trim();
  var headerVideo = "";
  var headerVideoMatch = sectionPlan ? sectionPlan.match(/Header Video:\s*(https?:\/\/[^\s\]]+)/i) : null;
  if (headerVideoMatch) headerVideo = headerVideoMatch[1].trim();
  var sectionBlock = "";
  var articleType = String(d["Article Type"] || "General").trim();
  
  // Amazon product block / COMM-flag logic retired — Section 4's TSM role
  // ("Maintenance Handover & Escalation Link") now owns maintenance advice
  // and reserved product-recommendation space for every Case Study,
  // regardless of COMM flag. This block previously produced a competing,
  // duplicate "Products Used In This Guide" / "Tailored Maintenance
  // Handover" section that fought with Section 4's own governed content.
  var amazonBlock = "";
  
  // Read Google optimization JSON from Column CY (103)
  var optimizationData = null;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('posts');
    var row = sheet.getActiveRange().getRow();
    var optimizationJson = sheet.getRange(row, 103).getValue();
    
    if (optimizationJson && optimizationJson.trim()) {
      // Strip markdown code fences if present
      var cleanedJson = optimizationJson.trim()
        .replace(/^```json\s*/i, '')
        .replace(/\s*```$/, '');
      optimizationData = JSON.parse(cleanedJson);
      Logger.log('W2B: Loaded ' + optimizationData.sections.length + ' optimization sections from Column CY');
    }
  } catch (optErr) {
    Logger.log('W2B: Could not read optimization JSON from Column CY: ' + optErr.message);
    // Continue without optimization data
  }
  if (sectionPlan) {
    var lines = sectionPlan.split(/\n/);
    var sections = [];
    var current = null;
    lines.forEach(function(line) {
      var secMatch = line.match(/^SECTION\s+(\d+):\s*(?:Heading\s+(H\d):\s*)?(.+)/i);
      if (secMatch) {
        if (current) sections.push(current);
        current = {
          title: secMatch[3].trim(),
          headingLevel: secMatch[2] ? secMatch[2].toUpperCase() : "H2",
          id: "",
          budget: "", 
          brief: "", 
          visualPattern: "",
          images: "",
          entities: "",
          internalLink: ""
        };
      } else if (current) {
        // Extract all fields from section plan
        var idMatch = line.match(/id="?([^"\s]+)"?/i) || line.match(/id:\s*([^\s]+)/i);
        if (idMatch) current.id = idMatch[1].trim();
        
        var budgetMatch = line.match(/Word Budget:\s*(\d+)/i) || line.match(/(\d+)\s*words?/i);
        if (budgetMatch) current.budget = budgetMatch[1];
        
        var briefMatch = line.match(/Content Brief:\s*(.+)/i);
        if (briefMatch) current.brief = briefMatch[1].trim();

        var tsmMatch = line.match(/TSM Requirement:\s*(.+)/i);
        if (tsmMatch) current.tsmRequirement = tsmMatch[1].trim();
        
        var visualMatch = line.match(/Visual Pattern:\s*(.+)/i);
        if (visualMatch) current.visualPattern = visualMatch[1].trim();
        
        var imagesMatch = line.match(/Images:\s*(.+)/i);
        if (imagesMatch) current.images = imagesMatch[1].trim();
        
        var videosMatch = line.match(/Videos:\s*(.+)/i);
        if (videosMatch) current.videos = videosMatch[1].trim();
        
        var entitiesMatch = line.match(/Entities to include:\s*(.+)/i);
        if (entitiesMatch) current.entities = entitiesMatch[1].trim();
        
        var linkMatch = line.match(/Internal link:\s*(.+)/i);
        if (linkMatch) current.internalLink = linkMatch[1].trim();
      }
    });
    if (current) sections.push(current);

    // Merge optimization data from Column CY JSON into sections
    if (optimizationData && optimizationData.sections) {
      sections.forEach(function(s, idx) {
        var sectionNumber = idx + 1;
        var opt = optimizationData.sections.find(function(o) {
          return o.section_number === sectionNumber;
        });
        if (opt) {
          s.googleOpt =
            "Opening anchor: " + (opt.opening_anchor || "") + " " +
            "Information gain directive: " + (opt.information_gain_directive || "") + " " +
            "Scannability pattern: " + (opt.scannability_pattern || "") + " " +
            "Featured snippet target: " + (opt.featured_snippet_target || "") + " " +
            "List opportunity: " + (opt.list_opportunity || "") + " " +
            "Blockquote flag: " + (opt.blockquote_flag || "");
        }
      });
    }

    if (sections.length > 0) {
      sectionBlock = "============================================================\n" +
        "H2 HEADING LOCK — CRITICAL INSTRUCTION (READ THIS TWICE)\n" +
        "============================================================\n" +
        "The H2 headings below are FINAL and APPROVED.\n" +
        "You MUST copy each H2 heading EXACTLY as written — character for character.\n" +
        "Do NOT rephrase, improve, simplify, shorten, or rewrite any H2.\n" +
        "Do NOT add words, remove words, or change word order.\n" +
        "Do NOT 'fix' grammar, punctuation, or capitalization.\n" +
        "Even if an H2 seems unclear or poorly worded — USE IT EXACTLY AS WRITTEN.\n\n" +
        "A rewritten H2 is a CRITICAL AUDIT FAILURE that will block publication.\n" +
        "If you change even ONE WORD of an H2, the entire output will be rejected.\n\n" +
        "CORRECT: Copy the H2 text verbatim from the section list below.\n" +
        "WRONG: Rewording 'Marble stains that remain after cleaning' to 'Why marble stains persist after cleaning'.\n" +
        "============================================================\n\n" +
        "APPROVED SECTION PLAN (Hard Lock — produce ONLY these sections):\n" +
        "Any section not listed below is prohibited. Do not add FAQ, CTA, or unlisted sections.\n\n";
      
      sections.forEach(function(s, i) {
        var tsmReq = String(s.tsmRequirement || "").toLowerCase();
        var wordBudgetNum = parseInt(s.budget) || 0;
        var processTerms = /honing|polishing|sealing|stripping|re-grouting|recolouring|burnishing/;
        var isProcess = processTerms.test(tsmReq) && wordBudgetNum > 180;
        
        sectionBlock += "SECTION " + (i + 1) + ":\n";
        sectionBlock += "  HEADING TAG: <" + (s.headingLevel || "h2").toLowerCase() + "> (use this exact tag level — not h3, not h1)\n";
        sectionBlock += "  H2 (COPY EXACTLY — DO NOT CHANGE): " + s.title + "\n";
        sectionBlock += "  id: \"" + (s.id || "section-" + (i+1)) + "\"\n";
        
        if (s.budget) {
          const minBudget = s.budget;
          const maxBudget = Math.round(s.budget * 1.5);
          sectionBlock += "  WORD BUDGET: " + minBudget + "–" + maxBudget + " words. Stay within range.\n";
        }
        
        if (s.brief) {
          sectionBlock += "  CONTENT BRIEF: " + s.brief + "\n";
        }
        
        if (i === 0) {
          sectionBlock += "  SECTION 1 AUTHORITY DISTRIBUTION RULE (Hard Lock):\n" +
                          "  Section 1 covers entry condition and problem identification ONLY.\n" +
                          "  Do NOT include technical mechanism explanation in Section 1.\n" +
                          "  Do NOT include process rationale or professional judgement in Section 1.\n" +
                          "  Do NOT include sympathetic restoration rationale in Section 1.\n" +
                          "  Technical mechanism explanation belongs in Section 2.\n" +
                          "  Professional judgement and process rationale belongs in Section 3.\n" +
                          "  A Section 1 that front-loads all technical authority is a governance failure.\n" +
                          "  Write Section 1 to the word budget only — do not expand beyond it.\n";
        }
        
        if (s.visualPattern && s.visualPattern !== "None") {
          sectionBlock += "  VISUAL PATTERN: " + s.visualPattern + " — Apply Rule 25 pattern as specified above.\n";
        }
        
        if (s.entities && s.entities !== "None") {
          sectionBlock += "  ENTITIES TO INCLUDE: " + s.entities + "\n" +
                          "  CRITICAL: Write entity names in LOWERCASE in body text (e.g., 'capillary action', not 'Capillary Action').\n" +
                          "  Only capitalize if it's a proper noun or at the start of a sentence.\n" +
                          "  ENTITY TRANSLATION RULE (Hard Lock): Internal pipeline entity labels must NEVER appear verbatim in body text.\n" +
                          "  Translate each entity into plain descriptive language that expresses the concept naturally.\n" +
                          "  WRONG: 'That routine protects heritage pigment protection by reducing abrasion.'\n" +
                          "  CORRECT: 'That routine protects the tile's original colour and surface character by reducing abrasion.'\n" +
                          "  WRONG: 'Honest outcome framing matters here because the fix is controlled cleaning.'\n" +
                          "  CORRECT: 'Setting realistic expectations matters here because the fix is controlled cleaning.'\n" +
                          "  The entity name tells you WHAT concept to include — not the exact words to use.\n" +
                          "  NOTE ON THE CONTENT BRIEF ABOVE: if this section's Content Brief happens to reference this entity\n" +
                          "  in Title Case or as a named proper-noun process, that phrasing is planning shorthand only —\n" +
                          "  it is NOT approved article text and must NOT be copied into the output. Apply the Entity\n" +
                          "  Translation Rule above regardless of how the entity appears in the Content Brief.\n";
        }
        
        if (s.images && s.images !== "None") {
          var imgUrls = s.images.split(",").map(function(u) { return u.trim(); }).filter(function(u) { return u.length > 0; });
          if (imgUrls.length === 1) {
            var imgUrl = imgUrls[0];
            if (!imgUrl.startsWith('http')) {
              imgUrl = 'https://www.abbeyfloorcare.co.uk/wp-content/uploads/' + imgUrl;
            }
            sectionBlock += "  IMAGES (CRITICAL): You MUST insert the image in this section.\n" +
                            "  Exact URL: " + imgUrl + "\n" +
                            "  Wrap in: <figure class=\"wp-caption aligncenter\" style=\"width: 700px\"><img src=\"" + imgUrl + "\" alt=\"[diagnostic alt text]\" /><figcaption class=\"wp-caption-text\">[diagnostic caption]</figcaption></figure>\n" +
                            "  Copy the URL EXACTLY as shown above. Do not modify or shorten it.\n" +
                            "  A missing or incorrect URL is an AUDIT FAILURE.\n";
          } else {
            sectionBlock += "  IMAGES (CRITICAL): You MUST insert ALL " + imgUrls.length + " images in this section as SEPARATE figure blocks.\n" +
                            "  Each image MUST have its own individual figure block — do NOT merge URLs into one src attribute.\n" +
                            "  WRONG: <img src=\"url1, url2, url3\" />\n" +
                            "  CORRECT: Three separate <figure> blocks each with one src URL.\n" +
                            "  Each figure block must follow this format:\n" +
                            "  <figure class=\"wp-caption aligncenter\" style=\"width: 700px\"><img src=\"[single URL]\" alt=\"[diagnostic alt text]\" /><figcaption class=\"wp-caption-text\">[diagnostic caption]</figcaption></figure>\n\n";
            imgUrls.forEach(function(imgUrl, idx) {
              if (!imgUrl.startsWith('http')) {
                imgUrl = 'https://www.abbeyfloorcare.co.uk/wp-content/uploads/' + imgUrl;
              }
              sectionBlock += "  IMAGE " + (idx + 1) + " Exact URL: " + imgUrl + "\n";
            });
            sectionBlock += "  Place images after substantive paragraphs. No two figures may appear consecutively.\n" +
                            "  A merged src URL is a CRITICAL AUDIT FAILURE.\n";
          }
        }
        
        if (s.internalLink && s.internalLink !== "None") {
        sectionBlock +=
          "  INTERNAL LINK — HARD LOCK:\n" +
          "  You MUST include this exact governed URL in this section:\n" +
          "  " + s.internalLink + "\n" +
          "  Use it as the href value of a normal contextual <a> link inside a paragraph.\n" +
          "  Copy the URL EXACTLY — do not shorten it, alter it, replace it, or omit it.\n" +
          "  A missing governed internal URL is a CRITICAL AUDIT FAILURE.\n";
      }
        
        if (s.videos && s.videos !== "None") {
          sectionBlock += "  VIDEO (CRITICAL): You MUST embed this video in this section.\n" +
                          "  Exact iframe src: " + s.videos + "\n" +
                          "  Use this EXACT format:\n" +
                          "  <p><center><iframe loading=\"lazy\" src=\"" + s.videos + "\" width=\"560\" height=\"315\" frameborder=\"0\" allowfullscreen=\"allowfullscreen\"></iframe></center></p>\n" +
                          "  Copy the src URL EXACTLY as shown. A missing or incorrect video is an AUDIT FAILURE.\n";
        }
        
        if (s.googleOpt && s.googleOpt.trim()) {
          sectionBlock += "  GOOGLE OPTIMIZATION: " + s.googleOpt.trim() + "\n";
        }
        
        if (isProcess) {
          sectionBlock += "  PROCESS SECTION: Minimum 3 paragraphs before any <figure>. Para 1: symptom. Para 2: mechanism. Para 3: outcome.\n";
        }
        
        sectionBlock += "\n";
      });
      
      sectionBlock += "============================================================\n" +
        "H2 VERIFICATION CHECKPOINT — BEFORE YOU OUTPUT\n" +
        "============================================================\n" +
        "Before outputting your HTML, verify EVERY H2 heading:\n" +
        "1. Open your output HTML in your internal buffer\n" +
        "2. Check each <h2> tag against the H2 list above\n" +
        "3. If ANY H2 differs by even one character — STOP and rewrite it\n" +
        "4. Only output when ALL H2s match EXACTLY\n\n" +
        "This is your final quality gate. A mismatched H2 fails audit.\n" +
        "============================================================\n";
    }
  } else {
    sectionBlock = "SECTION PLAN: Not stored — write to approved Stage 1.5B plan only.\n" +
      "------------------------------------------------------------\n";
  }

  return "ALL CONTEXT LOADED — EXECUTE NOW. Output HTML only starting with <header>.\n\n" +
    "SECTION 3B — EXECUTE NOW\n" +
    "You confirmed READY in your previous message.\n" +
    "Now write the HTML body content following the section plan below.\n\n" +
    amazonBlock +
    (headerVideo ? "------------------------------------------------------------\n" +
    "HEADER VIDEO (Hard Lock — CRITICAL):\n" +
    "The section plan assigns a video to the header. You MUST embed it inside <header> using this EXACT format:\n" +
    "<figure class=\"video-embed\" style=\"margin:24px 0;text-align:center\"><iframe width=\"560\" height=\"315\" src=\"" + headerVideo + "\" title=\"[article title]\" frameborder=\"0\" allow=\"accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture\" allowfullscreen></iframe><figcaption>[one sentence describing the video]</figcaption></figure>\n" +
    "Place it after the intro <p> tag and before the closing </header> tag.\n" +
    "Follow it immediately with a short bridging <p> tag introducing the article below.\n" +
    "A missing header video is a CRITICAL audit failure.\n" +
    "------------------------------------------------------------\n\n" : "") +
    sectionBlock + "\n" +
    "------------------------------------------------------------\n" +
    "VISUAL RHYTHM RULES (Rule 25 — Wall-of-Text Prevention):\n" +
    "To prevent cognitive fatigue across mechanism-heavy sections, apply these patterns where specified in the Section Plan:\n\n" +
    "1. SYMPTOM-ANCHOR OPENINGS\n" +
    "   When Section Plan specifies 'Symptom-Anchor' pattern:\n" +
    "   Open the FIRST paragraph only with: <strong>If your [material] shows [symptom]</strong>, [explanation].\n" +
    "   All subsequent paragraphs in the same section must open in plain text without bold.\n" +
    "   This counts as ONE bold instance toward the page total of 4.\n" +
    "   If applying Symptom-Anchor would push the page bold total above 4 — use plain text instead.\n" +
    "   Example: <strong>If your marble shows a dark patch after a spill</strong>, the liquid has soaked into the stone.\n\n" +
    "2. DIAGNOSTIC-SEQUENCE\n" +
    "   When Section Plan specifies 'Diagnostic-Sequence' pattern:\n" +
    "   Use numbered <ol> list for self-tests and procedural checks.\n" +
    "   This is the ONLY permitted list type in mechanism content.\n" +
    "   Example:\n" +
    "   <p>After cleaning and drying, test whether staining is below the surface:</p>\n" +
    "   <ol>\n" +
    "   <li>Place water droplets on three areas</li>\n" +
    "   <li>Wait 5 minutes</li>\n" +
    "   <li>If stone darkens: moisture is penetrating</li>\n" +
    "   </ol>\n\n" +
    "3. COMPARISON-PARAGRAPHS\n" +
    "   When Section Plan specifies 'Comparison-Paragraphs' pattern:\n" +
    "   Use parallel paragraph structure with condition names in <strong> tags.\n" +
    "   Example:\n" +
    "   <p>If your marble shows [symptom A], this is <strong>staining</strong> — [mechanism].</p>\n" +
    "   <p>If your marble shows [symptom B], this is <strong>etching</strong> — [mechanism].</p>\n\n" +
    "4. MECHANISM-BLOCKQUOTE\n" +
    "   When Section Plan specifies 'Mechanism-Blockquote' pattern:\n" +
    "   Extract critical mechanism principle into <blockquote> visual island.\n" +
    "   Example:\n" +
    "   <p>Poultice extraction reverses the absorption process:</p>\n" +
    "   <blockquote>As the paste dries, it pulls contamination upward through capillary action.</blockquote>\n\n" +
    "5. H3-SUBHEADINGS\n" +
    "   When Section Plan specifies 'H3-Subheadings' pattern:\n" +
    "   Use <h3> subheadings to chunk sections covering 3+ distinct entities.\n" +
    "   Each H3 section opens with symptom-anchor paragraph.\n" +
    "   Example:\n" +
    "   <h2>Why different stains behave differently</h2>\n" +
    "   <h3>Oil-based contamination</h3>\n" +
    "   <p><strong>If your marble has dark patches near cooking areas</strong>, [mechanism].</p>\n\n" +
    "APPLY ONLY PATTERNS SPECIFIED IN SECTION PLAN.\n" +
      "If Visual Pattern field is 'None' — write pure prose with no special formatting.\n" +
      "Maximum 3-4 visual patterns per article.\n" +
      "Do NOT use bullet lists, tables, or comparison boxes.\n" +
      "------------------------------------------------------------\n" +
      "IMAGE CONTEXT RULE (Hard Lock):\n" +
      "Every <figure> block must be preceded by at least one substantive paragraph in the same section.\n" +
      "No two <figure> blocks may appear consecutively — a substantive paragraph must separate them.\n" +
      "A substantive paragraph is a minimum of 2 sentences explaining the condition, stage, or process shown in the image.\n" +
      "WRONG: Two <figure> blocks with no paragraph between them.\n" +
      "WRONG: A <figure> block as the first element in a section with no preceding paragraph.\n" +
      "CORRECT: One or more paragraphs → <figure> → one or more paragraphs → <figure>\n" +
      "------------------------------------------------------------\n" +
    "PHRASE VARIATION RULE:\n" +
    "Avoid repeating the same outcome phrase more than twice in the entire article.\n" +
    "Vary language when describing results:\n" +
    "  - 'will look significantly better' → use once or twice maximum\n" +
    "  - Then vary: 'regains clarity', 'restores appearance', 'improves dramatically', 'returns to original condition'\n" +
    "  - 'easier to maintain' → use once, then vary: 'simpler to care for', 'requires less effort', 'stays cleaner longer'\n" +
    "Apply the same principle to any recurring phrase about outcomes or process.\n" +
    "Natural variation prevents template-like repetition.\n" +
    "------------------------------------------------------------\n" +
    "OUTPUT REMINDER:\n" +
    "Output the revised HTML body only.\n" +
    "No tab-delimited columns. No row structure. No column numbers.\n" +
    "Start with the first HTML tag. End with the last HTML tag. Nothing else.\n" +
    "Your first token must be <header>\n" +
    "\n" +
    "FINAL OUTPUT SEQUENCE (Hard Lock — follow this order exactly):\n" +
"1. <header>...</header>\n" +
(articleType === "Hub Page"
  ? "2. <section id=\"hub-intro\">...</section>\n" +
    "3. <section id=\"section-1\">...</section>\n" +
    "4. ... all sections ...\n" +
    "5. <section id=\"section-N\">...</section> (last section)\n" +
    "6. BIO_PARAGRAPH: [your 2-3 sentence bio text — plain text only, no HTML tags]\n" +
    "7. <footer>...</footer>\n"
  : "2. <section id=\"section-1\">...</section>\n" +
    "3. ... all sections ...\n" +
    "4. <section id=\"section-N\">...</section> (last section)\n" +
    "5. BIO_PARAGRAPH: [your 2-3 sentence bio text — plain text only, no HTML tags]\n" +
    "6. <footer>...</footer>\n") +
    "Missing the BIO_PARAGRAPH line is a critical output failure.\n" +
    "------------------------------------------------------------\n" +
  "PRE-OUTPUT VALIDATION CHECK (MANDATORY — run this before outputting):\n" +
    "Before writing your first token, verify:\n" +
    "  1. No <h1> tag exists anywhere in your output\n" +
    "  2. <header> contains the intro <p> tag and, if a Header Video was assigned, a video figure block and bridging <p> — no heading tags\n" +
    (articleType === "Hub Page"
      ? "  3. <section id=\"hub-intro\"> exists immediately after <header>\n" +
        "  4. BIO_PARAGRAPH: line exists immediately before <footer>\n" +
        "  5. BIO_PARAGRAPH: contains plain text only — no HTML tags\n"
      : "  3. BIO_PARAGRAPH: line exists immediately before <footer>\n" +
        "  4. BIO_PARAGRAPH: contains plain text only — no HTML tags\n") +
    "If ANY check fails — correct it internally before outputting.\n" +
        "Do NOT explain. Do NOT apologise. Only output the final corrected HTML.";
}

/* ============================================================
   GOVERNANCE RELATED HUB DATA
   Renamed copy of getRelatedHubData() — avoids conflict with
   existing function in other .gs files.
============================================================ */


function getGovernanceRelatedHubData(currentStone, currentRow) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("posts");
  const data  = sheet.getDataRange().getValues();

  // Article types that use #service fragment — all others use #article
  const serviceFragmentTypes = ["Service Page", "Geo Service Page"];

  let hubLinks = [];
  for (let i = 1; i < data.length; i++) {
    if (i + 1 === currentRow) continue;
    if (data[i][6] === currentStone) {
      const url         = data[i][2];
      const title       = data[i][79];
      const articleType = String(data[i][7] || "").trim();
      if (url && title) {
        const fragment = serviceFragmentTypes.indexOf(articleType) > -1 ? "#service" : "#article";
        hubLinks.push({
          url:         url,
          title:       title,
          articleType: articleType,
          fragment:    fragment
        });
      }
    }
  }
  return JSON.stringify(hubLinks.slice(0, 7), null, 2);
}

function getGovernancePrompt() {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const sheet      = ss.getSheetByName("posts");
  const activeRow  = sheet.getActiveRange().getRow();

  const articleType = sheet.getRange(activeRow, 8).getValue();
  const pageUrl     = sheet.getRange(activeRow, 3).getValue();
  const pageHtml    = sheet.getRange(activeRow, 98).getValue();
  const stoneType   = sheet.getRange(activeRow, 7).getValue();
  const primaryTerm = sheet.getRange(activeRow, 93).getValue();
  const featuredImg = sheet.getRange(activeRow, 85).getValue();
  const logoUrl     = sheet.getRange(activeRow, 86).getValue();
  const feedsHub    = sheet.getRange(activeRow, 15).getValue();  // col O — Feeds Hub

  if (!primaryTerm || !pageHtml) {
    return "ERROR: Missing Primary Search Term (col 93) or HTML (col 98) for this row.";
  }

  // Read schema alignment rules for this article type
  const savSheet = ss.getSheetByName("Schema Alignment Validation Sheet");
  let schemaRulesBlock = "Schema Alignment Validation Sheet not found — use Article as primary @type by default.";
  let relationshipBlock = "";

  if (savSheet) {
    const savData    = savSheet.getDataRange().getValues();
    const savHeaders = savData[0].map(function(h) { return String(h).trim(); });
    const artTypeIdx = savHeaders.indexOf("Article Type");
    let matchRow = null;
    for (let i = 1; i < savData.length; i++) {
      if (String(savData[i][artTypeIdx]).trim().toLowerCase() === String(articleType).trim().toLowerCase()) {
        matchRow = savData[i]; break;
      }
    }
    if (matchRow) {
      const get = function(col) {
        const idx = savHeaders.indexOf(col);
        return idx > -1 ? String(matchRow[idx] || "").trim() : "Not specified";
      };
      const allowedPrimary    = get("Allowed Primary Schema Type");
      const allowedSecondary  = get("Allowed Secondary Schema Types");
      const prohibited        = get("Prohibited Schema Types");
      const areaServed        = get("Must Include Area Served");
      const expansion         = get("Schema Expansion Allowed");
      const primaryRel        = get("Primary Relationship Property");
      const secondaryRel      = get("Secondary Relationship Property");

      schemaRulesBlock =
        "SCHEMA TYPE RULES FOR THIS ARTICLE TYPE (" + articleType + ") — HARD LOCK:\n" +
        "- Allowed Primary @type: " + allowedPrimary + "\n" +
        "- Allowed Secondary @type: " + (allowedSecondary === "None" ? "None — do not add secondary types" : allowedSecondary) + "\n" +
        "- Prohibited @types — DO NOT USE: " + prohibited + "\n" +
        "- Schema type change allowed: No\n" +
        "- Schema expansion allowed: " + expansion + "\n" +
        (areaServed === "Yes"
          ? '- Must include areaServed: "Surrey and South East England"\n'
          : "- Do NOT include areaServed\n") +
        "\nYour primary @type MUST be: " + allowedPrimary + ". Any other primary type is a validation failure.";

      // Build relationship instructions based on article type
      // Hub pages declare spoke pages as children via hasPart
      // Spoke pages declare their parent hub via isPartOf
      // Service pages link to supporting content via isRelatedTo
      if (primaryRel && primaryRel !== "Not specified") {
        relationshipBlock = "SCHEMA RELATIONSHIP RULES — HARD LOCK:\n";

        if (primaryRel === "hasPart") {
          // Hub Page — list related pages as WebPage children with correct fragment per article type
          relationshipBlock +=
            "- This is a HUB PAGE. Use hasPart to declare spoke pages as children of this Article.\n" +
            "- Each spoke page in RELATED HUB ENTITIES includes an 'articleType' and 'fragment' field.\n" +
            "- Use the 'fragment' field to construct the @id for each spoke entry — do not guess or default to #article.\n" +
            "- Each spoke entry must follow this exact pattern:\n" +
            '  { "@type": "WebPage", "@id": "[url][fragment]", "url": "[url]", "name": "[title]" }\n' +
            "- Example: if fragment is #service → @id is https://www.abbeyfloorcare.co.uk/page/#service\n" +
            "- Example: if fragment is #article → @id is https://www.abbeyfloorcare.co.uk/page/#article\n" +
            "- Do NOT use mentions or isRelatedTo for hub-spoke relationships on this page type.\n";
        } else if (primaryRel === "isPartOf") {
          // Spoke pages — declare parent hub
          const hubId = feedsHub
            ? feedsHub.replace(/\/$/, "") + "/#article"
            : "Hub URL not set — omit isPartOf";
          relationshipBlock +=
            "- This page is a SPOKE PAGE that belongs to a hub. Use isPartOf to declare the parent hub.\n" +
            "- Add isPartOf to the primary Article entity:\n" +
            '  "isPartOf": { "@type": "Article", "@id": "' + hubId + '" }\n' +
            "- HARD RULE: DO NOT add hasPart to this schema. hasPart is only used on Hub Pages.\n" +
            "  This page is a spoke — adding hasPart here is a structural error.\n" +
            (secondaryRel
              ? "- Use " + secondaryRel + " to reference related pages or the topic this page covers.\n"
              : "");
        } else if (primaryRel === "isRelatedTo") {
          // Service / Geo Service pages
          relationshipBlock +=
            "- Use isRelatedTo on the Service entity to link to supporting content pages.\n" +
            "- Each entry in RELATED HUB ENTITIES should appear as an isRelatedTo reference:\n" +
            '  { "@id": "[url]#service" }\n' +
            "- Use provider to link Service to the LocalBusiness entity via @id.\n";
        }
      }

    } else {
      schemaRulesBlock = "No match found for article type '" + articleType + "' — use Article as primary @type by default.";
    }
  }

  const hubData = getGovernanceRelatedHubData(stoneType, activeRow);

  return (
    "Act as a Stone Restoration SEO Specialist and Schema Architect.\n" +
    "Generate H1, Meta Title, Meta Description, and a VALIDATED JSON-LD Graph.\n\n" +
    "PAGE DATA:\n" +
    "- URL: " + pageUrl + "\n" +
    "- Article Type: " + articleType + "\n" +
    "- Stone: " + stoneType + "\n" +
    "- Primary Keyword: " + primaryTerm + "\n" +
    "- Featured Image: " + featuredImg + "\n" +
    "- Publisher Logo: " + logoUrl + "\n\n" +
    "════════════════════════════════════════\n" +
    schemaRulesBlock + "\n" +
    "════════════════════════════════════════\n\n" +
    (relationshipBlock ? relationshipBlock + "\n" : "") +
    "STRICT SCHEMA FORMAT RULES (Stop Google Errors):\n" +
    "1. USE FLAT GRAPH: All schema objects as separate entries in the \"@graph\" array.\n" +
    "2. PRIMARY ENTITY must use mainEntityOfPage linking to the current URL.\n" +
    "3. If Article is primary @type — link to LocalBusiness via publisher field, not provider.\n" +
    "4. If Service is primary @type:\n" +
    "   - DO NOT use 'about' or 'hasPart' inside Service.\n" +
    "   - Use 'provider' to link to the '@id' of the LocalBusiness.\n" +
    "5. If HowTo is permitted as secondary:\n" +
    "   - Use '@type': 'HowToSupply' for all supplies (NEVER use 'Product' inside 'supply').\n" +
    "   - Link HowTo to the page via 'mainEntityOfPage'.\n" +
    "6. Always include BreadcrumbList.\n" +
    "7. JSON must be minified — single line, no internal line breaks.\n" +
    "8. All URLs must start with https://www.abbeyfloorcare.co.uk\n" +
    "9. No HTML-encoded characters (%22, %3A etc). No bracket-wrapped URLs.\n" +
    "10. PUBLISHER ANCHOR — HARD RULE: publisher @id must always be exactly\n" +
    "    \"https://www.abbeyfloorcare.co.uk/#localbusiness\".\n" +
    "    Do NOT use #organization or any other @id string. This site\'s established\n" +
    "    publisher entity is LocalBusiness. Introducing a different @id breaks\n" +
    "    entity consistency across the schema graph.\n" +
    "11. AUTHOR — HARD RULE: author must always be\n" +
    "    {\"@type\": \"Person\", \"name\": \"David Allen\"}.\n" +
    "    Do NOT use Organization as author type.\n" +
    "12. Do NOT add a \'url\' field directly on the Article entity.\n" +
    "    The page URL is declared via mainEntityOfPage. A duplicate url field\n" +
    "    on Article is redundant and creates conflicting signals.\n\n" +
    "RELATED HUB ENTITIES (Select 5 most relevant):\n" +
    hubData + "\n\n" +
    "PAGE CONTENT (HTML):\n" +
    '"""' + pageHtml.substring(0, 3200) + '"""\n\n' +
    "Return exactly:\n" +
    "New H1: [Result]\n" +
    "New Meta Title: [Result]\n" +
    "New Meta Description: [Result]\n" +
    "Schema:\n" +
    '[OUTPUT EVERYTHING INSIDE A SINGLE CODE WINDOW]\n' +
    '<script type="application/ld+json">\n' +
    "[VALIDATED JSON-LD GRAPH]\n" +
    "</script>"
  );
}

/* ============================================================
   H1 + META PROMPT GENERATOR  (replaces Stage 2C)
   Generates H1, Meta Title, Meta Description only.
   Strict character length rules — no schema, no ambiguity.
============================================================ */
function filterAuthorityBriefForMeta(rawBrief) {
  if (!rawBrief) return "";

  var text = rawBrief;

  // Sections to strip entirely — each runs from its header to the next
  // "---...---" divider or the next ALL-CAPS section header, whichever comes first.
  var sectionsToStrip = [
    "SEMANTIC FINGERPRINT DEPLOYMENT RULES",
    "ENTITY-SECTION ARCHITECTURE",
    "ENTITY RELATIONSHIP MAPPING (CRITICAL)",
    "MAXIMUM MARGIN SEPARATION (CRITICAL)",
    "GOVERNANCE CONSTRAINTS",
    "--- ARTICLE CLASSIFICATION (Narrative Governance) ---",
    "EXACT COVERAGE QUANTIZATION - SUMMARY BOX (Featured Snippet Optimization)"
  ];

  sectionsToStrip.forEach(function(header) {
    var startIdx = text.indexOf(header);
    if (startIdx === -1) return;

    // Find the next section header after this one, or end of string
    var searchFrom = startIdx + header.length;
    var nextHeaderIdx = text.length;

    var allHeaders = [
      "PAGE IDENTITY", "QUERY AUTHORITY", "PROBLEM ANGLE", "ENTITY COVERAGE REQUIRED",
      "SEMANTIC FINGERPRINT DEPLOYMENT RULES", "ENTITY-SECTION ARCHITECTURE",
      "ENTITY RELATIONSHIP MAPPING (CRITICAL)", "MAXIMUM MARGIN SEPARATION (CRITICAL)",
      "GOVERNANCE CONSTRAINTS", "REWRITE INSTRUCTION", "SCOPE BOUNDARIES",
      "OBSERVED PROBLEM", "GSC AUTHORITY GAP ANALYSIS", "EXACT COVERAGE QUANTIZATION",
      "--- ARTICLE CLASSIFICATION", "--- END ARTICLE CLASSIFICATION ---"
    ];

    allHeaders.forEach(function(h) {
      if (h === header) return;
      var idx = text.indexOf(h, searchFrom);
      if (idx > -1 && idx < nextHeaderIdx) nextHeaderIdx = idx;
    });

    text = text.substring(0, startIdx) + text.substring(nextHeaderIdx);
  });

  // Strip raw Benchmark/Monitor query line-dumps between their start marker
  // and the section's closing divider — keep only the summary stats above.
  var dumpStart = text.indexOf("Benchmark query clusters");
  if (dumpStart > -1) {
    var dividerAfter = text.indexOf("------------------------------------------------------------", dumpStart);
    if (dividerAfter > -1) {
      text = text.substring(0, dumpStart) + text.substring(dividerAfter + "------------------------------------------------------------".length);
    } else {
      text = text.substring(0, dumpStart);
    }
  }

  // Remove orphaned closing marker left behind if its opening section was stripped
  text = text.replace(/--- END ARTICLE CLASSIFICATION ---\n?/g, "");

  // Collapse resulting multiple blank lines
  text = text.replace(/\n{3,}/g, "\n\n");

  return text.trim();
}

function getMetaPrompt() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const sheet     = ss.getSheetByName("posts");
  const activeRow = sheet.getActiveRange().getRow();
  const d         = getActiveRowDataMap();

  const postID      = d["Post ID"]      || "Unknown";
  const material    = d["Stone Type"]   || "Unknown";
  const articleType = d["Article Type"] || "General";
  const url         = d["URL"]          || "";
 const location = String(d["Locality"] || d["Location"] || "").trim();

  const primaryTerm    = String(d["Primary Search Term"]         || "").trim();
  const queryCluster   = String(d["Primary Query Cluster Owned"] || "").trim();
  const problemAngle   = String(d["Problem Angle"]               || "").trim();
  const authorityBrief = String(d["Authority Brief"]             || "").trim();
  const competitorSerpNotes = String(d["Competitor SERP Notes"]  || "").trim();

  const systemLabels  = ["marble hub","travertine hub","limestone hub","slate hub",
                         "terrazzo hub","victorian hub","sandstone hub","quartzite hub",
                         "porcelain hub","terracotta hub","granite hub",
                         "marble hub page","travertine hub page","limestone hub page"];
  const isSystemLabel = systemLabels.indexOf(primaryTerm.toLowerCase()) > -1;
  const effectiveTerm = (!primaryTerm || isSystemLabel) ? queryCluster : primaryTerm;

  const materialLower = material.toLowerCase();

  const isMethodGuideType = (articleType === "Method Guide");

  // Detect a safety/risk constraint word from Problem Angle or Rewrite Brief, if present
  var safetyConstraintWord = "";
  var combinedForConstraint = (problemAngle + " " + authorityBrief).toLowerCase();
  if (/\bsafely\b/.test(combinedForConstraint)) safetyConstraintWord = "safely";
  else if (/\bsafest\b/.test(combinedForConstraint)) safetyConstraintWord = "safest";
  else if (/\bsafe\b/.test(combinedForConstraint)) safetyConstraintWord = "safe";

  const pagePurposeLock = isMethodGuideType
    ? "--- PAGE-PURPOSE LOCK (Hard Lock — Method Guide) ---\n" +
      "This page's defining purpose is the task named in the Primary Search Term: '" + effectiveTerm + "'.\n" +
      "The H1, the Meta Title, and the first sentence of the Meta Description must each explicitly communicate this task —\n" +
      "not only the symptoms, contamination, or surrounding conditions that motivate it.\n" +
      (safetyConstraintWord
        ? "The Problem Angle / Rewrite Brief for this page centres on a safety or risk constraint (the word '" + safetyConstraintWord + "' or a close variant).\n" +
          "That safety constraint MUST also appear explicitly in the H1 and the Meta Title — not only implied by context.\n"
        : "") +
      "A field that describes only the coating condition, contamination, or symptom — without naming the task itself\n" +
      "('" + effectiveTerm + "' or a natural close variant) — has failed this lock and must be rewritten before any other rule is applied.\n\n" +
      "RULE PRIORITY FOR METHOD GUIDES (Hard Lock):\n" +
      "Apply rules in this strict order when they conflict:\n" +
      "1. State the exact task and material named in the Primary Search Term.\n" +
      "2. State the governing safety or risk constraint, if one is specified above.\n" +
      "3. Reflect the primary query / GSC language naturally.\n" +
      "4. Add recognition, tension, or curiosity through phrasing.\n" +
      "Lower-priority stylistic rules (symptom-led framing, counterintuitive tension) must never obscure\n" +
      "the task and material named in steps 1 and 2.\n\n"
    : "";

  const groundingStepBlock = isMethodGuideType
    ? "--- GROUNDING STEP (Hard Lock — Method Guide — complete before writing anything) ---\n" +
      "This is a process/method article, not a project case study. Do NOT invent or search for a unique floor,\n" +
      "location, coating brand, or individual project condition unless one is actually present in the ARTICLE CONTENT above.\n" +
      "Before writing any field, write one internal sentence (do not include it in your output) stating the exact task\n" +
      "this article teaches, the material involved, and the principal safety or scope constraint that governs it.\n" +
      "For this page, the grounding statement must be equivalent to:\n" +
      "\"This article explains how to " + effectiveTerm + (safetyConstraintWord ? " " + safetyConstraintWord : "") + ", covering identification, controlled application, recovery, rinsing and drying.\"\n" +
      "Every field you write below must preserve this central task and constraint. If a candidate H1, Meta Title, or\n" +
      "Meta Description could apply to a different task on the same material, it has failed this step — rewrite it\n" +
      "using the task and constraint named above.\n\n"
    : "--- GROUNDING STEP (Hard Lock — complete before writing anything) ---\n" +
      "Before writing any field, write one internal sentence (do not include it in your output) stating what is UNIQUE about this specific page — not what the material generally involves, but what actually happened on THIS floor, in THIS location, with THIS specific problem.\n" +
      "Use only facts drawn from the ARTICLE CONTENT above — the specific coating brand, the specific defect, the specific outcome.\n" +
      "Every field you write below must trace back to this grounding sentence. If a candidate H1, Meta Title, or Meta Description could apply to any other article about the same material, it has failed the grounding step — discard it and rewrite using a more specific detail from the article content.\n\n";

  const methodGuidePurposeCheck = isMethodGuideType
    ? "19. PURPOSE-LOSS TEST (Method Guide): After drafting each field, ask — could a reader identify from\n" +
      "    this field alone that the page explains how to " + effectiveTerm + (safetyConstraintWord ? (" " + safetyConstraintWord) : "") + "?\n" +
      "    If the answer is no, rewrite the field before checking length or stylistic quality, then re-run checks 1-18.\n"
    : "";

  var articleHtml = "";
  try {
    var col104  = String(sheet.getRange(activeRow, 104).getValue() || "").trim();
    var col98   = String(sheet.getRange(activeRow, 98).getValue()  || "").trim();
    var rawHtml = col104 || col98 || "";

    if (rawHtml) {
      var headerMatch = rawHtml.match(/<header[\s\S]*?<\/header>/i);
      var h2Matches   = rawHtml.match(/<h2[^>]*>([\s\S]*?)<\/h2>/gi) || [];
      var headerText  = headerMatch
        ? headerMatch[0].replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim()
        : "";
      var h2Text = h2Matches
        .map(function(h) { return h.replace(/<[^>]+>/g, "").trim(); })
        .join("\n");

      if (headerText || h2Text) {
        articleHtml = (headerText ? "INTRO:\n" + headerText + "\n\n" : "") +
                      (h2Text     ? "SECTION HEADINGS:\n" + h2Text   : "");
      } else {
        articleHtml = rawHtml.substring(0, 2000);
      }
    }
  } catch(e) {
    articleHtml = "";
  }

  var descExample1 = material + " floors that never look clean usually have surface damage, " +
    "not dirt. Correcting the surface rather than cleaning it restores clarity and simplifies maintenance.";

  var descExample2 = material + " tiles that deteriorate despite careful cleaning are losing " +
    "surface integrity, not accumulating dirt. Professional restoration recovers the finish and " +
    "stabilises the surface.";

  var descWrong1 = material + " floors that never look clean usually have surface damage, " +
    "not dirt. Understand how " + materialLower + " behaves and choose the next step to restore " +
    "a consistent finish.";

  var descWrong2 = material + " floors that never look clean usually have surface damage, " +
    "not dirt. So correcting the surface restores clarity.";

  var descWrong3 = material + " floors that never look clean are usually surface damage, not dirt.";

  var descWrong4 = material + " Floors That Never Look Clean Have Surface Damage. " +
    "Correcting The Surface Restores Clarity.";

  // Location rule blocks — triggered whenever a Locality value is present
  // on the active row, regardless of article type. Previously scoped only
  // to Geo Service Page, which left Case Studies (and any other located
  // article type) with no protection against the locality being silently
  // dropped or generalised away during H1/Meta Title generation.
  const isLocationPage = Boolean(location);

  const geoH1Rule = isLocationPage
    ? "- LOCATION PRESERVATION RULE (Hard Lock): The H1 MUST include the target location (" + location + ").\n" +
      "  An H1 without the location name fails the self-check, regardless of article type.\n" +
      "  CORRECT: '" + material + " Honing And Polishing In " + location +
        " — What It Actually Involves'\n" +
      "  WRONG:   '" + material + " Floors That Need More Than Cleaning'\n"
    : "";

  const geoMetaTitleRule = isLocationPage
    ? "- LOCATION PRESERVATION RULE (Hard Lock): The Meta Title MUST include the target location (" +
        location + ").\n" +
      "  CORRECT: '" + material + " Honing & Polishing in " + location +
        " | Abbey Floor Care'\n" +
      "  WRONG:   '" + material + " Honing & Polishing Services | Abbey Floor Care'\n"
    : "";

  const geoSelfCheck = isLocationPage
    ? "17. LOCALITY IS SET FOR THIS PAGE — verify H1 contains '" + location + "'.\n" +
      "    If absent — rewrite H1 to include the location name. This applies regardless of article type.\n" +
      "18. LOCALITY IS SET FOR THIS PAGE — verify Meta Title contains '" +
        location + "'.\n" +
      "    If absent — rewrite Meta Title to include the location name. This applies regardless of article type.\n"
    : "";

  return (
    "H1 + META GENERATION\n" +
    "ROLE: Senior UK SEO Strategist\n" +
    "TASK: Write exactly four fields for this page. " +
    "Output only the four labelled fields — no preamble, no explanation, no markdown.\n\n" +

    "--- PAGE CONTEXT ---\n" +
    "POST ID: "             + postID        + "\n" +
    "URL: "                 + url           + "\n" +
    "MATERIAL: "            + material      + "\n" +
    "ARTICLE TYPE: "        + articleType   + "\n" +
    "PRIMARY SEARCH TERM: " + effectiveTerm + "\n" +
    (location       ? "LOCATION: "        + location       + "\n" : "") +
    (problemAngle   ? "PROBLEM ANGLE: "   + problemAngle   + "\n" : "") +
    (authorityBrief ? "AUTHORITY BRIEF: " + filterAuthorityBriefForMeta(authorityBrief) + "\n" : "") +
    (function() {
      var postIdForQueries = String(d["Post ID"] || "").trim();
      var montr  = postIdForQueries ? bc_getQueriesFromSheet(postIdForQueries, 'GSC Monitor')   : String(d["Montr Queries"]  || "").trim();
      var benmrk = postIdForQueries ? bc_getQueriesFromSheet(postIdForQueries, 'GSC Benchmark') : String(d["Benmrk Queries"] || "").trim();
      var basel  = postIdForQueries ? bc_getQueriesFromSheet(postIdForQueries, 'GSC Baseline')  : String(d["Basel Queries"]  || "").trim();
      var queries = montr || benmrk || basel;
      if (!queries) return "";
      var lines = queries.split(/\n/).map(function(l) { return l.trim(); }).filter(function(l) { return l.length > 3; });
      if (lines.length === 0) return "";

      // Sort by impressions (not source order) so high-impression, low/zero-click
      // queries surface first — these reveal the exact CTR failure this stage
      // is meant to fix. Taking the first 10 lines in raw sheet order can silently
      // exclude the query actually driving the impression volume.
      var parsedLines = lines.map(function(l) {
        var parts = l.split(/\s{2,}/).map(function(p) { return p.trim(); });
        var impressions = parts.length >= 3 ? (parseInt(parts[2].replace(/,/g, '')) || 0) : 0;
        return { line: l, impressions: impressions };
      });
      parsedLines.sort(function(a, b) { return b.impressions - a.impressions; });
      var top = parsedLines.slice(0, 10).map(function(p) { return p.line; }).join("\n");
      return "GSC REAL SEARCHER QUERIES (use these to identify emotional entry points and natural language):\n" +
             "These are the actual phrases real homeowners have used to find this page or pages like it.\n" +
             "Use this language to inform the H1 tension, meta title phrasing, and meta description specificity.\n" +
             "Do NOT copy these phrases verbatim — use them to understand what the homeowner is worried about.\n\n" +
             top + "\n\n";
    })() +
    "\n" +

    (competitorSerpNotes
      ? "--- COMPETITOR SERP CONTEXT ---\n" +
        "The following titles/descriptions are already ranking for this page's primary search term.\n" +
        "Use this to differentiate — do not duplicate the pattern below.\n" +
        "This is a qualitative snapshot, not exhaustive data — weigh it alongside the other rules, not above them.\n\n" +
        competitorSerpNotes + "\n\n"
      : "") +
    (articleHtml
      ? "--- ARTICLE CONTENT ---\n" +
        "Use the intro and section headings below to inform tone, angle, and specificity.\n" +
        "Do not copy phrases verbatim — use the content to understand what the page covers.\n\n" +
        articleHtml + "\n\n"
      : "") +

    "--- VOICE & TONE ---\n" +
    "Write as a knowledgeable professional speaking directly to a frustrated homeowner.\n" +
    "The tone is authoritative but plainspoken — no corporate language, no marketing superlatives.\n" +
    "The H1 should sound like something a specialist would say to a client, not a page title.\n" +
    "The Meta Description should read like the opening of an expert explanation — not an advert.\n\n" +
    "--- GSC LANGUAGE SIGNAL (apply before writing any field) ---\n" +
    "If GSC Real Searcher Queries are provided above, read them before writing anything.\n" +
    "Identify the dominant emotional concern in those queries — what is the homeowner worried about?\n" +
    "That emotional concern must be the entry point for the H1 and the first sentence of the Meta Description.\n" +
    "The H1 must create recognition — the homeowner should read it and think 'that is exactly my situation'.\n" +
    "The Meta Title must read like a phrase a person would actually say — not a keyword string assembled for a crawler.\n" +
    "The Meta Description second sentence must give a specific reason to click — what will the reader understand\n" +
    "or be able to do after reading this article that they could not before?\n\n" +
    "--- AI CITATION OPTIMISATION ---\n" +
    "AI systems cite pages that directly answer a specific question with a concrete finding.\n" +
    "The Meta Description must contain at least one specific claim or finding — not a generic outcome statement.\n" +
    "WRONG: 'Professional restoration recovers colour and stabilises the surface.'\n" +
    "CORRECT: 'This Penkhull hallway shows what carpet-covered Victorian tiles look like after careful restoration — and why the bedding matters as much as the surface.'\n" +
    "The specific finding should reference the location, the material condition, or the unexpected discovery\n" +
    "where possible — these are the signals that make AI select one page over another as a citation source.\n\n" +

    "--- CAPITALISATION RULES (HARD LOCK — apply before any other rule) ---\n" +
    "H1: Title Case — THE FIRST LETTER OF EVERY WORD MUST BE CAPITALISED. NO EXCEPTIONS.\n" +
    "    CORRECT: '" + material + " Floor Never Feels Clean Despite Careful Care'\n" +
    "    WRONG:   '" + material + " floor never feels clean despite careful care'\n" +
    "    Every word including short words like 'a', 'of', 'in', 'to' must be capitalised.\n\n" +
    "Meta Title: Title Case — capitalise the first letter of every word except minor\n" +
    "    prepositions and articles (a, an, the, in, on, at, for, of, to, and, but, or)\n" +
    "    unless they are the first word.\n" +
    "    CORRECT: '" + material + " Floors Not Staying Clean | Abbey Floor Care'\n" +
    "    WRONG:   '" + materialLower + " floors not staying clean | Abbey Floor Care'\n\n" +
    "Meta Description: Sentence case ONLY.\n" +
    "    Capitalise ONLY the first word of each sentence and proper nouns\n" +
    "    (Abbey Floor Care, David Allen).\n" +
    "    Do NOT capitalise the stone or tile type mid-sentence: " + materialLower + ".\n" +
    "    Do NOT capitalise service words mid-sentence: cleaning, polishing, honing,\n" +
    "    sealing, restoration, repair, grinding.\n" +
    "    CORRECT: '" + material + " floors that never look clean have surface damage, not dirt.'\n" +
    "    WRONG:   '" + material + " Floors That Never Look Clean Have Surface Damage, Not Dirt.'\n\n" +
    "Yoast Keyphrase: Lowercase only — no capitalisation, no punctuation.\n" +
    "    CORRECT: '" + materialLower + " floor not staying clean'\n" +
    "    WRONG:   '" + material + " Floor Not Staying Clean'\n\n" +

    pagePurposeLock +
    groundingStepBlock +

    "--- RULES ---\n\n" +

    "H1:\n" +
    "- CAPITALISATION: Title Case — first letter of EVERY word capitalised — no exceptions\n" +
    "- Write fresh — do not copy any heading from the article content above\n" +
    "- EXACTLY 40-60 characters — count every character before outputting\n" +
    "- Must name the material (" + material + ")\n" +
    "- Lead with the user's symptom or problem, OR use a counterintuitive framing\n" +
    (isMethodGuideType
      ? "- METHOD GUIDE OVERRIDE (Hard Lock): For this article type, symptom-led or counterintuitive framing is\n" +
        "  permitted only AFTER the core task from the Page-Purpose Lock above is explicit in the H1. The H1 must\n" +
        "  state that '" + effectiveTerm + "'" + (safetyConstraintWord ? (" is being done " + safetyConstraintWord) : "") + " — tension or symptom framing may be layered around that task, but must never replace it.\n"
      : "") +
    "- Specific and concrete — no generic labels like 'Guide', 'Overview', 'Diagnosis'\n" +
    "- No verb forms ending in '-ing' as the first word\n" +
    "- Must create curiosity or tension — avoid obvious or predictable phrasing\n" +
    "- Prohibited patterns: 'looks dull after cleaning', 'everything you need to know',\n" +
    "  'the complete guide', 'marks still visible', 'still visible after'\n" +
    "- No question marks unless the question is genuinely surprising or counterintuitive\n" +
    "- NEAR ME RULE (Hard Lock): Never use the phrase 'near me' in the H1.\n" +
    "  'Near me' is a search modifier — no homeowner would say it in conversation.\n" +
    "  Strip it from any candidate H1 and rephrase around the core intent.\n" +
    "  WRONG: 'Victorian Tile Restoration Near Me Found Hidden Damage'\n" +
    "  CORRECT: 'What Carpet Hid In This Penkhull Victorian Tile Hallway'\n" +
    "- SOLUTION INTEGRITY RULE (Hard Lock): The H1 must never cast doubt on, or imply\n" +
    "  the inadequacy of, the solution this page actually delivers. If the Primary\n" +
    "  Search Term or article content centres on a product, method, or kit that this\n" +
    "  page confirms works for its stated scope, the H1 must build tension around the\n" +
    "  READER'S SYMPTOM (e.g. dirt-trapping holes, failed old filler, a neglected look)\n" +
    "  — never around doubt in the product or method itself.\n" +
    "  Any escalation boundary (e.g. cracks needing resin, or a professional handoff\n" +
    "  point) belongs in the Meta Description's specific detail only — it must NOT\n" +
    "  become the central claim of the H1.\n" +
    "  WRONG: '" + material + " Floor Holes Need More Than A Repair Kit' — casts doubt on\n" +
    "  the page's own solution.\n" +
    "  CORRECT: 'Dirty " + material + " Holes Trap More Than Dust' — tension from the\n" +
    "  symptom, not the solution.\n" +
    (geoH1Rule ? geoH1Rule : "") +
    "- Plain text only — no HTML tags, no markdown\n\n" +

    "META TITLE:\n" +
    "- CAPITALISATION: Title Case — first letter of every word except minor prepositions\n" +
    "  and articles unless first word — see Capitalisation Rules above\n" +
    "- TARGET 50-60 characters. Never exceed 60 characters.\n" +
    "- Must contain the primary search term or a natural close variant\n" +
    "- Must include 'Abbey Floor Care' as the brand signal — never truncate to 'Abbey' alone\n" +
    "- Must read naturally in UK English\n" +
    "- CTR PURPOSE (Hard Lock): The Meta Title is not merely a label for the service or article.\n" +
    "  It must give the searcher a reason to choose this result over otherwise similar SERP results.\n" +
    "- Build the click-driving angle around ONE of the following, whichever is genuinely supported by\n" +
    "  the Page Context, Article Content, Problem Angle, Authority Brief or Competitor SERP Context:\n" +
    "  * a visible problem the homeowner recognises\n" +
    "  * a decision they are unsure about\n" +
    "  * a meaningful limitation or trade-off\n" +
    "  * a consequence of choosing the wrong treatment\n" +
    "  * a repair-versus-replacement distinction\n" +
    "  * an unexpected or counterintuitive fact\n" +
    "  * a specific 'what actually matters' distinction\n" +
    "- Prefer tension, contrast or useful specificity over generic service wording.\n" +
    "- The title must make an implicit promise that the page resolves a real uncertainty.\n" +
    "- BORING TITLE TEST (Hard Lock): Reject any candidate that is essentially only:\n" +
    "  [service] + [location] + [brand].\n" +
    "  Example to reject: 'Ceramic Tile Repair West End Edinburgh | Abbey Floor Care'.\n" +
    "- Also reject generic modifiers that add no useful reason to click, including:\n" +
    "  'professional', 'quality', 'trusted', 'local', 'reliable', 'specialist service',\n" +
    "  'services', 'solutions', 'experts', 'expert', 'guide', 'overview', 'tips', 'hub'.\n" +
    "- Strong patterns may include natural phrases such as:\n" +
    "  'What Can Actually Be Repaired',\n" +
    "  'Repair or Replace?',\n" +
    "  'Why the Damage Keeps Returning',\n" +
    "  'When Cleaning Will Not Fix It',\n" +
    "  'What the Damage Is Really Telling You',\n" +
    "  but ONLY when that angle is supported by this page.\n" +
    "- Do not manufacture drama. No fake urgency, fear, exaggeration, guarantees or unsupported claims.\n" +
    "- Do not use empty clickbait such as 'You Won't Believe', 'Shocking', 'Secret', 'Amazing',\n" +
    "  'Game-Changing', 'Must See' or similar tabloid phrasing.\n" +
    "- Use the COMPETITOR SERP CONTEXT above to avoid repeating the dominant title pattern.\n" +
    "  If competitors are mostly using plain service + location titles, deliberately choose a more\n" +
    "  useful problem-, decision-, consequence- or distinction-led angle that this article supports.\n" +
    "- Word order must follow natural English syntax\n" +
    "- NEAR ME RULE (Hard Lock): Never use the phrase 'near me' in the Meta Title.\n" +
    "  Replace with the location name where available, or use the core phrase without the modifier.\n" +
    "  WRONG: 'Victorian Tile Restoration Near Me in Penkhull | Abbey Floor Care'\n" +
    "  CORRECT: 'Victorian Tile Restoration Penkhull | Abbey Floor Care'\n" +
    (geoMetaTitleRule ? geoMetaTitleRule : "") +
    "- Plain text only — no markdown, no punctuation padding\n\n" +

    "META DESCRIPTION:\n" +
    "- CAPITALISATION: Sentence case only — see Capitalisation Rules above\n" +
    "- EXACTLY 140-160 characters — count every character before outputting\n" +
    "- Exactly two sentences — a full stop ends sentence one, sentence two ends the description\n" +
    "- SENTENCE ONE RULES:\n" +
    "  * States the core problem of this specific page as established fact\n" +
    "  * Must be specific to THIS article — not a generic stone care statement\n" +
    "  * Must use correct grammar — '" + materialLower + " floors have surface damage'\n" +
    "    not '" + materialLower + " floors are surface damage'\n" +
    "  * SUBJECT RULE (Hard Lock): If the Primary Search Term names a product, kit, or\n" +
    "    tool (e.g. 'repair kit', 'filler kit'), the grammatical subject of sentence one\n" +
    "    must still be the FLOOR or the MATERIAL — never the product itself.\n" +
    "    WRONG: '" + material + " repair kit holes fail when...' — implies the kit has holes.\n" +
    "    CORRECT: '" + materialLower + " floors with dirty holes can be too open for a repair kit alone.'\n" +
    "  * Prohibited openers: 'Learn', 'Discover', 'Find out', 'This guide',\n" +
    "    'This page', 'So', 'And', 'But'\n" +
    "  * Must NOT open with a question\n" +
    "  * Must NOT use a connector word as the first word\n" +
    "- SENTENCE TWO RULES:\n" +
    "  * Must state what WILL HAPPEN TO THE FLOOR — not what the reader will do,\n" +
    "    understand, learn, or choose\n" +
    "  * Must describe a concrete physical outcome — clarity restored, finish recovered,\n" +
    "    surface stabilised, maintenance simplified\n" +
    "  * Must NOT be a continuation of sentence one joined by 'so' or 'and'\n" +
    "  * Must NOT begin with a connector word: 'So', 'And', 'But', 'This means'\n" +
    "  * Prohibited phrases: 'contact us', 'find out more', 'click here',\n" +
    "    'get in touch', 'choose the right', 'the right next step',\n" +
    "    'understand how', 'learn how', 'discover how'\n" +
    "  * AI CITATION SPECIFICITY RULE (Hard Lock):\n" +
    "    The second sentence must contain at least one specific detail that could not\n" +
    "    appear on any other article about the same material.\n" +
    "    IF a location is available: reference the specific location AND the specific\n" +
    "    condition found or corrected — not just a generic outcome.\n" +
    "    CORRECT with location: 'This Penkhull hallway regained stable colour after\n" +
    "    carpet concealed movement and trapped residue for decades.'\n" +
    "    IF no location is available: reference the specific material condition,\n" +
    "    the unexpected finding, or the specific intervention that makes this article\n" +
    "    different from a generic article on the same topic.\n" +
    "    CORRECT without location: 'Victorian tile floors where linseed oil has\n" +
    "    darkened the clay body permanently need stripping before colour can return.'\n" +
    "    WRONG in both cases: 'Professional restoration recovers colour and stabilises\n" +
    "    the surface.' — too generic to be citable as evidence of a specific finding.\n" +
    "    " + ce_getLocationVarietyWithoutFabricationLock().replace(/\n/g, '\n    ') + "\n" +
    "- Names the material (" + material + ") at least once across both sentences\n" +
    "- Write with authority — not reassurance\n" +
    "- Plain text only — no markdown\n\n" +

    "YOAST KEYPHRASE:\n" +
    "- The focus keyphrase Yoast SEO will use to evaluate this page\n" +
    "- Derive from the Primary Search Term provided in Page Context\n" +
    "- If the Primary Search Term is longer than 6 words, distil it to its\n" +
    "  core 2-6 word search intent — the phrase a homeowner would actually type\n" +
    "- Must appear naturally in the Meta Title and Meta Description you have already written — if it\n" +
    "  does not, adjust the keyphrase until it matches what is already in those fields\n" +
    (isMethodGuideType
      ? "- METHOD GUIDE EXCEPTION (Hard Lock): The H1 may use a grammatically natural close variant of the\n" +
        "  keyphrase with the same task intent, rather than the exact phrase — natural H1 phrasing takes\n" +
        "  priority over exact keyphrase matching in the H1 only. Exact matching is still required in the\n" +
        "  Meta Title and Meta Description.\n"
      : "- Must appear naturally in the H1 as well — if it does not, adjust the keyphrase until it matches.\n") +
    "- Must name the material (" + materialLower + ")\n" +
    "- Format: lowercase only — no capitalisation, no punctuation\n" +
    "- 2-6 words maximum\n" +
    "- MINIMUM LENGTH RULE: For Hub Page articles, the keyphrase must be\n" +
    "  a minimum of 3 words. A single material name (e.g. '" + materialLower + "')\n" +
    "  is too generic for a hub page and will not pass Yoast evaluation.\n" +
    "  Use a phrase that reflects the hub's specific scope — for example:\n" +
    "  '" + materialLower + " floor restoration' or '" + materialLower + " floor care guide'\n" +
    "  rather than just '" + materialLower + "' alone.\n" +
    "- Prohibited: brand names (abbey, abbey floor care),\n" +
    "  article type labels (hub, guide, overview)\n" +
    "- Plain text only — no markdown\n\n" +

    "--- META DESCRIPTION CONSTRUCTION METHOD ---\n" +
    "Build it in this exact sequence:\n" +
    "Step 1: Write sentence one — the problem stated as fact (aim for 70-85 characters)\n" +
    "Step 2: Write sentence two — a concrete outcome that happens to the floor\n" +
    "        (aim for 65-80 characters)\n" +
    "Step 3: Count total characters. If under 140 — expand the shorter sentence.\n" +
    "        If over 160 — tighten the longer sentence. Repeat until 140-160.\n" +
    "Step 4: Verify sentence two states what happens to the floor, not what the reader does.\n" +
    "Step 5: Verify sentence two does not begin with a connector word.\n" +
    "Step 6: Verify no mid-sentence capitalisation of stone or service nouns.\n\n" +

    "--- QUALITY TARGET ---\n" +
    "The H1 must be specific, surprising, or tension-creating — not a label for the page.\n" +
    "The Meta Title must read like a real phrase — not assembled from keywords.\n" +
    "The Meta Description first sentence states the problem as established fact.\n" +
    "The Meta Description second sentence names a concrete outcome for the floor.\n" +
    "The Yoast Keyphrase must be the phrase a homeowner would actually search —\n" +
    "short, specific, lowercase, and present in the H1, Meta Title, and Meta Description.\n\n" +

    "--- EXAMPLES OF CORRECT META DESCRIPTION STRUCTURE ---\n" +
    "CORRECT: '" + descExample1 + "'\n\n" +
    "CORRECT: '" + descExample2 + "'\n\n" +
    "WRONG — weak sentence two: '" + descWrong1 + "' " +
    "— sentence two tells the reader what to do, not what outcome the floor achieves.\n\n" +
    "WRONG — connector opener: '" + descWrong2 + "' " +
    "— sentence two must not open with 'So'.\n\n" +
    "WRONG — grammar error: '" + descWrong3 + "' " +
    "— floors cannot 'be' damage. Use 'have surface damage' or 'show surface damage'.\n\n" +
    "WRONG — title case in description: '" + descWrong4 + "' " +
    "— Meta Description must be sentence case only.\n\n" +

    "--- SELF-CHECK BEFORE OUTPUTTING ---\n" +
    "1.  Count H1 characters. If not 40-60 — rewrite.\n" +
    "2.  Verify H1 is Title Case — every word capitalised including short words.\n" +
    "    If any word is lowercase — rewrite the entire H1.\n" +
    "3.  Count Meta Title characters. If over 60 — rewrite. Aim for 50-60 where possible.\n" +
    "4.  Verify Meta Title is Title Case with preposition exceptions. If not — rewrite.\n" +
    "5.  Verify Meta Title word order follows natural English syntax. If inverted — rewrite.\n" +
    "6.  BORING TITLE TEST: Does the Meta Title amount to little more than service + location + brand?\n" +
    "    If yes — REJECT it and rewrite with a supported problem, decision, consequence,\n" +
    "    limitation, contrast or 'what actually matters' angle.\n" +
    "    The title must give a searcher a concrete reason to choose this result over a generic competitor listing.\n" +
    "7.  Verify Meta Title contains no prohibited or empty promotional words. If it does — rewrite.\n" +
    "8.  Verify the CTR angle is genuinely supported by the article, Problem Angle, Authority Brief\n" +
    "    or Competitor SERP Context. If the angle is invented, exaggerated or unsupported — rewrite.\n" +
    "7.  Count Meta Description characters. If not 140-160 — rewrite.\n" +
    "8.  Verify Meta Description has exactly two sentences. If not — rewrite.\n" +
    "9.  Verify sentence one does not open with a prohibited opener. If it does — rewrite.\n" +
    "10. Verify sentence one uses correct grammar — floors have/show damage,\n" +
    "    not floors are damage. If wrong — rewrite.\n" +
    "11. Verify sentence two states a concrete floor outcome — not a reader action.\n" +
    "    If vague or reader-led — rewrite.\n" +
    "12. Verify sentence two does not begin with a connector word. If it does — rewrite.\n" +
    "13. Verify Meta Description uses sentence case — no mid-sentence capitals on\n" +
    "    stone type (" + materialLower + ") or service words. If any found — rewrite.\n" +
    "14. Verify Yoast Keyphrase is lowercase only. If any capitals — rewrite.\n" +
    "15. Verify Yoast Keyphrase is 2-6 words. If longer — distil.\n" +
    "16. Verify Yoast Keyphrase appears naturally in Meta Title and Meta Description. If absent from either —\n" +
    "    adjust keyphrase to match. " + (isMethodGuideType ? "For the H1 only, a natural close variant with the same task intent is acceptable per the Method Guide Exception above.\n" : "For the H1, verify it appears naturally as well — if absent, adjust keyphrase to match.\n") +
    (geoSelfCheck ? geoSelfCheck : "") +
    "17. Verify Meta Description sentence two contains at least one specific detail\n" +
    "    that could not appear on any other article about the same material and article type.\n" +
    "    Use the first available detail from this priority list:\n" +
    "    PRIORITY 1 — Location and condition: If a location is present (" + (location ? location : "none") + ")\n" +
    "    sentence two must name the location AND the specific condition found or corrected.\n" +
    (location
      ? "    The location '" + location + "' must appear in sentence two.\n" +
        "    If it does not — rewrite sentence two to include it.\n"
      : "    No location is available — skip to Priority 2.\n") +
    "    PRIORITY 2 — Material condition: If no location, name the specific behaviour\n" +
    "    of this material (" + material + ") that is central to THIS article's actual\n" +
    "    content (see ARTICLE CONTENT section above) — not a condition borrowed from\n" +
    "    the examples below if it falls outside this page's scope.\n" +
    "    HARD CHECK: Before using any condition, verify it does NOT fall within a\n" +
    "    category excluded by the Cannibalisation Guardrail in Scope Boundaries\n" +
    "    above (e.g. a Cleaning-scoped page must not cite a Sealing, Honing/Polishing,\n" +
    "    or Restoration-level defect as its specific detail — that belongs on the\n" +
    "    Safe Handoff Page instead, and using it here re-creates the exact scope\n" +
    "    drift the Rewrite Brief was written to remove).\n" +
    "    ESCALATION BOUNDARY AS SPECIFICITY SOURCE: If the AUTHORITY BRIEF above contains\n" +
    "    a Rewrite Brief with an 'ESCALATION BOUNDARY:' labelled section, that is the\n" +
    "    authoritative, article-verified statement of where DIY or basic intervention\n" +
    "    stops and professional/next-stage work begins — use its exact stated action,\n" +
    "    do not substitute a generic assumption (e.g. do not default to naming resin,\n" +
    "    sealing, or any specific material/technique unless the ESCALATION BOUNDARY\n" +
    "    text itself names it as the correct next step).\n" +
    "    CRITICAL PHRASING CHECK: If the ESCALATION BOUNDARY text states that a\n" +
    "    material/technique is professional-only or not a homeowner action, your\n" +
    "    Meta Description sentence must make that professional-only status explicit —\n" +
    "    never compress it into a phrase implying the reader will do it themselves.\n" +
    "    This check applies to whatever material or technique the ESCALATION BOUNDARY\n" +
    "    names for THIS row's material — not any single fixed example. Apply the same\n" +
    "    logic to any professional-only material/technique this article's ESCALATION\n" +
    "    BOUNDARY identifies (e.g. resin, acid cleaning, specialist sealants,\n" +
    "    mechanical honing — whatever applies here).\n" +
    "    WRONG: 'cracks or honeycomb holes need resin' — reads as the reader's next\n" +
    "    DIY step, contradicting an ESCALATION BOUNDARY that says resin is\n" +
    "    professional-only.\n" +
    "    CORRECT: 'cracks or extensive damage need a professional' or 'larger repairs\n" +
    "    need a professional, not a DIY resin fix' — preserves the professional-only\n" +
    "    fact even under character-count pressure.\n" +
    "    POSITIVE STATEMENT REQUIREMENT (Hard Lock): Sentence two must state the\n" +
    "    escalation fact as a positive claim about what DOES happen — not merely what\n" +
    "    does NOT happen. Avoiding the wrong claim is not sufficient on its own. This\n" +
    "    applies regardless of material or which specific technique the ESCALATION\n" +
    "    BOUNDARY names for this row.\n" +
    "    WEAK (avoidance only, insufficient — pattern, not literal wording): stating\n" +
    "    only that a technique is NOT used, without saying what happens instead.\n" +
    "    CORRECT (positive fact): state the actual next step named in the ESCALATION\n" +
    "    BOUNDARY — e.g. 'extensive damage needs professional assessment', or name the\n" +
    "    specific escalation action this article's ESCALATION BOUNDARY identifies.\n" +
    "    If no ESCALATION BOUNDARY label is present, use the Cannibalisation Guardrail\n" +
    "    in Scope Boundaries as the specificity source instead — state the scope limit\n" +
    "    itself (e.g. 'localised repair only') without inventing an escalation material.\n" +
    "    The examples below are illustrative only, not a menu to select from —\n" +
    "    only use a condition that genuinely appears in this article's content.\n" +
    "    This must be a condition specific to " + material + " — not a generic stone care statement.\n" +
    "    Example for marble: etching from acid contact, heat ring damage, or vinegar staining.\n" +
    "    Example for travertine: void collapse, surface cap failure, or filler breakdown.\n" +
    "    Example for victorian tile: linseed oil darkening, bedding movement, or bitumen residue.\n" +
    "    Example for slate: delamination, cleavage separation, or iron staining.\n" +
    "    PRIORITY 3 — Specific intervention: If neither location nor condition is distinctive,\n" +
    "    name the specific professional decision that governed this job and why it differed\n" +
    "    from a standard approach for this material.\n" +
    "    WRONG in all cases: 'Professional restoration recovered colour and stabilised the surface.'\n" +
    "    This is too generic — it could appear on any article for any material.\n" +
    "    WRONG: Any sentence two that does not contain at least one detail specific to\n" +
    "    this exact article — rewrite until it does.\n" +
    "18. Verify total Meta Description character count is 140-160 after any rewrite\n" +
    "    triggered by check 17.\n" +
    "    If over 160 — tighten sentence one first, preserving the specific detail in sentence two.\n" +
    "    If under 140 — expand the specific detail in sentence two.\n" +
    "    Recount after every edit. Do not output until character count is within 140-160.\n" +
    methodGuidePurposeCheck +
    "\n" +

    "--- MANDATORY SELF-CHECK EXECUTION (Hard Lock) ---\n" +
    "The 18-point self-check above is not optional and not a formality — it is the step that catches the most common failures in this task: character counts outside range, lowercase words in Title Case fields, prohibited openers like 'Learn', and stray question marks fused into a sentence.\n" +
    "Before writing your final answer, go through checks 1–18 one at a time, in order, against your current draft. For each check, if it fails, rewrite the relevant field immediately and re-check it — do not move to the next check until the current one passes.\n" +
    "Do not skip this process because the draft 'looks right'. Draft output that has not been checked against all 18 items is not permitted to be your final answer.\n" +
    "Do not show your self-check work in the output — only the four final fields below, after every check has passed silently.\n\n" +

    "--- OUTPUT FORMAT (EXACT LABELS, ALL FOUR ON CONSECUTIVE LINES) ---\n" +
    "Before outputting, verify each field traces back to the grounding sentence — the specific fact that makes this page different from any other article on the same material.\n" +
    "New H1: [value]\n" +
    "New Meta Title: [value]\n" +
    "New Meta Description: [value]\n" +
    "Yoast Keyphrase: [value]\n\n" +

    "--- GENERATE NOW ---"
  );
}
/* ============================================================
   SCHEMA PROMPT GENERATOR  (replaces Stage 2C schema section)
   Generates full <script type="application/ld+json"> block.
   Alignment rules read from Schema Alignment Validation Sheet.
============================================================ */
function getSchemaPrompt() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const d           = getActiveRowDataMap();
  const postID      = d["Post ID"]      || "Unknown";
  const material    = d["Stone Type"]   || "Unknown";
  const articleType = d["Article Type"] || "General";
  const url         = d["URL"]          || "";

  const savSheet = ss.getSheetByName("Schema Alignment Validation Sheet");
  let schemaRules = "Schema Alignment Validation Sheet not found — use Article schema by default.";

  if (savSheet) {
    const savData    = savSheet.getDataRange().getValues();
    const savHeaders = savData[0].map(function(h) { return String(h).trim(); });
    const artTypeIdx = savHeaders.indexOf("Article Type");
    let matchRow = null;
    for (let i = 1; i < savData.length; i++) {
      if (String(savData[i][artTypeIdx]).trim().toLowerCase() === articleType.toLowerCase()) {
        matchRow = savData[i]; break;
      }
    }
    if (matchRow) {
      const get = function(col) {
        const idx = savHeaders.indexOf(col);
        return idx > -1 ? String(matchRow[idx] || "").trim() : "Not specified";
      };
      schemaRules =
        "Allowed Primary Schema Type: "       + get("Allowed Primary Schema Type")       + "\n" +
        "Allowed Secondary Schema Types: "    + get("Allowed Secondary Schema Types")    + "\n" +
        "Prohibited Schema Types: "           + get("Prohibited Schema Types")            + "\n" +
        "Multiple Primary Schema Allowed: "   + get("Multiple Primary Schema Allowed")   + "\n" +
        "Primary Entity Required In Schema: " + get("Primary Entity Required In Schema")  + "\n" +
        "Must Include Area Served: "          + get("Must Include Area Served")           + "\n" +
        "Schema Type Change Allowed: "        + get("Schema Type Change Allowed")         + "\n" +
        "Schema Expansion Allowed: "          + get("Schema Expansion Allowed")           + "\n" +
        "Maximum Allowed Schema Risk: "       + get("Maximum Allowed Schema Risk")        + "\n" +
        "Assigned Schema Risk: "              + get("Assigned Schema Risk");
    } else {
      schemaRules = "No match for article type '" + articleType + "' in Schema Alignment Validation Sheet.";
    }
  }

  return (
    "SCHEMA GENERATION\n" +
    "ROLE: Senior UK SEO & Stone Restoration Strategist\n" +
    "TASK: Generate the JSON-LD schema block for this page. Output only the schema block — no preamble, no explanation.\n\n" +
    "--- PAGE CONTEXT ---\n" +
    "POST ID: " + postID + "\n" +
    "URL: " + url + "\n" +
    "MATERIAL: " + material + "\n" +
    "ARTICLE TYPE: " + articleType + "\n" +
    "Site root: https://www.abbeyfloorcare.co.uk/\n\n" +
    "--- SCHEMA ALIGNMENT RULES (HARD LOCK) ---\n" +
    schemaRules + "\n\n" +
    "--- FORMAT RULES (HARD LOCK) ---\n" +
    '- Output as a complete HTML block: <script type="application/ld+json">{json}</script>\n' +
    "- JSON must be minified — single line, no internal line breaks whatsoever\n" +
    "- No string value may contain a line break or newline character\n" +
    '- @context must be exactly "https://schema.org"\n' +
    "- @type must comply with Allowed Primary Schema Type above\n" +
    "- Always include BreadcrumbList\n" +
    "- Add secondary schema type only if listed as Allowed Secondary above\n" +
    '- If Must Include Area Served is Yes — add areaServed: "Surrey and South East England"\n' +
    "- Material entity (" + material + ") must appear in name field\n" +
    "- All URLs must start with https://www.abbeyfloorcare.co.uk\n" +
    "- No bracket-wrapped URLs\n" +
    "- No HTML-encoded characters (%22, %3A etc)\n\n" +
    "--- GENERATE NOW ---"
  );
}



/* ============================================================
   SCHEMA-ONLY AUDIT PROMPT GENERATOR
   Called by W5c "Generate Schema Audit Prompt" button.
   Receives the full <script> wrapped block pasted from W5b.
   Strips wrapper to extract JSON, then builds audit prompt.
   Reads alignment rules from sheet.
============================================================ */
function getSchemaAuditPrompt(schemaBlock) {
  // Strip <script> wrapper if present to get raw JSON
  var schemaJson = schemaBlock
    .replace(/<script[^>]*>/gi, '')
    .replace(/<\/script>/gi, '')
    .trim();
  // Strip embedded newlines inside JSON string values
  schemaJson = schemaJson.replace(/(\r\n|\n|\r)\s*/g, ' ').trim();
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const d           = getActiveRowDataMap();
  const postID      = d["Post ID"]      || "Unknown";
  const material    = d["Stone Type"]   || "Unknown";
  const articleType = d["Article Type"] || "General";
  const url         = d["URL"]          || "";

  // Read Schema Alignment Validation Sheet
  const savSheet = ss.getSheetByName("Schema Alignment Validation Sheet");
  let schemaRules = "Schema Alignment Validation Sheet not found — validate against Article schema defaults.";

  if (savSheet) {
    const savData    = savSheet.getDataRange().getValues();
    const savHeaders = savData[0].map(function(h) { return String(h).trim(); });
    const artTypeIdx = savHeaders.indexOf("Article Type");
    let matchRow = null;
    for (let i = 1; i < savData.length; i++) {
      if (String(savData[i][artTypeIdx]).trim().toLowerCase() === articleType.toLowerCase()) {
        matchRow = savData[i];
        break;
      }
    }
    if (matchRow) {
      const get = function(col) {
        const idx = savHeaders.indexOf(col);
        return idx > -1 ? String(matchRow[idx] || "").trim() : "Not specified";
      };
      schemaRules =
        "Allowed Primary Schema Type: "      + get("Allowed Primary Schema Type")      + "\n" +
        "Allowed Secondary Schema Types: "   + get("Allowed Secondary Schema Types")   + "\n" +
        "Prohibited Schema Types: "          + get("Prohibited Schema Types")           + "\n" +
        "Multiple Primary Schema Allowed: "  + get("Multiple Primary Schema Allowed")  + "\n" +
        "Primary Entity Required In Schema: "+ get("Primary Entity Required In Schema") + "\n" +
        "Must Include Area Served: "         + get("Must Include Area Served")          + "\n" +
        "Schema Type Change Allowed: "       + get("Schema Type Change Allowed")        + "\n" +
        "Schema Expansion Allowed: "         + get("Schema Expansion Allowed")          + "\n" +
        "Maximum Allowed Schema Risk: "      + get("Maximum Allowed Schema Risk")       + "\n" +
        "Assigned Schema Risk: "             + get("Assigned Schema Risk");
    } else {
      schemaRules = "No match for article type '" + articleType + "' in Schema Alignment Validation Sheet.";
    }
  }

  return `SCHEMA SELF-AUDIT
ROLE: Senior UK SEO & Stone Restoration Strategist
TASK: Audit the schema JSON below against the rules provided. This is the schema you generated — audit your own output.

--- PAGE CONTEXT ---
POST ID: ${postID}
URL: ${url}
MATERIAL: ${material}
ARTICLE TYPE: ${articleType}

--- SCHEMA ALIGNMENT RULES (HARD LOCK) ---
${schemaRules}

--- SCHEMA FORMAT RULES ---
- Output must be a complete HTML block: <script type="application/ld+json">{json}</script>
- JSON must be minified — single line, no internal line breaks
- No string value may contain a line break or newline character
- @context must be exactly "https://schema.org" — no brackets, no encoding
- @type must comply with Allowed Primary Schema Type above
- All URLs must start with https://www.abbeyfloorcare.co.uk
- No bracket-wrapped URLs — never "[https://..."
- No HTML-encoded characters (%22, %3A, %2F etc)
- Must include BreadcrumbList
- Material entity (${material}) must appear in name field
- If Must Include Area Served is Yes — areaServed must be "Surrey and South East England"

--- SCHEMA TO AUDIT ---
${schemaJson}

--- OUTPUT INSTRUCTION ---
Audit every rule above against the schema provided.
If ALL rules pass: respond with exactly "SCHEMA AUDIT: PASS" then stop.
If ANY rule fails: list each failure briefly, then return the complete corrected schema as a full <script type="application/ld+json">...</script> block — minified JSON, no markdown fences.
`.trim();
}

/* ============================================================
   STAGE 2B — RE-HYDRATION
   Replaces [REF:PostID] tokens with real URLs.
   Returns clean HTML only — no TDF row, no column structure.
============================================================ */

function generateFinalTDFRowFromSidebarGPT(stage2HTML) {
  if (!stage2HTML || stage2HTML.trim() === "") return "ERROR: Paste Stage 2B HTML first.";

  let rehydratedHTML = stage2HTML.replace(/REF:(\d+)/g, function(match, id) {
    const realUrl = getUrlByPostID(id);
    return realUrl ? realUrl : match;
  });

  return rehydratedHTML.replace(/\r?\n|\r/g, " ").replace(/\s+/g, " ").trim();
}



function buildCompetitorVocabBlock(vocabTerms, wordCount) {
  if (!vocabTerms && !wordCount) return "";

  var block = "";
  block += "------------------------------------------------------------\n";
  block += "COMPETITOR VOCABULARY & LENGTH BENCHMARK (Hard Lock)\n";
  block += "------------------------------------------------------------\n";
  block += "The following data is derived from the top 10 competitor pages\n";
  block += "for this article's primary search term.\n\n";

  if (wordCount) {
    block += "COMPETITOR AVERAGE WORD COUNT: " + wordCount + " words\n";
    block += "Competitor average: " + wordCount + " words.\n";
    block += "RULE: The Tier Structural Coverage Matrix word budgets per section are the minimum floor — never go below them.\n";
    block += "The competitor average is a ceiling guide only — write toward it where content depth justifies it.\n";
    block += "Do NOT pad sections to reach this number. Do NOT cut governed content to stay below it.\n\n";
  }

  if (vocabTerms && vocabTerms.length > 0) {
    block += "COMPETITOR VOCABULARY GAP TERMS:\n";
    block += "These terms appear on 3 or more competitor pages but are not in the governed entity list.\n";
    block += "Include them naturally where contextually relevant — do not force.\n\n";
    block += vocabTerms.join(", ") + "\n\n";
    block += "COVERAGE RULE:\n";
    block += "Aim to include at least 60% of these terms naturally across the article.\n";
    block += "They represent the vocabulary Google associates with authoritative pages on this topic.\n";
  }

  block += "------------------------------------------------------------\n";
  return block;
}