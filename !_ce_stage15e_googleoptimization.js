/**
 * ================================================================================
 * ce_Stage15E_GoogleOptimization.gs - STAGE 1.5E — GOOGLE OPTIMIZATION LAYER (SIMPLIFIED - DATA ONLY)
 * ================================================================================
 * 
 * Purpose: Generate Google optimization instructions without touching section plan
 * 
 * Reads:  Col 102 (CX) — W1.5D Final Plan
 * Saves:  Col 103 (CY) — W1.5E Optimization JSON (via saveStage15EOptimizations)
 * 
 * Design Philosophy:
 * - W1.5E never sees Internal link fields
 * - Outputs minimal structured data (not full section plan)
 * - W2B merges Col 102 (structure + links) + Col 103 (optimization)
 * 
 * Part of Abbey Floor Care Content Pipeline v77+
 * Updated: March 2026
 * ================================================================================
 */

function buildStage15EPrompt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('posts');
  const row = sheet.getActiveRange().getRow();
  if (row < 2) return 'ERROR: Select a data row (row 2 or below).';

  const d = getActiveRowDataMap();
  
  // Read approved section plan from W1.5D (Column CX — 102)
  const sectionPlanCol = 102;
  const sectionPlan = String(sheet.getRange(row, sectionPlanCol).getValue() || "").trim();
  
  if (!sectionPlan || sectionPlan.length === 0) {
    return 'ERROR: No section plan found in Column CX. Run W1.5D first.';
  }
  
  // Validate section plan has expected structure
  if (!sectionPlan.includes('SECTION 1:')) {
    return 'ERROR: Section plan format invalid. Expected "SECTION 1:" but not found.';
  }
  
  const sectionCount = (sectionPlan.match(/SECTION \d+:/g) || []).length;
  const expectedSectionNumbers = [];

    for (let i = 1; i <= sectionCount; i++) {
      expectedSectionNumbers.push(i);
    }
  if (sectionCount < 3) {
    return 'ERROR: Section plan has only ' + sectionCount + ' sections. Expected at least 3.';
  }

  const articleType = d["Article Type"] || "General";
  const primaryEntity = d["Primary Entity"] || "Unknown";
  const material = d["Stone Type"] || "Unknown";

  // Fetch material behaviour reference from TSM (Common Issues / Diagnosis section)
  // so information_gain_directive claims can be checked against actual wear/failure mechanisms.
  let materialBehaviourBlock = "";
  try {
    const techSheet = bc_getSheetByConfigKeys_(["techSpec"]);
    if (techSheet) {
      const data = techSheet.getDataRange().getValues();
      const materialLower = material.toLowerCase().trim();
      let matchRow = null;
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][1] || "").trim().toLowerCase() === materialLower) {
          matchRow = data[i];
          break;
        }
      }
      if (matchRow) {
        const fullTSM = String(matchRow[2] || "").trim();
        if (fullTSM) {
          materialBehaviourBlock = fullTSM;
        }
      }
    }
  } catch(e) {
    materialBehaviourBlock = "";
  }

  const prompt = `

  STAGE 1.5E — GOOGLE OPTIMIZATION LAYER (DATA ONLY)

  You are a Google search optimization specialist focusing on user engagement signals, featured snippet optimization, and natural language query matching.

  CRITICAL INSTRUCTION:
  You will receive a section plan below.
  DO NOT output the section plan.
  DO NOT output H2 headings, word budgets, content briefs, TSM requirements, or internal links.
  ONLY output optimization instructions in VALID JSON format.

  GOVERNANCE PRESERVATION GATE — HARD LOCK

  W1.5E is an optimization layer only. It MUST NOT change what a section is fundamentally about.

  For every opening_anchor, information_gain_directive, scannability_pattern, featured_snippet_target, list_opportunity and blockquote_flag:

  1. Preserve the section's TSM Requirement and Content Brief from W1.5D.
  2. Preserve the Primary Entity and Article Type.
  3. Obey the Page Rewrite Brief, especially RE-ANCHOR FOCUS and SCOPE BOUNDARY.
  4. Do not promote a secondary or adjacent intent into the section's main task.
  5. Do not convert:
    - contractor evaluation into a cleaning process,
    - diagnosis into a repair tutorial,
    - repair into routine maintenance,
    - service selection into a technical treatment workflow,
    - a local service page into general material education.
  6. Technical entities may support the governed section role, but must not redefine it.
  7. Do not introduce a new material, treatment, service, locality, customer/project narrative or factual claim that is not already supported by the supplied W1.5D plan.
  8. If an optimization would require changing the section's governed role, do not make that optimization.

  FINAL SELF-CHECK:
  Before returning the JSON, compare each proposed optimization with the corresponding W1.5D section.
  If the reader task, section role, material, locality or scope has changed, rewrite the optimization until the original governed task is preserved.

  YOUR TASK:
  For each numbered SECTION in the W1.5D plan below, generate optimization instructions as a JSON object.

  SECTION NUMBER LOCK — HARD RULE:
  - W1.5D contains exactly ${sectionCount} governed sections.
  - The ONLY permitted section_number values are: ${expectedSectionNumbers.join(', ')}.
  - Return exactly ${sectionCount} objects in the "sections" array.
  - Return exactly ONE object for each permitted section number.
  - Do NOT create SECTION 0.
  - Do NOT create an object for the header, introduction, metadata, notes, or any unnumbered content.
  - Do NOT omit, duplicate, renumber, merge or add sections.

  OUTPUT FORMAT REQUIREMENT:
  You MUST output ONLY valid JSON — no preamble, no explanation, no markdown code fences.
  Start your response with { and end with }

    JSON STRUCTURE:

        {
      "sections": [
        {
          "section_number": 1,
          "opening_anchor": "[section-specific opening anchor]",
          "information_gain_directive": "[section-specific practical observation or decision aid]",
          "scannability_pattern": "[section-specific formatting instruction or None]",
          "featured_snippet_target": "Yes or No",
          "list_opportunity": "[specific list opportunity or None]",
          "blockquote_flag": "[Yes — suggested quote, or No]"
        }
      ]
    }

    IMPORTANT:
    The text inside square brackets is a schema placeholder only.
    Do NOT copy or paraphrase the placeholder wording.
    Generate every field from the actual SECTION PLAN supplied below.

  FIELD DEFINITIONS:

  1. opening_anchor (string):
    Generate a varied section opener using the FOUR-WAY PIVOT rotation below.
    The opener must be specific to this section's topic and the material.

    FIRST-HAND EXPERIENCE RULE (Hard Lock):

        Every opening_anchor must include at least one first-hand observational cue appropriate to the section.

        The cue may be:
        - visual;
        - tactile;
        - sensory;
        - or hands-on.

        Choose the type of cue that naturally fits the section.

        Do NOT force touch-based language where visual inspection is the more appropriate way for a homeowner to recognise the condition.

        Examples of suitable cues include:
        - "under raking light"
        - "compare the grout with the surrounding tile"
        - "look across the surface from a low angle"
        - "run your hand over"
        - "feels rough underfoot"
        - "catches a fingernail"

        The purpose is to make the opener feel grounded in real inspection or use of the floor, not to require physical touching in every section.

        WRONG:
        "Run your fingertips across the grout joint..." where the section is primarily about visible discolouration.

        BETTER:
        "Compare the grout joints with the surrounding tile under the same light..."

        The observational cue must support the actual section purpose and must not introduce a diagnostic test that the section does not require.

    ROTATION RULE — no two consecutive sections may use the same style:

    Style A — CONDITIONAL (diagnostic): "If your ${material.toLowerCase()} shows [symptom]..."
    Use for: sections where reader recognition of a visible symptom is the priority.
    Example: "If your ${material.toLowerCase()} shows dull patches that will not shift after cleaning..."

    Style B — FACT-FIRST: Open with the physical reality or named condition.
    Use for: mechanism explanations, cause-and-effect sections.
    Example: "Dull patches that refuse to shift after cleaning are almost never dirt — they are physical change in the stone surface."

    Style C — EXPERT WARNING: Open with the consequence of the wrong action.
    Use for: sections where misdiagnosis or incorrect treatment is a common mistake.
    Example: "Repeatedly scrubbing a patchy ${material.toLowerCase()} floor often makes the condition worse, not better."

    Style D — NOUN-FIRST: Lead with the technical entity to anchor authority.
    Use for: mechanism sections where the entity name carries SEO weight.
    Example: "Capillary action is the reason spills appear to grow even after wiping up — liquids are drawn sideways through crystal pathways."

    SECTION 1: May use any style — Conditional is appropriate if the article opens diagnostically.
    SECTIONS 2 onward: Must vary. Never use Conditional for two consecutive sections.
    HEADER: Never use a Conditional opener — headers must open with a direct statement or observation.
    ROUTING SECTIONS: Sections whose sole purpose is routing the reader to other pages must use Style B or Style C only. Style A is prohibited on routing sections — the reader is not diagnosing a symptom. Style D is also inappropriate for pure routing sections.

  2. information_gain_directive (string):

      Provide ONE useful, specific practical observation or decision aid for this section.

      It must:
      - add something useful beyond simply repeating the Content Brief;
      - be supported by the supplied section plan and material governance;
      - help the reader recognise, compare, interpret or respond to the condition more accurately.

      Prefer genuine information gain, but NEVER invent:
      - a new diagnostic test;
      - a new wear or failure mechanism;
      - an unsupported cause;
      - a treatment rule or treatment sequence;
      - a diagnostic test based on applying a product or chemical;
      - a cleaning chemistry decision tree or product-selection pathway unless the SECTION PLAN explicitly requires it;
      - a recommendation to try progressively stronger cleaning products;
      - a fixed maintenance, resealing or renewal schedule unless the SECTION PLAN explicitly provides one;
      - a wetting, darkening or absorption test to decide whether sealing is needed;
      - a physical behaviour;
      - a measurement requirement or threshold;
      - or a material response

      merely to make the advice sound more expert or non-obvious.

      If the supplied evidence does not support a genuinely new technical insight, give a useful observational or comparison instruction instead.

      Example:
      "Compare the affected area with an undisturbed part of the same floor under the same lighting before deciding the surface itself has changed."

      WRONG:
      Inventing a tactile test, chemical reaction, moisture behaviour or failure mechanism that is not supported by the supplied material governance.

      This directive is internal guidance for the writer. It does NOT need to be quoted verbatim in the final article.
      HARD RULE:
        The information_gain_directive and scannability_pattern must NOT prescribe, sequence, compare or recommend cleaning chemicals, chemical classes, pH types, alkaline treatment, acidic treatment, stripping agents or product-selection logic unless that exact treatment sequence is explicitly stated in the SECTION PLAN.

        Material reference data may be used to prevent factual errors, but must NOT be promoted into new treatment instructions that are absent from the SECTION PLAN.

  3. scannability_pattern (string):
    Suggest where to use comparison structures, if/then patterns, or diagnostic sequences.
    Example: "Use 'If polished, it resists stains; if honed, it needs more care' comparison"
    Write "None" if section doesn't need special formatting.

  4. featured_snippet_target (string):

      Write "Yes" only when the SECTION PLAN clearly supports a concise diagnostic sequence, step-by-step process, checklist, or direct comparison that could stand alone as a useful answer.

      Do NOT mark a section "Yes" merely because it contains a list.

      If the section is mainly explanatory, narrative, routing, or advisory, write "No".
      HARD RULE: If the section's primary purpose is routing the reader to another guide, page or topic, featured_snippet_target MUST be "No".
      HARD RULE: Maintenance, long-term care, general advisory and routine upkeep sections MUST use "No" unless the SECTION PLAN explicitly defines a step-by-step procedure, diagnostic test or checklist as the section's primary purpose.

  5. list_opportunity (string):
    When List eligible = "YES", specify what type of list would work:
    - "Diagnostic steps for identifying [problem]"
    - "Symptom checklist for [condition]"
    - "Comparison points between [X] and [Y]"
    Write "None" if no list needed or List eligible = "NO"

  6. blockquote_flag (string):

      Use a blockquote only when the SECTION PLAN itself clearly contains a concise insight, mechanism or warning that benefits from visual emphasis.

      Do NOT create a new technical claim just to justify a blockquote.

      If Visual Pattern = "Mechanism-Blockquote" and the section plan contains a clearly supported statement suitable for emphasis:
      Write "Yes — [suggested quote under 20 words]"

      Otherwise write:
      "No"

      The suggested quote must:
      - stay within the supplied section purpose;
      - use only supported material behaviour;
      - avoid absolute claims unless the source explicitly supports them;
      - avoid introducing new terminology or mechanisms.

  CONSTRAINTS:

  ❌ DO NOT suggest new entity names or technical terminology
  ❌ DO NOT add academic or scientific language  
  ❌ DO NOT include H2 headings, TSM requirements, word budgets, content briefs in your output
  ❌ DO NOT include images, entities, or internal links in your output
  ❌ DO NOT include markdown code fences (\`\`\`json) around your JSON
  ❌ DO NOT include any preamble or explanation text
  ❌ DO NOT let information_gain_directive restate the Content Brief — it must add a
    genuinely new, specific detail
  ❌ DO NOT turn MATERIAL BEHAVIOUR REFERENCE data into new cleaning methods, chemical sequences, product choices or treatment instructions unless those exact instructions already appear in the SECTION PLAN

  ✅ DO output ONLY valid JSON starting with { and ending with }
  ✅ DO include all sections in the JSON a✅ DO include exactly ${sectionCount} section objects in the JSON array
  ✅ DO use only these section_number values: ${expectedSectionNumbers.join(', ')}
  ✅ DO NOT include SECTION 0, headers, introductions, metadata, notes, or unnumbered contentrray
  ✅ DO provide specific, actionable instructions in each field
  ✅ DO use "None" or "No" where appropriate (these are valid string values)
  ✅ DO include at least one appropriate first-hand observational cue in every opening_anchor — visual, tactile, sensory or hands-on
  ✅ DO choose the cue that naturally fits the section; do not force touch-based language

  MATERIAL CONTEXT:
  Material: ${material}
  Article Type: ${articleType}
  Primary Entity: ${primaryEntity}

  --- MATERIAL BEHAVIOUR REFERENCE (Hard Lock — factual ground truth) ---
  ${materialBehaviourBlock || "No material behaviour data available for this material — do not invent wear, failure, or damage mechanisms. Keep information_gain_directive claims generic and verifiable rather than asserting a specific cause."}

  PER-SECTION FACTUAL SELF-CHECK (Hard Lock — run before writing each information_gain_directive):
  Before writing the information_gain_directive for EACH section, check any factual claim
  you are about to make — about where damage appears, what causes wear, which areas fail
  first, or how a defect progresses — against the MATERIAL BEHAVIOUR REFERENCE above.
  If your claim contradicts or is unsupported by the reference (for example, claiming edges,
  static furniture zones, or thresholds show wear before high-traffic areas, when the
  reference states wear is traffic- and abrasion-driven) — do NOT write that claim.
  Rewrite the directive using a mechanism the reference actually supports, or fall back to
  advice that does not depend on an unverified location/cause claim.
  This check applies independently to every section — passing it for one section does not
  exempt any other section's information_gain_directive from the same check.
  Do not include this check or any reasoning about it in your output. Only output the final JSON.

  --- SECTION PLAN (extract optimization needs only) ---
  ${sectionPlan}

  --- OUTPUT INSTRUCTION ---
  Output a valid JSON object containing optimization instructions for all sections.

  CRITICAL: Output ONLY the JSON object.
  Do NOT include:
  - Any text before the opening {
  - Any text after the closing }
  - Markdown code fences like \`\`\`json
  - Explanations or preambles

  Start your response with { and end with }
  The JSON must be parseable by JSON.parse()
  `.trim();

  return prompt;
}