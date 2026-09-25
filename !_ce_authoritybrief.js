/* ============================================================
   ce_AuthorityBrief
   W-1 — AUTHORITY BRIEF GENERATOR
   Reads entity analysis columns from the active posts row,
   builds a concise governing brief for the rewrite pipeline,
   and stores it in col DN (118) — "Authority Brief".

   The brief is injected into Stage 1A as the first governing
   block — distilled from 24 entity analysis columns into a
   compact ~300-400 word instruction the LLM can reliably hold.

   Column sources (by header name):
     Stone Type              col G (7)
     Article Type            col H (8)
     Primary Entity          col I (9)
     Supporting Entities Core  col L (12)
     Peripheral Entities Link Out  col M (13)
     Feeds Hub               col N (14)
     Page Governance Summary col O (15)
     Entity Governance Status  col P (16)
     Observed Query Cluster  col T (20)
     GSC Intent Evidence     col U (21)
     Primary Query Cluster Owned  col V (22)
     Drift Status            col W (23)
     Rewrite Status          col X (24)
     Rewrite Governance Summary  col Y (25)
     Cannibalisation Guardrail  col Z (26)
     Safe Handoff Pages      col AA (27)
     Rewrite Role Lock       col AB (28)
     Page Rewrite Brief      col AC (29)
     Rewrite Brief Date      col AD (30)
     Authority Brief         col DN (118)
============================================================ */

function buildAuthorityBrief(targetRow) {
  try {
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const sheet   = ss.getSheetByName("posts");
    const row = targetRow || sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "Select a data row first." };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
                         .getValues()[0]
                         .map(function(h) { return String(h).trim(); });

    const rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

    function getCol(name) {
      var idx = headers.indexOf(name);
      return idx > -1 ? String(rowValues[idx] || "").trim() : "";
    }

    // Read entity analysis columns
    const stoneType          = getCol("Stone Type");
    const articleType        = getCol("Article Type");
    const primaryEntity      = getCol("Primary Entity");
    const supportingEntities = getCol("Supporting Entities Core");
    const peripheralEntities = getCol("Peripheral Entities Link Out");
    const feedsHub           = getCol("Feeds Hub");
    const govSummary         = getCol("Page Governance Summary");
    const entityStatus       = getCol("Entity Governance Status");
    const observedCluster    = getCol("Observed Query Cluster");
    const gscEvidence        = getCol("GSC Intent Evidence");
    const primaryCluster     = getCol("Primary Query Cluster Owned");
    const driftStatus        = getCol("Drift Status");
    const rewriteStatus      = getCol("Rewrite Status");
    const rewriteRoleLock    = getCol("Rewrite Role Lock");
    const rewriteGovSummary  = getCol("Rewrite Governance Summary");
    const cannibGuardrail    = getCol("Cannibalisation Guardrail");
    const safeHandoff        = getCol("Safe Handoff Pages");
    const rewriteBrief       = getCol("Page Rewrite Brief");
    const rewriteBriefDate   = getCol("Rewrite Brief Date");
    const pageUrl            = getCol("URL") || getCol("Canonical URL");
    const postTitle          = getCol("Title");

    // Validate minimum required fields
    if (!articleType) return { success: false, message: "Article Type column empty — run entity analysis first." };
    if (!primaryEntity) return { success: false, message: "Primary Entity column empty — run entity analysis first." };

    // Build the authority brief
    var brief = "";
    brief += "Generated: " + new Date().toLocaleDateString("en-GB") + "\n";
    // Format rewrite brief date cleanly — strip GMT timezone string if raw JS Date
    var rbd = String(rewriteBriefDate || "").trim();
    if (rbd) {
      // If it's a raw JS Date string, extract DD/MM/YYYY
      var rbdMatch = rbd.match(/(\d{1,2})\/((\d{1,2})\/(\d{2,4}))/);
      if (!rbdMatch) {
        // Try parsing as Date object string e.g. "Tue Mar 17 2026..."
        var rbdDate = new Date(rbd);
        if (!isNaN(rbdDate.getTime())) {
          rbd = ("0" + rbdDate.getDate()).slice(-2) + "/" +
                ("0" + (rbdDate.getMonth()+1)).slice(-2) + "/" +
                rbdDate.getFullYear();
        }
      }
      brief += "Rewrite Brief Date: " + rbd + "\n";
    }
    brief += "\n";

    const locality   = getCol("Locality");
    const parentArea = getCol("Parent Area");

    brief += "PAGE IDENTITY\n";
    if (postTitle)  brief += "Title: " + postTitle + "\n";
    if (pageUrl)    brief += "URL: " + pageUrl + "\n";
    brief += "Material: " + stoneType + "\n";
    brief += "Article Type: " + articleType + "\n";
    brief += "Primary Entity: " + primaryEntity + "\n";
    if (locality)   brief += "Locality: " + locality + "\n";
    if (parentArea) brief += "Parent Area: " + parentArea + "\n";
    brief += "\n";

    brief += "QUERY AUTHORITY\n";
    if (primaryCluster)   brief += "Primary Query Cluster: " + primaryCluster + "\n";
    if (observedCluster)  brief += "Observed Query Cluster: " + observedCluster + "\n";
    if (gscEvidence)      brief += "GSC Signal: " + gscEvidence + "\n";
    brief += "\n";

    // Read stored Problem Angle
    var problemAngle = getCol("Problem Angle");
    if (problemAngle) {
      brief += "\n";
      brief += "PROBLEM ANGLE\n";
      brief += "------------------------------------------------------------\n";
      brief += "Selected angle: " + problemAngle + "\n";
      brief += "Opening sentence instruction: The article must open by addressing the specific problem " +
               "identified in the Primary Query Cluster above from this angle.\n";
      brief += "------------------------------------------------------------\n";
    }

    brief += "ENTITY COVERAGE REQUIRED\n";
    if (supportingEntities) brief += "Core Entities: " + supportingEntities + "\n";
    if (peripheralEntities) brief += "Peripheral Entities: " + peripheralEntities + "\n";
    if (feedsHub)           brief += "Feeds Hub: " + feedsHub + "\n";
    brief += "\n";
    
    brief += "SEMANTIC FINGERPRINT DEPLOYMENT RULES\n";
    brief += "------------------------------------------------------------\n";
    brief += "For each Core Entity listed above:\n";
    brief += "1. Create ONE dedicated focused paragraph explaining the entity mechanism\n";
    brief += "2. Use the COMPLETE semantic fingerprint vocabulary for that entity\n";
    brief += "3. Deploy all fingerprint terms as a cluster within 150 words\n";
    brief += "4. Do NOT scatter 2-3 terms across multiple sections\n";
    brief += "5. Show how the entity connects to the reader's observed problem\n";
    brief += "\n";
    brief += "ENTITY NAME REPETITION LIMIT (CRITICAL):\n";
    brief += "- Entity names may appear 1-2 times MAXIMUM in the entire article\n";
    brief += "- ALL other references use semantic fingerprint vocabulary instead\n";
    brief += "- Example: 'Efflorescence' appears once; all other references use:\n";
    brief += "  'white crystalline deposits', 'salt migration', 'dissolved minerals'\n";
    brief += "\n";
    brief += "VOCABULARY DISTRIBUTION PATTERN (MULTINOMIAL SCORING):\n";
    brief += "Target 8-12 instances of semantic fingerprint vocabulary per Core Entity:\n";
    brief += "- Opening section: 2-3 instances\n";
    brief += "- Main content sections: 3-5 instances\n";
    brief += "- Conclusion/outcome sections: 2-3 instances\n";
    brief += "Distribution strengthens topical authority through term frequency signals.\n";
    brief += "------------------------------------------------------------\n";
    brief += "\n";
    
    brief += "ENTITY-SECTION ARCHITECTURE\n";
    brief += "------------------------------------------------------------\n";
    brief += "Structure the article by anchoring sections to Core Entities:\n";
    brief += "- Each Core Entity should anchor ONE major section or subsection\n";
    brief += "- Section heading should reflect the entity's reader-facing problem\n";
    brief += "- Section content explains the entity mechanism using complete fingerprint\n";
    brief += "- Show entity relationships: how entities interact to create symptoms\n";
    brief += "Example: If Core Entities are 'Unglazed Clay Porosity' and 'Efflorescence',\n";
    brief += "create sections showing how porosity enables salt movement (entity relationship)\n";
    brief += "------------------------------------------------------------\n";
    brief += "\n";
    
    brief += "ENTITY RELATIONSHIP MAPPING (CRITICAL)\n";
    brief += "------------------------------------------------------------\n";
    brief += "Do NOT treat Core Entities as isolated topics.\n";
    brief += "REQUIRED: Explain how entities interact mechanically to create symptoms.\n";
    brief += "Build causal chains showing entity relationships:\n";
    brief += "- Which entity creates the condition for another entity to act\n";
    brief += "- How multiple entities combine to produce the reader's observed problem\n";
    brief += "- Why addressing one entity without understanding its relationship fails\n";
    brief += "Example causal chain:\n";
    brief += "'Efflorescence appears when [Unglazed Clay Porosity] allows [moisture movement]\n";
    brief += "in the absence of [DPM], carrying [dissolved salts] that crystallize at surface.'\n";
    brief += "This shows 4 entities working as one connected system.\n";
    brief += "------------------------------------------------------------\n";
    brief += "\n";

    brief += "MAXIMUM MARGIN SEPARATION (CRITICAL)\n";
    brief += "------------------------------------------------------------\n";
    brief += "Content must maintain clear separation from unrelated intents while allowing\n";
    brief += "cross-material TSM entities that apply to this material.\n";
    brief += "\n";
    brief += "ALLOWED CROSS-MATERIAL CONTENT:\n";
    brief += "- Entities already in THIS material's TSM sheet (check the Recognised Technical\n";
    brief += "  Entities section provided in your context)\n";
    brief += "- Entities from other materials' TSMs that legitimately apply to THIS material's\n";
    brief += "  chemistry, structure, or restoration needs (shared porous masonry behaviour,\n";
    brief += "  universal cleaning chemistry, common tool types)\n";
    brief += "- Universal restoration principles (alkaline cleaning, mechanical extraction,\n";
    brief += "  moisture management, pH neutralisation)\n";
    brief += "\n";
    brief += "PROHIBITED (creates noisy features that blur material boundaries):\n";
    brief += "- Entities incompatible with THIS material's chemistry (acid cleaning on limestone,\n";
    brief += "  mechanical polishing on ceramic glaze, bleach on cementitious grout)\n";
    brief += "- Processes that don't apply to THIS material's structure (honing unglazed clay,\n";
    brief += "  sealing non-porous vitrified tiles, waxing moisture-active floors without DPM)\n";
    brief += "- Methods from unrelated material categories (wood finishing in stone articles,\n";
    brief += "  vinyl installation in tile articles, carpet cleaning in masonry)\n";
    brief += "\n";
    brief += "VALIDATION FOR CROSS-MATERIAL ENTITIES:\n";
    brief += "When considering an entity that appears in another material's TSM:\n";
    brief += "\n";
    brief += "1. CHEMISTRY COMPATIBILITY CHECK:\n";
    brief += "   Does this material share the chemical constraint?\n";
    brief += "   ✅ YES: Efflorescence (applies to all porous masonry - Quarry, Victorian, Sandstone)\n";
    brief += "   ✅ YES: Acid sensitivity (applies to all calcium-carbonate - Marble, Limestone, Travertine)\n";
    brief += "   ❌ NO: Mechanical polishing (applies to marble, NOT to ceramic glaze)\n";
    brief += "\n";
    brief += "2. STRUCTURAL COMPATIBILITY CHECK:\n";
    brief += "   Does this material share the structural behaviour?\n";
    brief += "   ✅ YES: Breathability requirement (all ground-bearing porous floors without DPM)\n";
    brief += "   ✅ YES: Grout porosity (all cementitious grout in tiled installations)\n";
    brief += "   ❌ NO: Surface cap fragility (travertine-specific, NOT limestone)\n";
    brief += "\n";
    brief += "3. MATERIAL-SPECIFIC CONTEXT REQUIRED:\n";
    brief += "   Even when an entity legitimately applies, explain WHY it matters for THIS material:\n";
    brief += "   - Don't just say 'pH-neutral cleaner prevents damage'\n";
    brief += "   - Say 'pH-neutral cleaner prevents alkaline residue cycling in quarry tile's open\n";
    brief += "     clay pores, where residues accelerate abrasion'\n";
    brief += "\n";
    brief += "NEW ENTITY DISCOVERY:\n";
    brief += "If you identify a technical entity during article generation that is NOT in the\n";
    brief += "provided TSM entity list but IS material-specific and technically accurate:\n";
    brief += "- Use it with proper semantic fingerprint vocabulary (8-12 surrounding technical terms)\n";
    brief += "- Ensure it describes THIS material's specific behaviour, not generic flooring\n";
    brief += "- The entity will be reviewed and added to the TSM during post-generation audit\n";
    brief += "\n";
    brief += "EXAMPLE - CORRECT (cross-material entity with compatibility validation):\n";
    brief += "'White crystalline deposits signal moisture movement through quarry tile's interconnected\n";
    brief += "clay pores - a diagnostic pattern called efflorescence that requires breathable sealing\n";
    brief += "to manage rather than film-forming coatings that trap the moisture below.'\n";
    brief += "(Efflorescence applies to quarry tile because: porous structure + ground-bearing + no DPM)\n";
    brief += "\n";
    brief += "EXAMPLE - WRONG (incompatible cross-material entity):\n";
    brief += "'Hone the surface with progressive diamond pads to remove scratches and restore clarity.'\n";
    brief += "(Diamond honing applies to marble/limestone, NOT to quarry tile - unglazed clay is too\n";
    brief += "soft and abrasion removes integral pigment irreversibly)\n";
    brief += "------------------------------------------------------------\n";
    brief += "\n";

    brief += "GOVERNANCE CONSTRAINTS\n";
    if (govSummary)       brief += "Governance Summary: " + govSummary + "\n";
    if (entityStatus)     brief += "Entity Status: " + entityStatus + "\n";
    if (rewriteRoleLock)  brief += "Role Lock: " + rewriteRoleLock + "\n";
    if (driftStatus)      brief += "Drift Status: " + driftStatus + "\n";
    if (rewriteStatus)    brief += "Rewrite Status: " + rewriteStatus + "\n";
    brief += "\n";

    brief += "REWRITE INSTRUCTION\n";
    if (rewriteGovSummary) brief += "Governance Summary: " + rewriteGovSummary + "\n";
    if (rewriteBrief)      brief += "Rewrite Brief: " + rewriteBrief + "\n";
    brief += "\n";

    brief += "SCOPE BOUNDARIES\n";
    if (cannibGuardrail) brief += "Cannibalisation Guardrail: " + cannibGuardrail + "\n";
    if (safeHandoff && safeHandoff !== pageUrl)     brief += "Safe Handoff Pages: " + safeHandoff + "\n";
    brief += "\n";

    // - Problem Angle section
    var existingAngleStmt = getCol("Problem Angle");
    var existingAngleCode = "";

    if (existingAngleStmt) {
      var angleMatch = existingAngleStmt.match(/^([A-Z]+)\s+—/);
      if (angleMatch) {
        existingAngleCode = angleMatch[1];
      }
    }

    brief += "OBSERVED PROBLEM\n";
    brief += "------------------------------------------------------------\n";
    brief += "Primary problem: " + (getCol("Primary Query Cluster Owned") || "not yet set") + "\n";

    if (existingAngleStmt) {
      brief += "Selected angle: " + (existingAngleCode || "SET") + "\n";
      brief += "Angle statement: " + existingAngleStmt + "\n";
      brief += "Status: ANGLE SET — proceed to W0\n";
    }

    brief += "\n";
  

    // ── GSC Gap Analysis — Benchmark vs Monitor (Pre/Post Migration) ──
    var benmrkClicks  = getCol("Benmrk Clicks");
    var benmrkImpr    = getCol("Benmrk Impr");
    var benmrkPos     = getCol("Benmrk Pos");
    var postIdForAuth = getCol("Post ID");
    var benmrkQueries = postIdForAuth ? bc_getQueriesFromSheet(postIdForAuth, 'GSC Benchmark') : getCol("Benmrk Queries");
    var montrQueries  = postIdForAuth ? bc_getQueriesFromSheet(postIdForAuth, 'GSC Monitor')   : getCol("Montr Queries");
    var montrClicks   = getCol("Montr Clicks");
    var montrImpr     = getCol("Montr Impr");
    var montrPos      = getCol("Montr Position");
    var baselClicks   = getCol("Basel Clicks");
    var baselImpr     = getCol("Basel Impr");

    if (benmrkImpr || montrImpr || baselImpr) {
      brief += "\n";
      brief += "GSC AUTHORITY GAP ANALYSIS (Pre vs Post URL Migration)\n";
      brief += "------------------------------------------------------------\n";

      if (baselImpr || baselClicks) {
        brief += "Baseline: " + (baselClicks || "0") + " clicks | " + (baselImpr || "0") + " impressions\n";
      }
      brief += "Benchmark (pre-migration): " + (benmrkClicks || "0") + " clicks | " + (benmrkImpr || "0") + " impressions" + (benmrkPos ? " | avg pos " + benmrkPos : "") + "\n";
      brief += "Monitor (post-migration):  " + (montrClicks  || "0") + " clicks | " + (montrImpr  || "0") + " impressions" + (montrPos  ? " | avg pos " + montrPos  : "") + "\n";

      var bImpr = parseFloat(String(benmrkImpr || "0").replace(/,/g,"")) || 0;
      var mImpr = parseFloat(String(montrImpr  || "0").replace(/,/g,"")) || 0;
      var bClk  = parseFloat(String(benmrkClicks || "0").replace(/,/g,"")) || 0;
      var mClk  = parseFloat(String(montrClicks  || "0").replace(/,/g,"")) || 0;
      var imprDelta = mImpr - bImpr;
      var clkDelta  = mClk  - bClk;

      if (bImpr > 0 || mImpr > 0) {
        var imprPct = bImpr > 0 ? " (" + (imprDelta >= 0 ? "+" : "") + Math.round((imprDelta/bImpr)*100) + "%)" : "";
        var clkPct  = bClk  > 0 ? " (" + (clkDelta  >= 0 ? "+" : "") + Math.round((clkDelta /bClk )*100) + "%)" : "";
        brief += "Migration impact: impressions " + (imprDelta >= 0 ? "+" : "") + imprDelta + imprPct +
                 " | clicks " + (clkDelta >= 0 ? "+" : "") + clkDelta + clkPct + "\n";
        if (imprDelta < -10 || clkDelta < -5) {
          brief += "Recovery status: RECOVERY REQUIRED — significant authority loss post-migration\n";
          brief += "Pipeline action: Strengthen entity coverage for benchmark query clusters listed below\n";
        } else if (imprDelta < 0 || clkDelta < 0) {
          brief += "Recovery status: GOVERNANCE TIGHTENING RECOMMENDED — minor authority loss post-migration\n";
        } else if (mImpr > 0 && mClk === 0) {
          brief += "Recovery status: CTR IMPROVEMENT REQUIRED — page has impressions but zero clicks\n";
          brief += "Pipeline action: Strengthen title, meta description, and decision-support framing to improve click conversion\n";
        } else if (mImpr > 0 && mClk > 0) {
          var ctr = Math.round((mClk / mImpr) * 1000) / 10;
          if (ctr < 2) {
            brief += "Recovery status: CTR IMPROVEMENT RECOMMENDED — CTR " + ctr + "% is below target\n";
          } else {
            brief += "Recovery status: STABLE — impressions and clicks within expected range\n";
          }
        } else {
          brief += "Recovery status: STABLE — no significant authority loss detected\n";
        }
      }

      if (benmrkQueries) {
        brief += "\nBenchmark query clusters (pre-migration — recover these):\n" + benmrkQueries + "\n";
      }
      if (montrQueries) {
        brief += "\nMonitor query clusters (post-migration — current visibility):\n" + montrQueries + "\n";
      }
      if (benmrkQueries && !montrQueries) {
        brief += "\nCRITICAL: Benchmark queries present but monitor queries absent — complete visibility loss post-migration\n";
      }
    brief += "------------------------------------------------------------\n";
    }

    // 👇 INSERT EXACT COVERAGE QUANTIZATION HERE (Block 3, Point 2) 👇
    
    brief += "\n";
    brief += "EXACT COVERAGE QUANTIZATION - SUMMARY BOX (Featured Snippet Optimization)\n";
    brief += "------------------------------------------------------------\n";
    brief += "The opening diagnostic paragraph must be optimized for featured snippet eligibility.\n";
    brief += "This is the ONLY paragraph in the article with strict quantization requirements.\n";
    brief += "\n";
    brief += "WORD COUNT REQUIREMENT: 45-75 words (hard limit)\n";
    brief += "- Below 45 words: Too thin for featured snippet selection\n";
    brief += "- Above 75 words: Exceeds Google's featured snippet display threshold\n";
    brief += "- Target sweet spot: 55-65 words\n";
    brief += "\n";
    brief += "HIGH-IDF TERM DENSITY: 8-12 technical terms (mandatory)\n";
    brief += "High-IDF terms are material-specific technical vocabulary that distinguish\n";
    brief += "expert content from generic advice. Examples per material:\n";
    brief += "\n";
    brief += "Quarry Tile: unglazed clay porosity, capillary absorption, efflorescence,\n";
    brief += "mineral oxide pigmentation, breathable impregnator, moisture-active substrate\n";
    brief += "\n";
    brief += "Marble: calcium carbonate, acid etching, diamond honing, calcite crystal structure,\n";
    brief += "mechanical refinement, impregnating sealer\n";
    brief += "\n";
    brief += "Travertine: void structure, surface cap, filler collapse, calcium carbonate,\n";
    brief += "alkaline residue, hot water extraction\n";
    brief += "\n";
    brief += "COMPLETE ANSWER STRUCTURE (all three components required):\n";
    brief += "1. PROBLEM: What the homeowner sees (symptom description)\n";
    brief += "2. CAUSE: Why it happens to this material (mechanism)\n";
    brief += "3. SOLUTION: What professional restoration does (intervention category)\n";
    brief += "\n";
    brief += "EXAMPLE - CORRECT Summary Box (Quarry Tile, 62 words, 10 high-IDF terms):\n";
    brief += "\"Quarry tiles absorb soil deeply into their unglazed clay body through capillary\n";
    brief += "action, darkening progressively as old wax and coating layers trap contamination\n";
    brief += "within the porous structure. Professional restoration removes decades of embedded\n";
    brief += "residue through controlled alkaline extraction, then applies breathable impregnating\n";
    brief += "sealers that prevent rapid re-soiling while allowing moisture-active substrates to\n";
    brief += "function as designed.\"\n";
    brief += "\n";
    brief += "High-IDF terms: unglazed clay body, capillary action, porous structure, alkaline\n";
    brief += "extraction, breathable impregnating sealers, moisture-active substrates (10 total)\n";
    brief += "\n";
    brief += "VALIDATION CHECKLIST:\n";
    brief += "Before moving past the Summary Box, confirm:\n";
    brief += "□ Word count between 45-75 (use word counter)\n";
    brief += "□ 8-12 material-specific technical terms present\n";
    brief += "□ Problem + Cause + Solution all included\n";
    brief += "□ No generic flooring language ('looks dirty', 'needs cleaning', 'lasts longer')\n";
    brief += "□ Reads as complete standalone answer (could be extracted as featured snippet)\n";
    brief += "\n";
    brief += "This quantization requirement applies ONLY to the Summary Box.\n";
    brief += "All other article sections follow natural prose flow without word count limits.\n";
    brief += "------------------------------------------------------------\n";
    brief += "\n";
    
    // 👆 END OF EXACT COVERAGE QUANTIZATION 👆

    // Append classification block if axis values exist
    var acEntryCondition     = getCol("ac_entry_condition");
    var acHomeownerPerception = getCol("ac_homeowner_perception");
    var acMaterialBehaviour  = getCol("ac_material_behaviour");
    var acConstraint         = getCol("ac_constraint");
    var acProcessEmphasis    = getCol("ac_process_emphasis");
    var acResultType         = getCol("ac_result_type");
    var acNarrativeArchetype = getCol("ac_narrative_archetype");

    if (acEntryCondition || acNarrativeArchetype) {
      brief += "\n";
      brief += "--- ARTICLE CLASSIFICATION (Narrative Governance) ---\n";
      brief += "These seven values govern the narrative shape of this article.\n";
      brief += "They determine the entry point, structural sequence, and emphasis.\n";
      brief += "All stages must comply with the narrative archetype assigned below.\n\n";
      if (acEntryCondition)      brief += "ENTRY CONDITION: " + acEntryCondition + "\n";
      if (acHomeownerPerception) brief += "HOMEOWNER PERCEPTION: " + acHomeownerPerception + "\n";
      if (acMaterialBehaviour)   brief += "MATERIAL BEHAVIOUR: " + acMaterialBehaviour + "\n";
      if (acConstraint)          brief += "CONSTRAINT: " + acConstraint + "\n";
      if (acProcessEmphasis)     brief += "PROCESS EMPHASIS: " + acProcessEmphasis + "\n";
      if (acResultType)          brief += "RESULT TYPE: " + acResultType + "\n";
      if (acNarrativeArchetype) {
        brief += "\nNARRATIVE ARCHETYPE: " + acNarrativeArchetype + "\n";
        brief += "ARCHETYPE INSTRUCTION:\n";
        if (acNarrativeArchetype === "Diagnostic-first") {
          brief += "Open with the homeowner observation. Follow with the hidden cause. Introduce the risk constraint. Then describe the treatment. Close with the outcome.\n";
        } else if (acNarrativeArchetype === "Material-first") {
          brief += "Open with the floor type behaviour. Explain why normal cleaning failed. Describe the contamination pattern. Cover the controlled restoration. Close with the maintenance outcome.\n";
        } else if (acNarrativeArchetype === "Constraint-first") {
          brief += "Open with what could not safely be done. Explain why aggressive cleaning was avoided. Describe the alternative approach. Cover extraction and drying. Close with the visual recovery.\n";
        } else if (acNarrativeArchetype === "Surface-history-first") {
          brief += "Open with the previous coatings and floor history. Describe the cumulative residue build-up. Explain the visual masking effect. Cover the removal strategy. Close with the recovered appearance.\n";
        } else if (acNarrativeArchetype.toLowerCase().includes("transformation")) {
          brief += "Open with the floor's condition before intervention and the homeowner's feelings about it. Show how the problem was worse than it first appeared. Describe the professional decisions that made recovery possible. Close with the transformation — what the floor became and what that meant to the homeowner.\n";
        } else if (acNarrativeArchetype.toLowerCase().includes("rescue")) {
          brief += "Open with the floor looking beyond saving. Establish what was at stake. Describe how careful assessment revealed recovery was possible. Cover the intervention sequence. Close with the rescued result and what would have been lost without it.\n";
        } else if (acNarrativeArchetype.toLowerCase().includes("preservation")) {
          brief += "Open with the original character of the floor and why it mattered. Describe the risks that threatened it. Explain how every decision protected the original material. Close with the preserved result — improved but unchanged in character.\n";
        } else if (acNarrativeArchetype.toLowerCase().includes("reassurance")) {
          brief += "Open with the homeowner's fear that intervention would cause damage. Establish what careful assessment revealed. Describe how the approach was governed by restraint. Close with the result that proved the concern was addressable without risk.\n";
        } else if (acNarrativeArchetype.toLowerCase().includes("discovery")) {
          brief += "Open with the moment the floor was uncovered or revealed. Describe the condition it was found in. Explain what professional assessment identified as possible. Cover the recovery sequence. Close with the revealed and restored result.\n";
        } else if (acNarrativeArchetype.toLowerCase().includes("recovery")) {
          brief += "Open with the damage caused by previous treatment. Establish what went wrong and why. Describe how the professional approach corrected the earlier failure. Close with the recovered floor and what correct treatment achieves.\n";
        } else if (acNarrativeArchetype.toLowerCase().includes("maintenance")) {
          brief += "Open with the floor that had become impossible to care for. Describe why normal cleaning had stopped working. Explain how professional intervention reset the surface. Close with the floor that is now simple to maintain.\n";
        } else if (acNarrativeArchetype.toLowerCase().includes("sympathetic")) {
          brief += "Open with the period character of the floor and why it required a different approach. Describe the constraints that governed every decision. Explain how each stage respected the original material. Close with the result that improved the floor without erasing its history.\n";
        } else if (acNarrativeArchetype.toLowerCase().includes("authority")) {
          brief += "Open by establishing what makes this material distinctive and why homeowners consistently struggle with it. Organise the content around the range of problems a homeowner might face rather than a single job narrative. Each section should help the reader identify which specific condition applies to their floor. Close by directing the reader toward the appropriate detailed guidance for their specific problem rather than summarising a completed job.\n";
        } else {
          brief += "This article follows the narrative direction: " + acNarrativeArchetype + ". Open by establishing the homeowner's situation and the floor's condition. Build through the professional decisions that shaped the approach. Close with the outcome that mattered most to the homeowner.\n";
        }
      }
      brief += "--- END ARTICLE CLASSIFICATION ---\n";
    }

    return { success: true, brief: brief };

  } catch(e) {
    return { success: false, message: "Build error: " + e.toString() };
  }
}

function pushAuthorityBriefToSheet(brief, targetRow) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("posts");
    const row   = targetRow || sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "Select a data row first." };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
                         .getValues()[0]
                         .map(function(h) { return String(h).trim(); });

    var colIdx = headers.indexOf("Authority Brief");
    if (colIdx === -1) {
      return { success: false, message: "Authority Brief column not found in posts sheet. Add header to col DN (118)." };
    }

    var cell = sheet.getRange(row, colIdx + 1);
    cell.setNumberFormat("@");
    cell.setValue(brief.trim());

    logPipelineResume("W1 — Authority Brief", "");
    return { success: true, message: "Authority Brief saved to col DN for row " + row };
  } catch(e) {
    return { success: false, message: "Push error: " + e.toString() };
  }
}

/* ============================================================
   PROBLEM ANGLE SYSTEM — TIER B
   Eight-code angle taxonomy governing how each article
   addresses its primary problem.
   Codes: WHY | ID | FIX | PREVENT | CHOOSE | LOCAL | PROVE | DECIDE
============================================================ */

var ANGLE_DEFINITIONS = {
  "WHY":     { label: "Material Explanation", readerQ: "Why is this happening to my floor?", articleTypes: ["Hub Page","Educational Guide"], openingTemplate: "explains why [PROBLEM] occurs at material level, so you understand what you are actually dealing with" },
  "ID":      { label: "Diagnosis",            readerQ: "Which specific condition do I have?", articleTypes: ["Diagnostic Guide"], openingTemplate: "helps you identify which specific condition is causing [PROBLEM], so you choose the right solution" },
  "FIX":     { label: "Method / Process",     readerQ: "How is this professionally corrected?", articleTypes: ["Method Guide","Service Page"], openingTemplate: "explains how [PROBLEM] is professionally corrected, stage by stage" },
  "PREVENT": { label: "Maintenance",          readerQ: "How do I stop this happening again?", articleTypes: ["Maintenance Entity"], openingTemplate: "explains how to prevent [PROBLEM] recurring through correct ongoing maintenance" },
  "CHOOSE":  { label: "Contractor Selection", readerQ: "Who should I hire and how do I evaluate them?", articleTypes: ["Buyer Guide"], openingTemplate: "helps you find and evaluate the right specialist to correct [PROBLEM]" },
  "LOCAL":   { label: "Geo Service",          readerQ: "Can someone fix this near me?", articleTypes: ["Geo Service Page"], openingTemplate: "offers professional correction of [PROBLEM] for homeowners in [LOCALITY]" },
  "PROVE":   { label: "Case Study / Proof",   readerQ: "Has this actually been fixed for someone like me?", articleTypes: ["Case Study"], openingTemplate: "shows how [PROBLEM] was corrected on a real floor in [LOCALITY], with before and after evidence" },
  "DECIDE":  { label: "Value / Comparison",   readerQ: "Is professional intervention worth it for my situation?", articleTypes: ["Buyer Guide","Educational Guide"], openingTemplate: "helps you decide whether professional intervention is the right choice for [PROBLEM]" }
};

function buildProblemAngleStatement(angleCode, primaryProblem, locality) {
  var def = ANGLE_DEFINITIONS[angleCode];
  if (!def) return "";
  var problem = primaryProblem || "this problem";
  var loc     = locality || "your area";
  var stmt    = def.openingTemplate
    .replace("[PROBLEM]",   problem)
    .replace("[LOCALITY]",  loc);
  return "This article " + stmt + ".";
}

function pushProblemAngleToSheet(angleCode, angleStatement) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");
    var row   = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "Select a data row first." };

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
                       .getValues()[0]
                       .map(function(h) { return String(h).trim(); });

    var stmtIdx = headers.indexOf("Problem Angle");
    if (stmtIdx === -1) return { success: false, message: "Problem Angle column not found." };

    // Build full block matching the format other pages use
    var def = {
      "WHY":     { label: "Material Explanation",   readerQ: "Why is this happening to my floor?" },
      "ID":      { label: "Diagnosis",              readerQ: "Which specific condition do I have?" },
      "FIX":     { label: "Method / Process",       readerQ: "How is this professionally corrected?" },
      "PREVENT": { label: "Maintenance",            readerQ: "How do I stop this happening again?" },
      "CHOOSE":  { label: "Contractor Selection",   readerQ: "Who should I hire and how do I evaluate them?" },
      "LOCAL":   { label: "Geo Service",            readerQ: "Can someone fix this near me?" },
      "PROVE":   { label: "Case Study / Proof",     readerQ: "Has this actually been fixed for someone like me?" },
      "DECIDE":  { label: "Value / Comparison",     readerQ: "Is professional intervention worth it?" }
    };

    var entry = def[angleCode];
    if (!entry) return { success: false, message: "Unknown angle code: " + angleCode };

    var fullBlock = angleCode + " — " + entry.label + "\n" +
                    "→ " + entry.readerQ + "\n" +
                    "" + angleStatement;
  

    var stmtCell = sheet.getRange(row, stmtIdx + 1);
    stmtCell.setNumberFormat("@");
    stmtCell.setValue(fullBlock);

    return { success: true, message: "Saved: Problem Angle (" + angleCode + ")" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function getProblemAngleData() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");
    var row   = sheet.getActiveRange().getRow();
    if (row < 2) return { angleCode: "", angleStatement: "", primaryCluster: "", articleType: "", locality: "" };

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
                       .getValues()[0]
                       .map(function(h) { return String(h).trim(); });

    function gc(name) {
      var idx = headers.indexOf(name);
      return idx > -1 ? String(sheet.getRange(row, idx + 1).getValue() || "").trim() : "";
    }

    return {
      angleCode:      gc("Problem Angle Code"),
      angleStatement: gc("Problem Angle"),
      primaryCluster: gc("Primary Query Cluster Owned"),
      articleType:    gc("Article Type"),
      locality:       gc("Locality")
    };
  } catch(e) {
    return { angleCode: "", angleStatement: "", primaryCluster: "", articleType: "", locality: "" };
  }
}

function getAuthorityBrief() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("posts");
    const row   = sheet.getActiveRange().getRow();
    if (row < 2) return "";

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
                         .getValues()[0]
                         .map(function(h) { return String(h).trim(); });

    var colIdx = headers.indexOf("Authority Brief");
    if (colIdx === -1) return "";
    return String(sheet.getRange(row, colIdx + 1).getValue() || "").trim();
  } catch(e) {
    return "";
  }
}

function autoRunAuthorityBrief(targetRow) {
  try {
    console.log("AUTO W1 STARTED");
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");

    if (!sheet) {
      return { success: false, message: "Posts sheet not found." };
    }

    var row = targetRow || SpreadsheetApp.getActiveRange().getRow();
    console.log("ACTIVE ROW: " + row);
    if (row < 2) {
      return { success: false, message: "Select a data row first." };
    }

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
                       .getValues()[0]
                       .map(function(h) { return String(h).trim(); });

    function getValue(name) {
      var idx = headers.indexOf(name);
      if (idx === -1) return "";
      return String(sheet.getRange(row, idx + 1).getValue() || "").trim();
    }

    // ----------------------------------------------------------
    // 1. ENSURE A PROBLEM ANGLE EXISTS
    // ----------------------------------------------------------
    var angleResult = autoAssignProblemAngle(row);

    if (!angleResult.success) {
     function autoRunAuthorityBrief(targetRow) {
  try {
    console.log("AUTO W1 STARTED");
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");

    if (!sheet) {
      return { success: false, message: "Posts sheet not found." };
    }

    var row = targetRow || SpreadsheetApp.getActiveRange().getRow();
    console.log("ACTIVE ROW: " + row);
    if (row < 2) {
      return { success: false, message: "Select a data row first." };
    }

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
                       .getValues()[0]
                       .map(function(h) { return String(h).trim(); });

    function getValue(name) {
      var idx = headers.indexOf(name);
      if (idx === -1) return "";
      return String(sheet.getRange(row, idx + 1).getValue() || "").trim();
    }

    // ----------------------------------------------------------
    // 1. ENSURE A PROBLEM ANGLE EXISTS
    // ----------------------------------------------------------
    var angleResult = autoAssignProblemAngle(row);

    if (!angleResult.success) {
      return {
        success:false,
        message:"W-1 stopped — Problem Angle could not be assigned: " + angleResult.message
      };
    }

    // ----------------------------------------------------------
    // 2. BUILD AUTHORITY BRIEF
    // ----------------------------------------------------------
    var buildResult = buildAuthorityBrief(row);

    if (!buildResult.success) {
      return {
        success: false,
        message: "W-1 stopped — Authority Brief could not be built: " +
                 buildResult.message
      };
    }

    // ----------------------------------------------------------
    // 3. SAVE AUTHORITY BRIEF
    // ----------------------------------------------------------
    var saveResult = pushAuthorityBriefToSheet(buildResult.brief, row);

    if (!saveResult.success) {
      return {
        success: false,
        message: "W-1 stopped — Authority Brief could not be saved: " +
                 saveResult.message
      };
    }

    return {
      success: true,
      message: "W-1 complete — Problem Angle confirmed and Authority Brief saved."
    };

  } catch (e) {
    return {
      success: false,
      message: "Automated W-1 error: " + e.toString()
    };
  }
}

    }

    // ----------------------------------------------------------
    // 2. BUILD AUTHORITY BRIEF
    // ----------------------------------------------------------
    var buildResult = buildAuthorityBrief(row);

    if (!buildResult.success) {
      return {
        success: false,
        message: "W-1 stopped — Authority Brief could not be built: " +
                 buildResult.message
      };
    }

    // ----------------------------------------------------------
    // 3. SAVE AUTHORITY BRIEF
    // ----------------------------------------------------------
    var saveResult = pushAuthorityBriefToSheet(buildResult.brief, row);

    if (!saveResult.success) {
      return {
        success: false,
        message: "W-1 stopped — Authority Brief could not be saved: " +
                 saveResult.message
      };
    }

    return {
      success: true,
      message: "W-1 complete — Problem Angle confirmed and Authority Brief saved."
    };

  } catch (e) {
    return {
      success: false,
      message: "Automated W-1 error: " + e.toString()
    };
  }
}

/* ============================================================
   PROBLEM ANGLE SYSTEM
   Reads Primary Query Cluster from active row.
   Returns angle selector data for W-1 sidebar.
   Saves selected angle code and assembled statement to sheet.
============================================================ */