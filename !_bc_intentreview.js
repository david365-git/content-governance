/**
 * ============================================================
 * bc_IntentReview.gs - v1.1
 * Abbey Floor Care — Governance Prompt Assembler
 * Intent Review Helper
 *
 * Purpose:
 * - Parse Prompt 1A output server-side
 * - Build a Step 2 review model for the sidebar
 * - Provide ranked intents, context, primary recommendation,
 *   and secondary Y/N recommendations with reasons
 * ============================================================
 */

function bc_buildIntentReviewModel(step1Text) {
  const text = String(step1Text || "").trim();

  const ranked  = bc_irExtractRankedIntents(text);
  const context = bc_irExtractContextBlocks(text);
  const q1rec   = bc_irExtractQ1Recommendation(text);
  let q2items   = bc_irExtractQ2Items(text);

  if (!ranked.length) {
    return {
      ok: false,
      message: "Could not find ranked intent list in Prompt 1A output."
    };
  }

  const primaryIntentName = q1rec.rec || ranked[0].name;

  if (!q2items.length && ranked.length > 1) {
    q2items = bc_irBuildFallbackSecondaryItems(ranked, primaryIntentName);
  }

  return {
    ok: true,
    ranked: ranked,
    context: context,
    q1rec: {
      rec: primaryIntentName,
      reason: q1rec.reason || ""
    },
    q2items: q2items
  };
}


function bc_irExtractRankedIntents(text) {
  const lines   = String(text || "").split(/\r?\n/);
  const intents = [];
  let inRanked  = false;

  for (let i = 0; i < lines.length; i++) {
    const raw  = lines[i] || "";
    const line = raw.trim();

    if (!inRanked) {
      if (
        line === "RANKED INTENT LIST:" ||
        line === "BLOCK 4 — RANKED INTENTS" ||
        line === "BLOCK 4 - RANKED INTENTS"
      ) {
        inRanked = true;
      }
      continue;
    }

    if (
      line === "BLOCK 5 — JUSTIFICATION" ||
      line === "BLOCK 5 - JUSTIFICATION" ||
      line === "CLASSIFICATION JUSTIFICATION:" ||
      /^QUESTION\b/i.test(line) ||
      /^H2 INTENT MAP\b/i.test(line) ||
      /^AMBIGUITY\b/i.test(line)
    ) {
      break;
    }

    if (!line) continue;
    if (line.startsWith("- Include ")) continue;
    if (line.startsWith("- Do NOT ")) continue;

    // Numbered format: "1. Intent Name"
    let m = line.match(/^(\d+)\.\s+(.+)$/);
    if (m) {
      intents.push({
        num: m[1],
        name: m[2].trim()
      });
      continue;
    }

    // Bulleted format: "- Intent Name"
    m = line.match(/^[-*]\s+(.+)$/);
    if (m) {
      intents.push({
        num: String(intents.length + 1),
        name: m[1].trim()
      });
      continue;
    }

    // Plain text format (new BLOCK 4 format): just the intent name on its own line
    // Accept any line that doesn't match instruction patterns
    if (
      !/^Include\b/i.test(line) &&
      !/^Do NOT\b/i.test(line) &&
      !/^-\s/.test(line) &&
      line.length > 2
    ) {
      intents.push({
        num: String(intents.length + 1),
        name: line
      });
    }
  }

  return intents.filter(function(item) {
    const name = String(item.name || "").trim();
    return name &&
      name !== "-" &&
      !/^Include\b/i.test(name) &&
      !/^Do NOT\b/i.test(name);
  });
}


function bc_irExtractContextBlocks(text) {
  const lines = String(text || "").split(/\r?\n/);

  let gscSignal      = "";
  let gscDominant    = "";
  let gscImpressions = "";
  let plainEnglish   = "";
  let confidence     = "";

  let inGsc           = false;
  let inJustification = false;

  for (let i = 0; i < lines.length; i++) {
    const line = (lines[i] || "").trim();

    if (
      line === "GSC SIGNAL ASSESSMENT:" ||
      line === "BLOCK 1 — GSC SIGNAL ASSESSMENT" ||
      line === "BLOCK 1 - GSC SIGNAL ASSESSMENT"
    ) {
      inGsc = true;
      continue;
    }

    if (inGsc) {
      if (line.startsWith("Signal strength:")) {
        gscSignal = line.replace("Signal strength:", "").trim();
        continue;
      }
      if (line.startsWith("Dominant intent category:")) {
        gscDominant = line.replace("Dominant intent category:", "").trim();
        continue;
      }
      if (line.startsWith("Dominant intent:")) {
        gscDominant = line.replace("Dominant intent:", "").trim();
        continue;
      }
      if (line.startsWith("Impression weight:")) {
        gscImpressions = line.replace("Impression weight:", "").trim();
        continue;
      }
      if (line.startsWith("Total impressions:")) {
        gscImpressions = line.replace("Total impressions:", "").trim();
        continue;
      }

      if (
        line === "BLOCK 2 — CONFIRMED ARTICLE TYPE + PRIMARY INTENT" ||
        line === "BLOCK 2 - CONFIRMED ARTICLE TYPE + PRIMARY INTENT" ||
        line.startsWith("CONFIRMED ARTICLE TYPE:") ||
        line.startsWith("QUERY-TO-INTENT") ||
        line.startsWith("GSC-DERIVED") ||
        line.startsWith("RECLASSIFICATION")
      ) {
        inGsc = false;
      }
    }

    if (
      line === "CLASSIFICATION JUSTIFICATION:" ||
      line === "BLOCK 5 — JUSTIFICATION" ||
      line === "BLOCK 5 - JUSTIFICATION"
    ) {
      inJustification = true;
      continue;
    }

    if (inJustification) {
      if (line.startsWith("CONFIDENCE:")) {
        confidence = line.replace("CONFIDENCE:", "").trim();
        continue;
      }

      if (line.startsWith("PLAIN ENGLISH SUMMARY FOR REVIEW:")) {
        const collected = [];
        for (let j = i + 1; j < lines.length; j++) {
          const next = (lines[j] || "").trim();
          if (!next) break;
          if (
            next.startsWith("QUESTION 1") ||
            next.startsWith("QUESTION 2") ||
            next.startsWith("CONFIRMED ARTICLE TYPE:")
          ) {
            break;
          }
          collected.push(next);
        }
        plainEnglish = collected.join(" ").trim();
        continue;
      }

      if (
        line.startsWith("QUESTION 1") ||
        line.startsWith("QUESTION 2") ||
        line.startsWith("CONFIRMED ARTICLE TYPE:")
      ) {
        inJustification = false;
        continue;
      }

      if (line && !plainEnglish) {
        plainEnglish = line;
      }
    }
  }

  return {
    gscSignal: gscSignal,
    gscDominant: gscDominant,
    gscImpressions: gscImpressions,
    confidence: confidence,
    plainEnglish: plainEnglish
  };
}


function bc_irExtractQ1Recommendation(text) {
  const lines = String(text || "").split(/\r?\n/);
  let rec     = "";
  let reason  = "";
  let inQ1    = false;

  for (let i = 0; i < lines.length; i++) {
    const line = (lines[i] || "").trim();

    if (line.startsWith("QUESTION 1 — PRIMARY INTENT:")) {
      inQ1 = true;
      continue;
    }

    if (inQ1) {
      if (line.startsWith("The decision tree and GSC data recommend:")) {
        rec = line.replace("The decision tree and GSC data recommend:", "").trim();
      }
      if (line.startsWith("Plain English reason:")) {
        reason = line.replace("Plain English reason:", "").trim();
      }
      if (line.startsWith("QUESTION 2") || line.match(/^\d+\./)) {
        break;
      }
    }

    if (!rec && line.startsWith("CONFIRMED PRIMARY INTENT:")) {
      rec = line.replace("CONFIRMED PRIMARY INTENT:", "").trim();
    }
  }

  return { rec: rec, reason: reason };
}


function bc_irExtractQ2Items(text) {
  const lines = String(text || "").split(/\r?\n/);
  const items = [];
  let inQ2    = false;
  let current = null;
  let field   = null;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();

    if (line.startsWith("QUESTION 2 — SECONDARY INTENT DECISIONS:")) {
      inQ2 = true;
      continue;
    }

    if (!inQ2) continue;

    if (line.startsWith("Reply with your answers") ||
        line.startsWith("Once you have your answers") ||
        line.startsWith("CONFIRMED INTENT:")) {
      if (current) items.push(current);
      break;
    }

    const intentHeader = line.match(/^(\d+)\.\s+(.+)$/);
    if (intentHeader &&
        !line.startsWith("Decision tree") &&
        !line.startsWith("GSC") &&
        !line.startsWith("Recommendation") &&
        !line.startsWith("Reason:") &&
        !line.startsWith("Permitted") &&
        !line.startsWith("Excluded") &&
        !line.startsWith("Your answer")) {
      if (current) items.push(current);
      current = {
        num: intentHeader[1],
        name: intentHeader[2].trim(),
        step: "",
        gscEvidence: "",
        gscModifier: "",
        rec: "",
        reason: "",
        permitted: "",
        excluded: "",
        answer: ""
      };
      field = null;
      continue;
    }

    if (!current) continue;

    if (line.startsWith("Decision tree step matched:")) {
      current.step = line.replace("Decision tree step matched:", "").trim();
      field = null;
      continue;
    }

    if (line.startsWith("GSC evidence:")) {
      current.gscEvidence = line.replace("GSC evidence:", "").trim();
      field = "gscEvidence";
      continue;
    }

    if (line.startsWith("GSC modifier:") ||
        line.startsWith("— CAUTION") ||
        line.startsWith("— No GSC visibility")) {
      if (line.startsWith("GSC modifier:")) {
        field = "gscModifier";
      } else {
        current.gscModifier += (current.gscModifier ? " " : "") + line;
      }
      continue;
    }

    if (line.startsWith("Recommendation:")) {
      current.rec = line.replace("Recommendation:", "").trim();
      field = null;
      continue;
    }

    if (line.startsWith("Reason:")) {
      current.reason = line.replace("Reason:", "").trim();
      field = "reason";
      continue;
    }

    if (line.startsWith("Permitted content if Y:")) {
      current.permitted = line.replace("Permitted content if Y:", "").trim();
      field = "permitted";
      continue;
    }

    if (line.startsWith("Excluded content if Y:")) {
      current.excluded = line.replace("Excluded content if Y:", "").trim();
      field = "excluded";
      continue;
    }

    if (line.startsWith("Your answer:")) {
      field = null;
      continue;
    }

    if (field && line && !line.startsWith("—")) {
      current[field] += " " + line;
    }

    if (line.startsWith("— CAUTION") || line.startsWith("— No GSC")) {
      current.gscModifier += (current.gscModifier ? " " : "") + line;
    }
  }

  if (current) items.push(current);
  return items;
}


function bc_irBuildFallbackSecondaryItems(ranked, primaryIntentName) {
  const items = [];

  for (let i = 0; i < ranked.length; i++) {
    const item = ranked[i];
    if (item.name === primaryIntentName) continue;

    const fallback = bc_irGetFallbackRecommendation(item.name);

    items.push({
      num: item.num,
      name: item.name,
      step: "",
      gscEvidence: "",
      gscModifier: "",
      rec: fallback.rec,
      reason: fallback.reason,
      permitted: fallback.permitted,
      excluded: fallback.excluded,
      answer: fallback.rec
    });
  }

  return items;
}


function bc_irGetFallbackRecommendation(intentName) {
  const map = {
    "Method Safety / Risk Validation": {
      rec: "Y",
      reason: "This usually supports a cleaning page well because it defines what methods are unsafe and protects the main intent from harmful misuse.",
      permitted: "Keep concise principle-led safety guidance that strengthens the cleaning advice.",
      excluded: "Do not let safety validation become a standalone article inside this page."
    },
    "DIY / Task-Led Instruction": {
      rec: "Y",
      reason: "This can remain where the page already contains light homeowner-safe execution guidance tied closely to the primary cleaning topic.",
      permitted: "Keep only basic safe household actions and realistic boundaries.",
      excluded: "Do not let DIY expand into restoration, honing, or aggressive treatment instruction."
    },
    "Stain Removal": {
      rec: "Y",
      reason: "This can remain as a supporting layer where it helps explain what to do when routine cleaning does not fully solve a mark.",
      permitted: "Keep brief stain-response guidance tied to safe cleaning boundaries.",
      excluded: "Do not let stain removal overtake the page’s main cleaning purpose."
    },
    "Damage Cause / Diagnosis": {
     rec: "Y",
     reason: "This supports a restoration page by clarifying the problem state, visible damage, and treatment boundary before or within the corrective workflow.",
     permitted: "Keep concise diagnosis that explains the observed condition only where it directly supports the restoration narrative.",
     excluded: "Do not turn the page into a standalone diagnosis article."
     },
    "Sealing / Resealing": {
      rec: "N",
      reason: "This often becomes a separate mechanism-led topic and is usually better handled by a dedicated linked page than retained inside a cleaning page.",
      permitted: "Only keep a brief myth-busting note if absolutely necessary.",
      excluded: "Do not retain a full sealing explanation inside this page."
    },
    "Honing / Polishing / Finish Recovery": {
      rec: "Y",
      reason: "This can work as a boundary-setting support layer by explaining what happens when cleaning stops being enough.",
      permitted: "Keep a brief referral-led explanation of the next corrective stage.",
      excluded: "Do not teach polishing or honing as a full process inside this page."
    },
    "Grout Cleaning / Sealing": {
      rec: "N",
      reason: "This is usually too narrow and side-branching unless grout behaviour is central to the main page purpose.",
      permitted: "Only keep very brief marble-safe caution if needed.",
      excluded: "Do not develop grout cleaning as its own branch inside this page."
    },
    "Maintenance / Aftercare": {
      rec: "Y",
      reason: "This usually supports cleaning naturally because it extends the page into sensible ongoing care without displacing the primary purpose.",
      permitted: "Keep concise long-term care guidance connected to routine cleaning.",
      excluded: "Do not broaden it into a separate maintenance hub."
    }
  };

  return map[intentName] || {
    rec: "N",
    reason: "This secondary intent is present, but it should be retained only if it materially strengthens the primary reader journey.",
    permitted: "Remove unless explicitly confirmed.",
    excluded: "Do not retain as a parallel topic branch by default."
  };
}