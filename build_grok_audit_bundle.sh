#!/usr/bin/env bash
set -euo pipefail
OUT="Grok_Generation_Rules_Review.txt"
files=("!_bc_case_study_evidence.js" "!_ce_stage2_contentgeneration.js" "!_ce_stage2_3_coherencecheck.js" "!_ce_stage15b_h2generation.js" "!_ce_stage2c_imagesuggestion.js" "!_bc_automatedchain.js" "!_ce_rewrite_comparison_audit.js")
for f in "${files[@]}"; do
  if [[ ! -f "$f" ]]; then echo "Missing script: $f"; exit 1; fi
done
cat > "$OUT" <<'GROK_PROMPT'
# GROK INDEPENDENT AUDIT — GENERATION RULES, EVIDENCE, MEDIA, HEADINGS
Audit only. DO NOT MODIFY CODE. All seven full script files follow this prompt.
Project: UK stone/tile floor restoration articles for homeowners; approximately 30 years of hands-on practitioner expertise.
Repository review branch: new-tier-matrix-adapter-review.
Live reference test: Penkhull Victorian floor Case Study, posts row 80, Post ID 22021.
Workflow: PreAC → W1.5A–F → W2B → W2B.3 → W8E; GG is W2B.2 HTML, GH is original W2B.3 audit, GI is repaired HTML, CT is W8E final HTML. GL–GV are source-quoted case-study evidence + Evidence Status.
A recent recovery run W2B.3→W8E succeeded: 2 coherence issues in 2 sections structurally repaired. The final original-versus-CT assessment nevertheless reports residual cross-section technical repetition, some overassertive inferred site conditions, captions claiming invisible results, dropped useful source details/links, and uneven h2/h3 hierarchy. Auditor also challenged handover and first-person actions not expressly stated in the old article; these may be genuine author-confirmed work, and absence from old prose does not establish fabrication.

PROPOSED CORRECTIONS TO AUDIT BEFORE ANY CODE CHANGES
P1 — Generation evidence claims: make the source-quoted GL–GV evidence block and separately confirmed author corrections the primary source for job-specific events, while allowing accurate general professional explanations clearly framed as general. Do not make unrecorded client observations, actions or tests into site facts. Review current actual implementation first; avoid duplicate safeguards.
P2 — Preserve chronology and factual nuance: protect carpet discovery, distinct observed floor problems, actual repair/clean/extract/dry/seal steps and observed results. Retain unique project evidence while removing truly redundant abstract exposition. Do not mandate rigid order or length.
P3 — Control cross-section repetition at first draft: require each section to add a distinct observation/decision/process/result; avoid restating the same fired-clay chemistry/acid warning or moisture principle without new purpose. Do not prohibit legitimate multi-stage references, or force shorter copy.
P4 — Media truthfulness: image alt and figure/video captions describe what is visually evidenced (or explicitly described in source) and avoid claiming invisible protection, slurry extraction, tests or a video demonstration. Retain actual image/video URLs, correct media associations, accessibility text and lazy loading where applicable. Identify which downstream W8 stages may overwrite captions, not just W2B.
P5 — Heading hierarchy: major governed stages h2; h3 only genuine subordinate subsections. Audit whether stage planners or later transforms force mixed h2/h3 and propose minimal changes that respect existing governed heading locks.
P6 — Link and local evidence decisions: keep useful genuine project comparison links and concise verified local context where they serve the reader; never reinstate links to hit quotas or invent local housing facts.

CHECK THESE INTEGRATION RISKS
- How NEW Tier Structural Coverage Matrix Required/Conditional, depth and evidence-to-protect are interpreted across stages, without old numeric word budgets.
- Whether GL–GV evidence reader really validates source fingerprints, handles missing/invalid evidence and distinguishes professionally known versus specifically observed facts.
- Whether first-person corrections and aftercare need an explicit owner-confirmed supplement rather than falsely treating source omission as proof of invention.
- Whether atomic W2B.3 repair checks are too brittle for harmless paraphrase; whether material repetition that reaches GI is due to W2B or downstream stages.
- Incomplete provided dependencies: report as UNKNOWN, not confirmed defects. Trace call sites and precise existing functions. The file ! _bc_case_study_evidence.js (without space) is INCLUDED.
- Google Apps Script runtime, prompts, token costs, fail-closed semantics, and how to improve first-time success without introducing multi-pass AI verification.

REQUIRED OUTPUT
1. Verdict PASS / PASS WITH CONDITIONS / FAIL.
2. For each P1–P6: ACCEPT / ADJUST / REJECT, why, actual code location(s), and conflict risks.
3. Confirmed defects versus assumptions or source-verification questions.
4. Minimal scripts/functions to edit; exact suggested change in behaviour (not full rewritten code).
5. Distinguish which issues originate at planning, W2B drafting, W2B.3 repair, and downstream W8E.
6. Recommended architecture prioritising factual preservation, natural British homeowner language, low cost and first-time success.
7. Small regression test set for row 80 plus another Case Study. Audit only, no code edits.

GROK_PROMPT
for f in "${files[@]}"; do
  printf '\n\n========== BEGIN SCRIPT: %s ==========\n' "$f" >> "$OUT"
  cat "$f" >> "$OUT"
  printf '\n========== END SCRIPT: %s ==========\n' "$f" >> "$OUT"
done
echo "Created: $OUT"
echo "Upload that ONE text file to Grok; its audit prompt is included at the top."
