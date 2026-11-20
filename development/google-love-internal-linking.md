# Google Love Sol internal-link integration — W1.5D_Sol

## Current architecture (20 November 2026)

Google Love Sol **continues generating its own article structure and editorial copy**, independent of the Content Generation W1.5A–W1.5D plan. The new **glb_w15d_sol.js** module acts on the finished HTML:

1. The existing Sol prompt generates the article unchanged (no CW/CX or W1.5A–C prerequisites).
2. The original Sol HTML is written to the recoverable **Google Love Production Drafts** sheet before link selection.
3. **W1.5D_Sol** extracts actual H2 headings and paragraph excerpts.
4. It calls the existing **buildEnhancedSiloMap** and **getMaxLateralLinks** helpers from the original W1.5D script, preserving their locality, material, eligibility and article-type rules. It verifies Feeds Hub separately.
5. A small **MODEL_CHEAP** call chooses eligible destinations, H2 numbers and an exact existing paragraph phrase.
6. A deterministic inserter adds only approved HTML anchor tags; it refuses ambiguous or missing phrase matches, changes to the original article HTML, wrong sections, invalid destinations, missing hub and exceeded lateral quotas.
7. If any validation fails, **New HTML remains unchanged**, the unlinked draft stays recoverable, and the queue marks **SOL REVIEW**. No automatic second paid Sol call.
8. If all checks pass, linked HTML is written into **posts → New HTML**. Original full-batch CI → Sol → W5 → W5C and Sol-only Sol-without-CI modes remain available.

**Original W1.5D remains unchanged.** The old **glb_internal_links.js** represents a legacy development experiment for the CW/CX approach and is NOT required by this architecture.

## Installation (after offline and one-post verification)

Update only:
- glb_batch_orchestrator.js (existing)
- glb_BatchSidebar.html (existing sidebar HTML)
- glb_w15d_sol.js (new Apps Script code file)

Keep the existing !_ce_stage15d_laterallinks.js and !_bc_automatedchain.js files; W1.5D_Sol calls their pre-existing helper functions but does not require running W1.5A–D or the Content Generation pipeline.

**Test constraints:** Only isolated syntax/unit checks have passed so far. A live Google Apps Script end-to-end test, model-output quality review and cost review are still needed. The Sol-only button requires a recorded successful **Content Intelligence completion in the Google Love Batch Queue**, not the unrelated `posts` column CI ("New Post Sent To Site"). All new Sol requests incur model cost; the independent link-selection call may incur a smaller cost.

## Known deliberate constraints

- Insertion is fail-closed: the model must select a unique, exact plain-text phrase within a single HTML paragraph text node. Formatting spans or repeated phrases stop for review rather than generating guessed links.
- No semantic rewrite and no new wording: the post-processing changes tags only.
- Original W1.5D rules are reused through shared helper calls. The model supplies content-specific relevance judgments from the actual article.
- First pilot should review whether candidate selection, locality and paragraph placement are natural.
