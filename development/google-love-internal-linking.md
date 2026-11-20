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
8. If all checks pass, linked HTML is written into **posts → New HTML**. Full Batch runs CI → Sol → W1.5D_Sol → W5 → W5C. Sol Only reuses completed CI and runs Sol → W1.5D_Sol → W5 → W5C.

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

## Sol-only W5/W5C regression check

A mocked two-post sidebar execution confirmed `PREPARE → SOL → W5 → W5C` for each Post ID in Sol Only mode, without CI. Full Batch still uses `PREPARE → CI → SOL → W5 → W5C`. These are offline behavioural checks, not live Apps Script execution.

## Stronger destination relevance (20 November 2026)

W1.5D_Sol now excludes geography-targeted educational, diagnostic, method, buyer and FAQ pages when the current article's explicit locality does not match. The check uses destination locality metadata where populated and recognisable UK geographic markers in titles/URLs where not. It is deliberately conservative: blank locality on a general article does not justify a Glasgow-targeted educational guide. **Case studies remain exempt from this added geographic educational filter**; their technical relevance is still judged from supplied evidence and the existing silo rules. Existing geo-service and service-page eligibility is unchanged.

Selection instructions require a distinct reader benefit for every chosen link, discourage repeating the source article's advice, and demand anchors that signal destination topic rather than a generic statement. Returned assignments must contain a nontrivial `readerBenefit` explanation or validation stops. This is a structural check, **not proof** that every claimed benefit is substantively correct; review the first pilots editorially. The source article and links previously saved to New HTML are not retroactively altered.

## W6 partial WordPress publication recovery (20 November 2026)

The W6 Apps Script error path now handles an HTTP 500, timeout, invalid receipt or post-request transport error as an **uncertain write**, never as proof that the post stayed unchanged. It performs authenticated read-only checks of the WordPress post, title, raw HTML, Yoast SEO fields, status/date and featured-image attachment URL; any inaccessible field is marked NOT VERIFIED. Schema in `aFhfc_head_code` is not exposed by the ordinary WordPress REST post response and is explicitly marked NOT VERIFIED, not assumed absent/present. The detailed read-back and any error code are shown to the operator. An uncertain response prevents W6 publication-log, site-export and schema-sheet writes and disables further Update/Republish actions for that post during the sidebar session, including after re-verification.

The patch does **not** automatically reissue the publish POST, alter WordPress, retry a syndication event, or silently replace the currently published article. No live site changes are made by deployment of these files. This is a safety/recovery fix, **not** a confirmed repair for the underlying WordPress PHP HTTP 500. The actual active WordPress bridge must be instrumented for phase checkpoints / PHP server errors to pinpoint the server-side cause. A disabled/test WPCode source is not proof of the active bridge implementation.

Files to manually deploy in Apps Script: `ce_W6_WordPress_Direct.js` and `ce_Sidebar_W6.html`. Regression test: `node development/test-w6-error-recovery.cjs`. The old post 48148 remains published and requires separate editorial review of the metadata before any reconciliation.
