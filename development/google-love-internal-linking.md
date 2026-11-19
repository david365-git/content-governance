# Google Love internal-linking integration — development specification

**Status:** development only; do not merge, deploy, or republish.
**Reference case:** Post 48148 — contact is present but the material hub and relevant internal links are missing.
**Baseline:** exact live-source GitHub `main` at branch creation.

## Confirmed existing code

- `!_ce_stage15d_laterallinks.js`: W1.5D reads **Feeds Hub** from the posts row and a governed W1.5C plan in **CW (101)**, assigns the hub and approved lateral routes, then saves its final plan in **CX (102)**. Selection relies on `buildEnhancedSiloMap` and article-type link budgets.
- `!_ce_stage2_contentgeneration.js`: legacy W2B consumes CX's exact section-linked destinations, instructs the writer to emit actual links, and refuses missing CX.
- `glb_batch_orchestrator.js`: Google Love CI → Sol → W5 → W5C does **not** currently run W1.5D or pass CX / Feeds Hub destinations to Sol. It saves Sol HTML without checking material-hub or lateral anchor completeness. It separately adds a contact footer.

## Requirements before any merge

1. **No invented destinations.** Feeds Hub is the authoritative material hub for non-hub articles when populated with a valid on-site URL. Existing W1.5D CX assignments are reused only when current and verifiable against the posts row and known eligible destinations. Do not trust a stale CX plan merely because it says `Stage 1.5D complete.`
2. **Article type rules.** Hub pages must not link to themselves; Hub pages may link to approved children. Service / geo-service regional and locality restrictions must remain intact. Lateral quotas come from Article Type Control Sheet. Contact is separate from hub and lateral quotas.
3. **Missing CX support.** Google Love should work for posts never run through W1.5, but it must not silently invent or auto-approve lateral links. It needs a controlled proposal/selection path, using `buildEnhancedSiloMap`, relevance checks and existing W1.5D eligibility rules. A missing or invalid hub cannot be fabricated.
4. **Correct placement.** Supply approved destinations to Sol with section context, and verify the **final generated HTML** has a real `<a href>` pointing to each required destination, not merely a URL in plain text, JSON-LD, script, image or footer. Resolve relative vs absolute paths carefully and reject self-links and duplicates.
5. **Fail clearly before production output.** On required-link failure, keep previous `posts!New HTML` intact; preserve draft/review evidence and label link review, with no automatic second paid API call. Do not retry publishing.
6. **Respect user-authored content.** Do not rewrite historical content, change W2C media, stage W5/5C, author footer, or WordPress publication solely to insert links.
7. **Explicit tests.** Good CX, stale CX, missing CX, missing hub, article-is-hub, zero lateral quota, approved case-study, disallowed geo service route, wrong locality/country, duplicate links, self-link, URL aliases/trailing slash, incorrect section, href missing but text present, contact-only output, case-study image retention, and recovery on failed link audit.

## Suggested implementation milestones

- **A — Discovery / read-only validation:** record current per-post link governance inputs and demonstrate gaps on representative rows including 48148.
- **B — Governed destination adapter:** reuse W1.5D's eligibility and preserve manually/previously approved assignments where still valid.
- **C — Sol prompt integration:** require approved links without inventing URLs or section-specific evidence.
- **D — Deterministic HTML audit:** prevent `SOL DONE` and production HTML replacement when required links are absent.
- **E — Offline regression tests:** no API calls, no spreadsheet writes and no WordPress actions; check sidebars and existing stages unaffected.
- **F — Owner review, then separate PR merge and controlled one-file/few-file Apps Script deployment.**

## Open questions to resolve from actual sheet row, not guess

- Post 48148's article type, material, Feeds Hub, populated CW/CX, current New HTML and current Site Export source links.
- Whether the existing CX plan is fresh for this Google Love rewrite, or whether W1.5D candidates must be selected anew.
- How the existing Sol article sections align with W1.5D's section identifiers; absent alignment, validation must not falsely claim in-section placement.

No code is authorised for production until these conditions are demonstrated.
