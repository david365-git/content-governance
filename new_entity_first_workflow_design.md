# New entity-first workflow — design contract (draft)

Status: design only; not implemented or approved for production.
Branch: `new-tier-matrix-adapter-review`.
Principle: **protect evidence, not inherited prose**.

## Success criteria
The resulting page serves its specific search intent and article type, gives useful, practical answers, demonstrates expertise grounded in real restoration work, and is clear to homeowners. Rankings are a goal, not a guaranteed output.

## Immutable evidence and governance
- Preserve the meaning of verified observations, client requirements, tests actually performed, materials, treatments, outcomes and limitations. Distinguish direct evidence from inference and general technical knowledge.
- Never invent inspections, equipment used, results, clients, dates, locations, repairs or reactions.
- For case studies, use the versioned evidence extraction package as a fact source; preserve evidence IDs and source references across all stages.
- Preserve the source media inventory (images, videos and meaningful captions/alt) from `site-export`; intentional omission must be explicitly recorded rather than silently dropped.
- Only W1.5D/CX authorises internal destinations and placement. Revalidate link destinations at every save; do not inject generic signposting paragraphs.
- Respect approved article type, scope, primary query ownership, Problem Angle and the Technical Matrix.
- Preserve approved keyphrase, metadata constraints, original source snapshots, exception reporting, resumability and cost accounting.

## Freely rewritable content
- Reorder sections and paragraphs, change headings, introductions and conclusions, delete repetition, replace weak prose, improve examples and causal explanations.
- An existing article is **context and evidence input**, not a compulsory writing template or word-count baseline.
- Do not reward similarity to original wording. A large rewrite is valid when verified evidence and meaningful claims survive.
- Do not impose arbitrary word targets, exact entity repetition, vocabulary proximity quotas or forced question headings.
- Explain material properties only where they help the reader understand a diagnosis, choice, treatment, limitation or outcome.
- Article-type rules govern emphasis: case studies show documented chronology/results; guides teach; buyer guides support choices; diagnostic pages help recognise causes; geo pages make only grounded local claims.

## Generation and verification boundaries
1. Collect W0 research, verified evidence, source-media baseline and article intent.
2. Build a governed article plan with clear section purposes and coverage requirements; retain W1.5D as separate link authority.
3. Generate **new prose** from the plan and evidence; do not mechanically preserve old wording.
4. Independently fact-check all specific material/project claims against evidence and technical references, recording a versioned **fact-lock ledger** of corrections and protected claims.
5. Editorially improve the current draft; edits may be extensive but must not change protected facts, media or approved links. Recheck the ledger after any rewrite.
6. Prepare publication assets, metadata and schema; run a final post-W8E rules-based gate on final HTML and metadata, with no unvalidated changes afterwards.

## What to measure in an old-vs-new test
- Factual claims supported / unsupported, and documented project details retained.
- Search-intent coverage, differentiated usefulness, clarity, unnecessary repetition and natural entity relationships.
- Images/videos retained; links authorised; headings, metadata and schema valid.
- Total AI calls, total API cost, repair attempts, manual interventions and pipeline stops.
- Record independent human/Claude quality assessments blind to version where practicable.

## Safe migration
- Keep production and the original workflow unchanged.
- New code files use `new_` prefix; execute in shadow mode against a test copy or `v2` output, never original pipeline columns until approved.
- Test all eight article types, including existing manual angles and cases with rich evidence/media, before any switch-over.
- Preserve existing W7–W9 and W8A–W8D manual tooling and site-export integration.
- A simple syntax pass does not prove end-to-end safety; require mock/unit tests, independent review and controlled live tests.

## Outstanding decisions before coding the generator
- Exact fact-lock ledger format, provenance and where immutable source snapshots are stored.
- Explicit media omission policy and final diff/approval mechanism.
- How to preserve article-type-specific checks without multiplying AI review calls.
- Shadow-column or separate-`v2` storage mapping.
