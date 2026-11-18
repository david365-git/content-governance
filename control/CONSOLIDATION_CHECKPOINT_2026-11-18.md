# Consolidation checkpoint — 18 November 2026

This is a **read-only audit checkpoint**, not an approved deployment or a claim of runtime functionality.

## Established from operator-run comparisons

- PC working copy: `F:\clasp-downloads\content-governance` (branch `new-tier-matrix-adapter-review`), 37 uncommitted status entries at audit time.
- Last known working deployment copy: `F:\clasp-downloads\content-governance-w6-w2c-deploy-20261117`.
- Isolated candidate: `F:\clasp-downloads\ce-control-reports\consolidated-candidate`.
- PC top-level safety copy: `F:\clasp-downloads\ce-control-reports\pc-safety-copy-20261117-185316` (210 top-level files; does **not** include Git history or nested folders).
- Candidate has 197 `.js/.html/.gs/.json` files, and 129 `.js` files passed `node --check` (0 syntax errors).
- PC/deployed filename mismatch: `glc_ArticleCompare.gs` vs `glc_ArticleCompare.js`, same content in diff. Four other files differ: `!_menu.js`, `ce_Sidebar_W6.html`, `ce_W6_WordPress_Direct.js`, `glb_batch_orchestrator.js`.
- Menu audit: 49 entries found, 48 target functions found, historical `openSelectedRowInViewerSheet` missing (unresolved). Do not remove it merely to obtain a pass.
- Sidebar check: 8 panels, 40 inline actions, 15 server-looking calls, 0 unmatched action references, 0 unmatched server candidates; 3 Featured Image calls resolved through shared HTML code. `ce_Sidebar_Main.html` includes both `ce_Sidebar_FI` and `ce_Sidebar_JS_1_Core`. These checks do not prove live clicks work.
- Register of protected capabilities: `PROTECTED_CAPABILITIES_REGISTER.csv`. Source inventory: `MASTER_SCRIPT_REGISTER.csv`, **taken from GitHub main, not yet reconciled with deployed candidate**.

## Before GitHub main can be authoritative

1. Compare full local uncommitted work, live deployment copy, and candidate. Preserve version provenance and avoid overwriting the development branch.
2. Resolve the orphaned viewer menu action deliberately (restore tested implementation or obtain explicit decision to retire).
3. Make an accurate generated master source inventory from the chosen candidate and repeat menu / sidebar / cross-file checks against it; improve parser coverage and test scripts themselves.
4. Create tests for high-risk publishing states and Google Love queue preservation; run relevant end-to-end smoke checks.
5. Commit **approved** consolidated source to `main` through a controlled review, establish a clean official PC working directory, then implement deploy with preflight checks and post-deployment readback.
6. Until then: **NO clasp push** from any of the working, candidate, or test branches.

No Apps Script, WordPress, live Google Sheet, or local PC files were changed by this GitHub-only checkpoint.
