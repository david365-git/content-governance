# Release candidate record — 18 November 2026

**Status: BACKED UP AND STATICALLY VERIFIED — NOT DEPLOYED**

This record preserves the operator's evidence without claiming runtime approval.

## Live baseline

- Successfully pulled from the live Apps Script project: 196 source/configuration files.
- Backup directory: `F:\clasp-downloads\ce-control-reports\live-backups\apps-script-live-20261118-133049-UTC`
- Verified ZIP: `F:\clasp-downloads\ce-control-reports\live-backups\apps-script-live-20261118-133049-UTC.zip`
- The backup script checked its Script ID against the candidate before downloading.
- The live script and WordPress site remained unchanged.

## Exact source comparison

Operator ran `control/build_candidate_report.py` at 2026-11-18T13:31:49Z.

- Candidate inventory: 197 items (includes `.clasp.json`).
- Live backup comparison inventory: 198 items (includes `.clasp.json` and generated `manifest.json`).
- Actual script differences: `!_menu.js` (obsolete viewer menu removed at owner's request) and `glb_BatchSidebar.html` (operator's Stop finishes current article rather than current stage).
- The third reported difference, `manifest.json`, is backup-only metadata.
- 48 menu commands resolved; 40 sidebar actions resolved; selected protected checks present.
- Google Love mock stop tests passed at CI, SOL, W5, W5C, confirming no next post begins under simulated callbacks.
- This does not prove live behaviour.

## Controls added after comparison

- Automatic GitHub workflow for source versions, protected capabilities, syntax and Google Love simulation.
- `control/check_live_release_candidate.py`: read-only exact live-baseline comparison, allowing only the two documented source differences.
- `control/verify_live_readback.py`: checks an independently pulled live snapshot against an approved Git commit; **does not itself pull or deploy**.
- Branch protection created for `main` by owner (screenshot confirmation 18 November 2026).

## Outstanding

1. Operator runs the exact candidate-vs-live backup check from the PC. The output must PASS.
2. Close review of source ownership and confirm the original development tree (including uncommitted changes) remains preserved and unaltered.
3. Approve and merge the draft pull request into `main` only after review; establish one dedicated `main` PC folder.
4. Create and validate an *approval-gated* deploy procedure that performs a fresh backup and independent readback; require explicit approval before changing live scripts.
5. After any approved deployment, check live menus and affected features (especially Google Love stop) and retain a documented recovery path.

**Never run `clasp push` from the consolidation branch, original PC development folder or a staging archive.**
