# Content Engine — Master Control (foundation)

**Status: FOUNDATION / NOT APPROVED FOR DEPLOYMENT.**

This folder is a starting point for consolidating the Content Engine. It has NOT merged the deployed 17 November code into GitHub `main`. Do not run `clasp push` from this branch.

## What is in this folder

- `MASTER_SCRIPT_REGISTER.csv`: 193 source/configuration file entries from the **GitHub main tree at the time of inventory**. Labels are provisional and approval is marked TO REVIEW. It does not yet represent a verified live deployment.
- `audit_content_engine.py`: a **read-only** check for missing/new script files, missing menu functions, and selected critical functionality. It deliberately does not modify files or connect to WordPress, Google Sheets, or Apps Script.

## The single-source-of-truth rule

1. **Preserve and inspect** the actual running Apps Script and the last successful controlled deployment before choosing a baseline.
2. Merge approved changes into **one verified main branch**; retain a recoverable snapshot of local uncommitted development.
3. Keep one normal PC working directory for that official branch; experimental branches are not deployed.
4. Audit **all** source files and protected capabilities before deployment. No individual repair may deploy from an old archive or unrelated branch.
5. Record the approved Git commit, project Script ID, timestamp, and deployment outcome. The live Apps Script content must be checked back against the approved source.
6. Give the user one simple deploy command **only after** the consolidation, safety checks, and readback mechanisms are tested.

## Run the preliminary audit (read only)

From the repository root:

```bash
python control/audit_content_engine.py
```

To examine the existing controlled deployment without altering it:

```bash
python control/audit_content_engine.py --root /f/clasp-downloads/content-governance-w6-w2c-deploy-20261117
```

This preliminary audit is expected to find differences between `main` and the deployed copy. Differences are **evidence**, not a reason to overwrite either copy. An audit PASS is not proof of working UI behaviour.

## Known capabilities to preserve (first pass)

- Row Analysis menu and Google Love Batch Queue live refresh
- W6 scheduled/published verification, Update, publish-now, recovery safeguard
- W2C media upload and Featured Image upload
- Content Intelligence and Google Love batch execution
- Existing article comparison and all other approved menu/sidebar actions

## Explicit safety constraints

**Do not:** reset, clean, or force-pull the existing PC development branch; force-update main; merge blindly; run `clasp push` from this control branch; change WordPress posts; or remove historical backups.

**Next:** reconcile the staging folder and local development changes against `main`, inspect the resulting menu/capability delta, approve a complete baseline, then build and test a guarded deploy command.

## Phase 2: Reconcile the actual file versions (read only)

Run `control/reconcile_versions.py` from an independent reports directory after fetching the relevant remote branches. It reads:

- the last known successful `clasp push` staging folder,
- all current PC working files **including uncommitted changes**,
- remote-tracking refs for `main`, Matrix, W6 and Google Love.

It produces `SUMMARY.json`, `source_comparison.csv`, `capability_comparison.csv`, and `pc_uncommitted_changes.txt`. **None of the inputs are modified.**

The per-file hashes show exact content matches and differences. They cannot tell us which changed code is preferable, and the capability checks are preliminary string-presence checks only. We must inspect and test the conflicting changes before choosing a consolidated baseline.

Do not mistake a 196-file `clasp push` for evidence that all 196 files contain the right features.
