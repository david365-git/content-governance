# Social Content Manager Abbey — GBP foundation (27 November 2026)

This folder is an **isolated staging copy** in the existing GitHub repository. It is not connected automatically to the separate local folder or Google Apps Script project. No credentials are stored here.

## What is implemented
- Explicitly send one selected Content Engine article into the Social Manager sheet.
- Update an existing article by WordPress Post ID without deleting its image/post records.
- Extract each unique published Abbey image URL from finished article HTML.
- Add image opportunities in REVIEW state, with duplicates suppressed.
- Create GBP draft placeholders only for image opportunities manually marked GENERATE **and** supplied with a distinct, fact-checked post angle.
- Preserve the manual-approval and no-auto-publishing safeguards.
- Validate GBP approval, UK schedule, Abbey article URL and prepared JPG availability; preview due/future/blocked posts without publishing.
- Record eligible due posts in Publication Log with an explicit DRY RUN result.
- Prepare temporary JPG copies of selected Abbey WordPress uploads through a protected WordPress REST endpoint; keep the original files untouched.
- Release prepared files only after all linked posts are marked published; WordPress then waits seven days before cleanup.

## Installation targets
1. **Social Content Manager Abbey** Google Apps Script project `1YJrDXIjPiWuK8uKIDtqImSZUDZsVVCe9goCEbTDIEk2i3rVtGeuf8b36`: use `Code.js`, `Schedule_DryRun.js` and `Image_Processing.js`. Avoid overwriting any unrelated functions you have already added.
2. **Content Engine** existing Apps Script project: use `ContentEngine_SendToSocial.js` there only. Run `ce_sendActiveArticleToSocialManager` on the desired active row. It needs permission to edit the separate workbook.
3. **WordPress administrator** installs and activates `wordpress/abbey-social-temporary-images.php` as a WordPress plugin (place it inside its own plugin folder first). The plugin exposes authenticated `/wp-json/abbey-social/v1/prepare` and `/release` endpoints. WordPress must have an image editor capable of reading WebP and writing JPEG. Do not make the endpoint public.
4. **Social Manager Apps Script settings**: under Project Settings > Script Properties, store `SCM_WP_USER` and `SCM_WP_APP_PASSWORD` for a WordPress account with upload permission. Do not put passwords or application passwords in GitHub or Sheet cells. This calls the WordPress REST API with the account's application password.
5. Destination workbook: `Social Content Manager - Abbey`, ID `1XQbRXDfesWPEKY_EmD8YXj0M5CVtgP-C5OH2bAx3ucE`.

### First safe test
- Confirm Settings > Auto Publish = NO.
- Send one selected article from the source spreadsheet.
- In the destination spreadsheet menu, run **Find article images**.
- Review `Image Opportunities`: add a distinct verified `Distinct Post Angle` and change Decision to `GENERATE` for one genuinely useful photograph.
- Run **Create GBP drafts**. Inspect the draft in `Social Posts Queue`.
- Set a valid scheduled date/time, add a prepared JPG URL and approve a test draft; run **Preview scheduled GBP posts**. No external posts are sent.
- Optionally run **Log GBP dry run** to record eligible posts in Publication Log.
- After WordPress installation and credential setup, choose **Prepare selected JPG images** for a reviewed opportunity marked GENERATE. Check `Image Processing` and the prepared JPG URL on linked queue rows. This processes at most 10 jobs per manual execution.

### Important limitations
- Drafts are deliberately conservative templates, **not AI-written final copy**. The supplied angle is reviewed by the owner and remains unverified by the script.
- No Google Business Profile API publishing, scheduling triggers, Facebook or X integration is active yet. JPG conversion is implemented but requires WordPress plugin installation and configuration.
- The schedule preview accepts ISO (YYYY-MM-DD) or UK (DD/MM/YYYY) dates, and 24-hour UK times (HH:MM).
- Nothing in this branch has been deployed to Apps Script or WordPress.
- `ContentEngine_SendToSocial.js` uses named columns where available and known fallback positions (D, B, H, G, DK, DL, CT). Confirm header mapping on a test row.
- This branch remains separate from the Content Engine runtime until deliberately installed.

## Next development increments
1. Validate source/target data and add quality gates for photo selection.
2. Scheduled-post eligibility checks, idempotent Publication Log and dry-run execution.
3. Test temporary WordPress WebP-to-JPEG conversion and clean-up on a non-critical image.
4. Explicit approval, authorised GBP API integration and controlled pilot.
5. Facebook and X integrations once GBP publishing is reliable.

The main Content Engine application's existing files on this branch have not been altered.
