# Social Content Manager Abbey — GBP foundation (27 November 2026)

This folder is an **isolated staging copy** in the existing GitHub repository. It is not connected automatically to the separate local folder or Google Apps Script project. No credentials are stored here.

## What is implemented
- Explicitly send one selected Content Engine article into the Social Manager sheet.
- Update an existing article by WordPress Post ID without deleting its image/post records.
- Extract each unique published Abbey image URL from finished article HTML.
- Add image opportunities in REVIEW state, with duplicates suppressed.
- Create GBP draft placeholders only for image opportunities manually marked GENERATE **and** supplied with a distinct, fact-checked post angle.
- Preserve the manual-approval and no-auto-publishing safeguards.

## Installation targets
1. **Social Content Manager Abbey** Google Apps Script project `1YJrDXIjPiWuK8uKIDtqImSZUDZsVVCe9goCEbTDIEk2i3rVtGeuf8b36`: use `Code.js`. Avoid overwriting any unrelated functions you have already added.
2. **Content Engine** existing Apps Script project: use `ContentEngine_SendToSocial.js` there only. Run `ce_sendActiveArticleToSocialManager` on the desired active row. It needs permission to edit the separate workbook.
3. Destination workbook: `Social Content Manager - Abbey`, ID `1XQbRXDfesWPEKY_EmD8YXj0M5CVtgP-C5OH2bAx3ucE`.

### First safe test
- Confirm Settings > Auto Publish = NO.
- Send one selected article from the source spreadsheet.
- In the destination spreadsheet menu, run **Find article images**.
- Review `Image Opportunities`: add a distinct verified `Distinct Post Angle` and change Decision to `GENERATE` for one genuinely useful photograph.
- Run **Create GBP drafts**. Inspect the draft in `Social Posts Queue`.

### Important limitations
- Drafts are deliberately conservative templates, **not AI-written final copy**. The supplied angle is reviewed by the owner and remains unverified by the script.
- No Google Business Profile API publishing, scheduling triggers, JPG conversion, Facebook or X integration is active yet.
- Nothing in this branch has been deployed to Apps Script or WordPress.
- `ContentEngine_SendToSocial.js` uses named columns where available and known fallback positions (D, B, H, G, DK, DL, CT). Confirm header mapping on a test row.
- This branch remains separate from the Content Engine runtime until deliberately installed.

## Next development increments
1. Validate source/target data and add quality gates for photo selection.
2. Scheduled-post eligibility checks, idempotent Publication Log and dry-run execution.
3. Temporary WordPress WebP-to-JPEG conversion and clean-up.
4. Explicit approval, authorised GBP API integration and controlled pilot.
5. Facebook and X integrations once GBP publishing is reliable.

The main Content Engine application's existing files on this branch have not been altered.
