/**
 * ================================================================================
 * ce_Stage5_JudgmentAudit.cs - STAGE5 JUDGMENTAUDIT
 * ================================================================================
 * 
 * LLM judgment audit for technical accuracy
 * 
 * Part of Abbey Floor Care Content Pipeline v77+
 * Reorganised: March 2026
 * ================================================================================
 */

function runSchemaAuditForCorrection() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('posts');
    const row   = sheet.getActiveRange().getRow();
    if (row < 2) return { status: 'error', message: 'ERROR: Select a data row first.' };

    // col 92 — Schema (JSON-LD) (CN)
    let raw = sheet.getRange(row, 92).getValue();
    if (!raw || String(raw).trim() === '') {
      return { status: 'error', message: 'ERROR: No schema found in this row. Run W5 first.' };
    }

    // Sanitise — same as runMechanicalAudit()
    let schema = String(raw)
      .replace(/%22/g, '"').replace(/%3A/g, ':').replace(/%2F/g, '/')
      .replace(/%7B/g, '{').replace(/%7D/g, '}').replace(/%26/g, '&')
      .replace(/%5B/g, '[').replace(/%5D/g, ']')
      .replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#39;/g, "'")
      .replace(/application\/ld\s+json/gi, 'application/ld+json');

    // Preserve full script wrapper for output, strip for parsing
    const fullSchema = schema.trim();
    const sm = fullSchema.match(/<script[^>]*>([\s\S]*?)<\/script>/i);
    const schemaJSON = sm ? sm[1].trim() : fullSchema;

    // Run checks
    const failures = [];
    const passes   = [];

    // Check 10 — valid JSON
    const c10 = auditSchemaValidJSON(schemaJSON);
    if (c10.pass) {
      passes.push('Check 10 — Valid JSON');
    } else {
      failures.push({ check: 10, label: 'Invalid JSON', detail: c10.detail });
      return buildCorrectionPrompt(failures, fullSchema, true);
    }

    // Parse — only if valid JSON
    const parsed = JSON.parse(schemaJSON);

    const checks = [
      { n: 11, fn: auditSchemaContext,              label: '@context value' },
      { n: 12, fn: auditSchemaType,                 label: '@type present' },
      { n: 13, fn: auditSchemaRequiredFields,       label: 'Required fields' },
      { n: 14, fn: auditSchemaEncodedChars,         label: 'HTML-encoded characters', raw: schemaJSON },
      { n: 15, fn: auditSchemaURLs,                 label: 'URL format' },
      { n: 16, fn: auditSchemaBracketURLs,          label: 'Bracket-wrapped URLs', raw: schemaJSON },
      { n: 20, fn: auditSchemaPublisherAnchor,      label: 'Publisher @id anchor' },
      { n: 21, fn: auditSchemaAuthorPerson,         label: 'Author @type Person' },
      { n: 22, fn: auditSchemaNoUrlOnArticle,       label: 'No url field on Article' },
      { n: 23, fn: auditSchemaHasPartFragmentTypes, label: 'hasPart fragment/@type consistency' },
      { n: 24, fn: auditSchemaRootFields,            label: 'Illegal root-level fields' }
    ];

    checks.forEach(function(c) {
      // Some checks take raw JSON string, others take parsed object
      const result = (c.n === 14 || c.n === 16) ? c.fn(c.raw) : c.fn(parsed);
      if (result.pass) {
        passes.push('Check ' + c.n + ' — ' + c.label);
      } else {
        failures.push({ check: c.n, label: c.label, detail: result.detail });
      }
    });

    if (failures.length === 0) {
      return {
        status: 'clear',
        message: 'SCHEMA AUDIT CLEAR — ' + passes.length + '/' + passes.length + ' checks passed.\nNo corrections needed.',
        prompt: null
      };
    }

    return buildCorrectionPrompt(failures, fullSchema, false);

  } catch(e) {
    return { status: 'error', message: 'AUDIT ERROR: ' + e.toString() };
  }
}

/* ── Build the correction prompt from failures ── */
function buildCorrectionPrompt(failures, fullSchema, jsonInvalid) {
  const lines = [];
  lines.push('You are a Schema.org JSON-LD correction specialist.');
  lines.push('The schema below has failed a structural and mechanical audit. Correct ONLY the listed failures.');
  lines.push('Return ONLY the minified JSON inside a <script type="application/ld+json"> block.');
  lines.push('No explanation. No preamble. No markdown.');
  lines.push('');
  lines.push('════ MANDATORY ROOT STRUCTURE ════');
  lines.push('The JSON MUST strictly follow this shape. Only two keys are permitted at root level.');
  lines.push('No fields (headline, url, @type, mainEntityOfPage etc.) are allowed outside @graph:');
  lines.push('{');
  lines.push('  "@context": "https://schema.org",');
  lines.push('  "@graph": [');
  lines.push('    { "@type": "Article", "@id": ".../#article", "headline": "...", "description": "...", ... },');
  lines.push('    { "@type": "BreadcrumbList", "@id": ".../#breadcrumb", ... },');
  lines.push('    { "@type": "LocalBusiness", "@id": ".../#localbusiness", ... }');
  lines.push('  ]');
  lines.push('}');
  lines.push('');
  lines.push('════ FAILURES TO CORRECT ════');

  failures.forEach(function(f) {
    lines.push('CHECK ' + f.check + ' — ' + f.label + ': ' + f.detail);
  });

  lines.push('');
  lines.push('════ HARD RULES ════');
  lines.push('- @context MUST be exactly "https://schema.org" — never a website URL');
  lines.push('- ROOT LEVEL: Only "@context" and "@graph" are permitted. Remove ALL other root-level fields.');
  lines.push('- publisher @id must be exactly "https://www.abbeyfloorcare.co.uk/#localbusiness"');
  lines.push('- author must be {"@type":"Person","name":"David Allen"}');
  lines.push('- Article entity: MUST have "description" field. Must NOT have a direct "url" field.');
  lines.push('- Use @id for internal graph references — never a bare "url" property on graph entities.');
  lines.push('- hasPart fragments: #article → @type: Article | #service → @type: WebPage');
  lines.push('- hasPart entries: do not add "url" fields — use @id only');
  lines.push('- JSON must be minified — single line, no internal line breaks');
  lines.push('- Wrap output in <script type="application/ld+json"> ... </script>');
  lines.push('');
  lines.push('════ SCHEMA TO CORRECT ════');
  lines.push(fullSchema);

  return {
    status: 'failures',
    failCount: failures.length,
    message: failures.length + ' structural failure(s) found — correction prompt ready to copy.',
    prompt: lines.join('\n')
  };
}

/* ============================================================
   CHECK 23 — HAPART FRAGMENT / @TYPE CONSISTENCY
   #article fragments must use @type: Article
   #service fragments must use @type: WebPage
============================================================ */
function auditSchemaHasPartFragmentTypes(parsed) {
  const graph = parsed['@graph'] || (parsed['@type'] ? [parsed] : []);
  const problems = [];

  graph.forEach(function(entity) {
    if (!entity.hasPart) return;
    const parts = Array.isArray(entity.hasPart) ? entity.hasPart : [entity.hasPart];
    parts.forEach(function(part) {
      const id   = part['@id'] || '';
      const type = part['@type'] || '';
      if (id.endsWith('#article') && type !== 'Article') {
        problems.push('@id "' + id + '" has #article fragment but @type is "' + type + '" — expected "Article"');
      }
      if (id.endsWith('#service') && type !== 'WebPage') {
        problems.push('@id "' + id + '" has #service fragment but @type is "' + type + '" — expected "WebPage"');
      }
    });
  });

  if (problems.length === 0) {
    return { pass: true, label: 'PASS', detail: 'All hasPart fragment/@type pairs are consistent' };
  }
  return {
    pass: false,
    label: 'FAIL',
    detail: problems.join('; ')
  };
}

/* ── Push corrected schema only to active row ── */
function pushCorrectedSchemaToActiveRow(raw) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('posts');
    const row   = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: 'ERROR: Select a data row first.' };

    // col 92 — Schema (JSON-LD) (CN)
    // Sanitise
    let cleaned = String(raw).trim()
      .replace(/application\/ld\s+json/gi, 'application/ld+json')
      .replace(/%22/g, '"');

    const cell = sheet.getRange(row, 92);
    try {
      cell.setPlainTextValue(cleaned);
    } catch(e) {
      cell.setNumberFormat('@');
      cell.setValue(cleaned);
    }

    logPipelineResume("W5B — Schema Corrected", "");
    return { success: true, message: 'Schema pushed to row ' + row + '. Re-run audit to verify.' };
  } catch(e) {
    return { success: false, message: 'PUSH ERROR: ' + e.toString() };
  }
}

/* ============================================================
   CHECK 24 — ILLEGAL ROOT-LEVEL FIELDS
   Only @context and @graph are permitted at root level.
   Detects hybrid schemas where entity fields (headline, url,
   @type, mainEntityOfPage etc.) are placed outside @graph.
============================================================ */


