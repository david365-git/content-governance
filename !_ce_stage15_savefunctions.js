/**
 * ================================================================================
 * ce_Stage15_SaveFunctions.gs - STAGE 1.5 SAVE FUNCTIONS
 * ================================================================================
 * 
 * Save functions for all W1.5 stages
 * 
 * Column map:
 *   Col 99  (CU) — W1.5A Structure Audit
 *   Col 100 (CV) — W1.5B H2 Generation
 *   Col 101 (CW) — W1.5C Enriched Plan
 *   Col 102 (CX) — W1.5D Lateral Links Final Plan
 *   Col 103 (CY) — W1.5E Google Optimization JSON (read by W2B)
 * 
 * Part of Abbey Floor Care Content Pipeline v77+
 * Updated: May 2026
 * ================================================================================
 */

function saveStage15AStructure(structure) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('posts');
    const row = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: 'Select a data row first.' };

      const cleanedStructure = String(structure || '').trim();

      if (!cleanedStructure) {
        return {
          success: false,
          message: 'No W1.5A structure provided.'
        };
      }

      // Never save a W1.5A structure while any audit check is still failing.
      const failedChecks =
        cleanedStructure.match(/AUDIT CHECK\s+\d+:\s*FAIL\b[^\n]*/gi) || [];

      if (failedChecks.length > 0) {
        return {
          success: false,
          message:
            'W1.5A contains ' +
            failedChecks.length +
            ' failed audit check(s) — structure not saved to CU. ' +
            failedChecks.join(' | ')
        };
      }

      // Reject unresolved location/template placeholders.
      if (/\[(?:Town\/City|Location|Area|Locality|Parent Area)\]/i.test(cleanedStructure)) {
        return {
          success: false,
          message:
            'W1.5A contains an unresolved location placeholder — structure not saved to CU.'
        };
      }

      // Hard-lock aggregate word budgets for repeated TSM Orders.
      const d = getActiveRowDataMap();
      const material = d["Stone Type"] || "UNKNOWN";
      const articleType = d["Article Type"] || "General";
      const secondaryIntents =
        String(d["Secondary Intent Decisions"] || "");

      const tierMap = {
        "Hub Page":          "Tier 1",
        "Educational Guide": "Tier 1",
        "Method Guide":      "Tier 2",
        "Service Page":      "Tier 2",
        "Geo Service Page":  "Tier 2",
        "Diagnostic Guide":  "Tier 3",
        "Buyer Guide":       "Tier 3",
        "FAQ Spoke":         "Tier 3",
        "Case Study":        "Tier 4"
      };

      const tierLabel = tierMap[articleType] || "Tier 2";
      let tsmData =
        getTierStructuralRequirements(
          material,
          tierLabel,
          articleType
        ) || [];

      tsmData = tsmData.filter(function(req) {
        return !secondaryIntents.includes(
          "N — " + req.name
        );
      });

      const expectedByOrder = {};
      tsmData.forEach(function(req, index) {
        const minWords = Number(req.wordCount || 0);
        expectedByOrder[index + 1] = {
          name: String(req.name || ""),
          min: minWords,
          max: Math.round(minWords * 1.2)
        };
      });

      const actualByOrder = {};
      const sectionLineRegex =
        /^SECTION\s+\d+:\s*Order\s+(\d+)\b[^\n]*?—\s*(\d+)\s*[–-]\s*(\d+)\s+words\b/gmi;

      let sectionMatch;
      while ((sectionMatch = sectionLineRegex.exec(cleanedStructure)) !== null) {
        const order = Number(sectionMatch[1]);
        const minWords = Number(sectionMatch[2]);
        const maxWords = Number(sectionMatch[3]);

        if (!actualByOrder[order]) {
          actualByOrder[order] = {
            min: 0,
            max: 0
          };
        }

        actualByOrder[order].min += minWords;
        actualByOrder[order].max += maxWords;
      }

      const budgetErrors = [];

      Object.keys(expectedByOrder).forEach(function(key) {
        const order = Number(key);
        const expected = expectedByOrder[order];
        const actual = actualByOrder[order];

        if (!actual) {
          budgetErrors.push(
            'Order ' + order +
            ' (' + expected.name + ') has no mapped word budget.'
          );
          return;
        }

        if (
          actual.min !== expected.min ||
          actual.max !== expected.max
        ) {
          budgetErrors.push(
            'Order ' + order +
            ' (' + expected.name + ') totals ' +
            actual.min + '–' + actual.max +
            ' words; governed total is ' +
            expected.min + '–' + expected.max + '.'
          );
        }
      });

      Object.keys(actualByOrder).forEach(function(key) {
        const order = Number(key);
        if (!expectedByOrder[order]) {
          budgetErrors.push(
            'Unexpected Order ' + order +
            ' appears in the section mapping.'
          );
        }
      });

      if (budgetErrors.length > 0) {
        return {
          success: false,
          message:
            'W1.5A word-budget validation failed — structure not saved to CU. ' +
            budgetErrors.join(' | ')
        };
      }

      // Save to Column CU (99)
      const cell = sheet.getRange(row, 99);
      cell.setValue(cleanedStructure);
    
    return { 
      success: true, 
      message: 'Structure saved to Column CU.' 
    };
  } catch(e) {
    return { 
      success: false, 
      message: 'Error saving: ' + e.message 
    };
  }
}

function saveStage15BH2s(h2s) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('posts');
    const row = sheet.getActiveRange().getRow();

    if (row < 2) {
      return {
        success: false,
        message: 'Select a data row first.'
      };
    }

    const raw = String(h2s || '').trim();

    if (!raw) {
      return {
        success: false,
        message: 'No W1.5B H2 output provided.'
      };
    }

    // CU is the governed W1.5A structure.
    const structure =
      String(sheet.getRange(row, 99).getValue() || '').trim();

    if (!structure) {
      return {
        success: false,
        message: 'Column CU is empty — W1.5A must complete before W1.5B.'
      };
    }

    // Determine exactly which numbered sections W1.5A requires
    // and preserve the governed H2/H3 level.
    const expectedSections = [];
    const expectedLevels = {};
    const sectionRegex = /^SECTION\s+(\d+):([^\n]*)$/gmi;
    let match;

    while ((match = sectionRegex.exec(structure)) !== null) {
      const number = Number(match[1]);
      const line = String(match[2] || '');
      const levelMatch = line.match(/(?:—|-)\s*(H[23])\s*:/i);
      const level = levelMatch ? levelMatch[1].toUpperCase() : 'H2';

      if (expectedSections.indexOf(number) === -1) {
        expectedSections.push(number);
      }

      expectedLevels[number] = level;
    }

    expectedSections.sort(function(a, b) {
      return a - b;
    });

    if (expectedSections.length === 0) {
      return {
        success: false,
        message: 'Could not determine required section numbers from W1.5A in CU.'
      };
    }

    const expectsHubIntro =
      /^HUB-INTRO:/mi.test(structure);

    const lines = raw
      .replace(/^```[a-z]*\s*/i, '')
      .replace(/\s*```$/i, '')
      .split('\n')
      .map(function(line) {
        return String(line).trim();
      })
      .filter(function(line) {
        return line !== '';
      });

    const foundSections = {};
    const errors = [];

    lines.forEach(function(line) {

      const sectionMatch =
        line.match(/^SECTION\s+(\d+):\s*(.+)$/i);

      if (!sectionMatch) {
        return;
      }

      const number = Number(sectionMatch[1]);
      let payload = String(sectionMatch[2] || '').trim();

      const explicitLevelMatch =
        payload.match(/^Heading\s+(H[23]):\s*(.+)$/i);

      let suppliedLevel = '';
      let heading = payload;

      if (explicitLevelMatch) {
        suppliedLevel = explicitLevelMatch[1].toUpperCase();
        heading = String(explicitLevelMatch[2] || '').trim();
      } else {
        heading = payload.replace(/^Heading\s+H\d:\s*/i, '').trim();
      }

      if (!heading) {
        errors.push(
          'SECTION ' + number + ' has no heading.'
        );
        return;
      }

      const expectedLevel =
        expectedLevels[number] || 'H2';

      if (
        suppliedLevel &&
        suppliedLevel !== expectedLevel
      ) {
        errors.push(
          'SECTION ' + number +
          ' must remain ' + expectedLevel +
          ' as governed by W1.5A.'
        );
        return;
      }

      if (foundSections[number]) {
        errors.push(
          'SECTION ' + number + ' appears more than once.'
        );
        return;
      }

      foundSections[number] =
        'SECTION ' +
        number +
        ': Heading ' +
        expectedLevel +
        ': ' +
        heading;
    });

    // Make sure every governed W1.5A section is present.
    expectedSections.forEach(function(number) {
      if (!foundSections[number]) {
        errors.push(
          'SECTION ' + number + ' is missing.'
        );
      }
    });

    // Reject section numbers that do not exist in W1.5A.
    Object.keys(foundSections).forEach(function(key) {
      const number = Number(key);

      if (expectedSections.indexOf(number) === -1) {
        errors.push(
          'Unexpected SECTION ' +
          number +
          ' is not present in W1.5A.'
        );
      }
    });

    if (errors.length > 0) {
      return {
        success: false,
        message:
          'W1.5B heading set rejected — CV not changed. ' +
          errors.join(' | ')
      };
    }

    const outputLines = [];

    // Preserve the governed HUB-INTRO from CU rather than trusting
    // the AI to reproduce it correctly.
    if (expectsHubIntro) {
      const hubMatch =
        structure.match(/^HUB-INTRO:[^\n]*/mi);

      if (hubMatch) {
        outputLines.push(hubMatch[0].trim());
        outputLines.push('');
      }
    }

    expectedSections.forEach(function(number) {
      outputLines.push(foundSections[number]);
    });

    outputLines.push('');
    outputLines.push(
      'Stage 1.5B complete. Waiting for Stage 1.5C.'
    );

    const cleanedH2s =
      outputLines.join('\n');

    // Only now overwrite CV.
    sheet.getRange(row, 100).setValue(cleanedH2s);

    return {
      success: true,
      message:
        'Validated H2/H3 heading set saved to Column CV — ' +
        expectedSections.length +
        ' governed section(s).'
    };

  } catch (e) {
    return {
      success: false,
      message: 'Error saving: ' + e.message
    };
  }
}

function saveStage15CEnrichedPlan(plan) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('posts');
    const row = sheet.getActiveRange().getRow();
    if (row < 2) {
      return {
        success: false,
        message: 'Select a data row first.'
      };
    }
    const governed =
      ce_getStage15CGovernedRequirements_();
    if (!governed.success) {
      return governed;
    }
    let cleanedPlan =
      ce_stripStage15CCodeFences_(plan);
    if (!cleanedPlan) {
      return {
        success: false,
        message:
          'W1.5C enriched plan is empty — CW not changed.'
      };
    }
    cleanedPlan =
      cleanedPlan
        .split('\n')
        .map(function(line) {
          if (
            line.trim().indexOf(
              'Internal link:'
            ) === 0
          ) {
            const prefix =
              line.substring(
                0,
                line.indexOf(
                  'Internal link:'
                ) +
                'Internal link:'.length
              );
            const value =
              line.substring(
                line.indexOf(
                  'Internal link:'
                ) +
                'Internal link:'.length
              ).trim();
            return (
              prefix +
              ' ' +
              bc_stripMarkdownUrl_(value)
            );
          }
          return line;
        })
        .join('\n')
        .trim();

    // Canonicalise governed fields from W1.5A before validation.
    // W1.5C may enrich a section, but it must never redefine
    // the TSM Requirement or Word Budget inherited from CU.
    governed.expectedSections.forEach(
      function(number) {
        const blockRegex =
          new RegExp(
            '(^|\\n)(SECTION\\s+' +
            number +
            ':[\\s\\S]*?)(?=\\nSECTION\\s+\\d+:|\\nStage 1\\.5C complete\\.|$)',
            'i'
          );

        cleanedPlan =
          cleanedPlan.replace(
            blockRegex,
            function(fullMatch, prefix, block) {
              let canonicalBlock =
                String(block || '');

              const governedTsm =
                governed.expectedTsmRequirements[number];

              const governedBudget =
                governed.expectedWordBudgets[number];

              if (
                /^TSM Requirement:/mi.test(
                  canonicalBlock
                )
              ) {
                canonicalBlock =
                  canonicalBlock.replace(
                    /^TSM Requirement:.*$/mi,
                    'TSM Requirement: ' +
                    governedTsm
                  );
              } else {
                const firstBreak =
                  canonicalBlock.indexOf(
                    '\n'
                  );

                if (firstBreak === -1) {
                  canonicalBlock +=
                    '\nTSM Requirement: ' +
                    governedTsm;
                } else {
                  canonicalBlock =
                    canonicalBlock.slice(
                      0,
                      firstBreak + 1
                    ) +
                    'TSM Requirement: ' +
                    governedTsm +
                    '\n' +
                    canonicalBlock.slice(
                      firstBreak + 1
                    );
                }
              }

              if (
                /^Word Budget:/mi.test(
                  canonicalBlock
                )
              ) {
                canonicalBlock =
                  canonicalBlock.replace(
                    /^Word Budget:.*$/mi,
                    'Word Budget: ' +
                    governedBudget
                  );
              } else {
                const tsmLineRegex =
                  /^TSM Requirement:.*$/mi;

                canonicalBlock =
                  canonicalBlock.replace(
                    tsmLineRegex,
                    function(tsmLine) {
                      return (
                        tsmLine +
                        '\nWord Budget: ' +
                        governedBudget
                      );
                    }
                  );
              }

              return (
                prefix +
                canonicalBlock
              );
            }
          );
      }
    );

    const errors = [];
    const foundSections = {};
    const sectionRegex =
      /^SECTION\s+(\d+):\s*Heading\s+(H[23]):\s*(.+)$/gmi;
    let match;
    while (
      (
        match =
          sectionRegex.exec(
            cleanedPlan
          )
      ) !== null
    ) {
      const number =
        Number(
          match[1]
        );
      const level =
        String(
          match[2] || 'H2'
        ).toUpperCase();
      const heading =
        String(
          match[3] || ''
        ).trim();
      if (
        foundSections[number]
      ) {
        errors.push(
          'SECTION ' +
          number +
          ' appears more than once.'
        );
      } else {
        foundSections[number] = {
          level: level,
          heading: heading
        };
      }
    }
    governed.expectedSections.forEach(
      function(number) {
        if (
          !foundSections[number]
        ) {
          errors.push(
            'SECTION ' +
            number +
            ' is missing.'
          );
          return;
        }
        if (
          foundSections[number].heading !==
          governed.expectedH2s[number]
        ) {
          errors.push(
            'SECTION ' +
            number +
            ' heading does not exactly match W1.5B.'
          );
        }

        if (
          foundSections[number].level !==
          governed.expectedHeadingLevels[number]
        ) {
          errors.push(
            'SECTION ' +
            number +
            ' heading level does not exactly match W1.5B.'
          );
        }
      }
    );
    Object.keys(
      foundSections
    ).forEach(
      function(key) {
        const number =
          Number(key);
        if (
          governed.expectedSections.indexOf(
            number
          ) === -1
        ) {
          errors.push(
            'Unexpected SECTION ' +
            number +
            ' is not present in W1.5A.'
          );
        }
      }
    );
    if (
      governed.expectsHubIntro &&
      !/^HUB-INTRO:/mi.test(
        cleanedPlan
      )
    ) {
      errors.push(
        'HUB-INTRO is required but missing.'
      );
    }
    const sectionBlockRegex =
      /(?:^|\n)SECTION\s+(\d+):[\s\S]*?(?=\nSECTION\s+\d+:|\nStage 1\.5C complete\. Waiting for Stage 1\.5D\.|$)/gi;
    let blockMatch;
    while (
      (
        blockMatch =
          sectionBlockRegex.exec(
            cleanedPlan
          )
      ) !== null
    ) {
      const number =
        Number(
          blockMatch[1]
        );
      const block =
        blockMatch[0];
      const requiredFields = [
        'TSM Requirement:',
        'Word Budget:',
        'Content Brief:',
        'Visual Pattern:',
        'List eligible:',
        'Images:',
        'Videos:',
        'Entities to include:',
        'Internal link:'
      ];
      requiredFields.forEach(
        function(field) {
          const escapedField =
            field.replace(
              /[.*+?^${}()|[\]\\]/g,
              '\\$&'
            );
          const fieldRegex =
            new RegExp(
              '^' +
              escapedField,
              'mi'
            );
          if (
            !fieldRegex.test(
              block
            )
          ) {
            errors.push(
              'SECTION ' +
              number +
              ' is missing required field "' +
              field +
              '".'
            );
          }
        }
      );
    }
        governed.expectedSections.forEach(
      function(number) {
        const blockRegex =
        new RegExp(
          '(?:^|\\n)SECTION\\s+' +
          number +
          ':[\\s\\S]*?(?=\\nSECTION\\s+\\d+:|\\nStage 1\\.5C complete\\.|$)',
          'i'
        );

        const blockMatch =
          cleanedPlan.match(blockRegex);

        if (!blockMatch) {
          return;
        }

        const block =
          blockMatch[0];

        const tsmMatch =
          block.match(
            /^TSM Requirement:\s*(.+)$/mi
          );

        const budgetMatch =
          block.match(
            /^Word Budget:\s*(.+)$/mi
          );

        if (
          tsmMatch &&
          tsmMatch[1].trim() !==
          governed.expectedTsmRequirements[number]
        ) {
          errors.push(
            'SECTION ' +
            number +
            ' TSM Requirement does not exactly match W1.5A.'
          );
        }

        if (
          budgetMatch &&
          budgetMatch[1].trim() !==
          governed.expectedWordBudgets[number]
        ) {
          errors.push(
            'SECTION ' +
            number +
            ' Word Budget does not exactly match W1.5A.'
          );
        }
      }
    );

    

    const completionLine =
  'Stage 1.5C complete. Waiting for Stage 1.5D.';

    if (
      !cleanedPlan.endsWith(
        completionLine
      )
    ) {
      errors.push(
        'W1.5C completion line is missing or incorrect.'
      );
    }
    if (
      errors.length > 0
    ) {
      return {
        success: false,
        message:
          'W1.5C enriched plan rejected — CW not changed. ' +
          errors.join(' | ')
      };
    }
    // Only overwrite CW after the whole plan validates.
    sheet
      .getRange(row, 101)
      .setValue(cleanedPlan);
    logPipelineResume(
      "W1.5C — Enriched Plan",
      ""
    );
    return {
      success: true,
      message:
        'Validated enriched plan saved to Column CW — ' +
        governed.expectedSections.length +
        ' governed section(s).'
    };
  } catch (e) {
    return {
      success: false,
      message:
        'Error saving W1.5C enriched plan: ' +
        e.message
    };
  }
}

// saveW15DUpdatedPlan lives in ce_Stage15D_LateralLinks.gs and saves to col 102 (CX)
function saveStage15EOptimizations(optimizations) {
  try {

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('posts');
    const row = sheet.getActiveRange().getRow();

    if (row < 2) {
      return {
        success: false,
        message: 'Select a data row first.'
      };
    }

    const sourcePlan =
      String(
        sheet.getRange(row, 102).getValue() || ''
      ).trim();

    if (!sourcePlan) {
      return {
        success: false,
        message:
          'Column CX is empty — W1.5D must complete before W1.5E.'
      };
    }

    let cleaned =
      String(optimizations || '')
        .trim()
        .replace(/^```json\s*/i, '')
        .replace(/^```\s*/i, '')
        .replace(/\s*```$/i, '')
        .trim();

    if (!cleaned) {
      return {
        success: false,
        message:
          'W1.5E output is empty — CY not changed.'
      };
    }

    let parsed;

    try {
      parsed =
        JSON.parse(cleaned);
    } catch (jsonError) {
      return {
        success: false,
        message:
          'Invalid W1.5E JSON — CY not changed. ' +
          jsonError.message
      };
    }

    if (
      !parsed ||
      !Array.isArray(parsed.sections)
    ) {
      return {
        success: false,
        message:
          'Invalid W1.5E JSON — missing "sections" array. CY not changed.'
      };
    }

    const expectedSections = [];

    const sectionRegex =
      /^SECTION\s+(\d+):/gmi;

    let sectionMatch;

    while (
      (
        sectionMatch =
          sectionRegex.exec(sourcePlan)
      ) !== null
    ) {

      const number =
        Number(sectionMatch[1]);

      if (
        expectedSections.indexOf(number) === -1
      ) {
        expectedSections.push(number);
      }
    }

    expectedSections.sort(function(a, b) {
      return a - b;
    });

    if (expectedSections.length === 0) {
      return {
        success: false,
        message:
          'Could not determine governed section numbers from W1.5D in CX.'
      };
    }

    const requiredFields = [
      'section_number',
      'opening_anchor',
      'information_gain_directive',
      'scannability_pattern',
      'featured_snippet_target',
      'list_opportunity',
      'blockquote_flag'
    ];

    const foundSections = {};
    const errors = [];

    parsed.sections.forEach(
      function(section, index) {

        if (
          !section ||
          typeof section !== 'object' ||
          Array.isArray(section)
        ) {
          errors.push(
            'JSON section entry ' +
            (index + 1) +
            ' is not a valid object.'
          );
          return;
        }

        requiredFields.forEach(
          function(field) {

            if (!(field in section)) {
              errors.push(
                'JSON section entry ' +
                (index + 1) +
                ' is missing required field "' +
                field +
                '".'
              );
            }
          }
        );

        const number =
          Number(
            section.section_number
          );

        if (
          !Number.isInteger(number)
        ) {
          errors.push(
            'JSON section entry ' +
            (index + 1) +
            ' has an invalid section_number.'
          );
          return;
        }

        if (
          foundSections[number]
        ) {
          errors.push(
            'SECTION ' +
            number +
            ' appears more than once in W1.5E JSON.'
          );
          return;
        }

        foundSections[number] = true;
      }
    );

    expectedSections.forEach(
      function(number) {

        if (!foundSections[number]) {
          errors.push(
            'SECTION ' +
            number +
            ' is missing from W1.5E JSON.'
          );
        }
      }
    );

    Object.keys(foundSections).forEach(
      function(key) {

        const number =
          Number(key);

        if (
          expectedSections.indexOf(number) === -1
        ) {
          errors.push(
            'Unexpected SECTION ' +
            number +
            ' is present in W1.5E JSON.'
          );
        }
      }
    );

    if (
      parsed.sections.length !==
      expectedSections.length
    ) {
      errors.push(
        'W1.5E JSON contains ' +
        parsed.sections.length +
        ' section object(s), but W1.5D requires ' +
        expectedSections.length +
        '.'
      );
    }

    if (errors.length > 0) {
      return {
        success: false,
        message:
          'W1.5E JSON rejected — CY not changed. ' +
          errors.join(' | ')
      };
    }

    // Only overwrite CY after full validation.
    sheet
      .getRange(row, 103)
      .setValue(cleaned);

    logPipelineResume(
      "W1.5E — Google Optimizations",
      ""
    );

    return {
      success: true,
      message:
        'Validated W1.5E JSON saved to Column CY — ' +
        expectedSections.length +
        ' governed section(s).'
    };

  } catch (e) {

    return {
      success: false,
      message:
        'Error saving W1.5E optimizations: ' +
        e.message
    };
  }
}

