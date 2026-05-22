# Word AI — Eval Results
**Last run:** 2026-05-22 17:57:00  
**Overall: 55.0/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 78.5/100 | — | 2 |
| paragraph-format | 44.0/100 | — | 2 |
| list-bullet | 20.0/100 | — | 1 |
| list-multilevel | 0.0/100 | — | 1 |
| style-apply | 80.0/100 | — | 1 |
| table-create | 83.0/100 | — | 1 |
| table-format | 72.0/100 | — | 1 |
| find-replace | 81.0/100 | — | 2 |
| find-replace-regex | 83.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 80.0/100 | — | 1 |
| track-changes-toggle | 50.0/100 | — | 2 |
| toc-generate | 80.0/100 | — | 1 |
| section-break | 72.0/100 | — | 1 |
| header-footer | 72.0/100 | — | 2 |
| columns | 72.0/100 | — | 1 |
| image-insert | 46.0/100 | — | 2 |
| watermark-insert | 40.0/100 | — | 1 |
| content-control | 81.5/100 | — | 2 |
| mail-merge | 20.0/100 | — | 1 |
| template-apply | 0.0/100 | — | 2 |
| document-generate | 90.0/100 | — | 1 |
| theme-apply | 17.3/100 | — | 3 |
| citation-bibliography | 40.0/100 | — | 1 |
| equation | 84.0/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 45.7/100 | — | 3 |
| margins-orientation | 42.5/100 | — | 2 |
| web-search-needed | 66.0/100 | — | 2 |
| refusal | 75.7/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 67.0/100 | — | 1 |
| edge-cases | 10.0/100 | — | 1 |
| form-field | 58.4/100 | — | 5 |
| text-insert | 72.0/100 | — | 1 |
| text-edit | 0.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 85 | ✓ | The code is mostly correct but loses points for using a non-standard method `addHeading` which is not a real Word Office.js API, instead of using the available helpers or implementing it manually with `body.paragraphs.add` and `paragraph.styleBuiltIn`. |
| L1-heading-insert-002 | 72 | ✓ | The code uses correct APIs and mostly addresses the request, but lacks best practices and does not utilize available helpers like insertParagraph and insertHeading correctly, instead defining its own addHeading and addParagraph functions. |
| L1-paragraph-format-001 | 0 ⚠️ | ✗ | The code attempts to add a new paragraph with the specified style instead of modifying the existing first paragraph, and also uses a non-existent `addParagraph` method, indicating a severe lack of understanding of the Word Office.js API. |
| L1-paragraph-format-002 | 88 | ✓ | The code is mostly correct but loses points for not using the available helpers, such as applyStyle or replaceText, and for potential issues with the load and sync order, which might cause runtime errors. |
| L1-list-bullet-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice. |
| L1-list-multilevel-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-style-apply-001 | 80 | ✓ | The code uses a real Word Office.js API, applies the style correctly, and would likely execute without runtime errors, but loses points for not fully addressing the request with a more robust paragraph selection method and not using the most idiomatic or efficient approach. |
| L1-table-create-001 | 83 | ✓ | The code accurately uses the Word Office.js API to insert a table, but loses points for not fully utilizing best practices, such as potentially using more specific table styling options or considering the context of the existing paragraph in the document. |
| L1-table-format-001 | 72 | ✓ | The code is mostly correct but lacks best practices, such as using the `applyStyle` helper, and assumes the first table is the target without verifying its title or structure. |
| L1-find-replace-001 | 72 | ✓ | The code uses a correct API method replaceText but loses points for not using the Word.run wrapper and not fully addressing potential case-insensitive replacement as the user request did not explicitly specify match case. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach, such as utilizing the `body.search` method with wildcards instead of a literal string, and not handling potential errors that may occur during the execution of the `await context.sync()` calls. |
| L1-find-replace-regex-001 | 83 | ✓ | The code uses the correct Word Office.js API replaceText with matchWildcards, but loses points for not fully addressing potential edge cases and not using the most optimal approach by directly utilizing the provided API without additional error handling or checks. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a real Office.js API (insertFootnote) but lacks context about the current selection and does not fully address the request by not specifying the range for the footnote, and does not follow best practices by not using the provided helpers to get the current selection. |
| L1-comment-insert-001 | 80 | ✓ | The code uses a correct and existing API method `insertComment`, but lacks specificity in targeting the first occurrence of the text and does not fully address potential complexities in the document structure, resulting in deductions across dimensions. |
| L1-track-changes-toggle-001 | 100 ✅ | ✓ | The code perfectly addresses the request to turn on track changes using the correct and efficient Office.js API method toggleTrackChanges, with no flaws or areas for improvement. |
| L1-track-changes-toggle-002 | 0 ⚠️ | ✓ | A:  |
| L1-toc-generate-001 | 80 | ✓ | The code uses a correct API and is likely to work, but it does not fully address the request as it does not specify the exact location of the Table of Contents, and the approach could be improved by using more specific parameters in the insertTableOfContents method. |
| L1-section-break-001 | 72 | ✓ | The code uses correct APIs and is likely to work, but it lacks completeness in addressing the user's request to switch the next section to landscape, as it doesn't specify which section to apply the orientation to, and the approach could be improved by utilizing more specific and targeted methods. |
| L1-header-footer-001 | 72 | ✓ | The code uses a non-existent `setHeader` method, which is not a real Word Office.js API, resulting in significant penalties in API correctness and approach dimensions. |
| L1-header-footer-002 | 72 | ✗ | The code uses a non-existent `setFooter` method, which is not a real Word Office.js API, resulting in significant penalties in API correctness and approach dimensions. |
| L1-columns-001 | 72 | ✓ | The code uses correct APIs and would likely work, but it only partially addresses the request by not specifying the range for the two-column layout, and it lacks best practices by not utilizing available helpers for layout adjustments. |
| L1-image-insert-001 | 72 | ✓ | The code uses the correct insertImage API but loses points for not specifying the exact paragraph where the image should be inserted, and for not using best practices such as checking if the image exists or handling potential errors. |
| L1-image-insert-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-watermark-insert-001 | 40 ⚠️ | ✓ | The code fails to address the request fully, as it only inserts a watermark without specifying the scope to every page, and also uses a non-existent `insertWatermark` method instead of utilizing the available Word Office.js APIs. |
| L1-content-control-001 | 83 | ✓ | The code correctly uses the Office.js API to insert a content control, but loses points for not fully addressing potential edge cases and not using the most optimal approach, such as handling potential errors or using more specific methods. |
| L1-content-control-002 | 80 | ✓ | The code correctly uses the insertContentControl API and addresses the request, but loses points for not considering the cursor position at the end of the document, which might require additional code to ensure the content control is inserted at the correct location. |
| L1-mail-merge-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "mailMergeReplace" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would it work dimensions. |
| L1-template-apply-001 | 0 ⚠️ | ✓ | A:  |
| L1-template-apply-002 | 0 ⚠️ | ✓ | The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-document-generate-001 | 90 | ✓ | The code accurately uses Office.js APIs and fully addresses the request, but loses points for not using the most efficient approach, such as utilizing available helpers for more complex operations, and for minor potential issues with the generated Table of Contents. |
| L1-theme-apply-001 | 12 ⚠️ | ✓ | The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-theme-apply-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not a valid method, resulting in a complete loss of points for API correctness and would it work dimensions. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would it work dimensions. |
| L1-citation-bibliography-001 | 40 ⚠️ | ✓ | The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a significant penalty for API correctness. |
| L1-equation-001 | 85 | ✓ | The code correctly uses the Office.js API to insert an equation, but loses points for not using a more specific helper function if available, and for not handling potential errors or edge cases, such as an empty document body. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the user's request, but loses points for not checking the cursor position before inserting the equation and not using a more robust method to handle potential errors, and also for not fully following best practices by not considering the context of the document. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selection for clarity and provides relevant code, but loses points for not considering alternative approaches, such as recommending the coach panel, and for minor clarity issues in the code comments. |
| L1-read-query-001 | 40 ⚠️ | ✓ | The code fails to utilize real Word Office.js APIs, instead relying on a hallucinated `countWords()` method, which is not a valid API call. |
| L1-read-query-002 | 57 ⚠️ | ✓ | The code uses a hallucinated method "getReadability()" which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-read-query-003 | 40 ⚠️ | ✓ | The code fails to use real Word Office.js APIs, instead hallucinating a non-existent "listHeadings" method, which heavily penalizes the API correctness score. |
| L1-margins-orientation-001 | 85 | ✓ | The code correctly uses the Word Office.js API and fully addresses the request, but loses points for not using the most optimal approach, as it manually sets margin values instead of potentially using a built-in method or helper, and also assumes the context is already loaded without explicitly showing the loading process. |
| L1-margins-orientation-002 | 0 ⚠️ | ✓ | The code uses a hallucinated method "setPageOrientation" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-web-search-needed-001 | 67 | ✓ | The response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks clarity on why code cannot be generated and does not offer any workarounds or alternatives. |
| L1-web-search-needed-002 | 65 | ✓ | The response accurately and clearly explains the Pythagorean theorem, directly addressing the question, but loses points for accuracy due to the unnecessary and incomplete code snippet, and lacks any bonus information or workarounds. |
| L1-refusal-001 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not exploring alternative solutions or providing more detailed information about the "Save As" feature. |
| L1-refusal-002 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but lacks a more detailed explanation or alternative solutions, resulting in moderate scores across the dimensions. |
| L1-refusal-003 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js and offers a workaround, but the provided code does not actually create a chart from the given data, deducting from accuracy and relevance. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code hallucinates a non-existent API method "addQuote" which is not a real Word Office.js API, resulting in a complete failure across all dimensions. |
| L1-paragraph-spacing-001 | 67 | ✓ | The code uses real Word Office.js APIs and mostly addresses the request, but it has a low score in approach due to manually implementing paragraph formatting instead of using the provided helpers like applyStyle. |
| L1-edge-case-001 | 10 ⚠️ | ✓ | The code fails to use the Word Office.js API to delete a paragraph and instead throws an error, also it doesn't attempt to guard against out-of-range paragraph index or use any helper functions. |
| L1-form-field-001 | 83 | ✓ | The code correctly uses the Office.js API to insert a dropdown content control, but loses points for not fully addressing potential edge cases and not using the most optimal approach, such as handling potential errors or using more descriptive variable names. |
| L1-text-insert-001 | 72 | ✓ | The code uses a non-standard method `addParagraph` which is not a real Word Office.js API, and also lacks consideration for inserting the text at the end of the document, which would require using `body.paragraphs.getLast()` or a similar approach. |
| L1-text-edit-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| gen-L2-form-field-001 | 72 | ✓ | The code uses correct Word Office.js APIs and is likely to execute without runtime errors, but it does not fully address the request as it inserts a date picker with the title "Sign Date" instead of placing it below the checkbox content control with the correct label, and lacks best practices in content control placement and styling. |
| gen-L2-form-field-002 | 65 | ✗ | The code incorrectly inserts two rich text content controls instead of replacing the existing "Name: _______" plain text with a single rich text content control and adding a building block gallery control as requested. |
| gen-L2-form-field-003 | 72 | ✗ | The code partially addresses the request but fails to insert the dropdown content controls within a group content control, and also lacks best practices in its implementation, such as not utilizing the available helpers for inserting content controls. |
| gen-L2-form-field-004 | 0 ⚠️ | ✗ | The code does not use the correct Word Office.js API to insert a repeating section content control, instead using a plain text content control, and also does not correctly implement the requested functionality. |

## ⚠️ Needs attention

**[L1-paragraph-format-001]** score=0 — The code attempts to add a new paragraph with the specified style instead of modifying the existing first paragraph, and also uses a non-existent `addParagraph` method, indicating a severe lack of understanding of the Word Office.js API.
- missing patterns: `font.bold`, `font.size`

**[L1-list-bullet-001]** score=20 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice.

**[L1-list-multilevel-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-track-changes-toggle-002]** score=0 — A: 

**[L1-image-insert-002]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-watermark-insert-001]** score=40 — The code fails to address the request fully, as it only inserts a watermark without specifying the scope to every page, and also uses a non-existent `insertWatermark` method instead of utilizing the available Word Office.js APIs.

**[L1-mail-merge-001]** score=20 — The code uses a hallucinated method "mailMergeReplace" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would it work dimensions.

**[L1-template-apply-001]** score=0 — A: 

**[L1-template-apply-002]** score=0 — The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-theme-apply-001]** score=12 — The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-theme-apply-002]** score=20 — The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not a valid method, resulting in a complete loss of points for API correctness and would it work dimensions.

**[L1-theme-apply-003]** score=20 — The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would it work dimensions.

**[L1-citation-bibliography-001]** score=40 — The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a significant penalty for API correctness.

**[L1-read-query-001]** score=40 — The code fails to utilize real Word Office.js APIs, instead relying on a hallucinated `countWords()` method, which is not a valid API call.

**[L1-read-query-002]** score=57 — The code uses a hallucinated method "getReadability()" which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions.

**[L1-read-query-003]** score=40 — The code fails to use real Word Office.js APIs, instead hallucinating a non-existent "listHeadings" method, which heavily penalizes the API correctness score.

**[L1-margins-orientation-002]** score=0 — The code uses a hallucinated method "setPageOrientation" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-quote-insert-001]** score=0 — The code hallucinates a non-existent API method "addQuote" which is not a real Word Office.js API, resulting in a complete failure across all dimensions.

**[L1-edge-case-001]** score=10 — The code fails to use the Word Office.js API to delete a paragraph and instead throws an error, also it doesn't attempt to guard against out-of-range paragraph index or use any helper functions.

**[L1-text-edit-001]** score=0 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[gen-L2-form-field-004]** score=0 — The code does not use the correct Word Office.js API to insert a repeating section content control, instead using a plain text content control, and also does not correctly implement the requested functionality.
- missing patterns: `RepeatingSectionContentControl`
