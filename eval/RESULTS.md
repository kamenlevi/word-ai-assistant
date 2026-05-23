# Word AI — Eval Results
**Last run:** 2026-05-23 17:06:57  
**Overall: 56.2/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 33.5/100 | — | 2 |
| paragraph-format | 41.5/100 | — | 2 |
| list-bullet | 20.0/100 | — | 1 |
| list-multilevel | 0.0/100 | — | 1 |
| style-apply | 80.0/100 | — | 1 |
| table-create | 90.0/100 | — | 1 |
| table-format | 0.0/100 | — | 1 |
| find-replace | 81.0/100 | — | 2 |
| find-replace-regex | 83.0/100 | — | 1 |
| footnote-insert | 0.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 88.5/100 | — | 2 |
| toc-generate | 72.0/100 | — | 1 |
| section-break | 57.0/100 | — | 1 |
| header-footer | 60.0/100 | — | 2 |
| columns | 83.0/100 | — | 1 |
| image-insert | 49.5/100 | — | 2 |
| watermark-insert | 0.0/100 | — | 1 |
| content-control | 81.5/100 | — | 2 |
| mail-merge | 79.0/100 | — | 1 |
| template-apply | 54.0/100 | — | 2 |
| document-generate | 79.0/100 | — | 1 |
| theme-apply | 15.0/100 | — | 3 |
| citation-bibliography | 40.0/100 | — | 1 |
| equation | 83.0/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 39.0/100 | — | 3 |
| margins-orientation | 77.5/100 | — | 2 |
| web-search-needed | 69.5/100 | — | 2 |
| refusal | 77.7/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 72.0/100 | — | 1 |
| edge-cases | 25.0/100 | — | 1 |
| form-field | 70.4/100 | — | 5 |
| text-insert | 80.0/100 | — | 1 |
| text-edit | 0.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 67 | ✓ | The code uses a non-standard method `addHeading` which is not a real Word Office.js API, indicating a lack of best practices and reliance on hallucinated methods. |
| L1-heading-insert-002 | 0 ⚠️ | ✓ | The code uses hallucinated methods "addHeading" and "addParagraph" which are not real Word Office.js APIs, resulting in a complete failure to meet the requirements. |
| L1-paragraph-format-001 | 0 ⚠️ | ✗ | The code uses a hallucinated method `addParagraph` which is not a real Word Office.js API, and also attempts to modify an existing paragraph's text and style in a way that is not supported by the API. |
| L1-paragraph-format-002 | 83 | ✓ | The code is mostly correct but loses points for not using the built-in applyStyle helper to set the font size and for not checking if the first paragraph is indeed the title before applying the changes. |
| L1-list-bullet-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice. |
| L1-list-multilevel-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the API correctness and best practices criteria. |
| L1-style-apply-001 | 80 | ✓ | The code uses a real Word Office.js API, applies the style to the correct paragraph, and would likely execute without runtime errors, but lacks completeness and best practices, such as using the paragraph.styleBuiltIn property directly. |
| L1-table-create-001 | 90 | ✓ | The code accurately uses the Word Office.js API to insert a table, but loses points for not fully addressing the request by only populating three of the four rows, and for not using the most optimal approach by directly specifying the table style instead of potentially using a more flexible method. |
| L1-table-format-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "styleTable" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions. |
| L1-find-replace-001 | 72 | ✓ | The code uses a non-existent `replaceText` method, which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach and for potential issues with the load and sync order, which could lead to runtime errors. |
| L1-find-replace-regex-001 | 83 | ✓ | The code uses the correct Word Office.js API and mostly addresses the request, but loses points for not handling potential edge cases and not using the most optimal approach, such as checking the scope of the replacement or using a more specific wildcard pattern. |
| L1-footnote-insert-001 | 0 ⚠️ | ✓ | A:  |
| L1-comment-insert-001 | 72 | ✓ | The code uses a correct API method to insert a comment, but lacks specificity in targeting the first occurrence of the text and does not fully address the request by not ensuring the comment is added to the correct location within the document. |
| L1-track-changes-toggle-001 | 97 ✅ | ✓ | The code is nearly perfect, using the correct toggleTrackChanges API and following best practices, but loses a few points for not being entirely necessary, as the toggleTrackChanges helper is already provided by the Office.js API. |
| L1-track-changes-toggle-002 | 80 | ✓ | The code is mostly correct but lacks the Word.run context and may not fully handle potential errors, also it does not use the best practice of wrapping the code in a Word.run function to ensure the context is properly loaded before executing the acceptAllRevisions method. |
| L1-toc-generate-001 | 72 | ✓ | The code uses a correct API but lacks best practices, such as using the provided helpers, and does not fully address the request as it does not specify the location of the Table of Contents, which should be at the top of the document. |
| L1-section-break-001 | 57 ⚠️ | ✓ | The code uses non-existent Word Office.js APIs, such as insertSectionBreak and setPageOrientation, which are not part of the real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-header-footer-001 | 80 | ✓ | The code is mostly correct but lacks consideration for the specific section or page where the header should be applied, and it does not follow best practices for setting headers in Word using the Office.js API. |
| L1-header-footer-002 | 40 ⚠️ | ✓ | The code uses a hallucinated method `addPageNumbers` which is not a real Word Office.js API, resulting in a severe penalty in API correctness. |
| L1-columns-001 | 83 | ✓ | The code is mostly correct but loses points for not using the Word.run context and potentially not handling the current selection or range, which might lead to unexpected behavior or errors. |
| L1-image-insert-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would_it_work dimensions. |
| L1-image-insert-002 | 79 | ✓ | The code uses the correct insertImage API and mostly addresses the request, but loses points for not fully handling potential errors or edge cases, such as image insertion failures or context synchronization issues. |
| L1-watermark-insert-001 | 0 ⚠️ | ✓ | The code does not use any real Word Office.js APIs, instead calling a non-existent `insertWatermark` method, which is a significant flaw in API correctness, completeness, and approach. |
| L1-content-control-001 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the request, but loses points for not using the Word.run method to ensure the code is executed in the correct context and for not handling potential errors, and also for not being a complete best practice implementation. |
| L1-content-control-002 | 80 | ✓ | The code correctly uses the Office.js API to insert a checkbox content control, but loses points for not fully addressing potential context and loading issues, and not using the most robust best practices for error handling and synchronization. |
| L1-mail-merge-001 | 79 | ✓ | The code uses the correct mailMergeReplace API but loses points for not fully addressing potential errors, missing input validation, and not explicitly handling the template context, while also not demonstrating the most optimal approach by directly using the provided records without considering potential edge cases. |
| L1-template-apply-001 | 25 ⚠️ | ✓ | The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a significant penalty for API correctness. |
| L1-template-apply-002 | 83 | ✓ | The code correctly applies the memo template with the specified fields, but loses points for API correctness and approach due to the use of a potentially non-existent `applyTemplate` method, which is not a standard Word Office.js API. |
| L1-document-generate-001 | 79 | ✓ | The code is mostly correct but lacks actual paragraph content and only includes headings, which reduces its completeness and approach scores, while its correctness and potential to work are still relatively high due to the proper use of Word Office.js APIs like insertTableOfContents and addHeading. |
| L1-theme-apply-001 | 5 ⚠️ | ✓ | The code hallucinates non-existent `designTheme` and `applyTheme` methods, which are not part of the real Word Office.js APIs, resulting in a complete failure in API correctness and overall approach. |
| L1-theme-apply-002 | 20 ⚠️ | ✓ | The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would it work dimensions. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, and also does not use the correct API "designTheme" to modify the theme. |
| L1-citation-bibliography-001 | 40 ⚠️ | ✓ | The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a significant penalty for API correctness. |
| L1-equation-001 | 83 | ✓ | The code uses a correct and existing helper method insertEquation, but loses points for not using Word.run to ensure the code is executed in the correct context and for not fully addressing potential errors or loading issues. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the user's request, but loses points for not using a more robust method to handle equation insertion, such as utilizing the Word JavaScript API's built-in equation builder or providing more context for the equation. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selected text for clarity, but loses points for not considering alternative approaches, such as recommending the coach panel, and for using a somewhat generic phrase "Rewriting the current selection for clarity" that could be more descriptive. |
| L1-read-query-001 | 20 ⚠️ | ✓ | The code hallucinates a non-existent `countWords()` method, which is not a real Word Office.js API, and does not provide a correct implementation to count the words in the document. |
| L1-read-query-002 | 57 ⚠️ | ✓ | The code uses a hallucinated method `getReadability()` which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-read-query-003 | 40 ⚠️ | ✓ | The code heavily relies on a hallucinated method "listHeadings()" which is not a real Word Office.js API, resulting in a significant penalty for API correctness. |
| L1-margins-orientation-001 | 83 | ✓ | The code correctly uses the Word Office.js API and addresses the user's request, but loses points for not using the most efficient or best practice approach, such as utilizing available helpers, and for potential issues with variable scope and load/sync order. |
| L1-margins-orientation-002 | 72 | ✓ | The code uses a non-existent `setPageOrientation` method, which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-web-search-needed-001 | 67 | ✓ | The response accurately identifies the need for a web search, is clear and relevant, but lacks a meaningful bonus, such as suggesting an alternative approach or noting a Word-specific limitation. |
| L1-web-search-needed-002 | 72 | ✓ | The response accurately and clearly explains the Pythagorean theorem, but the provided code is unnecessary and somewhat confusing since the user didn't request code, and there's no notable bonus information offered. |
| L1-refusal-001 | 83 | ✓ | The response is mostly accurate and clear, but loses points for not providing more detailed workaround instructions or alternatives, and for a slight lack of clarity in conveying the limitation of the Word Office.js API. |
| L1-refusal-002 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides clear workarounds, but loses points for not offering more detailed or creative solutions in the bonus category. |
| L1-refusal-003 | 67 | ✓ | The response accurately identifies the limitation of Word Office.js in creating Excel-style charts, but loses points for not providing a more detailed or relevant workaround, such as suggesting how to generate the chart image from the given data. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-paragraph-spacing-001 | 72 | ✓ | The code uses correct APIs and addresses the request, but has issues with approach, as it manually iterates over paragraphs and applies line spacing instead of using available helpers, and also calls an undefined `applyStyle` method. |
| L1-edge-case-001 | 25 ⚠️ | ✓ | The code fails to use the Word Office.js API and instead throws a generic error, not attempting to access or delete the paragraph using the available APIs. |
| L1-form-field-001 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the user's request, but loses points for not using a more descriptive title and for not handling potential errors or edge cases, and also for not fully following best practices in terms of code organization and readability. |
| L1-text-insert-001 | 80 | ✓ | The code uses a real Word Office.js API (addParagraph) but loses points for not fully addressing the request (inserting at the end of the document, not necessarily as a new paragraph) and not using the most optimal approach (inserting text directly into the last paragraph). |
| L1-text-edit-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| gen-L2-form-field-001 | 72 | ✓ | The code uses correct Office.js APIs and is mostly complete, but loses points for not fully addressing the request to place the date picker content control below the checkbox, and for not using best practices such as inserting a new paragraph for the date picker content control. |
| gen-L2-form-field-002 | 65 | ✗ | The code partially addresses the request but fails to correctly replace the existing plain text with a rich text content control and incorrectly adds a building block gallery control with the title "Address" instead of adding it to the document as a separate control. |
| gen-L2-form-field-003 | 60 | ✗ | The code fails to fully address the request by not grouping the two dropdown content controls together as required, instead inserting them separately, and also does not utilize the best practices and available helpers for inserting content controls. |
| gen-L2-form-field-004 | 72 | ✗ | The code partially addresses the request but fails to fully implement a repeating section content control, instead inserting a rich text content control, and also lacks best practices in using available helpers. |

## ⚠️ Needs attention

**[L1-heading-insert-002]** score=0 — The code uses hallucinated methods "addHeading" and "addParagraph" which are not real Word Office.js APIs, resulting in a complete failure to meet the requirements.

**[L1-paragraph-format-001]** score=0 — The code uses a hallucinated method `addParagraph` which is not a real Word Office.js API, and also attempts to modify an existing paragraph's text and style in a way that is not supported by the API.
- missing patterns: `font.bold`, `font.size`

**[L1-list-bullet-001]** score=20 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice.

**[L1-list-multilevel-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the API correctness and best practices criteria.

**[L1-table-format-001]** score=0 — The code uses a hallucinated method "styleTable" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions.

**[L1-footnote-insert-001]** score=0 — A: 

**[L1-section-break-001]** score=57 — The code uses non-existent Word Office.js APIs, such as insertSectionBreak and setPageOrientation, which are not part of the real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions.

**[L1-header-footer-002]** score=40 — The code uses a hallucinated method `addPageNumbers` which is not a real Word Office.js API, resulting in a severe penalty in API correctness.

**[L1-image-insert-001]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would_it_work dimensions.

**[L1-watermark-insert-001]** score=0 — The code does not use any real Word Office.js APIs, instead calling a non-existent `insertWatermark` method, which is a significant flaw in API correctness, completeness, and approach.

**[L1-template-apply-001]** score=25 — The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a significant penalty for API correctness.

**[L1-theme-apply-001]** score=5 — The code hallucinates non-existent `designTheme` and `applyTheme` methods, which are not part of the real Word Office.js APIs, resulting in a complete failure in API correctness and overall approach.

**[L1-theme-apply-002]** score=20 — The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would it work dimensions.

**[L1-theme-apply-003]** score=20 — The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, and also does not use the correct API "designTheme" to modify the theme.

**[L1-citation-bibliography-001]** score=40 — The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a significant penalty for API correctness.

**[L1-read-query-001]** score=20 — The code hallucinates a non-existent `countWords()` method, which is not a real Word Office.js API, and does not provide a correct implementation to count the words in the document.

**[L1-read-query-002]** score=57 — The code uses a hallucinated method `getReadability()` which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions.

**[L1-read-query-003]** score=40 — The code heavily relies on a hallucinated method "listHeadings()" which is not a real Word Office.js API, resulting in a significant penalty for API correctness.

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-edge-case-001]** score=25 — The code fails to use the Word Office.js API and instead throws a generic error, not attempting to access or delete the paragraph using the available APIs.

**[L1-text-edit-001]** score=0 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.
