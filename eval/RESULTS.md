# Word AI — Eval Results
**Last run:** 2026-05-21 03:46:14  
**Overall: 58.1/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 42.5/100 | — | 2 |
| paragraph-format | 46.0/100 | — | 2 |
| list-bullet | 0.0/100 | — | 1 |
| list-multilevel | 0.0/100 | — | 1 |
| style-apply | 83.0/100 | — | 1 |
| table-create | 90.0/100 | — | 1 |
| table-format | 67.0/100 | — | 1 |
| find-replace | 85.0/100 | — | 2 |
| find-replace-regex | 83.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 91.5/100 | — | 2 |
| toc-generate | 72.0/100 | — | 1 |
| section-break | 72.0/100 | — | 1 |
| header-footer | 41.5/100 | — | 2 |
| columns | 72.0/100 | — | 1 |
| image-insert | 20.0/100 | — | 2 |
| watermark-insert | 40.0/100 | — | 1 |
| content-control | 81.5/100 | — | 2 |
| mail-merge | 20.0/100 | — | 1 |
| template-apply | 80.0/100 | — | 2 |
| document-generate | 90.0/100 | — | 1 |
| theme-apply | 16.7/100 | — | 3 |
| citation-bibliography | 20.0/100 | — | 1 |
| equation | 83.0/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 32.3/100 | — | 3 |
| margins-orientation | 82.5/100 | — | 2 |
| web-search-needed | 63.5/100 | — | 2 |
| refusal | 79.3/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 72.0/100 | — | 1 |
| edge-cases | 15.0/100 | — | 1 |
| form-field | 72.8/100 | — | 5 |
| text-insert | 80.0/100 | — | 1 |
| text-edit | 25.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 85 | ✓ | The code uses a non-standard method `addHeading` which is not a real Word Office.js API, but it correctly addresses the request and would likely work with minimal modifications to use the actual Office.js API. |
| L1-heading-insert-002 | 0 ⚠️ | ✓ | The code uses hallucinated methods "addHeading" and "addParagraph" which are not real Word Office.js APIs, resulting in a complete loss of points across all dimensions. |
| L1-paragraph-format-001 | 0 ⚠️ | ✗ | The code uses a hallucinated method `addParagraph` which does not exist in the Word Office.js API, and also attempts to modify a paragraph's text and style in a way that is not supported by the API. |
| L1-paragraph-format-002 | 92 | ✓ | The code accurately addresses the user's request, uses correct APIs, and is likely to execute without errors, but loses points for not using the available helpers, such as `paragraph.styleBuiltIn` or `applyStyle`, to set the font size and alignment. |
| L1-list-bullet-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions. |
| L1-list-multilevel-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-style-apply-001 | 83 | ✓ | The code is mostly correct but loses points for using a generic applyStyle method which may not be a built-in Word Office.js API, and also for not fully utilizing the best practices and available helpers, such as using paragraph.styleBuiltIn instead of a generic applyStyle method. |
| L1-table-create-001 | 90 | ✓ | The code accurately uses the Word Office.js API to insert a table, but loses points for not using the most efficient approach, such as utilizing the `body.insertTable` method directly with the provided parameters, and for not handling potential errors or edge cases. |
| L1-table-format-001 | 67 | ✓ | The code uses a non-existent API method "styleTable" which is not a real Word Office.js API, resulting in significant penalties across multiple dimensions. |
| L1-find-replace-001 | 80 | ✓ | The code uses the correct replaceText API but loses points for not using Word.run to initiate the operation and for not handling potential errors, and also for not fully addressing the request in terms of providing a complete solution with error handling and context. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach and for potential issues with the load and sync order, which could lead to runtime errors. |
| L1-find-replace-regex-001 | 83 | ✓ | The code uses the correct Word Office.js API replaceText with matchWildcards, but loses points for not fully addressing potential edge cases and not using the most optimal approach by directly utilizing the provided API without considering additional error handling or context checks. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a correct API but lacks consideration for the current selection and paragraph context, and does not utilize the best practice helpers provided by the Office.js API. |
| L1-comment-insert-001 | 72 | ✓ | The code uses a correct API but lacks specificity in searching for the first occurrence of "$4.2M" and does not fully address the request by not utilizing the `body.search` method with a wildcard pattern to find the exact text, resulting in potential inaccuracies and a less optimal approach. |
| L1-track-changes-toggle-001 | 100 ✅ | ✓ | The code perfectly addresses the request to turn on track changes using the correct and efficient `toggleTrackChanges` API, ensuring correctness, completeness, and adherence to best practices. |
| L1-track-changes-toggle-002 | 83 | ✓ | The code is mostly correct, but it lacks the surrounding Word.run context and does not handle potential errors, also it does not use the best practice of using Word.run to execute the acceptAllRevisions method. |
| L1-toc-generate-001 | 72 | ✓ | The code uses a correct API method to insert a Table of Contents, but lacks best practices and available helpers, and does not fully address the request as it does not specify the exact location of the Table of Contents, which is supposed to be at the top. |
| L1-section-break-001 | 72 | ✓ | The code uses correct APIs and is likely to work, but it lacks completeness in addressing the request as it doesn't specify the section to apply the landscape orientation to, and the approach could be improved by utilizing more specific and targeted methods for section manipulation. |
| L1-header-footer-001 | 83 | ✓ | The code is mostly correct but loses points for not using the Word.run method to encapsulate the API calls and for not checking if the primary header already exists before setting it, which could lead to unexpected behavior. |
| L1-header-footer-002 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-columns-001 | 72 | ✓ | The code uses correct APIs and would likely work, but it doesn't fully address the request as it only inserts a section break and columns without ensuring the rest of the document is actually in the new layout, and it lacks best practices in its approach. |
| L1-image-insert-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-image-insert-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice. |
| L1-watermark-insert-001 | 40 ⚠️ | ✓ | The code is incomplete and does not use the correct Word Office.js APIs, as it hallucinates a non-existent `insertWatermark` method instead of using the available APIs to insert a watermark, such as `range.insertOoxml` or `body.insertContentControl`. |
| L1-content-control-001 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the user's request, but loses points for not using a more specific best practice approach, such as utilizing a built-in helper for content control insertion if available, and for potential issues with context synchronization. |
| L1-content-control-002 | 80 | ✓ | The code correctly uses the insertContentControl API and addresses the user's request, but loses points for not considering the cursor position at the end of the document and not using a more robust method to ensure the content control is inserted at the correct location. |
| L1-mail-merge-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "mailMergeReplace" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-template-apply-001 | 80 | ✓ | The code correctly applies the resume template with the provided details, but lacks error handling and does not utilize the full range of available Word Office.js APIs and best practices, such as using Word.run for context. |
| L1-template-apply-002 | 80 | ✓ | The code correctly applies the memo template with the specified fields, but loses points for not using the Word.run method to ensure the context is loaded before applying the template, and for not fully following best practices by directly manipulating the template properties instead of using available helpers. |
| L1-document-generate-001 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach, such as potentially reusing the `addHeading` and `addParagraph` methods in a loop to generate the sections, instead of repeating the code for each section. |
| L1-theme-apply-001 | 10 ⚠️ | ✓ | The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also lacks any actual implementation to address the request of designing a calm trustworthy theme for a healthcare startup and applying it to the document. |
| L1-theme-apply-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not used correctly, resulting in a complete loss of points for API correctness and would it work dimensions. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-citation-bibliography-001 | 20 ⚠️ | ✓ | The code heavily hallucinates methods, specifically "insertCitation" and "insertBibliography", which are not real Word Office.js APIs, resulting in a significant penalty. |
| L1-equation-001 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the request, but loses points for not using a more robust approach, such as checking the current selection or range before inserting the equation, and not handling potential errors that may occur during the context.sync() call. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the user's request, but loses points for not using a more robust method to handle equation insertion, such as checking the current selection or range, and for not following best practices in terms of error handling and code organization. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selected text for clarity, but loses points for not considering alternative approaches, such as recommending the coach panel, and for using a generic phrase "Rewriting the selected text for clarity" that doesn't add much value. |
| L1-read-query-001 | 40 ⚠️ | ✓ | The code fails to utilize real Word Office.js APIs, instead relying on a hallucinated `countWords()` method, which is not a valid API call. |
| L1-read-query-002 | 57 ⚠️ | ✓ | The code uses a hallucinated method `getReadability()` which is not a real Word Office.js API, heavily penalizing API correctness, and also lacks best practices, such as using `Word.run` and `body` properties to access document content. |
| L1-read-query-003 | 0 ⚠️ | ✓ | The code uses a hallucinated method "listHeadings()" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-margins-orientation-001 | 85 | ✓ | The code correctly sets page margins using the Word Office.js API, but lacks best practices and does not utilize available helpers, and also assumes the context is already loaded without explicitly showing the Word.run() method to ensure the context is synchronized. |
| L1-margins-orientation-002 | 80 | ✓ | The code is mostly correct but lacks the surrounding Word.run context and does not explicitly specify the section or document scope for the page orientation change, which might lead to issues with the code's completeness and robustness. |
| L1-web-search-needed-001 | 67 | ✓ | The response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks clarity on how this information will be used to add a section to the document, and does not provide any workarounds or Word-specific caveats. |
| L1-web-search-needed-002 | 60 | ✓ | The response accurately explains the Pythagorean theorem and provides a relevant equation, but generates unnecessary code and lacks clarity in separating the static knowledge explanation from the code insertion. |
| L1-refusal-001 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not offering more detailed guidance or alternatives, and for a slightly abrupt tone in the refusal to generate code. |
| L1-refusal-002 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not offering more detailed alternatives or caveats, and for a slightly abrupt tone in the refusal to provide code. |
| L1-refusal-003 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js in creating Excel-style charts, but the workaround provided is incomplete and doesn't fully address the user's request for a chart based on specific data. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-paragraph-spacing-001 | 72 | ✓ | The code uses correct APIs and addresses the request, but has issues with approach, using a manual loop instead of available helpers like a range or body iteration, and also applies a style unnecessarily before setting line spacing. |
| L1-edge-case-001 | 15 ⚠️ | ✓ | The code fails to use the Word Office.js API correctly, instead throwing a generic error and not attempting to delete the paragraph or handle the out-of-range scenario using available API methods. |
| L1-form-field-001 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the user's request, but loses points for not using a more specific title or tag that directly relates to the "Yes / No / Maybe" options, and for not demonstrating best practices in terms of error handling or content control configuration. |
| L1-text-insert-001 | 80 | ✓ | The code uses a real Word Office.js API and is likely to work, but it doesn't fully address the request as it inserts a new paragraph instead of appending text to the end of the document, and it doesn't use the most appropriate helper method. |
| L1-text-edit-001 | 25 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice. |
| gen-L2-form-field-001 | 72 | ✓ | The code partially addresses the request but loses points for using a non-matching title for the date picker content control and not fully following best practices for inserting content controls. |
| gen-L2-form-field-002 | 65 | ✗ | The code uses a non-existent `insertContentControl` method, which is not a real Word Office.js API, and also fails to properly replace the existing plain text "Name: _______" with the new rich text content control. |
| gen-L2-form-field-003 | 72 | ✗ | The code fails to insert a group content control containing the two dropdown content controls as requested, instead inserting separate content controls, and does not utilize the best practices and available helpers provided by the Word Office.js API. |
| gen-L2-form-field-004 | 72 | ✗ | The code uses some correct APIs but lacks proper integration of the repeating section content control with the table, and does not follow best practices for inserting content controls and tables in Word. |

## ⚠️ Needs attention

**[L1-heading-insert-002]** score=0 — The code uses hallucinated methods "addHeading" and "addParagraph" which are not real Word Office.js APIs, resulting in a complete loss of points across all dimensions.

**[L1-paragraph-format-001]** score=0 — The code uses a hallucinated method `addParagraph` which does not exist in the Word Office.js API, and also attempts to modify a paragraph's text and style in a way that is not supported by the API.
- missing patterns: `font.bold`, `font.size`

**[L1-list-bullet-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions.

**[L1-list-multilevel-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-header-footer-002]** score=0 — The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-image-insert-001]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-image-insert-002]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice.

**[L1-watermark-insert-001]** score=40 — The code is incomplete and does not use the correct Word Office.js APIs, as it hallucinates a non-existent `insertWatermark` method instead of using the available APIs to insert a watermark, such as `range.insertOoxml` or `body.insertContentControl`.

**[L1-mail-merge-001]** score=20 — The code uses a hallucinated method "mailMergeReplace" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-theme-apply-001]** score=10 — The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also lacks any actual implementation to address the request of designing a calm trustworthy theme for a healthcare startup and applying it to the document.

**[L1-theme-apply-002]** score=20 — The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not used correctly, resulting in a complete loss of points for API correctness and would it work dimensions.

**[L1-theme-apply-003]** score=20 — The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-citation-bibliography-001]** score=20 — The code heavily hallucinates methods, specifically "insertCitation" and "insertBibliography", which are not real Word Office.js APIs, resulting in a significant penalty.

**[L1-read-query-001]** score=40 — The code fails to utilize real Word Office.js APIs, instead relying on a hallucinated `countWords()` method, which is not a valid API call.

**[L1-read-query-002]** score=57 — The code uses a hallucinated method `getReadability()` which is not a real Word Office.js API, heavily penalizing API correctness, and also lacks best practices, such as using `Word.run` and `body` properties to access document content.

**[L1-read-query-003]** score=0 — The code uses a hallucinated method "listHeadings()" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-edge-case-001]** score=15 — The code fails to use the Word Office.js API correctly, instead throwing a generic error and not attempting to delete the paragraph or handle the out-of-range scenario using available API methods.

**[L1-text-edit-001]** score=25 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice.
