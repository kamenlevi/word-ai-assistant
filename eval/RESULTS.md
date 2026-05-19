# Word AI — Eval Results
**Last run:** 2026-05-19 11:24:16  
**Overall: 52.4/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 33.5/100 | — | 2 |
| paragraph-format | 44.0/100 | — | 2 |
| list-bullet | 20.0/100 | — | 1 |
| list-multilevel | 0.0/100 | — | 1 |
| style-apply | 80.0/100 | — | 1 |
| table-create | 92.0/100 | — | 1 |
| table-format | 72.0/100 | — | 1 |
| find-replace | 81.0/100 | — | 2 |
| find-replace-regex | 83.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 90.0/100 | — | 2 |
| toc-generate | 80.0/100 | — | 1 |
| section-break | 67.0/100 | — | 1 |
| header-footer | 41.5/100 | — | 2 |
| columns | 85.0/100 | — | 1 |
| image-insert | 46.0/100 | — | 2 |
| watermark-insert | 0.0/100 | — | 1 |
| content-control | 81.5/100 | — | 2 |
| mail-merge | 55.0/100 | — | 1 |
| template-apply | 41.5/100 | — | 2 |
| document-generate | 90.0/100 | — | 1 |
| theme-apply | 15.0/100 | — | 3 |
| citation-bibliography | 20.0/100 | — | 1 |
| equation | 77.5/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 26.7/100 | — | 3 |
| margins-orientation | 42.5/100 | — | 2 |
| web-search-needed | 70.0/100 | — | 2 |
| refusal | 75.7/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 5.0/100 | — | 1 |
| edge-cases | 25.0/100 | — | 1 |
| form-field | 51.6/100 | — | 5 |
| text-insert | 80.0/100 | — | 1 |
| text-edit | 0.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 67 | ✓ | The code uses a non-standard method "addHeading" which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-heading-insert-002 | 0 ⚠️ | ✓ | The code uses hallucinated methods "addHeading" and "addParagraph" which are not real Word Office.js APIs, resulting in a complete failure across all dimensions. |
| L1-paragraph-format-001 | 0 ⚠️ | ✗ | The code attempts to add a new paragraph with the specified style instead of modifying the existing first paragraph, and it uses the `addParagraph` method which is not a real Word Office.js API, resulting in a complete failure to address the user's request. |
| L1-paragraph-format-002 | 88 | ✓ | The code is mostly correct but loses points for not using the available helpers, such as applyStyle or replaceText, to center-align and resize the title, instead opting for manual property assignment. |
| L1-list-bullet-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a severe penalty for API correctness and would not work in Word without runtime errors. |
| L1-list-multilevel-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-style-apply-001 | 80 | ✓ | The code correctly applies the Quote style to the third paragraph, but uses a generic applyStyle method which may not be the most efficient or Word-specific approach, and assumes the paragraph index is 2, which may not always be the case. |
| L1-table-create-001 | 92 | ✓ | The code accurately uses the Office.js API to insert a table, but loses points for completeness as it doesn't fully address the user's request by not handling potential edge cases or providing any error handling, and also for approach as it directly uses the insertTable method without considering any potential customizations or helper functions that could be used. |
| L1-table-format-001 | 72 | ✓ | The code uses a non-existent `styleTable` method, which is not a real Word Office.js API, resulting in significant penalties in API correctness and approach dimensions. |
| L1-find-replace-001 | 72 | ✓ | The code uses a correct API method replaceText, but lacks the context of a Word.run call and does not handle potential errors, also it does not follow best practices as it does not utilize the Word.run method which is the entry point for most Word JavaScript API operations. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient or best-practice approach, such as utilizing the `range.font` properties directly on the search results without explicitly loading the font property. |
| L1-find-replace-regex-001 | 83 | ✓ | The code uses the correct Word Office.js API replaceText with matchWildcards, but loses points for not fully addressing potential edge cases and not using the most optimal approach by directly utilizing the provided API without unnecessary complexity. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a correct API to insert a footnote but lacks specificity in targeting the current selection and does not fully follow best practices for handling the insertion within the context of the user's request. |
| L1-comment-insert-001 | 72 | ✓ | The code uses a correct API method `insertComment` but lacks proper context and range handling, and does not fully address the request as it does not specify the exact location of the comment, which may lead to incorrect placement. |
| L1-track-changes-toggle-001 | 97 ✅ | ✓ | The code perfectly addresses the request using the correct API, but loses a few points for not explicitly handling potential errors or edge cases, although the provided code snippet is otherwise flawless. |
| L1-track-changes-toggle-002 | 83 | ✓ | The code is mostly correct but loses points for not using the Word.run method to encapsulate the API calls and for not checking if the document has tracked revisions before accepting them, which could lead to unnecessary sync calls. |
| L1-toc-generate-001 | 80 | ✓ | The code uses a correct API method `insertTableOfContents` but loses points for not fully addressing the request with explicit heading level handling and potential issues with load order and variable scope. |
| L1-section-break-001 | 67 | ✓ | The code uses correct APIs and is likely to work, but it lacks completeness in fully addressing the user's request to switch the next section to landscape, and it does not follow best practices by using lower-level methods instead of available helpers. |
| L1-header-footer-001 | 83 | ✓ | The code is mostly correct but loses points for not using the best practice approach of utilizing available helpers and for potential issues with variable scope and method signatures. |
| L1-header-footer-002 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions. |
| L1-columns-001 | 85 | ✓ | The code is mostly correct and would work, but it lacks consideration for the current selection or range, and does not use the most idiomatic or efficient API methods, such as using `Word.run` to batch the operations. |
| L1-image-insert-001 | 72 | ✓ | The code uses a correct API method `insertImage` but lacks consideration for the cursor position at paragraph 3 and does not utilize best practices for image insertion, such as using the `paragraph` object to specify the insertion location. |
| L1-image-insert-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would_it_work dimensions. |
| L1-watermark-insert-001 | 0 ⚠️ | ✓ | The code hallucinates a non-existent method "insertWatermark" which is not a part of the Word Office.js APIs, resulting in a complete failure to address the request correctly. |
| L1-content-control-001 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the request, but loses points for not using the Word.run method to ensure the code is executed in the correct context and for not handling potential errors, and also for not using the most idiomatic approach to inserting a content control. |
| L1-content-control-002 | 80 | ✓ | The code is mostly correct but loses points for not specifying the exact location of the content control insertion, which could lead to unexpected behavior, and for not using the provided helpers to handle potential edge cases. |
| L1-mail-merge-001 | 55 ⚠️ | ✓ | The code fails to use real Word Office.js APIs, specifically hallucinating the "mailMergeReplace" method, which is not a valid API, resulting in a significant penalty for API correctness. |
| L1-template-apply-001 | 83 | ✓ | The code correctly applies the resume template with the provided details, but loses points for not using the Word.run method to execute the API calls and for not handling potential errors, and also for not fully utilizing the available helpers and best practices. |
| L1-template-apply-002 | 0 ⚠️ | ✓ | The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-document-generate-001 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach, such as potentially reusing the same paragraph or heading styles instead of re-creating them for each section. |
| L1-theme-apply-001 | 5 ⚠️ | ✓ | The code uses completely hallucinated and non-existent methods "designTheme" and "applyTheme", which are not part of the Word Office.js API, resulting in a severe penalty for API correctness and overall score. |
| L1-theme-apply-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not a valid method, resulting in a complete loss of points for API correctness and would it work dimensions. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, and also does not follow the correct approach for applying a theme, which should use the "designTheme" and "applyTheme" methods. |
| L1-citation-bibliography-001 | 20 ⚠️ | ✓ | The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-equation-001 | 72 | ✓ | The code uses a correct API method, is mostly complete, and would likely work, but it lacks best practices, such as using Word.run to ensure context and load the document, and does not handle potential errors or loading issues. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the user's request, but loses points for not using a more robust method to handle equation insertion, such as checking the cursor position and paragraph context before inserting the equation. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selection for clarity, is easy to understand, and directly addresses the question, but loses points for not considering the coach panel as an alternative and not providing additional Word-specific caveats or workarounds. |
| L1-read-query-001 | 40 ⚠️ | ✓ | The code hallucinates a non-existent `countWords()` method, which is not a real Word Office.js API, and does not utilize the available APIs to count the words in the document. |
| L1-read-query-002 | 40 ⚠️ | ✓ | The code fails to utilize actual Word Office.js APIs, instead relying on a hallucinated `getReadability()` method, which is not a real API, resulting in a significant penalty for API correctness. |
| L1-read-query-003 | 0 ⚠️ | ✓ | The code hallucinates a non-existent `listHeadings()` method, which is not a real Word Office.js API, and does not provide any actual implementation to list the headings in the document. |
| L1-margins-orientation-001 | 85 | ✓ | The code correctly uses the Word Office.js API and fully addresses the request, but loses points for not using the most efficient approach and potentially having issues with the load/sync order. |
| L1-margins-orientation-002 | 0 ⚠️ | ✓ | The code uses a hallucinated method "setPageOrientation" which is not a real Word Office.js API, resulting in a complete failure across all dimensions. |
| L1-web-search-needed-001 | 67 | ✓ | The response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks clarity on how this information will be used to address the user's request to add a section about the stock price to the document. |
| L1-web-search-needed-002 | 73 | ✓ | The response accurately and clearly explains the Pythagorean theorem, directly addressing the question, but lacks any notable bonus features such as workarounds or Word-specific caveats, and unnecessarily includes code for inserting the information into the document. |
| L1-refusal-001 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not offering more detailed guidance or alternatives, and for a slightly vague explanation of the API's capabilities. |
| L1-refusal-002 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but lacks additional relevant information or alternatives, resulting in a moderate bonus score. |
| L1-refusal-003 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js in creating charts, but loses points for not providing a more detailed or relevant workaround, such as suggesting how to use Excel to create the chart and then insert it as an image. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-paragraph-spacing-001 | 5 ⚠️ | ✓ | The code fails to address the user's request to set line spacing to 1.5 on existing paragraphs, instead attempting to add new paragraphs with the specified line spacing, and also uses incorrect methods and logic. |
| L1-edge-case-001 | 25 ⚠️ | ✓ | The code does not use any real Word Office.js APIs and instead throws a generic error, failing to attempt the deletion of the paragraph or guard against out-of-range indices using the available APIs. |
| L1-form-field-001 | 0 ⚠️ | ✓ | A:  |
| L1-text-insert-001 | 80 | ✓ | The code uses a real Word Office.js API and is likely to work, but it doesn't fully address the request as it inserts a new paragraph instead of just inserting the text at the end of the document, and it doesn't use the most appropriate helper method. |
| L1-text-edit-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions. |
| gen-L2-form-field-001 | 72 | ✓ | The code uses correct Office.js APIs and is mostly complete, but loses points for not exactly matching the requested label for the date picker and not using best practices for content control insertion, such as specifying the exact location of the controls. |
| gen-L2-form-field-002 | 72 | ✗ | The code uses correct APIs and is mostly complete, but it lacks proper replacement of the existing plain text and does not fully follow best practices, such as using the `insertContentControl` method correctly and not addressing the removal of the original text. |
| gen-L2-form-field-003 | 57 ⚠️ | ✗ | The code fails to fully address the request by not inserting the dropdown content controls within a group content control, and also lacks best practices by not utilizing the available helpers for inserting content controls. |
| gen-L2-form-field-004 | 57 ⚠️ | ✗ | The code fails to fully address the request by inserting a plain text content control instead of a repeating section content control, and also lacks best practices in using available helpers. |

## ⚠️ Needs attention

**[L1-heading-insert-002]** score=0 — The code uses hallucinated methods "addHeading" and "addParagraph" which are not real Word Office.js APIs, resulting in a complete failure across all dimensions.

**[L1-paragraph-format-001]** score=0 — The code attempts to add a new paragraph with the specified style instead of modifying the existing first paragraph, and it uses the `addParagraph` method which is not a real Word Office.js API, resulting in a complete failure to address the user's request.
- missing patterns: `font.bold`, `font.size`

**[L1-list-bullet-001]** score=20 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a severe penalty for API correctness and would not work in Word without runtime errors.

**[L1-list-multilevel-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-header-footer-002]** score=0 — The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions.

**[L1-image-insert-002]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would_it_work dimensions.

**[L1-watermark-insert-001]** score=0 — The code hallucinates a non-existent method "insertWatermark" which is not a part of the Word Office.js APIs, resulting in a complete failure to address the request correctly.

**[L1-mail-merge-001]** score=55 — The code fails to use real Word Office.js APIs, specifically hallucinating the "mailMergeReplace" method, which is not a valid API, resulting in a significant penalty for API correctness.

**[L1-template-apply-002]** score=0 — The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-theme-apply-001]** score=5 — The code uses completely hallucinated and non-existent methods "designTheme" and "applyTheme", which are not part of the Word Office.js API, resulting in a severe penalty for API correctness and overall score.

**[L1-theme-apply-002]** score=20 — The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not a valid method, resulting in a complete loss of points for API correctness and would it work dimensions.

**[L1-theme-apply-003]** score=20 — The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, and also does not follow the correct approach for applying a theme, which should use the "designTheme" and "applyTheme" methods.

**[L1-citation-bibliography-001]** score=20 — The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-read-query-001]** score=40 — The code hallucinates a non-existent `countWords()` method, which is not a real Word Office.js API, and does not utilize the available APIs to count the words in the document.

**[L1-read-query-002]** score=40 — The code fails to utilize actual Word Office.js APIs, instead relying on a hallucinated `getReadability()` method, which is not a real API, resulting in a significant penalty for API correctness.

**[L1-read-query-003]** score=0 — The code hallucinates a non-existent `listHeadings()` method, which is not a real Word Office.js API, and does not provide any actual implementation to list the headings in the document.

**[L1-margins-orientation-002]** score=0 — The code uses a hallucinated method "setPageOrientation" which is not a real Word Office.js API, resulting in a complete failure across all dimensions.

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-paragraph-spacing-001]** score=5 — The code fails to address the user's request to set line spacing to 1.5 on existing paragraphs, instead attempting to add new paragraphs with the specified line spacing, and also uses incorrect methods and logic.

**[L1-edge-case-001]** score=25 — The code does not use any real Word Office.js APIs and instead throws a generic error, failing to attempt the deletion of the paragraph or guard against out-of-range indices using the available APIs.

**[L1-form-field-001]** score=0 — A: 

**[L1-text-edit-001]** score=0 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions.

**[gen-L2-form-field-003]** score=57 — The code fails to fully address the request by not inserting the dropdown content controls within a group content control, and also lacks best practices by not utilizing the available helpers for inserting content controls.
- missing patterns: `GroupContentControl`

**[gen-L2-form-field-004]** score=57 — The code fails to fully address the request by inserting a plain text content control instead of a repeating section content control, and also lacks best practices in using available helpers.
- missing patterns: `RepeatingSectionContentControl`
