# Word AI — Eval Results
**Last run:** 2026-05-22 10:58:13  
**Overall: 59.2/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 69.5/100 | — | 2 |
| paragraph-format | 93.5/100 | — | 2 |
| list-bullet | 20.0/100 | — | 1 |
| list-multilevel | 0.0/100 | — | 1 |
| style-apply | 83.0/100 | — | 1 |
| table-create | 89.0/100 | — | 1 |
| table-format | 67.0/100 | — | 1 |
| find-replace | 78.5/100 | — | 2 |
| find-replace-regex | 83.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 92.5/100 | — | 2 |
| toc-generate | 0.0/100 | — | 1 |
| section-break | 0.0/100 | — | 1 |
| header-footer | 50.0/100 | — | 2 |
| columns | 72.0/100 | — | 1 |
| image-insert | 50.0/100 | — | 2 |
| watermark-insert | 0.0/100 | — | 1 |
| content-control | 87.5/100 | — | 2 |
| mail-merge | 72.0/100 | — | 1 |
| template-apply | 81.5/100 | — | 2 |
| document-generate | 83.0/100 | — | 1 |
| theme-apply | 15.0/100 | — | 3 |
| citation-bibliography | 40.0/100 | — | 1 |
| equation | 75.0/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 32.3/100 | — | 3 |
| margins-orientation | 81.5/100 | — | 2 |
| web-search-needed | 70.0/100 | — | 2 |
| refusal | 77.7/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 72.0/100 | — | 1 |
| edge-cases | 25.0/100 | — | 1 |
| form-field | 65.8/100 | — | 5 |
| text-insert | 80.0/100 | — | 1 |
| text-edit | 0.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 67 | ✓ | The code uses a non-standard method "addHeading" which is not a real Word Office.js API, indicating a lack of best practices and reliance on potentially hallucinated helpers. |
| L1-heading-insert-002 | 72 | ✓ | The code uses a hallucinated method addHeading and addParagraph, which are not real Word Office.js APIs, resulting in significant deductions in API correctness and approach dimensions. |
| L1-paragraph-format-001 | 92 | ✓ | The code accurately uses Word Office.js APIs, fully addresses the request, and would likely execute without runtime errors, but could be improved by using more best practices and available helpers, such as applying a style instead of manually setting font properties. |
| L1-paragraph-format-002 | 95 ✅ | ✓ | The code is almost perfect, but loses points on approach for not using the available helper methods, such as paragraph.styleBuiltIn or applyStyle, to set the font size and alignment, instead opting for direct property manipulation. |
| L1-list-bullet-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice. |
| L1-list-multilevel-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure across all dimensions. |
| L1-style-apply-001 | 83 | ✓ | The code is mostly correct but loses points for not using the provided helpers like applyStyle correctly in the context of the Word object model, and the variable "paragraphs" should be accessed through the body object, such as "body.paragraphs". |
| L1-table-create-001 | 89 | ✓ | The code accurately uses the Word Office.js API to insert a table, but loses points for not fully addressing the request by not considering the document context and paragraph placement, and for not using the most optimal approach by directly specifying the table style instead of potentially using a more flexible method. |
| L1-table-format-001 | 67 | ✓ | The code uses a hallucinated method `styleTable` which is not a real Word Office.js API, resulting in significant penalties across multiple dimensions. |
| L1-find-replace-001 | 67 | ✓ | The code uses a non-existent `replaceText` method, which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach and potentially having issues with the load and sync order, which could lead to runtime errors. |
| L1-find-replace-regex-001 | 83 | ✓ | The code uses the correct Word Office.js API replaceText with matchWildcards, but loses points for not fully addressing potential edge cases and not using the most optimal approach by directly utilizing the provided API without considering additional error handling or context checks. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a correct API but lacks completeness in addressing the request, has some potential issues with variable scope and load order, and does not use the best practices or available helpers, such as specifying the range for the footnote insertion. |
| L1-comment-insert-001 | 72 | ✓ | The code uses a non-existent `insertComment` method with two parameters, which is not a real Word Office.js API, and also lacks proper range or context for the comment insertion. |
| L1-track-changes-toggle-001 | 100 ✅ | ✓ | The code perfectly addresses the request to turn on track changes using the correct and available Word Office.js API, specifically the toggleTrackChanges helper method. |
| L1-track-changes-toggle-002 | 85 | ✓ | The code is mostly correct and complete, but loses points for not using the Word.run method to wrap the API calls, which is a best practice, and also for not checking if there are actually tracked revisions before accepting them. |
| L1-toc-generate-001 | 0 ⚠️ | ✓ | A:  |
| L1-section-break-001 | 0 ⚠️ | ✓ | The code uses hallucinated methods "insertSectionBreak" and "setPageOrientation" which are not real Word Office.js APIs, resulting in a complete failure across all dimensions. |
| L1-header-footer-001 | 80 | ✓ | The code is mostly correct but lacks consideration for the specific section or page range where the header should be applied, and uses a generic `setHeader` method which might not be a real Word Office.js API, deducting points for API correctness and completeness. |
| L1-header-footer-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-columns-001 | 72 | ✓ | The code uses correct APIs and would likely work, but it only partially addresses the request by not specifying the range for the two-column layout, and it does not use best practices such as selecting the rest of the document before applying the layout change. |
| L1-image-insert-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-image-insert-002 | 80 | ✓ | The code uses the correct insertImage API and addresses the request, but loses points for not using a more robust approach to handling image insertion and potential errors, and for not fully considering the context of the single-page document. |
| L1-watermark-insert-001 | 0 ⚠️ | ✓ | The code does not use real Word Office.js APIs, as there is no built-in method called "insertWatermark" in the Office.js API, and the code does not address the request to add a watermark to every page of the document. |
| L1-content-control-001 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the user's request, but loses points for not using a more specific best practice approach and potentially lacking error handling or loading checks. |
| L1-content-control-002 | 92 | ✓ | The code is mostly correct, but loses points for not specifying the location of the content control insertion, which might lead to unexpected behavior, and for not using the provided helpers to handle potential edge cases. |
| L1-mail-merge-001 | 72 | ✓ | The code uses the correct mailMergeReplace API but lacks proper error handling, does not fully address potential edge cases, and does not follow best practices for handling data and template integration. |
| L1-template-apply-001 | 80 | ✓ | The code correctly applies the resume template with the provided information, but loses points for not using the Word.run method to ensure the context is properly loaded and for not handling potential errors, and also for not fully utilizing the available helpers and best practices. |
| L1-template-apply-002 | 83 | ✓ | The code correctly applies the memo template with the specified fields, but loses points for not using the Word.run method to ensure the context is properly loaded before applying the template, and for not fully following best practices in using the available helpers. |
| L1-document-generate-001 | 83 | ✓ | The code correctly uses real Word Office.js APIs and fully addresses the request, but loses points for not using the most efficient approach, such as utilizing the `insertTableOfContents` method more effectively and not demonstrating best practices in code organization and variable handling. |
| L1-theme-apply-001 | 5 ⚠️ | ✓ | The code heavily hallucinates non-existent methods "designTheme" and "applyTheme", which are not part of the real Word Office.js APIs, resulting in a significant penalty. |
| L1-theme-apply-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not used correctly, resulting in a complete loss of points for API correctness and would it work dimensions. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, and also does not use the correct API "designTheme" and "applyTheme" to modify the theme. |
| L1-citation-bibliography-001 | 40 ⚠️ | ✓ | The code uses hallucinated methods `insertCitation` and `insertBibliography` which are not real Word Office.js APIs, resulting in a significant penalty for API correctness. |
| L1-equation-001 | 67 | ✓ | The code uses a non-existent insertEquation method, which is not a real Word Office.js API, resulting in significant penalties for API correctness and approach. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the user's request, but loses points for not using a more robust method to handle equation insertion and not considering potential errors or edge cases, such as invalid input or equation formatting issues. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selection for clarity and directly addresses the question, but loses points for not considering alternative approaches, such as recommending the coach panel, and for minor clarity issues in the surrounding text. |
| L1-read-query-001 | 40 ⚠️ | ✓ | The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `countWords()` method, which heavily penalizes the API correctness score. |
| L1-read-query-002 | 57 ⚠️ | ✓ | The code uses a hallucinated method "getReadability()" which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-read-query-003 | 0 ⚠️ | ✓ | The code uses a hallucinated method "listHeadings()" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-margins-orientation-001 | 83 | ✓ | The code correctly sets page margins using the Office.js API, but loses points for not using a more explicit best practice approach, such as using a predefined constant for the margin value, and for not handling potential errors that may occur during the execution of the code. |
| L1-margins-orientation-002 | 80 | ✓ | The code is mostly correct but lacks the surrounding Word.run context and may not fully handle potential errors, also it does not utilize the best practices for setting page orientation which might involve using section properties. |
| L1-web-search-needed-001 | 67 | ✓ | The response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks clarity on how this information will be used to add a section to the document, and does not offer any workarounds or Word-specific guidance. |
| L1-web-search-needed-002 | 73 | ✓ | The response accurately and clearly explains the Pythagorean theorem, directly addressing the question, but lacks bonus points for not providing additional workarounds or Word-specific caveats, and the code provided is straightforward but not particularly noteworthy. |
| L1-refusal-001 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not offering more detailed guidance on the "Save As" feature or alternative solutions. |
| L1-refusal-002 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not offering more detailed alternatives or settings options, and for the code sections being unnecessary and empty. |
| L1-refusal-003 | 67 | ✓ | The response accurately identifies the limitation of Word Office.js in creating charts, but lacks a clear and relevant workaround for generating a chart from the provided data, resulting in low bonus points. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-paragraph-spacing-001 | 72 | ✓ | The code uses a correct API but lacks best practices, such as using a loop based on the actual number of paragraphs in the document instead of hardcoding the number, and not utilizing the provided helpers like applyStyle correctly, as it's applied before setting the line spacing. |
| L1-edge-case-001 | 25 ⚠️ | ✓ | The code does not use any Word Office.js APIs and does not attempt to delete the paragraph, instead simply stating that no action is needed, which is a significant flaw in addressing the user's request. |
| L1-form-field-001 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the request, but loses points for not using a more descriptive title and tag, and for not handling potential errors or edge cases, which affects its completeness and adherence to best practices. |
| L1-text-insert-001 | 80 | ✓ | The code uses a real Word Office.js API (addParagraph) and is likely to execute without runtime errors, but loses points for not fully addressing the request (inserting at the end of the document, not necessarily after the last paragraph) and not using the most optimal approach (e.g., using range.insertText to insert text at a specific location). |
| L1-text-edit-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure across all dimensions. |
| gen-L2-form-field-001 | 72 | ✓ | The code uses correct Word Office.js APIs and is likely to execute without runtime errors, but it does not fully address the request as it inserts a date picker with the title "Sign Date" instead of placing it below the checkbox content control with the correct label, and it lacks best practices in content control placement and styling. |
| gen-L2-form-field-002 | 72 | ✗ | The code uses correct Office.js APIs and is mostly complete, but it lacks proper replacement of the existing plain text and does not fully follow best practices, such as using the `replaceText` helper to replace the existing text with the content control. |
| gen-L2-form-field-003 | 30 ⚠️ | ✗ | The code fails to use the real Word Office.js API to insert a group content control, instead using a non-existent `insertContentControl` method with incorrect parameters, and also does not address the request to insert a group content control containing two dropdown content controls. |
| gen-L2-form-field-004 | 72 | ✗ | The code uses correct APIs and is mostly complete, but it lacks the specific implementation of a repeating section content control and instead inserts a regular content control, and also does not fully follow best practices in terms of using available helpers for table styling and content control configuration. |

## ⚠️ Needs attention

**[L1-list-bullet-001]** score=20 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice.

**[L1-list-multilevel-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure across all dimensions.

**[L1-toc-generate-001]** score=0 — A: 

**[L1-section-break-001]** score=0 — The code uses hallucinated methods "insertSectionBreak" and "setPageOrientation" which are not real Word Office.js APIs, resulting in a complete failure across all dimensions.

**[L1-header-footer-002]** score=20 — The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-image-insert-001]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-watermark-insert-001]** score=0 — The code does not use real Word Office.js APIs, as there is no built-in method called "insertWatermark" in the Office.js API, and the code does not address the request to add a watermark to every page of the document.

**[L1-theme-apply-001]** score=5 — The code heavily hallucinates non-existent methods "designTheme" and "applyTheme", which are not part of the real Word Office.js APIs, resulting in a significant penalty.

**[L1-theme-apply-002]** score=20 — The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not used correctly, resulting in a complete loss of points for API correctness and would it work dimensions.

**[L1-theme-apply-003]** score=20 — The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, and also does not use the correct API "designTheme" and "applyTheme" to modify the theme.

**[L1-citation-bibliography-001]** score=40 — The code uses hallucinated methods `insertCitation` and `insertBibliography` which are not real Word Office.js APIs, resulting in a significant penalty for API correctness.

**[L1-read-query-001]** score=40 — The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `countWords()` method, which heavily penalizes the API correctness score.

**[L1-read-query-002]** score=57 — The code uses a hallucinated method "getReadability()" which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions.

**[L1-read-query-003]** score=0 — The code uses a hallucinated method "listHeadings()" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-edge-case-001]** score=25 — The code does not use any Word Office.js APIs and does not attempt to delete the paragraph, instead simply stating that no action is needed, which is a significant flaw in addressing the user's request.

**[L1-text-edit-001]** score=0 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure across all dimensions.

**[gen-L2-form-field-003]** score=30 — The code fails to use the real Word Office.js API to insert a group content control, instead using a non-existent `insertContentControl` method with incorrect parameters, and also does not address the request to insert a group content control containing two dropdown content controls.
- missing patterns: `GroupContentControl`
