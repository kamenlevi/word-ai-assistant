# Word AI — Eval Results
**Last run:** 2026-05-17 17:03:56  
**Overall: 42.0/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 42.5/100 | — | 2 |
| paragraph-format | 88.0/100 | — | 2 |
| list-bullet | 0.0/100 | — | 1 |
| list-multilevel | 0.0/100 | — | 1 |
| style-apply | 80.0/100 | — | 1 |
| table-create | 89.0/100 | — | 1 |
| table-format | 0.0/100 | — | 1 |
| find-replace | 20.5/100 | — | 2 |
| find-replace-regex | 83.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 18.0/100 | — | 1 |
| track-changes-toggle | 90.0/100 | — | 2 |
| toc-generate | 83.0/100 | — | 1 |
| section-break | 72.0/100 | — | 1 |
| header-footer | 40.0/100 | — | 2 |
| columns | 20.0/100 | — | 1 |
| image-insert | 37.0/100 | — | 2 |
| watermark-insert | 0.0/100 | — | 1 |
| content-control | 58.5/100 | — | 2 |
| mail-merge | 18.0/100 | — | 1 |
| template-apply | 51.5/100 | — | 2 |
| document-generate ✓ | 95.0/100 | — | 1 |
| theme-apply | 14.0/100 | — | 3 |
| citation-bibliography | 20.0/100 | — | 1 |
| equation | 46.0/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 32.3/100 | — | 3 |
| margins-orientation | 40.0/100 | — | 2 |
| web-search-needed | 44.5/100 | — | 2 |
| refusal | 52.3/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 0.0/100 | — | 1 |
| edge-cases | 5.0/100 | — | 1 |
| form-field | 49.0/100 | — | 5 |
| text-insert | 0.0/100 | — | 1 |
| text-edit | 0.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 85 | ✓ | The code is mostly correct but lacks the Word.run() context and may not handle potential errors, also it doesn't use the best practice of using Word.run() to wrap the API calls. |
| L1-heading-insert-002 | 0 ⚠️ | ✓ | The code uses hallucinated methods "addHeading" and "addParagraph" which are not real Word Office.js APIs, resulting in a complete failure to meet the requirements. |
| L1-paragraph-format-001 | 93 | ✓ | The code accurately uses real Word Office.js APIs and fully addresses the request, but loses points for approach due to manually setting font properties instead of using available helpers like applyStyle. |
| L1-paragraph-format-002 | 83 | ✓ | The code correctly uses Office.js APIs and addresses the request, but lacks best practices, such as using paragraph.styleBuiltIn or applyStyle, and has potential issues with load and sync order, which may cause runtime errors. |
| L1-list-bullet-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure across all dimensions. |
| L1-list-multilevel-001 | 0 ⚠️ | ✓ | The code usesadd real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions. |
| L1-style-apply-001 | 80 | ✓ | The code correctly applies the Quote style to the third paragraph using the applyStyle method, but loses points for not using the Word.run method to ensure context and for not checking if the paragraph exists before applying the style. |
| L1-table-create-001 | 89 | ✓ | The code accurately uses the Word Office.js API to insert a table, but loses points for not fully addressing the request by not considering the document context and paragraph placement, and for not using the most optimal approach by directly specifying the table style instead of potentially using a more flexible method. |
| L1-table-format-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "styleTable" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-find-replace-001 | 18 ⚠️ | ✓ | A: 18
B: |
| L1-find-replace-002 | 23 ⚠️ | ✓ | A: 23
B: |
| L1-find-replace-regex-001 | 83 | ✓ | The code uses the correct Word Office.js API replaceText with matchWildcards, but loses points for not explicitly using Word.run to execute the replacement, and for not fully addressing potential edge cases in email address formatting. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a correct API method to insert a footnote, but lacks consideration for the current selection and does not fully address the request, also it does not use the best practice approach as it manually implements the insertion of a footnote without utilizing the available helpers. |
| L1-comment-insert-001 | 18 ⚠️ | ✓ | A: 18
B: |
| L1-track-changes-toggle-001 | 97 ✅ | ✓ | The code perfectly addresses the request using the correct API, but loses a few points for not explicitly handling potential errors that might occur during the execution of the toggleTrackChanges method. |
| L1-track-changes-toggle-002 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the Word.run method to execute the acceptAllRevisions action, and for not fully following best practices by directly calling acceptAllRevisions without wrapping it in a Word.run context. |
| L1-toc-generate-001 | 83 | ✓ | The code correctly uses the insertTableOfContents API and addresses the request, but loses points for not fully utilizing best practices and available helpers, and for not perfectly handling potential edge cases in the generated code. |
| L1-section-break-001 | 72 | ✓ | The code uses correct APIs and is likely to work, but it does not fully address the request as it switches the orientation of the entire document instead of just the next section, and lacks best practices in its approach. |
| L1-header-footer-001 | 80 | ✓ | The code is mostly correct but lacks consideration for the specific section or page where the header should be applied, and it does not follow best practices for setting headers in Word using the Office.js API. |
| L1-header-footer-002 | 0 ⚠️ | ✗ | The AI response failed to generate any code, making it impossible to evaluate its correctness, completeness, executability, or approach to adding page numbers to the footer. |
| L1-columns-001 | 20 ⚠️ | ✓ | A: 20
B: |
| L1-image-insert-001 | 54 ⚠️ | ✓ | The code usesImage API lacks consideration for the cursor but position at does not fully adhere to best paragraph 3 and practices, the provided helpers such as utilizing for more |
| L1-image-insert-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method `insertImage` which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-watermark-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "insertWatermark" which is not a real Word Office.js API, and also does not address the request to add the watermark to every page. |
| L1-content-control-001 | 97 ✅ | ✓ | The code is nearly flawless, using the correct Office.js API to insert a content control, but loses a few points for not explicitly handling potential errors or edge cases, such as the cursor not being at the end of the document. |
| L1-content-control-002 | 20 ⚠️ | ✓ | A: 20
B: |
| L1-mail-merge-001 | 18 ⚠️ | ✓ | A: 18
B: |
| L1-template-apply-001 | 20 ⚠️ | ✓ | A: 20
B: |
| L1-template-apply-002 | 83 | ✓ | The code correctly applies the memo template with the provided details, but loses points for API correctness and approach due to the use of a potentially non-existent `applyTemplate` method, which is not a standard Word Office.js API. |
| L1-document-generate-001 | 95 ✅ | ✓ | The code is mostly correct and complete, but loses points for not handling potential errors and not using the most efficient approach for inserting multiple headings and paragraphs. |
| L1-theme-apply-001 | 2 ⚠️ | ✓ | A: 2
B: |
| L1-theme-apply-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also does not correctly apply the theme to the document content. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, and also does not use the correct API "designTheme" and "applyTheme" to modify the theme. |
| L1-citation-bibliography-001 | 20 ⚠️ | ✓ | The code usesinsertBibliography" which are not real Word Office.js APIs, resultingCitation" and "insert in a complete of points for API loss correctness and would-work dimensions. |
| L1-equation-001 | 72 | ✓ | The code uses a non-existent insertEquation method, which is not a real Word Office.js API, resulting in significant penalties for API correctness and approach. |
| L1-equation-002 | 20 ⚠️ | ✓ | A: 20
B: |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selected text for clarity, but loses points for not considering alternative approaches, such as recommending the coach panel, and for using a somewhat generic phrase "Rewriting the current selection for clarity" that doesn't add much value. |
| L1-read-query-001 | 40 ⚠️ | ✓ | The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `countWords()` method, which heavily penalizes the API correctness score. |
| L1-read-query-002 | 57 ⚠️ | ✓ | The code uses a hallucinated method "getReadability()" which is not a real Word Office.js API, resulting in significant penalties in API correctness and approach dimensions. |
| L1-read-query-003 | 0 ⚠️ | ✓ | The code hallucinates a non-existent `listHeadings()` method, which is not a real Word Office.js API, and does not provide any actual implementation to list the headings in the document. |
| L1-margins-orientation-001 | 0 ⚠️ | ✓ | The AI response failed to generate any code, leaving the request to set page margins to 1 inch on all sides unaddressed and without a viable solution. |
| L1-margins-orientation-002 | 80 | ✓ | The code correctly uses the.js API and addresses the user Office's request, but lacks context about the document and the `PageOrientation` method is notset a.js API, which standard Word Office reduces API correctness and approach the score for |
| L1-web-search-needed-001 | 67 | ✓ | The response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks clarity on how this information will be used to add a section to the Word document and does not offer any workarounds or Word-specific guidance. |
| L1-web-search-needed-002 | 22 ⚠️ | ✓ | A: 22
B: |
| L1-refusal-001 | 20 ⚠️ | ✓ | A: 20
B: |
| L1-refusal-002 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but lacks a detailed explanation or alternative solutions, resulting in moderate scores across the dimensions. |
| L1-refusal-003 | 65 | ✓ | The response accurately identifies the limitation of Word Office.js in creating charts, but loses points for not providing a more detailed workaround, such as generating the chart image using a separate library or service, and for inserting a generic image instead of a chart based on the provided data. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, and does not provide any actual implementation to insert a block quote. |
| L1-paragraph-spacing-001 | 0 ⚠️ | ✗ | The code is completely incorrect, with hallucinated methods and syntax errors, such as "add;" and "line.body.parSpacing", which are not real Word Office.js APIs, and the loop condition "i <" is incomplete and will cause a runtime error. |
| L1-edge-case-001 | 5 ⚠️ | ✓ | The response does not generate any code, failing to utilize Word Office.js APIs, address the request, or demonstrate best practices, resulting in a complete lack of functionality. |
| L1-form-field-001 | 83 | ✓ | The code correctly uses the insertContentControl API and fully addresses the request, but loses points for not using a more descriptive title and not handling potential errors, and also for not using the most robust and flexible approach by directly utilizing the available helpers. |
| L1-text-insert-001 | 0 ⚠️ | ✗ | The AI response failed to generate any code, completely neglecting the user's request to insert text at the end of the document, resulting in a total lack of correctness, completeness, functionality, and adherence to best practices. |
| L1-text-edit-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| gen-L2-form-field-001 | 72 | ✓ | The code uses correct Office.js APIs and is likely to work, but loses points for not fully addressing the request by inserting the date picker content control with the exact label specified, and for not using best practices in terms of content control placement and styling. |
| gen-L2-form-field-002 | 65 | ✗ | The code uses a non-existent `insertContentControl` method, which is not a real Word Office.js API, and also fails to properly replace the existing plain text "Name: _______" with the new rich text content control, resulting in incomplete and imperfect implementation. |
| gen-L2-form-field-003 | 20 ⚠️ | ✗ | A: 20
B: |
| gen-L2-form-field-004 | 5 ⚠️ | ✗ | The code does not use real Word Office.js APIs, instead hallucinating methods like "insertContent a repeating section content control" and "insertTable" with incorrect syntax and parameters. |

## ⚠️ Needs attention

**[L1-heading-insert-002]** score=0 — The code uses hallucinated methods "addHeading" and "addParagraph" which are not real Word Office.js APIs, resulting in a complete failure to meet the requirements.

**[L1-list-bullet-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure across all dimensions.

**[L1-list-multilevel-001]** score=0 — The code usesadd real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions.

**[L1-table-format-001]** score=0 — The code uses a hallucinated method "styleTable" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-find-replace-001]** score=18 — A: 18
B:

**[L1-find-replace-002]** score=23 — A: 23
B:

**[L1-comment-insert-001]** score=18 — A: 18
B:

**[L1-header-footer-002]** score=0 — The AI response failed to generate any code, making it impossible to evaluate its correctness, completeness, executability, or approach to adding page numbers to the footer.
- missing patterns: `addPageNumbers`

**[L1-columns-001]** score=20 — A: 20
B:

**[L1-image-insert-001]** score=54 — The code usesImage API lacks consideration for the cursor but position at does not fully adhere to best paragraph 3 and practices, the provided helpers such as utilizing for more

**[L1-image-insert-002]** score=20 — The code uses a hallucinated method `insertImage` which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-watermark-insert-001]** score=0 — The code uses a hallucinated method "insertWatermark" which is not a real Word Office.js API, and also does not address the request to add the watermark to every page.

**[L1-content-control-002]** score=20 — A: 20
B:

**[L1-mail-merge-001]** score=18 — A: 18
B:

**[L1-template-apply-001]** score=20 — A: 20
B:

**[L1-theme-apply-001]** score=2 — A: 2
B:

**[L1-theme-apply-002]** score=20 — The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also does not correctly apply the theme to the document content.

**[L1-theme-apply-003]** score=20 — The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, and also does not use the correct API "designTheme" and "applyTheme" to modify the theme.

**[L1-citation-bibliography-001]** score=20 — The code usesinsertBibliography" which are not real Word Office.js APIs, resultingCitation" and "insert in a complete of points for API loss correctness and would-work dimensions.

**[L1-equation-002]** score=20 — A: 20
B:

**[L1-read-query-001]** score=40 — The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `countWords()` method, which heavily penalizes the API correctness score.

**[L1-read-query-002]** score=57 — The code uses a hallucinated method "getReadability()" which is not a real Word Office.js API, resulting in significant penalties in API correctness and approach dimensions.

**[L1-read-query-003]** score=0 — The code hallucinates a non-existent `listHeadings()` method, which is not a real Word Office.js API, and does not provide any actual implementation to list the headings in the document.

**[L1-margins-orientation-001]** score=0 — The AI response failed to generate any code, leaving the request to set page margins to 1 inch on all sides unaddressed and without a viable solution.

**[L1-web-search-needed-002]** score=22 — A: 22
B:

**[L1-refusal-001]** score=20 — A: 20
B:

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, and does not provide any actual implementation to insert a block quote.

**[L1-paragraph-spacing-001]** score=0 — The code is completely incorrect, with hallucinated methods and syntax errors, such as "add;" and "line.body.parSpacing", which are not real Word Office.js APIs, and the loop condition "i <" is incomplete and will cause a runtime error.
- missing patterns: `lineSpacing`

**[L1-edge-case-001]** score=5 — The response does not generate any code, failing to utilize Word Office.js APIs, address the request, or demonstrate best practices, resulting in a complete lack of functionality.

**[L1-text-insert-001]** score=0 — The AI response failed to generate any code, completely neglecting the user's request to insert text at the end of the document, resulting in a total lack of correctness, completeness, functionality, and adherence to best practices.
- missing patterns: `Approved by management`

**[L1-text-edit-001]** score=0 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[gen-L2-form-field-003]** score=20 — A: 20
B:
- missing patterns: `GroupContentControl`

**[gen-L2-form-field-004]** score=5 — The code does not use real Word Office.js APIs, instead hallucinating methods like "insertContent a repeating section content control" and "insertTable" with incorrect syntax and parameters.
- missing patterns: `RepeatingSectionContentControl`
