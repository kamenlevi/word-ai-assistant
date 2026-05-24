# Word AI — Eval Results
**Last run:** 2026-05-24 09:58:53  
**Overall: 63.5/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 76.0/100 | — | 2 |
| paragraph-format | 88.0/100 | — | 2 |
| list-bullet | 25.0/100 | — | 1 |
| list-multilevel | 0.0/100 | — | 1 |
| style-apply | 80.0/100 | — | 1 |
| table-create | 83.0/100 | — | 1 |
| table-format | 88.0/100 | — | 1 |
| find-replace | 81.0/100 | — | 2 |
| find-replace-regex | 83.0/100 | — | 1 |
| footnote-insert | 80.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 91.5/100 | — | 2 |
| toc-generate | 83.0/100 | — | 1 |
| section-break | 72.0/100 | — | 1 |
| header-footer | 51.5/100 | — | 2 |
| columns | 85.0/100 | — | 1 |
| image-insert | 72.0/100 | — | 2 |
| watermark-insert | 0.0/100 | — | 1 |
| content-control | 90.0/100 | — | 2 |
| mail-merge | 72.0/100 | — | 1 |
| template-apply | 36.0/100 | — | 2 |
| document-generate | 85.0/100 | — | 1 |
| theme-apply | 12.3/100 | — | 3 |
| citation-bibliography | 20.0/100 | — | 1 |
| equation | 83.0/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 39.0/100 | — | 3 |
| margins-orientation | 84.0/100 | — | 2 |
| web-search-needed | 69.5/100 | — | 2 |
| refusal | 75.7/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 72.0/100 | — | 1 |
| edge-cases | 50.0/100 | — | 1 |
| form-field | 71.0/100 | — | 5 |
| text-insert | 83.0/100 | — | 1 |
| text-edit | 25.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 85 | ✓ | The code is mostly correct but loses points for using a potentially non-standard `addHeading` method, which is not a built-in Word Office.js API, although it seems to be a helper function that might be defined elsewhere in the codebase. |
| L1-heading-insert-002 | 67 | ✓ | The code uses a hallucinated method signature with addHeading and addParagraph, which are not real Word Office.js APIs, resulting in significant deductions in API correctness and approach dimensions. |
| L1-paragraph-format-001 | 93 | ✓ | The code accurately uses real Word Office.js APIs, fully addresses the request, and would execute without runtime errors, but loses points for approach due to manual font property modification instead of using a style or helper method. |
| L1-paragraph-format-002 | 83 | ✓ | The code is mostly correct but loses points for not using the provided helpers like applyStyle correctly and instead accessing paragraph properties directly, which is not the best practice. |
| L1-list-bullet-001 | 25 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice. |
| L1-list-multilevel-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-style-apply-001 | 80 | ✓ | The code correctly applies the Quote style to the third paragraph, but loses points for not using the most straightforward helper method and potential indexing issues, as paragraphs are typically 0-indexed, so the third paragraph would be at index 2, which is correct in this case, but the code does not explicitly handle potential paragraph indexing errors. |
| L1-table-create-001 | 83 | ✓ | The code is mostly correct and complete, but loses points for not fully utilizing best practices, such as potentially using a more specific table style or handling potential errors, and for not including the exact paragraph context 'Below is the revenue breakdown' in the code. |
| L1-table-format-001 | 88 | ✓ | The code is mostly correct but loses points for not using the available helper methods, such as applyStyle, and instead directly modifying the style property of the table. |
| L1-find-replace-001 | 72 | ✓ | The code uses a non-existent `replaceText` method, which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient or best practice approach, such as utilizing the `range.font` properties directly on the search results without explicitly loading the font property. |
| L1-find-replace-regex-001 | 83 | ✓ | The code uses the correct replaceText API with matchWildcards, but loses points for not fully addressing potential edge cases and not using the Word.run context to ensure asynchronous execution. |
| L1-footnote-insert-001 | 80 | ✓ | The code is mostly correct but loses points for not fully addressing the request by not specifying the range for the footnote insertion and not using the provided helpers like insertFootnote with a range object. |
| L1-comment-insert-001 | 72 | ✓ | The code uses a correct API to insert a comment, but it lacks specificity in searching for the exact occurrence of "$4.2M" and does not follow best practices for searching and inserting comments in the Word document. |
| L1-track-changes-toggle-001 | 100 ✅ | ✓ | The code perfectly addresses the request to turn on track changes using the correct and efficient `toggleTrackChanges` API, ensuring correctness, completeness, and adherence to best practices without any flaws. |
| L1-track-changes-toggle-002 | 83 | ✓ | The code is mostly correct, but it lacks a Word.run wrapper and does not handle potential errors, also it does not check if there are actually revisions to accept, which prevents it from being a complete and robust solution. |
| L1-toc-generate-001 | 83 | ✓ | The code correctly uses the insertTableOfContents API and addresses the request, but loses points for not fully utilizing best practices and available helpers, and for not perfectly handling potential edge cases in the generated code. |
| L1-section-break-001 | 72 | ✓ | The code uses correct APIs and is likely to work, but it lacks completeness in fully addressing the user's request to switch the next section to landscape, and it does not follow best practices by using lower-level APIs instead of available helpers. |
| L1-header-footer-001 | 83 | ✓ | The code is mostly correct but loses points for not using the best practice approach of utilizing available helpers, and the setHeader method is not a standard Word Office.js API, although it seems to be a custom wrapper around the actual API. |
| L1-header-footer-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-columns-001 | 85 | ✓ | The code correctly uses real Word Office.js APIs and would execute without runtime errors, but loses points for not fully addressing the request to switch the rest of the document into a two-column layout, as it only inserts a section break and columns without considering the existing content. |
| L1-image-insert-001 | 72 | ✓ | The code uses the correct insertImage API but loses points for not specifying the exact paragraph where the image should be inserted, and for not using best practices such as checking if the paragraph exists before inserting the image. |
| L1-image-insert-002 | 72 | ✓ | The code uses the correct insertImage API but loses points for not using the Word.run context and for not fully addressing the request with potential issues like image placement and sizing. |
| L1-watermark-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "insertWatermark" which is not a real Word Office.js API, and does not address the request to add a watermark to every page of a 4-page document. |
| L1-content-control-001 | 83 | ✓ | The code correctly uses the Office.js API to insert a content control, but loses points for not handling potential errors and not using the most descriptive variable names, and also for not fully addressing the request in terms of providing a complete date picker implementation. |
| L1-content-control-002 | 97 ✅ | ✓ | The code is nearly flawless, using the correct Office.js API to insert a content control, but loses a few points on approach for not explicitly checking if the cursor is at the end of the document before inserting the control, although the provided context implies this. |
| L1-mail-merge-001 | 72 | ✓ | The code uses the correct mailMergeReplace API but lacks proper error handling and context initialization, and does not fully follow best practices for Office.js development. |
| L1-template-apply-001 | 72 | ✓ | The code uses a non-standard `applyTemplate` method which is not a real Word Office.js API, and also lacks proper error handling and context loading, but it does address the user's request and attempts to use a template. |
| L1-template-apply-002 | 0 ⚠️ | ✓ | The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-document-generate-001 | 85 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach, such as potentially re-implementing the table of contents manually instead of relying on the built-in `insertTableOfContents` method, and not fully utilizing the available helpers for adding headings and paragraphs. |
| L1-theme-apply-001 | 5 ⚠️ | ✓ | The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would not work in practice. |
| L1-theme-apply-002 | 12 ⚠️ | ✓ | The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not a valid method, resulting in a complete loss of points for API correctness and would it work dimensions. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-citation-bibliography-001 | 20 ⚠️ | ✓ | The code uses hallucinated methods `insertCitation` and `insertBibliography` which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-equation-001 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the request, but loses points for not using a more specific best practice approach, such as using a range to specify the insertion point, and for potential issues with variable scope and load/sync order. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the user's request, but loses points for not checking the cursor position before inserting the equation and not using a more robust method to handle potential errors, and also for not fully following best practices in terms of error handling and code organization. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selection for clarity, is mostly clear, and directly addresses the question, but loses points for not considering alternative approaches, such as recommending the coach panel, and for using a simple replacement without acknowledging potential context-dependent nuances. |
| L1-read-query-001 | 40 ⚠️ | ✓ | The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `countWords()` method, which heavily penalizes the API correctness score. |
| L1-read-query-002 | 57 ⚠️ | ✓ | The code uses a hallucinated method `getReadability()` which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-read-query-003 | 20 ⚠️ | ✓ | The code hallucinates a non-existent `listHeadings()` method, which is not a real Word Office.js API, resulting in a significant penalty for API correctness. |
| L1-margins-orientation-001 | 85 | ✓ | The code correctly sets page margins using the Word Office.js API, but lacks best practices and does not utilize available helpers, and also assumes the context is already loaded without explicitly showing the Word.run call. |
| L1-margins-orientation-002 | 83 | ✓ | The code is mostly correct but lacks the surrounding Word.run context and uses a generic setPageOrientation method which might not be a real Word Office.js API, deducting points for potential API correctness and best practice issues. |
| L1-web-search-needed-001 | 67 | ✓ | The response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks clarity on how this information will be used to add a section to the Word document and does not provide any workarounds or Word-specific caveats. |
| L1-web-search-needed-002 | 72 | ✓ | The response accurately explains the Pythagorean theorem and directly addresses the question, but lacks clarity in the code implementation and does not provide any workarounds or Word-specific caveats, resulting in a low bonus score. |
| L1-refusal-001 | 83 | ✓ | The response is accurate, clear, and relevant, but loses points for not providing a more detailed workaround or alternative solution, such as using other APIs or third-party libraries to achieve the PDF export. |
| L1-refusal-002 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but lacks a more detailed explanation or alternative solutions, resulting in moderate scores across the dimensions. |
| L1-refusal-003 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js in creating charts, but loses points for not providing a more detailed workaround or alternative solution, such as using Excel APIs or suggesting a third-party library. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-paragraph-spacing-001 | 72 | ✓ | The code correctly uses Office.js APIs and addresses the request, but has unnecessary `applyStyle` calls, lacks error handling, and doesn't utilize best practices such as using `body.paragraphs.forEach` instead of manual looping. |
| L1-edge-case-001 | 50 ⚠️ | ✓ | The code fails to attempt the deletion of the paragraph or provide a meaningful guard against out-of-range paragraph numbers, instead immediately throwing an error without utilizing Word Office.js APIs to interact with the document's paragraphs. |
| L1-form-field-001 | 83 | ✓ | The code correctly uses the Word Office.js API to insert a dropdown content control, but loses points for not fully addressing potential errors or edge cases, and not using the most optimal approach by directly calling the insertContentControl method without considering the current context or potential existing content controls. |
| L1-text-insert-001 | 83 | ✓ | The code uses a real Word Office.js API (addParagraph) and is likely to execute without runtime errors, but loses points for not fully addressing the request (inserting at the end of the document, not necessarily after the last paragraph) and not using the most optimal approach (e.g., using body.paragraphs.getLast() to ensure insertion at the end). |
| L1-text-edit-001 | 25 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a severe penalty for API correctness. |
| gen-L2-form-field-001 | 72 | ✓ | The code uses correct Office.js APIs and is likely to work, but loses points for not fully addressing the request (missing label and incorrect title for date picker) and not using best practices (e.g., no explicit paragraph or range handling). |
| gen-L2-form-field-002 | 72 | ✗ | The code uses correct Office.js APIs and is mostly complete, but it lacks proper replacement of the existing plain text "Name: _______" and does not fully follow best practices, such as using the `insertContentControl` method correctly and not addressing the building block gallery control properly. |
| gen-L2-form-field-003 | 83 | ✗ | The code is mostly correct and complete, but loses points for not using the `Word.run` context to ensure asynchronous operations are properly synchronized, and for not utilizing the available helpers, such as `insertContentControl` with a more robust options object. |
| gen-L2-form-field-004 | 45 ⚠️ | ✗ | The code fails to use the correct Office.js API for inserting a repeating section content control, instead using a plain text content control, and also does not fully address the request by not specifying the repeating section properties. |

## ⚠️ Needs attention

**[L1-list-bullet-001]** score=25 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice.

**[L1-list-multilevel-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-header-footer-002]** score=20 — The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-watermark-insert-001]** score=0 — The code uses a hallucinated method "insertWatermark" which is not a real Word Office.js API, and does not address the request to add a watermark to every page of a 4-page document.

**[L1-template-apply-002]** score=0 — The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-theme-apply-001]** score=5 — The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would not work in practice.

**[L1-theme-apply-002]** score=12 — The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not a valid method, resulting in a complete loss of points for API correctness and would it work dimensions.

**[L1-theme-apply-003]** score=20 — The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-citation-bibliography-001]** score=20 — The code uses hallucinated methods `insertCitation` and `insertBibliography` which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-read-query-001]** score=40 — The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `countWords()` method, which heavily penalizes the API correctness score.

**[L1-read-query-002]** score=57 — The code uses a hallucinated method `getReadability()` which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions.

**[L1-read-query-003]** score=20 — The code hallucinates a non-existent `listHeadings()` method, which is not a real Word Office.js API, resulting in a significant penalty for API correctness.

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-edge-case-001]** score=50 — The code fails to attempt the deletion of the paragraph or provide a meaningful guard against out-of-range paragraph numbers, instead immediately throwing an error without utilizing Word Office.js APIs to interact with the document's paragraphs.

**[L1-text-edit-001]** score=25 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a severe penalty for API correctness.

**[gen-L2-form-field-004]** score=45 — The code fails to use the correct Office.js API for inserting a repeating section content control, instead using a plain text content control, and also does not fully address the request by not specifying the repeating section properties.
- missing patterns: `RepeatingSectionContentControl`
