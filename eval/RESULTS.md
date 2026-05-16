# Word AI — Eval Results
**Last run:** 2026-05-16 09:44:21  
**Overall: 57.7/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 69.5/100 | — | 2 |
| paragraph-format | 47.5/100 | — | 2 |
| list-bullet | 20.0/100 | — | 1 |
| list-multilevel | 72.0/100 | — | 1 |
| style-apply | 80.0/100 | — | 1 |
| table-create | 90.0/100 | — | 1 |
| table-format | 72.0/100 | — | 1 |
| find-replace | 81.0/100 | — | 2 |
| find-replace-regex | 83.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 90.0/100 | — | 2 |
| toc-generate | 72.0/100 | — | 1 |
| section-break | 80.0/100 | — | 1 |
| header-footer | 51.5/100 | — | 2 |
| columns | 72.0/100 | — | 1 |
| image-insert | 46.0/100 | — | 2 |
| watermark-insert | 40.0/100 | — | 1 |
| content-control | 81.5/100 | — | 2 |
| mail-merge | 80.0/100 | — | 1 |
| template-apply | 10.0/100 | — | 2 |
| document-generate | 23.0/100 | — | 1 |
| theme-apply | 32.3/100 | — | 3 |
| citation-bibliography | 20.0/100 | — | 1 |
| equation | 75.0/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 20.0/100 | — | 3 |
| margins-orientation | 40.0/100 | — | 2 |
| web-search-needed | 70.0/100 | — | 2 |
| refusal | 73.3/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 67.0/100 | — | 1 |
| edge-cases | 40.0/100 | — | 1 |
| form-field | 70.4/100 | — | 5 |
| text-insert | 80.0/100 | — | 1 |
| text-edit | 20.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 67 | ✓ | The code uses a non-standard method "addHeading" which is not a real Word Office.js API, indicating a lack of adherence to best practices and available helpers. |
| L1-heading-insert-002 | 72 | ✓ | The code uses a hallucinated method addHeading and addParagraph which are not real Word Office.js APIs, losing significant points in API correctness and approach dimensions. |
| L1-paragraph-format-001 | 0 ⚠️ | ✗ | The code attempts to add a new paragraph with the specified style instead of modifying the existing first paragraph, and also uses the `addParagraph` method which is not a real Word Office.js API, resulting in a complete failure to address the user's request. |
| L1-paragraph-format-002 | 95 ✅ | ✓ | The code is mostly correct and complete, but loses points on approach for not using the available helper methods, such as `paragraph.styleBuiltIn` or `applyStyle`, to set the font size and alignment, instead opting for direct property manipulation. |
| L1-list-bullet-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice. |
| L1-list-multilevel-001 | 72 | ✓ | The code uses a non-standard method "addList" which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-style-apply-001 | 80 | ✓ | The code correctly applies the Quote style to the third paragraph using the applyStyle method, but loses points for not using the Word.run method to execute the operation and for not checking if the paragraph exists before applying the style. |
| L1-table-create-001 | 90 | ✓ | The code accurately uses the Word Office.js API to insert a table with headers, but loses points for not fully addressing potential edge cases and not using the most optimal approach by directly utilizing the `body.insertTable` method with a more robust configuration. |
| L1-table-format-001 | 72 | ✓ | The code uses a non-existent `styleTable` method, which is not a real Word Office.js API, and lacks the proper `Word.run` context, but it attempts to address the request and has a correct sync order. |
| L1-find-replace-001 | 72 | ✓ | The code uses a correct API method replaceText, but lacks the context of a Word.run call and does not handle potential errors, also it does not follow best practices as it does not utilize the Word.run method which is the entry point for most Word JavaScript API operations. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach, such as using `range.font` properties directly on the search results without loading the entire `font` property, and not handling potential errors that may occur during the `await context.sync()` calls. |
| L1-find-replace-regex-001 | 83 | ✓ | The code uses the correct Word Office.js API replaceText with matchWildcards, but loses points for not fully addressing potential edge cases and not using the most robust best practices, such as wrapping the code in a Word.run context to ensure asynchronous execution. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a correct API but lacks completeness in addressing the request, has potential issues with variable scope and method signatures, and does not use the best practice approach of utilizing the provided helpers, specifically the insertFootnote helper is not a standard Word Office.js API. |
| L1-comment-insert-001 | 72 | ✓ | The code uses a non-existent `insertComment` method with two parameters, which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-track-changes-toggle-001 | 97 ✅ | ✓ | The code perfectly addresses the request using the correct API, but loses a few points for not being wrapped in a Word.run context, which is a best practice for ensuring asynchronous operations are properly handled. |
| L1-track-changes-toggle-002 | 83 | ✓ | The code is mostly correct and complete, but it lacks the surrounding Word.run context and does not explicitly check if there are tracked revisions before accepting them, which might lead to minor issues with best practices and robustness. |
| L1-toc-generate-001 | 72 | ✓ | The code uses a correct API method to insert a Table of Contents, but loses points for not using the provided helpers and for not fully addressing the request with more specific heading level handling. |
| L1-section-break-001 | 80 | ✓ | The code is mostly correct but lacks specificity in targeting the next section for landscape orientation and does not fully follow best practices for inserting section breaks and setting page orientation in Word Office.js APIs. |
| L1-header-footer-001 | 83 | ✓ | The code is mostly correct but loses points for not using the best practice approach of directly manipulating the header section, instead relying on a generic setHeader method which may not be a real Word Office.js API. |
| L1-header-footer-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work at runtime. |
| L1-columns-001 | 72 | ✓ | The code uses correct APIs and would likely work, but it only partially addresses the request by not specifying the range for the two-column layout, and it does not follow best practices by not using a more targeted approach to apply the layout change. |
| L1-image-insert-001 | 72 | ✓ | The code uses a non-standard `insertImage` method which is not a real Word Office.js API, and also lacks proper error handling and context setup, but it attempts to address the user's request and uses a correct width parameter. |
| L1-image-insert-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a significant penalty in API correctness and would not work in Word without runtime errors. |
| L1-watermark-insert-001 | 40 ⚠️ | ✓ | The code is incomplete and does not fully address the request, as it only inserts a watermark without specifying the scope (every page) and does not use the correct Word Office.js API, such as Word.run, to achieve the desired result. |
| L1-content-control-001 | 83 | ✓ | The code uses a correct Office.js API to insert a content control, but loses points for not using the `Word.run` context and for not fully addressing potential errors or edge cases, such as checking if the content control was successfully inserted. |
| L1-content-control-002 | 80 | ✓ | The code correctly uses the insertContentControl API and addresses the user's request, but loses points for not fully considering the cursor position at the end of the document and not using a more specific best practice approach. |
| L1-mail-merge-001 | 80 | ✓ | The code correctly uses the mailMergeReplace API and addresses the user's request, but loses points for not using Word.run to execute the mail merge and for not handling potential errors, and also for not using the most efficient approach by directly using the provided template placeholders. |
| L1-template-apply-001 | 20 ⚠️ | ✓ | A: 20
B: |
| L1-template-apply-002 | 0 ⚠️ | ✓ | The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions. |
| L1-document-generate-001 | 23 ⚠️ | ✓ | A: 23
B: |
| L1-theme-apply-001 | 5 ⚠️ | ✓ | The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would not work in practice. |
| L1-theme-apply-002 | 72 | ✓ | The code uses the correct designTheme and applyTheme APIs, but lacks completeness in addressing the request, as it doesn't specify how the theme will be applied to the existing body content, and the approach could be improved by utilizing more specific helpers for styling the document. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code heavily hallucinates a non-existent API method "tweakTheme" which is not a real Word Office.js API, resulting in a significant penalty. |
| L1-citation-bibliography-001 | 20 ⚠️ | ✓ | The code hallucinates non-existent methods insertCitation and insertBibliography, which are not part of the real Word Office.js APIs, resulting in a significant penalty for API correctness and would-it-work dimensions. |
| L1-equation-001 | 67 | ✓ | The code uses a non-existent `insertEquation` method, which is not a real Word Office.js API, resulting in significant penalties for API correctness and approach. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the user's request, but loses points for not using a more robust method to handle equation insertion, such as utilizing the range object to specify the exact location of the equation. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selected text for clarity, but loses points for not considering alternative approaches, such as recommending the coach panel, and for using a somewhat generic phrase "Rewriting the selected text for clarity" that could be more descriptive. |
| L1-read-query-001 | 20 ⚠️ | ✓ | The code hallucinates a non-existent `countWords()` method, which is not a real Word Office.js API, and does not use the correct API to count words, such as `body.getText()` or `range.getText()`. |
| L1-read-query-002 | 40 ⚠️ | ✓ | The code uses a hallucinated method "getReadability()" which is not a real Word Office.js API, resulting in a significant penalty for API correctness. |
| L1-read-query-003 | 0 ⚠️ | ✓ | The code hallucinates a non-existent `listHeadings()` method, which is not a real Word Office.js API, and does not provide any actual implementation to list the headings in the document. |
| L1-margins-orientation-001 | 0 ⚠️ | ✓ | The AI response failed to generate any code, leaving the request to set page margins to 1 inch on all sides unaddressed and without a viable solution. |
| L1-margins-orientation-002 | 80 | ✓ | The code is mostly correct but lacks consideration for the document's current state and potential existing content, and does not use the most efficient or best practice approach for setting page orientation. |
| L1-web-search-needed-001 | 67 | ✓ | The response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks a clear explanation or workaround for integrating this information into a Microsoft Word document using Office JavaScript API. |
| L1-web-search-needed-002 | 73 | ✓ | The response accurately and clearly explains the Pythagorean theorem, directly addressing the question, but loses points for not providing any workarounds or Word-specific caveats, and also for including unnecessary code in the response. |
| L1-refusal-001 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not offering more detailed guidance on the "Save As" feature or alternative solutions, resulting in a moderate bonus score. |
| L1-refusal-002 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but lacks a more detailed explanation or alternative solutions, resulting in moderate scores across the dimensions. |
| L1-refusal-003 | 65 | ✓ | The response accurately identifies the limitation of Word Office.js but loses points for not providing a more detailed workaround or alternative solution that directly addresses the user's request for an Excel-style bar chart. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions. |
| L1-paragraph-spacing-001 | 67 | ✓ | The code uses correct Word Office.js APIs and addresses the request, but has issues with approach, using a manual loop instead of available helpers, and lacks consideration for potential paragraph additions or removals beyond the initial 6 paragraphs. |
| L1-edge-case-001 | 40 ⚠️ | ✓ | The code does not use any real Word Office.js APIs, instead simply stating that no action will be taken, which results in a complete loss of points for API correctness and approach, while still partially addressing the request and being executable without runtime errors. |
| L1-form-field-001 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the request, but loses points for not using a more descriptive title and tag, and for not handling potential errors or edge cases, which affects its completeness and adherence to best practices. |
| L1-text-insert-001 | 80 | ✓ | The code uses a correct and existing API method `addParagraph` but loses points for not explicitly addressing the "at the end of the document" requirement and for not using the most precise method to ensure the text is inserted at the very end, such as manipulating the `body` object directly. |
| L1-text-edit-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice. |
| gen-L2-form-field-001 | 75 | ✓ | The code partially addresses the request but fails to exactly match the label for the date picker content control and does not ensure the date picker is inserted below the checkbox, while also using a reasonable approach with available APIs. |
| gen-L2-form-field-002 | 50 ⚠️ | ✗ | The code incorrectly inserts two rich text content controls instead of replacing the existing plain text "Name: _______" with a single rich text content control and adding a building block gallery control as requested. |
| gen-L2-form-field-003 | 72 | ✗ | The code is mostly correct but loses points for not properly grouping the dropdown content controls within the group content control as requested, and for not using best practices such as setting the parent of the dropdown content controls to the group content control. |
| gen-L2-form-field-004 | 72 | ✗ | The code partially addresses the request but lacks the use of a repeating section content control and instead uses a rich text content control, also it does not fully follow best practices and available helpers. |

## ⚠️ Needs attention

**[L1-paragraph-format-001]** score=0 — The code attempts to add a new paragraph with the specified style instead of modifying the existing first paragraph, and also uses the `addParagraph` method which is not a real Word Office.js API, resulting in a complete failure to address the user's request.
- missing patterns: `font.bold`, `font.size`

**[L1-list-bullet-001]** score=20 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice.

**[L1-header-footer-002]** score=20 — The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work at runtime.

**[L1-image-insert-002]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a significant penalty in API correctness and would not work in Word without runtime errors.

**[L1-watermark-insert-001]** score=40 — The code is incomplete and does not fully address the request, as it only inserts a watermark without specifying the scope (every page) and does not use the correct Word Office.js API, such as Word.run, to achieve the desired result.

**[L1-template-apply-001]** score=20 — A: 20
B:

**[L1-template-apply-002]** score=0 — The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions.

**[L1-document-generate-001]** score=23 — A: 23
B:

**[L1-theme-apply-001]** score=5 — The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would not work in practice.

**[L1-theme-apply-003]** score=20 — The code heavily hallucinates a non-existent API method "tweakTheme" which is not a real Word Office.js API, resulting in a significant penalty.

**[L1-citation-bibliography-001]** score=20 — The code hallucinates non-existent methods insertCitation and insertBibliography, which are not part of the real Word Office.js APIs, resulting in a significant penalty for API correctness and would-it-work dimensions.

**[L1-read-query-001]** score=20 — The code hallucinates a non-existent `countWords()` method, which is not a real Word Office.js API, and does not use the correct API to count words, such as `body.getText()` or `range.getText()`.

**[L1-read-query-002]** score=40 — The code uses a hallucinated method "getReadability()" which is not a real Word Office.js API, resulting in a significant penalty for API correctness.

**[L1-read-query-003]** score=0 — The code hallucinates a non-existent `listHeadings()` method, which is not a real Word Office.js API, and does not provide any actual implementation to list the headings in the document.

**[L1-margins-orientation-001]** score=0 — The AI response failed to generate any code, leaving the request to set page margins to 1 inch on all sides unaddressed and without a viable solution.

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions.

**[L1-edge-case-001]** score=40 — The code does not use any real Word Office.js APIs, instead simply stating that no action will be taken, which results in a complete loss of points for API correctness and approach, while still partially addressing the request and being executable without runtime errors.

**[L1-text-edit-001]** score=20 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice.

**[gen-L2-form-field-002]** score=50 — The code incorrectly inserts two rich text content controls instead of replacing the existing plain text "Name: _______" with a single rich text content control and adding a building block gallery control as requested.
- missing patterns: `RichTextContentControl`
