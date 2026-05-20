# Word AI — Eval Results
**Last run:** 2026-05-20 10:56:11  
**Overall: 55.9/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 33.5/100 | — | 2 |
| paragraph-format | 41.5/100 | — | 2 |
| list-bullet | 0.0/100 | — | 1 |
| list-multilevel | 72.0/100 | — | 1 |
| style-apply | 80.0/100 | — | 1 |
| table-create | 89.0/100 | — | 1 |
| table-format | 67.0/100 | — | 1 |
| find-replace | 81.0/100 | — | 2 |
| find-replace-regex | 83.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 91.5/100 | — | 2 |
| toc-generate | 72.0/100 | — | 1 |
| section-break | 72.0/100 | — | 1 |
| header-footer | 76.0/100 | — | 2 |
| columns | 72.0/100 | — | 1 |
| image-insert | 57.5/100 | — | 2 |
| watermark-insert | 30.0/100 | — | 1 |
| content-control | 84.0/100 | — | 2 |
| mail-merge | 0.0/100 | — | 1 |
| template-apply | 20.0/100 | — | 2 |
| document-generate | 83.0/100 | — | 1 |
| theme-apply | 16.7/100 | — | 3 |
| citation-bibliography | 20.0/100 | — | 1 |
| equation | 75.0/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 32.3/100 | — | 3 |
| margins-orientation | 77.5/100 | — | 2 |
| web-search-needed | 66.0/100 | — | 2 |
| refusal | 72.0/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 72.0/100 | — | 1 |
| edge-cases | 10.0/100 | — | 1 |
| form-field | 57.0/100 | — | 5 |
| text-insert | 80.0/100 | — | 1 |
| text-edit | 0.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 0 ⚠️ | ✓ | A:  |
| L1-heading-insert-002 | 67 | ✓ | The code uses a hallucinated method signature with addHeading and addParagraph, which are not real Word Office.js APIs, indicating a lack of best practices and incorrect approach. |
| L1-paragraph-format-001 | 0 ⚠️ | ✗ | The code attempts to add a new paragraph with the specified style instead of modifying the existing first paragraph, and it uses the `addParagraph` method which is not a real Word Office.js API, resulting in a complete failure to address the user's request. |
| L1-paragraph-format-002 | 83 | ✓ | The code is mostly correct but loses points for not using the built-in `applyStyle` method to set the font size and for not checking if the first paragraph is indeed the title before applying the changes. |
| L1-list-bullet-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions. |
| L1-list-multilevel-001 | 72 | ✓ | The code uses a non-standard method "addList" which is not a real Word Office.js API, and also lacks proper context and error handling, but it attempts to address the user's request and might work with some modifications. |
| L1-style-apply-001 | 80 | ✓ | The code correctly applies the Quote style to the third paragraph, but uses a generic applyStyle method which may not be the most efficient or Word-specific approach, and assumes the paragraph index is 2, which may not always be the case. |
| L1-table-create-001 | 89 | ✓ | The code is mostly correct and complete, but loses points for not fully addressing the request by only inserting three rows of data instead of four as specified, and not using the most optimal approach by directly using the insertTable method without considering the existing document context. |
| L1-table-format-001 | 67 | ✓ | The code uses a non-standard method `styleTable` which is not a real Word Office.js API, and lacks best practices, resulting in significant deductions in API correctness and approach dimensions. |
| L1-find-replace-001 | 72 | ✓ | The code uses a correct API method replaceText, but lacks the context of a Word.run call and does not handle potential errors, also it does not follow best practices as it does not utilize the Word.run method which is the entry point for most Word JavaScript API operations. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach and for potential issues with the load and sync order, which could lead to runtime errors. |
| L1-find-replace-regex-001 | 83 | ✓ | The code correctly uses the Word Office.js API and addresses the request, but loses points for not handling potential errors and not using the most efficient approach, such as checking if the replacement is necessary before calling replaceText. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a correct API but lacks completeness in addressing the request, has some potential issues with variable scope and load order, and does not use the best practice approach of utilizing the available helpers, specifically the insertFootnote helper is not wrapped in a Word.run context. |
| L1-comment-insert-001 | 72 | ✓ | The code uses a correct API to insert a comment, but lacks specificity in targeting the first occurrence of the text and does not utilize best practices for searching and replacing text in the document. |
| L1-track-changes-toggle-001 | 100 ✅ | ✓ | The code perfectly addresses the request to turn on track changes using the correct and available Word Office.js API, specifically the toggleTrackChanges helper, with no flaws or issues identified. |
| L1-track-changes-toggle-002 | 83 | ✓ | The code is mostly correct but loses points for not using the Word.run method to execute the acceptAllRevisions action, and for not fully following best practices by directly calling acceptAllRevisions without considering the context of the document. |
| L1-toc-generate-001 | 72 | ✓ | The code uses a correct API but loses points for not fully addressing the request with a more tailored approach, such as specifying the exact location of the Table of Contents, and for not utilizing the best practices and available helpers. |
| L1-section-break-001 | 72 | ✓ | The code uses correct APIs and is mostly complete, but lacks best practices and does not fully address the request as it doesn't specify the section to apply the landscape orientation to, and uses generic methods instead of available helpers. |
| L1-header-footer-001 | 80 | ✓ | The code is mostly correct but lacks consideration for the specific section or page range where the header should be applied, and uses a non-standard method name `setHeader` instead of the correct Word API methods like `section.getHeader` or `section.addHeader`, and also does not handle potential errors or loading issues. |
| L1-header-footer-002 | 72 | ✓ | The code uses a non-standard method "addPageNumbers" which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-columns-001 | 72 | ✓ | The code correctly uses real Word Office.js APIs and would likely execute without runtime errors, but it lacks completeness in addressing the user's request to switch the rest of the document into a two-column layout, and does not use best practices or available helpers. |
| L1-image-insert-001 | 80 | ✓ | The code correctly uses the insertImage API and addresses the request, but loses points for not specifying the exact paragraph where the image should be inserted and potential issues with the context.sync() order. |
| L1-image-insert-002 | 35 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a severe penalty for API correctness. |
| L1-watermark-insert-001 | 30 ⚠️ | ✓ | The code is incomplete and does not fully address the request, as it only inserts a watermark without specifying the scope to every page, and also uses a non-existent `insertWatermark` method instead of a real Word Office.js API. |
| L1-content-control-001 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the request, but loses points for not using the Word.run method to ensure the code is executed in the correct context and for not handling potential errors, and also for not using the most idiomatic approach to inserting a content control. |
| L1-content-control-002 | 85 | ✓ | The code correctly uses the Office.js API to insert a content control, but loses points for not using the more specific and recommended helper method, and for not addressing potential issues with the cursor position at the end of the document. |
| L1-mail-merge-001 | 0 ⚠️ | ✓ | A:  |
| L1-template-apply-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work at runtime, although it partially addresses the request. |
| L1-template-apply-002 | 20 ⚠️ | ✓ | The code hallucinates a non-existent API method "applyTemplate" which is not a real Word Office.js API, resulting in a significant penalty for API correctness. |
| L1-document-generate-001 | 83 | ✓ | The code effectively uses real Word Office.js APIs and addresses the request, but loses points for not fully utilizing best practices, such as potentially using more specific paragraph styles or content controls, and for minor potential issues with load and sync order. |
| L1-theme-apply-001 | 10 ⚠️ | ✓ | The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-theme-apply-002 | 20 ⚠️ | ✓ | The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work at runtime. |
| L1-citation-bibliography-001 | 20 ⚠️ | ✓ | The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-equation-001 | 67 | ✓ | The code uses a non-existent `insertEquation` method, which is not a real Word Office.js API, resulting in significant penalties in the API correctness and approach dimensions. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the request, but loses points for not considering the cursor position at the end of the paragraph, which might require additional code to handle properly. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selection for clarity, is easy to understand, and directly addresses the question, but loses points for not considering alternative approaches, such as recommending the coach panel or handling potential exceptions. |
| L1-read-query-001 | 40 ⚠️ | ✓ | The code fails to use real Word Office.js APIs, instead hallucinating a non-existent "countWords" method, which heavily penalizes the API correctness score. |
| L1-read-query-002 | 57 ⚠️ | ✓ | The code uses a hallucinated method "getReadability()" which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-read-query-003 | 0 ⚠️ | ✓ | The code hallucinates a non-existent `listHeadings()` method, which is not a real Word Office.js API, and does not provide any actual implementation to list the headings in the document. |
| L1-margins-orientation-001 | 83 | ✓ | The code correctly sets page margins using the Office.js API, but loses points for not using a more explicit method to specify inch units and for not fully utilizing best practices in its approach. |
| L1-margins-orientation-002 | 72 | ✓ | The code uses a non-existent setPageOrientation method, which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-web-search-needed-001 | 67 | ✓ | The AI response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks clarity on how this information will be used to add a section to the Word document and does not offer any workarounds or suggestions. |
| L1-web-search-needed-002 | 65 | ✓ | The response accurately and clearly explains the Pythagorean theorem, directly addressing the question, but loses points for accuracy due to the incomplete explanation of the theorem's context and application, and lacks any bonus information or workarounds. |
| L1-refusal-001 | 72 | ✓ | The response is mostly accurate and clear, but lacks a more detailed workaround or alternative solution, resulting in lower scores for relevance and bonus. |
| L1-refusal-002 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but lacks a more detailed explanation or alternative solutions, resulting in a moderate score for relevance and bonus. |
| L1-refusal-003 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js in creating Excel-style charts, but loses points for not providing a more detailed workaround, such as using the Excel REST API or a library to generate the chart, and for inserting a generic image instead of attempting to create a chart with the provided data. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-paragraph-spacing-001 | 72 | ✓ | The code uses real Word Office.js APIs and fully addresses the request, but loses points for not using the best practices, such as directly accessing paragraphFormat without ensuring the paragraph is properly loaded, and applying a style unnecessarily before setting line spacing. |
| L1-edge-case-001 | 10 ⚠️ | ✓ | The code fails to use the Word Office.js API to delete a paragraph and instead throws an error, also it does not attempt to guard against out-of-range paragraph numbers or provide a helper function to handle such cases. |
| L1-form-field-001 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the request, but loses points for not using a more descriptive title and tag, and for not handling potential errors or edge cases, which affects its completeness and adherence to best practices. |
| L1-text-insert-001 | 80 | ✓ | The code uses a real Office.js API and is mostly complete, but loses points for not fully addressing the request to insert text at the end of the document, as addParagraph may not necessarily add to the end, and for not using the most appropriate best practice approach. |
| L1-text-edit-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions. |
| gen-L2-form-field-001 | 0 ⚠️ | ✓ | A:  |
| gen-L2-form-field-002 | 47 ⚠️ | ✗ | The code fails to accurately address the user's request by inserting two rich text content controls instead of replacing the existing plain text "Name: _______" with one and adding a building block gallery control. |
| gen-L2-form-field-003 | 83 | ✗ | The code accurately addresses the request and uses correct APIs, but loses points for not fully utilizing best practices and available helpers, and for minor potential issues with variable scope and load order. |
| gen-L2-form-field-004 | 72 | ✗ | The code fails to fully address the request by not actually creating a repeating section content control, instead inserting a plain text content control and a table, which loses points for completeness and approach. |

## ⚠️ Needs attention

**[L1-heading-insert-001]** score=0 — A: 

**[L1-paragraph-format-001]** score=0 — The code attempts to add a new paragraph with the specified style instead of modifying the existing first paragraph, and it uses the `addParagraph` method which is not a real Word Office.js API, resulting in a complete failure to address the user's request.
- missing patterns: `font.bold`, `font.size`

**[L1-list-bullet-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions.

**[L1-image-insert-002]** score=35 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a severe penalty for API correctness.

**[L1-watermark-insert-001]** score=30 — The code is incomplete and does not fully address the request, as it only inserts a watermark without specifying the scope to every page, and also uses a non-existent `insertWatermark` method instead of a real Word Office.js API.

**[L1-mail-merge-001]** score=0 — A: 

**[L1-template-apply-001]** score=20 — The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work at runtime, although it partially addresses the request.

**[L1-template-apply-002]** score=20 — The code hallucinates a non-existent API method "applyTemplate" which is not a real Word Office.js API, resulting in a significant penalty for API correctness.

**[L1-theme-apply-001]** score=10 — The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-theme-apply-002]** score=20 — The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-theme-apply-003]** score=20 — The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work at runtime.

**[L1-citation-bibliography-001]** score=20 — The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-read-query-001]** score=40 — The code fails to use real Word Office.js APIs, instead hallucinating a non-existent "countWords" method, which heavily penalizes the API correctness score.

**[L1-read-query-002]** score=57 — The code uses a hallucinated method "getReadability()" which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions.

**[L1-read-query-003]** score=0 — The code hallucinates a non-existent `listHeadings()` method, which is not a real Word Office.js API, and does not provide any actual implementation to list the headings in the document.

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-edge-case-001]** score=10 — The code fails to use the Word Office.js API to delete a paragraph and instead throws an error, also it does not attempt to guard against out-of-range paragraph numbers or provide a helper function to handle such cases.

**[L1-text-edit-001]** score=0 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions.

**[gen-L2-form-field-001]** score=0 — A: 

**[gen-L2-form-field-002]** score=47 — The code fails to accurately address the user's request by inserting two rich text content controls instead of replacing the existing plain text "Name: _______" with one and adding a building block gallery control.
- missing patterns: `RichTextContentControl`
