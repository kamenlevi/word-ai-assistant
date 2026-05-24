# Word AI — Eval Results
**Last run:** 2026-05-24 03:45:11  
**Overall: 60.1/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 78.5/100 | — | 2 |
| paragraph-format | 88.0/100 | — | 2 |
| list-bullet | 0.0/100 | — | 1 |
| list-multilevel | 0.0/100 | — | 1 |
| style-apply | 80.0/100 | — | 1 |
| table-create | 90.0/100 | — | 1 |
| table-format | 72.0/100 | — | 1 |
| find-replace | 85.0/100 | — | 2 |
| find-replace-regex | 0.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 65.0/100 | — | 1 |
| track-changes-toggle | 91.5/100 | — | 2 |
| toc-generate | 83.0/100 | — | 1 |
| section-break | 67.0/100 | — | 1 |
| header-footer | 69.5/100 | — | 2 |
| columns | 72.0/100 | — | 1 |
| image-insert | 20.0/100 | — | 2 |
| watermark-insert | 0.0/100 | — | 1 |
| content-control | 83.0/100 | — | 2 |
| mail-merge | 20.0/100 | — | 1 |
| template-apply | 82.5/100 | — | 2 |
| document-generate | 73.0/100 | — | 1 |
| theme-apply | 17.3/100 | — | 3 |
| citation-bibliography | 20.0/100 | — | 1 |
| equation | 77.5/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 32.3/100 | — | 3 |
| margins-orientation | 83.0/100 | — | 2 |
| web-search-needed | 70.0/100 | — | 2 |
| refusal | 75.7/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 92.0/100 | — | 1 |
| edge-cases | 58.0/100 | — | 1 |
| form-field | 68.6/100 | — | 5 |
| text-insert | 80.0/100 | — | 1 |
| text-edit | 25.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 85 | ✓ | The code is mostly correct but lacks the surrounding Word.run context and uses a potentially undefined `addHeading` function, which is not a standard Word Office.js API, instead of using the built-in `paragraph.styleBuiltIn` and `paragraph.insertText` methods. |
| L1-heading-insert-002 | 72 | ✓ | The code uses a hallucinated method addHeading and addParagraph which are not real Word Office.js APIs, losing significant points in API correctness and approach dimensions. |
| L1-paragraph-format-001 | 93 | ✓ | The code accurately uses real Word Office.js APIs, fully addresses the request, and would execute without runtime errors, but loses points for approach due to manually setting font properties instead of using a style or a more robust method. |
| L1-paragraph-format-002 | 83 | ✓ | The code is mostly correct but loses points for not using the built-in `applyStyle` method to set the font size and for not checking if the first paragraph is indeed the title before applying the changes. |
| L1-list-bullet-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, and the provided code does not demonstrate a correct implementation of adding a bullet list using actual Word APIs. |
| L1-list-multilevel-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, and the provided code does not demonstrate a correct understanding of the Office.js API. |
| L1-style-apply-001 | 80 | ✓ | The code uses a real Word Office.js API, applies the style correctly, and would likely execute without runtime errors, but loses points for not fully addressing the request due to potential indexing issues and not using the most idiomatic approach with the available helpers. |
| L1-table-create-001 | 90 | ✓ | The code accurately uses the Word Office.js API to insert a table, but loses points for not fully addressing the request by not handling potential errors and not using the most efficient approach by directly using the `body.insertTable` method without unnecessary variables. |
| L1-table-format-001 | 72 | ✓ | The code uses a non-existent `styleTable` method, which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-find-replace-001 | 80 | ✓ | The code is mostly correct but lacks the Word.run wrapper and uses a generic replaceText function without specifying the context or range, which might lead to issues with the code's completeness and adherence to best practices. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient or best practice approach, and for potential issues with the load and sync order, which could lead to runtime errors. |
| L1-find-replace-regex-001 | 0 ⚠️ | ✗ | The AI response failed to generate any code, making it impossible to evaluate its correctness, completeness, executability, or approach to solving the user's request. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a correct API method to insert a footnote, but lacks consideration for the current selection and does not utilize the provided helpers, such as getting the current range and then calling range.insertFootnote. |
| L1-comment-insert-001 | 65 | ✓ | The code uses a correct API method to insert a comment, but it lacks specificity in searching for the first occurrence of "$4.2M" and may not work as expected due to potential issues with the search and insertion process. |
| L1-track-changes-toggle-001 | 100 ✅ | ✓ | The code perfectly addresses the request to turn on track changes using the correct and existing Word Office.js API method toggleTrackChanges, making it a complete, correct, and well-structured solution. |
| L1-track-changes-toggle-002 | 83 | ✓ | The code is mostly correct but loses points for not using the Word.run method to encapsulate the API calls and for not checking the document state before accepting revisions, which could lead to potential issues with the approach and completeness of the solution. |
| L1-toc-generate-001 | 83 | ✓ | The code correctly uses the insertTableOfContents API and addresses the user's request, but loses points for not fully utilizing best practices and available helpers, and for not perfectly handling potential edge cases in the generated code. |
| L1-section-break-001 | 67 | ✓ | The code uses non-existent Word Office.js APIs, such as "insertSectionBreak" and "setPageOrientation", which are not part of the real Word Office.js APIs, resulting in significant deductions in API correctness and approach dimensions. |
| L1-header-footer-001 | 67 | ✓ | The code uses a hallucinated method "setHeader" which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-header-footer-002 | 72 | ✓ | The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-columns-001 | 72 | ✓ | The code correctly uses real Word Office.js APIs and would likely execute without runtime errors, but it lacks completeness in addressing the user's request to switch the "rest of the document" into a two-column layout, as it only inserts a section break and applies the column layout without considering the existing content. |
| L1-image-insert-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-image-insert-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice. |
| L1-watermark-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "insertWatermark" which is not a real Word Office.js API, and also does not address the request to add the watermark to every page, resulting in a complete failure across all dimensions. |
| L1-content-control-001 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the request, but loses points for not using the Word.run method to ensure the code is executed in a context that supports asynchronous operations and for not checking if the content control can be inserted at the current cursor position. |
| L1-content-control-002 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the request, but loses points for not using the most idiomatic approach, such as providing more context about the insertion point, and not handling potential errors or edge cases. |
| L1-mail-merge-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "mailMergeReplace" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work at runtime. |
| L1-template-apply-001 | 80 | ✓ | The code correctly applies the resume template with the provided details, but loses points for not using the Word.run method to execute the template application and for not handling potential errors, and also for not following best practices in terms of error handling and code organization. |
| L1-template-apply-002 | 85 | ✓ | The code correctly applies the memo template with the requested fields, but loses points for API correctness due to the use of a potentially non-existent `applyTemplate` method, and for approach due to not using the `Word.run` context to ensure asynchronous execution. |
| L1-document-generate-001 | 73 | ✓ | The code correctly uses real Word Office.js APIs and would likely execute without runtime errors, but it only partially addresses the request by not including any paragraphs or content in the 5 sections, and does not fully utilize best practices such as using addParagraph to add content to each section. |
| L1-theme-apply-001 | 12 ⚠️ | ✓ | The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a significant penalty for API correctness and would not work in a real-world scenario. |
| L1-theme-apply-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not a valid method, resulting in a complete loss of points for API correctness and would it work dimensions. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work at runtime, although it partially addresses the user's request. |
| L1-citation-bibliography-001 | 20 ⚠️ | ✓ | The code heavily hallucinates methods, specifically "insertCitation" and "insertBibliography", which are not real Word Office.js APIs, resulting in a significant penalty. |
| L1-equation-001 | 72 | ✓ | The code uses a correct API but lacks best practices, such as using Word.run to ensure context and load the document, and does not handle potential errors or loading states, also it does not specify the location of the equation insertion which might lead to unexpected behavior. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the user's request, but loses points for not using a more robust method to handle equation insertion, such as checking the current selection or range, and for not providing any error handling or feedback to the user. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selection for clarity, but loses points for not considering alternative approaches, such as recommending the coach panel, and for using a generic phrase "Rewriting the selection for clarity" that could be more descriptive. |
| L1-read-query-001 | 40 ⚠️ | ✓ | The code fails to utilize real Word Office.js APIs, instead relying on a hallucinated `countWords()` method, which is not a valid API call. |
| L1-read-query-002 | 57 ⚠️ | ✓ | The code uses a hallucinated method "getReadability()" which is not a real Word Office.js API, resulting in heavy penalties in API correctness and approach dimensions. |
| L1-read-query-003 | 0 ⚠️ | ✓ | The code uses a hallucinated method "listHeadings()" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-margins-orientation-001 | 83 | ✓ | The code correctly uses the Office.js API and addresses the user's request, but loses points for not using the most efficient or best practice approach, such as directly utilizing the available margin settings in the Word API. |
| L1-margins-orientation-002 | 83 | ✓ | The code correctly uses the Office.js API and addresses the user's request, but lacks the surrounding Word.run context and does not utilize the best practice helpers, resulting in a straightforward yet not fully optimized solution. |
| L1-web-search-needed-001 | 67 | ✓ | The response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks clarity on how this information will be used to add a section to the document and does not provide any workarounds or Word-specific guidance. |
| L1-web-search-needed-002 | 73 | ✓ | The response accurately and clearly explains the Pythagorean theorem, but the code provided is unnecessary and doesn't add any value, resulting in a low bonus score. |
| L1-refusal-001 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but lacks a detailed explanation or alternative solutions, resulting in a slight deduction in accuracy and bonus points. |
| L1-refusal-002 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but lacks a more detailed explanation or alternative solutions, resulting in moderate scores across the dimensions. |
| L1-refusal-003 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js in creating charts, but the workaround provided is incomplete and lacks specificity, such as how to generate the chart image from the given data. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-paragraph-spacing-001 | 92 | ✓ | The code is mostly correct, but loses points for not using the `Word.run` method to wrap the code and ensure that the context is synced properly, and also for not checking if the paragraphs exist before trying to access their properties. |
| L1-edge-case-001 | 58 ⚠️ | ✓ | The code fails to fully address the request by not attempting to delete the paragraph and instead immediately throwing an error, and also lacks the use of best practices and available helpers in the Office JavaScript API. |
| L1-form-field-001 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the request, but loses points for not using a more descriptive title and for potential issues with the load and sync order, as well as not following best practices for content control insertion. |
| L1-text-insert-001 | 80 | ✓ | The code uses a real Word Office.js API (addParagraph) but loses points for not fully addressing the request to insert text at the end of the document, as addParagraph may not necessarily add the paragraph at the very end, and for not using the most appropriate approach, such as using range.insertText to insert the text at the end of the document. |
| L1-text-edit-001 | 25 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice. |
| gen-L2-form-field-001 | 72 | ✓ | The code partially addresses the request but loses points for not exactly matching the requested label for the date picker content control and not using the best practices for inserting content controls in a specific location, such as below the checkbox. |
| gen-L2-form-field-002 | 65 | ✗ | The code partially addresses the request but fails to correctly replace the existing plain text with a rich text content control and also uses a comboBox content control instead of a building block gallery control. |
| gen-L2-form-field-003 | 83 | ✗ | The code accurately uses Office.js APIs and addresses the request, but loses points for not using the most efficient approach, such as directly accessing the newly inserted group content control instead of using `doc.contentControls.getItem(0)`. |
| gen-L2-form-field-004 | 40 ⚠️ | ✗ | The code fails to use real Word Office.js APIs, instead hallucinating non-existent methods like "insertContentControl" and "insertTable" with incorrect signatures, resulting in a significant penalty for API correctness. |

## ⚠️ Needs attention

**[L1-list-bullet-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, and the provided code does not demonstrate a correct implementation of adding a bullet list using actual Word APIs.

**[L1-list-multilevel-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, and the provided code does not demonstrate a correct understanding of the Office.js API.

**[L1-find-replace-regex-001]** score=0 — The AI response failed to generate any code, making it impossible to evaluate its correctness, completeness, executability, or approach to solving the user's request.
- missing patterns: `matchWildcards`, `redacted`

**[L1-image-insert-001]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-image-insert-002]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice.

**[L1-watermark-insert-001]** score=0 — The code uses a hallucinated method "insertWatermark" which is not a real Word Office.js API, and also does not address the request to add the watermark to every page, resulting in a complete failure across all dimensions.

**[L1-mail-merge-001]** score=20 — The code uses a hallucinated method "mailMergeReplace" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work at runtime.

**[L1-theme-apply-001]** score=12 — The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a significant penalty for API correctness and would not work in a real-world scenario.

**[L1-theme-apply-002]** score=20 — The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not a valid method, resulting in a complete loss of points for API correctness and would it work dimensions.

**[L1-theme-apply-003]** score=20 — The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work at runtime, although it partially addresses the user's request.

**[L1-citation-bibliography-001]** score=20 — The code heavily hallucinates methods, specifically "insertCitation" and "insertBibliography", which are not real Word Office.js APIs, resulting in a significant penalty.

**[L1-read-query-001]** score=40 — The code fails to utilize real Word Office.js APIs, instead relying on a hallucinated `countWords()` method, which is not a valid API call.

**[L1-read-query-002]** score=57 — The code uses a hallucinated method "getReadability()" which is not a real Word Office.js API, resulting in heavy penalties in API correctness and approach dimensions.

**[L1-read-query-003]** score=0 — The code uses a hallucinated method "listHeadings()" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-edge-case-001]** score=58 — The code fails to fully address the request by not attempting to delete the paragraph and instead immediately throwing an error, and also lacks the use of best practices and available helpers in the Office JavaScript API.

**[L1-text-edit-001]** score=25 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice.

**[gen-L2-form-field-004]** score=40 — The code fails to use real Word Office.js APIs, instead hallucinating non-existent methods like "insertContentControl" and "insertTable" with incorrect signatures, resulting in a significant penalty for API correctness.
- missing patterns: `RepeatingSectionContentControl`
