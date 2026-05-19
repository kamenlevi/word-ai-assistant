# Word AI — Eval Results
**Last run:** 2026-05-19 03:40:05  
**Overall: 61.3/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 78.5/100 | — | 2 |
| paragraph-format | 88.0/100 | — | 2 |
| list-bullet | 20.0/100 | — | 1 |
| list-multilevel | 0.0/100 | — | 1 |
| style-apply | 60.0/100 | — | 1 |
| table-create | 83.0/100 | — | 1 |
| table-format | 72.0/100 | — | 1 |
| find-replace | 85.0/100 | — | 2 |
| find-replace-regex | 83.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 91.5/100 | — | 2 |
| toc-generate | 72.0/100 | — | 1 |
| section-break | 72.0/100 | — | 1 |
| header-footer | 43.5/100 | — | 2 |
| columns | 80.0/100 | — | 1 |
| image-insert | 77.5/100 | — | 2 |
| watermark-insert | 0.0/100 | — | 1 |
| content-control | 81.5/100 | — | 2 |
| mail-merge | 72.0/100 | — | 1 |
| template-apply | 68.5/100 | — | 2 |
| document-generate | 83.0/100 | — | 1 |
| theme-apply | 21.7/100 | — | 3 |
| citation-bibliography | 20.0/100 | — | 1 |
| equation | 83.0/100 | — | 2 |
| writing-coach | 0.0/100 | — | 1 |
| read-query | 32.3/100 | — | 3 |
| margins-orientation | 82.5/100 | — | 2 |
| web-search-needed | 66.0/100 | — | 2 |
| refusal | 79.3/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 72.0/100 | — | 1 |
| edge-cases | 40.0/100 | — | 1 |
| form-field | 74.0/100 | — | 5 |
| text-insert | 38.0/100 | — | 1 |
| text-edit | 20.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 85 | ✓ | The code is mostly correct but loses points for using a hypothetical `addHeading` method which is not a real Word Office.js API, instead of using the actual `paragraph.styleBuiltIn` and `paragraph.insertText` methods. |
| L1-heading-insert-002 | 72 | ✓ | The code uses a hallucinated method addHeading and addParagraph, which are not real Word Office.js APIs, resulting in significant deductions in API correctness and approach dimensions. |
| L1-paragraph-format-001 | 93 | ✓ | The code accurately uses real Word Office.js APIs, fully addresses the request, and would execute without runtime errors, but loses points for approach due to manual font property modification instead of using available style or formatting helpers. |
| L1-paragraph-format-002 | 83 | ✓ | The code is mostly correct but loses points for not using the best practice approach of applying a built-in style for font size and alignment, instead manually setting these properties. |
| L1-list-bullet-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-list-multilevel-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the API correctness and best practices criteria. |
| L1-style-apply-001 | 60 | ✓ | The code uses a real Word Office.js API, but it does not fully address the request as it applies the style to the last paragraph instead of the third paragraph, and it lacks the necessary logic to select the correct paragraph. |
| L1-table-create-001 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the `Word.run` method to execute the insertion of the table, and for not fully addressing the request by not handling potential errors or edge cases, and for not using the most optimal approach by directly inserting the table without considering the document context. |
| L1-table-format-001 | 72 | ✓ | The code uses a non-standard method `styleTable` which is not a real Word Office.js API, and also lacks the use of best practices and available helpers, such as `table.styleBuiltIn`, which would be a more appropriate and standard way to apply a table style. |
| L1-find-replace-001 | 80 | ✓ | The code uses a correct API method replaceText, but lacks the context of a Word.run call and does not specify the scope of the replacement, which might lead to issues with execution and completeness. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach, such as using `range.font` properties directly on the search results, and for not handling potential errors that may occur during the execution of the code. |
| L1-find-replace-regex-001 | 83 | ✓ | The code correctly uses the Word Office.js API and should work without runtime errors, but loses points for completeness and approach due to not explicitly handling potential errors and not using more specific wildcard patterns for email addresses. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a correct API method `insertFootnote` but lacks specificity in targeting the current selection and does not fully adhere to best practices, such as ensuring the range is properly set before inserting the footnote. |
| L1-comment-insert-001 | 72 | ✓ | The code uses a correct API to insert a comment, but it lacks specificity in searching for the exact occurrence of "$4.2M" and does not follow best practices for searching and inserting comments in the Word document. |
| L1-track-changes-toggle-001 | 100 ✅ | ✓ | The code perfectly addresses the request to turn on track changes using the correct and existing Word Office.js API method toggleTrackChanges, with no flaws or deductions in any dimension. |
| L1-track-changes-toggle-002 | 83 | ✓ | The code is mostly correct but loses points for not using the Word.run helper to encapsulate the API calls and for not checking if there are actually tracked revisions to accept before calling acceptAllRevisions. |
| L1-toc-generate-001 | 72 | ✓ | The code uses a correct API but loses points for not fully addressing the request with provided document context and not using the best practice helper method, instead directly calling insertTableOfContents. |
| L1-section-break-001 | 72 | ✓ | The code partially addresses the request but lacks best practices and uses non-existent APIs like setPageOrientation, which is not a real Word Office.js API. |
| L1-header-footer-001 | 67 | ✓ | The code uses a non-existent `setHeader` method, which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-header-footer-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice. |
| L1-columns-001 | 80 | ✓ | The code is mostly correct but loses points for not fully addressing the request of switching the "rest of the document" into a two-column layout, as it only inserts a section break and columns without considering the existing content. |
| L1-image-insert-001 | 72 | ✓ | The code uses a correct API method `insertImage` but lacks consideration for the cursor position at paragraph 3 and does not utilize best practices for image insertion, such as specifying the exact location or handling potential errors. |
| L1-image-insert-002 | 83 | ✓ | The code correctly uses the insertImage API and addresses the user's request, but loses points for not using a more robust approach, such as checking if an image with the specified characteristics already exists in the document or handling potential errors during the insertion process. |
| L1-watermark-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "insertWatermark" which is not a real Word Office.js API, and also does not address the request to add the watermark to every page, resulting in a complete failure across all dimensions. |
| L1-content-control-001 | 83 | ✓ | The code correctly uses the Office.js API to insert a content control, but loses points for not handling potential errors and not using the most robust way to ensure the content control is inserted at the correct location, and also for not fully following best practices in terms of error handling and code organization. |
| L1-content-control-002 | 80 | ✓ | The code correctly uses the insertContentControl API and addresses the request, but loses points for not considering the cursor position at the end of the document, which might require additional code to ensure the content control is inserted at the correct location. |
| L1-mail-merge-001 | 72 | ✓ | The code uses a correct mail merge approach but lacks proper Word Office.js API usage, such as utilizing the `Word.run` method to execute the mail merge, and instead relies on a hallucinated `mailMergeReplace` function. |
| L1-template-apply-001 | 72 | ✓ | The code uses a correct API method applyTemplate but lacks best practices and available helpers, and also assumes the existence of a "resume" template without ensuring its availability. |
| L1-template-apply-002 | 65 | ✓ | The code correctly the requested points for API correctness due to the use of a potentially non-existent `apply fields, but losesTemplate` method, and for due to not approach using the `Word to ensure asynchronous.run` context execution. |
| L1-document-generate-001 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the best practices, such as not checking if the table of contents is already inserted before calling insertTableOfContents, and not handling potential errors that may occur during the execution of the code. |
| L1-theme-apply-001 | 10 ⚠️ | ✓ | The code heavily hallucinates methods, using non-existent `designTheme` and `applyTheme` functions, which are not part of the real Word Office.js APIs. |
| L1-theme-apply-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not used correctly as it should be applied to a specific object like the document body. |
| L1-theme-apply-003 | 35 ⚠️ | ✓ | The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a severe penalty for API correctness. |
| L1-citation-bibliography-001 | 20 ⚠️ | ✓ | The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would_it_work dimensions. |
| L1-equation-001 | 83 | ✓ | The code is mostly correct but loses points for not using the Word.run method to execute the insertion of the equation, and for not checking if the cursor is at the end of the paragraph before inserting the equation, which could lead to incorrect placement. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the user's request, but loses points for not using a more robust method to handle equation insertion, such as checking the current selection or range, and for not following best practices in terms of error handling and code organization. |
| L1-writing-coach-001 | 0 ⚠️ | ✗ | A:  |
| L1-read-query-001 | 40 ⚠️ | ✓ | The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `countWords()` method, which heavily penalizes the API correctness score. |
| L1-read-query-002 | 57 ⚠️ | ✓ | The code uses a hallucinated method "getReadability()" which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-read-query-003 | 0 ⚠️ | ✓ | The code hallucinates a non-existent `listHeadings()` method, which is not a real Word Office.js API, and does not provide any actual implementation to list the headings. |
| L1-margins-orientation-001 | 85 | ✓ | The code correctly sets page margins using the Word Office.js API, but lacks best practices and does not utilize available helpers, and also assumes the context is already loaded without explicitly showing the Word.run() method to ensure the context is properly initialized. |
| L1-margins-orientation-002 | 80 | ✓ | The code correctly uses the Office.js API and addresses the user's request, but lacks best practices and does not utilize available helpers, and also assumes the existence of a setPageOrientation method which is not a standard Word Office.js API. |
| L1-web-search-needed-001 | 67 | ✓ | The AI response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks clarity on how this information will be used to add a section to the Word document and does not provide any workarounds or suggestions. |
| L1-web-search-needed-002 | 65 | ✓ | The response accurately and clearly explains the Pythagorean theorem, but loses points for accuracy due to the unnecessary and incomplete code snippet, and lacks any bonus information or workarounds. |
| L1-refusal-001 | 83 | ✓ | The response is mostly accurate and clear, but loses points for not providing more detailed workaround steps or alternatives, and for a slight lack of clarity in the explanation of the limitation. |
| L1-refusal-002 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not explicitly stating that the workaround requires user interaction, which slightly reduces clarity and bonus scores. |
| L1-refusal-003 | 72 | ✓ | The response is mostly clear and relevant, but loses accuracy points for not explicitly stating that Word Office.js cannot create Excel-style charts with dynamic data, and bonus points for not providing a more detailed workaround or alternative solution. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions. |
| L1-paragraph-spacing-001 | 72 | ✓ | The code uses a correct API but lacks best practices, such as using a loop based on the actual number of paragraphs in the document instead of hardcoding the number, and applying a style without checking if it's necessary, also it doesn't handle potential errors and doesn't use the most efficient way to achieve the result. |
| L1-edge-case-001 | 40 ⚠️ | ✓ | The code fails to attempt the deletion of the paragraph or provide a helpful guard against out-of-range paragraph numbers, instead immediately throwing an error without utilizing Word Office.js APIs or best practices. |
| L1-form-field-001 | 83 | ✓ | The code correctly uses the Word Office.js API to insert a dropdown content control, but loses points for not fully addressing potential edge cases and not using the most optimal approach, such as handling potential errors or using more specific methods. |
| L1-text-insert-001 | 38 ⚠️ | ✓ | The code is incomplete and incorrect, as it only adds a new paragraph without inserting the requested text, and uses a non-existent `addParagraph` method instead of the correct `paragraph.insertText` API. |
| L1-text-edit-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| gen-L2-form-field-001 | 72 | ✓ | The code uses a correct API to insert content controls, but it does not fully address the request as it inserts a date picker with the title "Sign Date" instead of placing it below the checkbox, and it lacks best practices in terms of positioning and layout. |
| gen-L2-form-field-002 | 65 | ✗ | The code uses a non-existent `insertContentControl` method, which is not a real Word Office.js API, and also fails to properly replace the existing plain text "Name: _______" with the rich text content control, resulting in incomplete and incorrect implementation. |
| gen-L2-form-field-003 | 67 | ✗ | The code uses Office.js APIs and accurately addresses the request, but loses using the most efficient approach points for not, such as directly accessing inserted group content control instead the newly of relyingdoc.content on `Controls.getItem(0) |
| gen-L2-form-field-004 | 83 | ✗ | The code is mostly correct but loses points for not explicitly using the Word.run method to execute the content control and table insertion, and for not fully addressing the request of adding a repeating section content control, instead inserting a regular content control. |

## ⚠️ Needs attention

**[L1-list-bullet-001]** score=20 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-list-multilevel-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the API correctness and best practices criteria.

**[L1-header-footer-002]** score=20 — The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice.

**[L1-watermark-insert-001]** score=0 — The code uses a hallucinated method "insertWatermark" which is not a real Word Office.js API, and also does not address the request to add the watermark to every page, resulting in a complete failure across all dimensions.

**[L1-theme-apply-001]** score=10 — The code heavily hallucinates methods, using non-existent `designTheme` and `applyTheme` functions, which are not part of the real Word Office.js APIs.

**[L1-theme-apply-002]** score=20 — The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not used correctly as it should be applied to a specific object like the document body.

**[L1-theme-apply-003]** score=35 — The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a severe penalty for API correctness.

**[L1-citation-bibliography-001]** score=20 — The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would_it_work dimensions.

**[L1-writing-coach-001]** score=0 — A: 
- missing patterns: `coach`

**[L1-read-query-001]** score=40 — The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `countWords()` method, which heavily penalizes the API correctness score.

**[L1-read-query-002]** score=57 — The code uses a hallucinated method "getReadability()" which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions.

**[L1-read-query-003]** score=0 — The code hallucinates a non-existent `listHeadings()` method, which is not a real Word Office.js API, and does not provide any actual implementation to list the headings.

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions.

**[L1-edge-case-001]** score=40 — The code fails to attempt the deletion of the paragraph or provide a helpful guard against out-of-range paragraph numbers, instead immediately throwing an error without utilizing Word Office.js APIs or best practices.

**[L1-text-insert-001]** score=38 — The code is incomplete and incorrect, as it only adds a new paragraph without inserting the requested text, and uses a non-existent `addParagraph` method instead of the correct `paragraph.insertText` API.

**[L1-text-edit-001]** score=20 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.
