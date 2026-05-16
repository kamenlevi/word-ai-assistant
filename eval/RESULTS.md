# Word AI — Eval Results
**Last run:** 2026-05-16 17:04:16  
**Overall: 61.2/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 78.5/100 | — | 2 |
| paragraph-format | 84.0/100 | — | 2 |
| list-bullet | 25.0/100 | — | 1 |
| list-multilevel | 0.0/100 | — | 1 |
| style-apply | 80.0/100 | — | 1 |
| table-create | 83.0/100 | — | 1 |
| table-format | 72.0/100 | — | 1 |
| find-replace | 85.0/100 | — | 2 |
| find-replace-regex | 83.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 89.0/100 | — | 2 |
| toc-generate | 83.0/100 | — | 1 |
| section-break | 67.0/100 | — | 1 |
| header-footer | 50.0/100 | — | 2 |
| columns | 83.0/100 | — | 1 |
| image-insert | 72.0/100 | — | 2 |
| watermark-insert | 0.0/100 | — | 1 |
| content-control | 83.0/100 | — | 2 |
| mail-merge | 55.0/100 | — | 1 |
| template-apply | 46.0/100 | — | 2 |
| document-generate | 85.0/100 | — | 1 |
| theme-apply | 15.0/100 | — | 3 |
| citation-bibliography | 40.0/100 | — | 1 |
| equation | 79.0/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 20.0/100 | — | 3 |
| margins-orientation | 82.5/100 | — | 2 |
| web-search-needed | 71.5/100 | — | 2 |
| refusal | 57.7/100 | — | 3 |
| quote-insert | 40.0/100 | — | 1 |
| paragraph-spacing | 72.0/100 | — | 1 |
| edge-cases | 35.0/100 | — | 1 |
| form-field | 73.2/100 | — | 5 |
| text-insert | 72.0/100 | — | 1 |
| text-edit | 0.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 85 | ✓ | The code is mostly correct, but it uses a custom `addHeading` function which is not a built-in Word Office.js API, and its implementation is not provided, which raises concerns about its correctness and potential for errors. |
| L1-heading-insert-002 | 72 | ✓ | The code uses correct APIs and mostly addresses the request, but lacks best practices, such as using Word.run to ensure the context is loaded before inserting content, and does not handle potential errors or edge cases. |
| L1-paragraph-format-001 | 88 | ✓ | The code is mostly correct and complete, but loses points for not using the best practice of loading properties in a single `load` call and for not checking if the paragraph exists before accessing its properties. |
| L1-paragraph-format-002 | 80 | ✓ | The code correctly uses real Word Office.js APIs and addresses the request, but loses points for not using the `applyStyle` method more effectively and not checking if the first paragraph is indeed the title before applying the style. |
| L1-list-bullet-001 | 25 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice. |
| L1-list-multilevel-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the API correctness and best practices criteria. |
| L1-style-apply-001 | 80 | ✓ | The code uses a real Word Office.js API, applies the style correctly, and would likely execute without runtime errors, but loses points for not fully addressing potential edge cases and not using the most idiomatic or efficient approach, such as directly using the `paragraph.styleBuiltIn` property. |
| L1-table-create-001 | 83 | ✓ | The code is mostly correct and complete, but loses points for not fully addressing the request by only including three rows of data instead of four as specified, and not using the most optimal approach by directly using the `body.insertTable` method with the correct parameters. |
| L1-table-format-001 | 72 | ✓ | The code uses a non-existent `styleTable` method, which is not a real Word Office.js API, resulting in significant penalties in API correctness and approach dimensions. |
| L1-find-replace-001 | 80 | ✓ | The code uses a correct API and mostly addresses the request, but lacks error handling and context setup, and does not fully follow best practices by not using Word.run to encapsulate the operation. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient or best practice approach, such as utilizing the `range.font` properties directly without loading them explicitly, and not handling potential errors or edge cases. |
| L1-find-replace-regex-001 | 83 | ✓ | The code is mostly correct but loses points for not using the Word.run method to ensure the context is loaded and synced before executing the replaceText operation, and for not fully addressing potential edge cases in email address replacement. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a correct API method to insert a footnote, but lacks context about the current selection and does not fully address best practices for implementing the solution. |
| L1-comment-insert-001 | 72 | ✓ | The code uses a correct API method `insertComment` but lacks specificity in targeting the first occurrence of "$4.2M" and does not utilize Word's built-in search functionality with wildcards or regular expressions as needed for precise text matching. |
| L1-track-changes-toggle-001 | 95 ✅ | ✓ | The code perfectly addresses the request using the correct API, but loses points for approach because it manually calls `toggleTrackChanges` instead of using a potentially more robust helper function, if available, although in this case, `toggleTrackChanges` is a correct and straightforward method to achieve the desired outcome. |
| L1-track-changes-toggle-002 | 83 | ✓ | The code is mostly correct but loses points for not using the Word.run method to wrap the API calls and for not checking if the document has tracked revisions before accepting them, which could lead to a minor loss of completeness and adherence to best practices. |
| L1-toc-generate-001 | 83 | ✓ | The code uses a correct API and is mostly complete, but loses points for not fully addressing the request to insert the Table of Contents at the top, and for not using the best practice of specifying the location of the Table of Contents. |
| L1-section-break-001 | 67 | ✓ | The code uses some correct APIs but lacks best practices and does not fully address the request as it does not specify the section to switch to landscape, only switching the page orientation. |
| L1-header-footer-001 | 80 | ✓ | The code correctly uses the Word Office.js API and addresses the request, but lacks consideration for best practices, such as using a more specific method for setting the header, and does not account for potential errors or edge cases, like handling multiple sections or existing headers. |
| L1-header-footer-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work at runtime. |
| L1-columns-001 | 83 | ✓ | The code is mostly correct but loses points for not fully addressing potential edge cases and not using the most efficient or best practice approaches, such as checking the current section's layout before inserting a new section break and columns. |
| L1-image-insert-001 | 72 | ✓ | The code uses the correct insertImage API, but loses points for not specifying the exact paragraph where the image should be inserted, and for not using best practices, such as using the Word.run method to queue the changes and not checking if the context is loaded before syncing. |
| L1-image-insert-002 | 72 | ✓ | The code uses a non-standard `insertImage` method which is not a real Word Office.js API, and also lacks error handling and context setup, but it attempts to address the user's request and has a correct sync order. |
| L1-watermark-insert-001 | 0 ⚠️ | ✓ | The code does not use any real Word Office.js APIs, instead calling a non-existent `insertWatermark` method, which is a significant flaw in API correctness, completeness, and approach. |
| L1-content-control-001 | 83 | ✓ | The code correctly uses the Office.js API to insert a content control, but loses points for not fully addressing the request by not specifying the exact location of the insertion, and for not using the most idiomatic approach by not utilizing the provided helpers. |
| L1-content-control-002 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the `Word.run` method to execute the content control insertion, and for not checking if the `context` is loaded before calling `context.sync()`, which could lead to runtime errors. |
| L1-mail-merge-001 | 55 ⚠️ | ✓ | The code uses a hallucinated method "mailMergeReplace" which is not a real Word Office.js API, resulting in a significant penalty for API correctness. |
| L1-template-apply-001 | 72 | ✓ | The code uses a correct API method applyTemplate but loses points for not using Word.run to execute the template application and for not checking if the template exists, and also for not following best practices in terms of error handling and variable scope. |
| L1-template-apply-002 | 20 ⚠️ | ✓ | A: 20
B: |
| L1-document-generate-001 | 85 | ✓ | The code is mostly correct and complete, but loses points for not using the best practices, such as not checking if the headings and paragraphs already exist before inserting them, and not handling potential errors that may occur during the execution of the code. |
| L1-theme-apply-001 | 5 ⚠️ | ✓ | The code heavily hallucinates methods, using non-existent "designTheme" and "applyTheme" functions, which are not part of the real Word Office.js APIs. |
| L1-theme-apply-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not a valid method, resulting in a complete loss of points for API correctness and would it work dimensions. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-citation-bibliography-001 | 40 ⚠️ | ✓ | The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a significant penalty for API correctness. |
| L1-equation-001 | 75 | ✓ | The code correctly uses the insertEquation API and addresses the request, but lacks context about the current selection or cursor position, and does not utilize the Word.run method to ensure the code is executed within a context that can handle asynchronous operations, which is a best practice. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the user's request, but loses points for not checking the cursor position before inserting the equation and not using a more robust method to handle potential errors, and also for not fully following best practices in terms of error handling and code organization. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selected text for clarity, but loses points for not considering the coach panel as an alternative and for using a somewhat generic phrase "Rewriting the current selection for clarity" that doesn't add much value. |
| L1-read-query-001 | 60 | ✓ | The code uses a hallucinated method "countWords()" which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-read-query-002 | 0 ⚠️ | ✓ | A: 0
B: |
| L1-read-query-003 | 0 ⚠️ | ✓ | The code uses a hallucinated method "listHeadings()" which is not a real Word Office.js API, and does not provide any actual implementation to list the headings in the document. |
| L1-margins-orientation-001 | 85 | ✓ | The code correctly sets page margins using the Word Office.js API, but lacks best practices and does not utilize available helpers, and also assumes the context is already loaded without explicitly showing the Word.run() method to ensure the context is properly synchronized. |
| L1-margins-orientation-002 | 80 | ✓ | The code is mostly correct but lacks a clear context and document object reference, and does not use the Word.run method to ensure the API is properly loaded and synchronized before making changes. |
| L1-web-search-needed-001 | 67 | ✓ | The AI response correctly identifies the need for a web search to get the current Apple stock price, but lacks clarity on how this information will be used to add a section to the document and does not provide any workarounds or Word-specific caveats. |
| L1-web-search-needed-002 | 76 | ✓ | The response accurately and clearly explains the Pythagorean theorem, directly addressing the question, but loses points for not providing any meaningful bonus information or workarounds specific to Microsoft Word or Office.js. |
| L1-refusal-001 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not offering more detailed guidance on the "Save As" feature or alternative solutions, resulting in a moderate bonus score. |
| L1-refusal-002 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but lacks a more detailed explanation or alternative solutions, resulting in moderate scores across the dimensions. |
| L1-refusal-003 | 18 ⚠️ | ✓ | A: 18
B: |
| L1-quote-insert-001 | 40 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a severe penalty for API correctness. |
| L1-paragraph-spacing-001 | 72 | ✓ | The code uses correct Word Office.js APIs and addresses the request, but it lacks best practices, has potential runtime errors due to the applyStyle method call, and does not account for dynamic paragraph counts. |
| L1-edge-case-001 | 35 ⚠️ | ✓ | The code fails to use the Word Office.js API correctly, instead directly throwing an error without attempting to access or manipulate the document's paragraphs using the available APIs. |
| L1-form-field-001 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the request, but loses points for not handling potential errors and not using the most descriptive variable names, and also for not fully following best practices in terms of code structure and readability. |
| L1-text-insert-001 | 72 | ✓ | The code uses a correct API method addParagraph, but loses points for not ensuring the paragraph is inserted at the end of the document, and for not using best practices such as checking the current selection or range before inserting the paragraph. |
| L1-text-edit-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| gen-L2-form-field-001 | 72 | ✓ | The code uses correct Office.js APIs and is likely to work, but it does not fully address the request as it inserts a "Sign Date" label for the date picker instead of no label, and lacks best practices such as using a more specific range for insertion and handling potential errors. |
| gen-L2-form-field-002 | 50 ⚠️ | ✗ | The code incorrectly inserts two rich text content controls, one for "Name" and one for "Address", instead of replacing the existing "Name: _______" plain text with a single rich text content control and adding a building block gallery control as requested. |
| gen-L2-form-field-003 | 89 | ✗ | The code is mostly correct and complete, but loses points for not using the `Word.run` method to ensure context synchronization and for not handling potential errors, and also for not using the most efficient approach by directly using the `insertContentControl` method on the `body` object instead of getting the first content control and then inserting into it. |
| gen-L2-form-field-004 | 72 | ✗ | The code partially addresses the request but fails to fully implement a repeating section content control, instead inserting a rich text content control, and also lacks best practices in using available helpers. |

## ⚠️ Needs attention

**[L1-list-bullet-001]** score=25 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice.

**[L1-list-multilevel-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the API correctness and best practices criteria.

**[L1-header-footer-002]** score=20 — The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work at runtime.

**[L1-watermark-insert-001]** score=0 — The code does not use any real Word Office.js APIs, instead calling a non-existent `insertWatermark` method, which is a significant flaw in API correctness, completeness, and approach.

**[L1-mail-merge-001]** score=55 — The code uses a hallucinated method "mailMergeReplace" which is not a real Word Office.js API, resulting in a significant penalty for API correctness.

**[L1-template-apply-002]** score=20 — A: 20
B:

**[L1-theme-apply-001]** score=5 — The code heavily hallucinates methods, using non-existent "designTheme" and "applyTheme" functions, which are not part of the real Word Office.js APIs.

**[L1-theme-apply-002]** score=20 — The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not a valid method, resulting in a complete loss of points for API correctness and would it work dimensions.

**[L1-theme-apply-003]** score=20 — The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-citation-bibliography-001]** score=40 — The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a significant penalty for API correctness.

**[L1-read-query-002]** score=0 — A: 0
B:

**[L1-read-query-003]** score=0 — The code uses a hallucinated method "listHeadings()" which is not a real Word Office.js API, and does not provide any actual implementation to list the headings in the document.

**[L1-refusal-003]** score=18 — A: 18
B:

**[L1-quote-insert-001]** score=40 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a severe penalty for API correctness.

**[L1-edge-case-001]** score=35 — The code fails to use the Word Office.js API correctly, instead directly throwing an error without attempting to access or manipulate the document's paragraphs using the available APIs.

**[L1-text-edit-001]** score=0 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[gen-L2-form-field-002]** score=50 — The code incorrectly inserts two rich text content controls, one for "Name" and one for "Address", instead of replacing the existing "Name: _______" plain text with a single rich text content control and adding a building block gallery control as requested.
- missing patterns: `RichTextContentControl`
