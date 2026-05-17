# Word AI — Eval Results
**Last run:** 2026-05-17 09:51:37  
**Overall: 55.4/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 9.0/100 | — | 2 |
| paragraph-format | 77.5/100 | — | 2 |
| list-bullet | 20.0/100 | — | 1 |
| list-multilevel | 72.0/100 | — | 1 |
| style-apply | 80.0/100 | — | 1 |
| table-create | 83.0/100 | — | 1 |
| table-format | 0.0/100 | — | 1 |
| find-replace | 81.0/100 | — | 2 |
| find-replace-regex | 83.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 90.0/100 | — | 2 |
| toc-generate | 83.0/100 | — | 1 |
| section-break | 80.0/100 | — | 1 |
| header-footer | 36.0/100 | — | 2 |
| columns | 72.0/100 | — | 1 |
| image-insert | 46.0/100 | — | 2 |
| watermark-insert | 0.0/100 | — | 1 |
| content-control | 81.5/100 | — | 2 |
| mail-merge | 83.0/100 | — | 1 |
| template-apply | 51.5/100 | — | 2 |
| document-generate | 83.0/100 | — | 1 |
| theme-apply | 34.0/100 | — | 3 |
| citation-bibliography | 20.0/100 | — | 1 |
| equation | 83.0/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 15.0/100 | — | 3 |
| margins-orientation | 52.5/100 | — | 2 |
| web-search-needed | 71.5/100 | — | 2 |
| refusal | 73.3/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 17.0/100 | — | 1 |
| edge-cases | 0.0/100 | — | 1 |
| form-field | 71.4/100 | — | 5 |
| text-insert | 72.0/100 | — | 1 |
| text-edit | 0.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addHeading" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions. |
| L1-heading-insert-002 | 18 ⚠️ | ✓ | A: 18
B: |
| L1-paragraph-format-001 | 72 | ✓ | The code uses a non-standard method `applyStyle(0, "Heading 1")` which is not a real Word Office.js API, and also applies a heading style instead of just making the text bold and 14pt as requested. |
| L1-paragraph-format-002 | 83 | ✓ | The code correctly uses real Word Office.js APIs and addresses the request, but loses points for not using the best practice approach of applying a built-in style for font size instead of manually setting it, and for not checking if the first paragraph is indeed the title before applying the changes. |
| L1-list-bullet-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-list-multilevel-001 | 72 | ✓ | The code uses a non-standard `addList` method which is not a real Word Office.js API, and also lacks the `Word.run` context which is required for most API calls, but it does attempt to address the user's request and is mostly syntactically correct. |
| L1-style-apply-001 | 80 | ✓ | The code correctly applies the Quote style to the third paragraph, but uses a generic applyStyle method which may not be the most efficient or Word-specific approach, and assumes the paragraph index is 2, which may not always be the case. |
| L1-table-create-001 | 83 | ✓ | The code correctly uses the Office.js API to insert a table, but loses points for not fully addressing the request by only inserting three rows of data instead of four as implied by the user's request for a 4x3 table. |
| L1-table-format-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method `styleTable` which is not a real Word Office.js API, resulting in a complete failure across all dimensions. |
| L1-find-replace-001 | 72 | ✓ | The code uses a correct API method replaceText but lacks the context of a Word.run batch and does not handle potential errors, also it does not follow best practices as it does not utilize the Word.run context to ensure the replacement is done in a single batch. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach, such as using `body.search` with `matchWildcards` instead of relying on the default search behavior, and not handling potential errors that may occur during the execution of the `await context.sync()` calls. |
| L1-find-replace-regex-001 | 83 | ✓ | The code correctly uses the Word Office.js API and addresses the request, but loses points for not handling potential errors and not using the most efficient approach, such as checking if the replacement is necessary before calling replaceText. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a correct API but lacks completeness in addressing the request, as it doesn't specify the range for the footnote insertion, and the approach could be improved by utilizing the `insertFootnote` method with a more precise range, such as the current selection. |
| L1-comment-insert-001 | 72 | ✓ | The code uses a correct API method `insertComment` but lacks specificity in targeting the first occurrence of the text "$4.2M" and does not fully address the request by not ensuring the comment is added to the exact location of the text in the document context. |
| L1-track-changes-toggle-001 | 97 ✅ | ✓ | The code is nearly flawless, using the correct toggleTrackChanges API and following best practices, but loses a few points for not being explicitly wrapped in a Word.run context, which is a minor deviation from optimal approach. |
| L1-track-changes-toggle-002 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the Word.run method to execute the acceptAllRevisions action, and for not demonstrating best practices in using the available helpers, such as toggleTrackChanges. |
| L1-toc-generate-001 | 83 | ✓ | The code correctly uses the insertTableOfContents API and addresses the request, but loses points for not using the most specific and efficient approach, such as specifying the exact heading levels present in the document, and for not handling potential errors or edge cases. |
| L1-section-break-001 | 80 | ✓ | The code is mostly correct but lacks consideration for the current section and paragraph context, and does not fully follow best practices for inserting section breaks and setting page orientation in a Word document using Office.js APIs. |
| L1-header-footer-001 | 0 ⚠️ | ✓ | The code is completely incorrect and does not use any real Word Office.js APIs, instead using hallucinated methods like "setHeader" which does not exist in the Office.js API. |
| L1-header-footer-002 | 72 | ✓ | The code uses a non-existent `addPageNumbers` method, which is not a real Word Office.js API, resulting in significant penalties in API correctness and approach dimensions. |
| L1-columns-001 | 72 | ✓ | The code correctly uses real Word Office.js APIs and would likely execute without runtime errors, but it lacks completeness in addressing the user's request to switch the "rest of the document" to a two-column layout, as it only inserts a section break and columns without considering the existing content. |
| L1-image-insert-001 | 72 | ✓ | The code uses a non-standard `insertImage` method which is not a real Word Office.js API, and also lacks proper error handling and context setup, but it attempts to address the request and has a correct width specification. |
| L1-image-insert-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-watermark-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "insertWatermark" which is not a real Word Office.js API, and does not address the request to add a watermark to every page of a 4-page document. |
| L1-content-control-001 | 83 | ✓ | The code is mostly correct and complete, but loses points for not perfectly matching the requested label and not using the most idiomatic way to insert a content control, and also for not handling potential errors or edge cases. |
| L1-content-control-002 | 80 | ✓ | The code correctly uses the Office.js API to insert a content control, but loses points for not specifying the location of the insertion, which might not always result in the desired placement at the end of the document. |
| L1-mail-merge-001 | 83 | ✓ | The code correctly uses the mailMergeReplace API and's request, but loses addresses the user points for not using Word execute the mail merge and.run to for not handling errors, and also potential for not demonstrating terms of error best practices in handling and code |
| L1-template-apply-001 | 20 ⚠️ | ✓ | A: 20
B: |
| L1-template-apply-002 | 83 | ✓ | The code correctly applies the memo template with the specified fields, but loses points for not using the Word.run method to ensure the context is loaded before applying the template, and for not demonstrating best practices in using the available helpers. |
| L1-document-generate-001 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach, such as potentially using a loop to add the sections instead of repeating the addHeading and addParagraph calls, and also for not handling potential errors that may occur during the execution of the code. |
| L1-theme-apply-001 | 10 ⚠️ | ✓ | The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also lacks any actual implementation to address the user's request to apply the theme to the document. |
| L1-theme-apply-002 | 72 | ✓ | The code uses the correct designTheme and applyTheme APIs, but lacks completeness in addressing the request, as it doesn't specify how the theme will be applied to the existing body content, and the approach could be improved by utilizing more specific helpers for styling the document. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, and also does not use the correct API "designTheme" to modify the theme. |
| L1-citation-bibliography-001 | 20 ⚠️ | ✓ | The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-equation-001 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the user's request, but loses points for not using a more robust approach, such as checking the current selection or range before inserting the equation, and not handling potential errors that may occur during the execution of the code. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the request, but loses points for not using a more robust method to handle equation insertion, such as utilizing the Word equation builder or providing more context for the equation. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selection for clarity, but loses points for not considering the coach panel recommendation as an alternative and for using a somewhat generic phrase "Rewriting the selection for clarity" that could be more descriptive. |
| L1-read-query-001 | 20 ⚠️ | ✓ | The code fails to utilize any real Word Office.js APIs, instead relying on a hallucinated `countWords()` method, which is not a valid API and would result in a runtime error. |
| L1-read-query-002 | 25 ⚠️ | ✓ | The code usesget a real Word Office.js API, heavilyReadability()` which is not penalizing correctness, and also API lacks best such as using `Word practices,.run` and `body` to access document properties content |
| L1-read-query-003 | 0 ⚠️ | ✓ | The code uses a hallucinated method "listHeadings()" which is not a real Word Office.js API, and does not provide any actual implementation to list the headings in the document. |
| L1-margins-orientation-001 | 85 | ✓ | The code correctly sets page margins using the Word Office.js API, but lacks best practices and does not utilize available helpers, and also assumes the context is already loaded without explicitly showing the Word.run call. |
| L1-margins-orientation-002 | 20 ⚠️ | ✓ | A: 20
B: |
| L1-web-search-needed-001 | 67 | ✓ | The response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks clarity on how this will be used to add a section to the Word document and does not provide any workarounds or Word-specific guidance. |
| L1-web-search-needed-002 | 76 | ✓ | The response accurately and clearly explains the Pythagorean theorem, directly addressing the question, but the bonus score is low because it doesn't offer any additional workarounds, settings options, or Word-specific caveats beyond simply inserting the information into the document. |
| L1-refusal-001 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not offering more detailed instructions or alternatives, and for a slightly abrupt tone in the refusal. |
| L1-refusal-002 | 72 | ✓ | The response accurately conveys the limitation of Word Office.js and provides a clear workaround, but loses points for not offering more detailed alternatives or settings options, and for including redundant information in the code snippet. |
| L1-refusal-003 | 65 | ✓ | The response accurately identifies the limitation of Word Office.js in creating Excel-style charts, but loses points for clarity and relevance due to the abrupt transition to inserting generic image, a not fully exploring and forarounds or providing alternative work more detailed information on using Excel create the chart. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions. |
| L1-paragraph-spacing-001 | 17 ⚠️ | ✓ | The code fails to address the user's request to set line spacing to 1.5 on every paragraph, instead attempting to add new paragraphs and replace text, resulting in a completely incorrect solution. |
| L1-edge-case-001 | 0 ⚠️ | ✓ | A: 0
B:
C: |
| L1-form-field-001 | 83 | ✓ | The code correctly uses the Office.js API to insert a dropdown content control, but loses points for not using a more specific best practice approach and potentially not handling the context of the cursor in the body. |
| L1-text-insert-001 | 72 | ✓ | The code uses a correct API method addParagraph but loses points for not ensuring the paragraph is inserted at the end of the document, and for not using best practices such as checking the current selection or range before inserting the paragraph. |
| L1-text-edit-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure across all dimensions. |
| gen-L2-form-field-001 | 72 | ✓ | The code uses correct Office.js APIs and is likely to work, but it does not fully address the request as it inserts a "Sign Date" label for the date picker instead of leaving it blank or using a more relevant label, and it lacks best practices in terms of content control placement and styling. |
| gen-L2-form-field-002 | 47 ⚠️ | ✗ | The code incorrectly inserts two rich text content controls instead of replacing the existing plain text "Name: _______" with a single rich text content control and adding a building block gallery control as requested. |
| gen-L2-form-field-003 | 72 | ✗ | The code is mostly correct but loses points for not properly nesting the dropdown content controls within the group content control, and for not using the best practices such as using a more specific method to insert the content controls at the correct location in the document. |
| gen-L2-form-field-004 | 83 | ✗ | The code is mostly correct and complete, but loses points for not explicitly using the `Word.run` API to execute the actions and for not perfectly aligning with best practices, such as potentially missing error handling and not using more specific built-in helpers for table styling. |

## ⚠️ Needs attention

**[L1-heading-insert-001]** score=0 — The code uses a hallucinated method "addHeading" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions.

**[L1-heading-insert-002]** score=18 — A: 18
B:

**[L1-list-bullet-001]** score=20 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-table-format-001]** score=0 — The code uses a hallucinated method `styleTable` which is not a real Word Office.js API, resulting in a complete failure across all dimensions.

**[L1-header-footer-001]** score=0 — The code is completely incorrect and does not use any real Word Office.js APIs, instead using hallucinated methods like "setHeader" which does not exist in the Office.js API.

**[L1-image-insert-002]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-watermark-insert-001]** score=0 — The code uses a hallucinated method "insertWatermark" which is not a real Word Office.js API, and does not address the request to add a watermark to every page of a 4-page document.

**[L1-template-apply-001]** score=20 — A: 20
B:

**[L1-theme-apply-001]** score=10 — The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also lacks any actual implementation to address the user's request to apply the theme to the document.

**[L1-theme-apply-003]** score=20 — The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, and also does not use the correct API "designTheme" to modify the theme.

**[L1-citation-bibliography-001]** score=20 — The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-read-query-001]** score=20 — The code fails to utilize any real Word Office.js APIs, instead relying on a hallucinated `countWords()` method, which is not a valid API and would result in a runtime error.

**[L1-read-query-002]** score=25 — The code usesget a real Word Office.js API, heavilyReadability()` which is not penalizing correctness, and also API lacks best such as using `Word practices,.run` and `body` to access document properties content

**[L1-read-query-003]** score=0 — The code uses a hallucinated method "listHeadings()" which is not a real Word Office.js API, and does not provide any actual implementation to list the headings in the document.

**[L1-margins-orientation-002]** score=20 — A: 20
B:

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions.

**[L1-paragraph-spacing-001]** score=17 — The code fails to address the user's request to set line spacing to 1.5 on every paragraph, instead attempting to add new paragraphs and replace text, resulting in a completely incorrect solution.

**[L1-edge-case-001]** score=0 — A: 0
B:
C:

**[L1-text-edit-001]** score=0 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure across all dimensions.

**[gen-L2-form-field-002]** score=47 — The code incorrectly inserts two rich text content controls instead of replacing the existing plain text "Name: _______" with a single rich text content control and adding a building block gallery control as requested.
- missing patterns: `RichTextContentControl`
