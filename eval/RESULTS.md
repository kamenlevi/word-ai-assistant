# Word AI — Eval Results
**Last run:** 2026-05-20 18:23:55  
**Overall: 60.0/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 69.5/100 | — | 2 |
| paragraph-format | 87.0/100 | — | 2 |
| list-bullet | 20.0/100 | — | 1 |
| list-multilevel | 0.0/100 | — | 1 |
| style-apply | 80.0/100 | — | 1 |
| table-create | 90.0/100 | — | 1 |
| table-format | 72.0/100 | — | 1 |
| find-replace | 81.0/100 | — | 2 |
| find-replace-regex | 83.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 91.5/100 | — | 2 |
| toc-generate | 72.0/100 | — | 1 |
| section-break | 67.0/100 | — | 1 |
| header-footer | 10.0/100 | — | 2 |
| columns | 72.0/100 | — | 1 |
| image-insert | 50.0/100 | — | 2 |
| watermark-insert | 0.0/100 | — | 1 |
| content-control | 83.0/100 | — | 2 |
| mail-merge | 83.0/100 | — | 1 |
| template-apply | 51.5/100 | — | 2 |
| document-generate | 83.0/100 | — | 1 |
| theme-apply | 15.0/100 | — | 3 |
| citation-bibliography | 20.0/100 | — | 1 |
| equation | 83.0/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 49.0/100 | — | 3 |
| margins-orientation | 82.5/100 | — | 2 |
| web-search-needed | 70.0/100 | — | 2 |
| refusal | 74.0/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 72.0/100 | — | 1 |
| edge-cases | 25.0/100 | — | 1 |
| form-field | 68.8/100 | — | 5 |
| text-insert | 80.0/100 | — | 1 |
| text-edit | 0.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 72 | ✓ | The code uses a non-standard method "addHeading" which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-heading-insert-002 | 67 | ✓ | The code uses non-standard helper methods like addHeading and addParagraph, which are not actual Word Office.js APIs, resulting in significant deductions in API correctness and approach dimensions. |
| L1-paragraph-format-001 | 91 | ✓ | The code is mostly correct and complete, but loses points on approach for not using the available helpers, such as paragraph.styleBuiltIn or range.font properties directly on the result of a body.paragraphs.getFirst() call, instead opting for direct property access. |
| L1-paragraph-format-002 | 83 | ✓ | The code is mostly correct but loses points for not using the available helper methods, such as applyStyle to set the font size, and for not checking if the first paragraph is indeed the title before applying the style. |
| L1-list-bullet-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-list-multilevel-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions. |
| L1-style-apply-001 | 80 | ✓ | The code correctly applies the Quote style to a paragraph using the applyStyle method, but loses points for not fully addressing potential edge cases and not using the most robust best practices, such as error handling and explicit paragraph indexing. |
| L1-table-create-001 | 90 | ✓ | The code accurately uses the Word Office.js API to insert a table, but loses points for not fully addressing the request by not considering the document context and paragraph positioning, and for not using the most optimal approach by directly specifying the table style instead of potentially using a more flexible method. |
| L1-table-format-001 | 72 | ✓ | The code uses real Word Office.js APIs and mostly addresses the request, but lacks error handling, assumes the first table is the target, and does not use the best practice of directly applying a style to the table object. |
| L1-find-replace-001 | 72 | ✓ | The code uses a correct API method replaceText but lacks the context of a Word.run function to execute the replacement, and does not utilize the best practice approach of using the search function with matchWildcards for text replacement. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient or best-practice approach, such as utilizing the `range.font` properties directly on the search results without explicitly loading the font property. |
| L1-find-replace-regex-001 | 83 | ✓ | The code correctly uses the Word Office.js API and addresses the request, but loses points for not handling potential errors and not using the most efficient approach, such as checking if the replacement is necessary before calling replaceText. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a correct API but lacks completeness in addressing the request, as it doesn't specify the range for the footnote insertion, and the approach could be improved by utilizing the `insertFootnote` method with a more precise range, such as the current selection. |
| L1-comment-insert-001 | 72 | ✓ | The code uses a correct API method `insertComment` but lacks specificity in targeting the first occurrence of "$4.2M" and does not utilize Word's built-in search functionality with wildcards or regular expressions as needed for precise text matching. |
| L1-track-changes-toggle-001 | 100 ✅ | ✓ | The code perfectly addresses the request to turn on track changes using the correct and existing Word Office.js API method toggleTrackChanges, and follows best practices. |
| L1-track-changes-toggle-002 | 83 | ✓ | The code is mostly correct, but it lacks a Word.run wrapper and does not handle potential errors, also it does not check if there are actually revisions to accept, which prevents it from being a complete and robust solution. |
| L1-toc-generate-001 | 72 | ✓ | The code uses a correct API but lacks best practices, such as not checking if a Table of Contents already exists, and the approach could be improved by using more specific parameters for the insertTableOfContents method. |
| L1-section-break-001 | 67 | ✓ | The code uses some correct APIs but lacks best practices and does not fully address the request as it doesn't specify the section to switch to landscape, and the approach is not optimal. |
| L1-header-footer-001 | 0 ⚠️ | ✓ | A:  |
| L1-header-footer-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-columns-001 | 72 | ✓ | The code correctly uses real Word Office.js APIs and would likely execute without runtime errors, but it lacks completeness in addressing the user's request to switch the rest of the document into a two-column layout, and does not use best practices or available helpers. |
| L1-image-insert-001 | 80 | ✓ | The code is mostly correct but loses points for not specifying the exact paragraph where the image should be inserted and not using the Word.run method to queue the changes, which is a best practice. |
| L1-image-insert-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a significant penalty for API correctness. |
| L1-watermark-insert-001 | 0 ⚠️ | ✓ | The code is completely incorrect because it uses a non-existent `insertWatermark` method, which is not a part of the Word Office.js API, and does not address the request to add a watermark to every page. |
| L1-content-control-001 | 83 | ✓ | The code correctly uses the Office.js API to insert a content control, but loses points for not fully addressing potential edge cases and not using the most optimal approach, such as handling potential errors or using more specific methods. |
| L1-content-control-002 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the Word.run() method to execute the content control insertion, and for not being wrapped in a best-practice helper function like insertContentControl is, which is available in the Office.js API. |
| L1-mail-merge-001 | 83 | ✓ | The code correctly uses the mailMergeReplace API and addresses the user's request, but loses points for not using the most efficient or best-practice approach, such as handling potential errors or using more specific typing for the replace operation. |
| L1-template-apply-001 | 83 | ✓ | The code correctly applies the resume template with the provided details, but loses points for not using the Word.run method to ensure the context is properly loaded and for not handling potential errors, and also for not fully utilizing the available helpers and best practices. |
| L1-template-apply-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would it work dimensions. |
| L1-document-generate-001 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient or best-practice approaches, such as potentially re-implementing table of contents logic instead of relying on built-in functionality, and not demonstrating a clear understanding of when to use context.sync(). |
| L1-theme-apply-001 | 5 ⚠️ | ✓ | The code uses completely hallucinated and non-existent methods "designTheme" and "applyTheme", which are not part of the Word Office.js API, resulting in a complete failure in API correctness and overall approach. |
| L1-theme-apply-002 | 20 ⚠️ | ✓ | The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a severe penalty for API correctness and would not work at runtime. |
| L1-citation-bibliography-001 | 20 ⚠️ | ✓ | The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-equation-001 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the request, but loses points for not using a more robust approach, such as checking the current selection or range before inserting the equation, and not handling potential errors or exceptions that may occur during execution. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the user's request, but loses points for not using a more robust approach to handle potential errors and for not providing additional context or formatting options, which are available through other Word Office.js APIs. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selection for clarity, but loses points for not considering alternative approaches, such as recommending the coach panel, and for minor clarity issues in the code comments. |
| L1-read-query-001 | 50 ⚠️ | ✓ | The code lacks correctness in using real Word Office.js APIs, as it hallucinates a non-existent `countWords()` method instead of utilizing the actual API, such as `Word.run` and `body.getText()`, to count the words in the document. |
| L1-read-query-002 | 57 ⚠️ | ✓ | The code uses a hallucinated method `getReadability()` which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-read-query-003 | 40 ⚠️ | ✓ | The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `listHeadings()` method, which heavily penalizes the API correctness score. |
| L1-margins-orientation-001 | 85 | ✓ | The code correctly sets page margins using the Office.js API, but lacks best practices and does not utilize available helpers, and also assumes the context is already loaded without explicitly showing the Word.run() method to ensure the context is properly initialized. |
| L1-margins-orientation-002 | 80 | ✓ | The code correctly uses the Office.js API and addresses the request, but lacks best practices and does not utilize available helpers, and also assumes the existence of a setPageOrientation method which is not a standard Word Office.js API. |
| L1-web-search-needed-001 | 67 | ✓ | The AI response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks clarity on how this information will be used to add a section to the Word document and does not provide any workarounds or Word-specific caveats. |
| L1-web-search-needed-002 | 73 | ✓ | The response accurately and clearly explains the Pythagorean theorem, but the code provided is unnecessary and doesn't add significant value, resulting in a lower bonus score. |
| L1-refusal-001 | 83 | ✓ | The response is mostly accurate and clear, but loses points for not providing more detailed workaround steps or alternative solutions, and for a slight lack of clarity in explaining the limitation of the Word Office.js API. |
| L1-refusal-002 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear VBA workaround, but loses points for relevance due to the introduction of an alternative approach that doesn't directly address the Office.js question, and for bonus due to not exploring Office.js-specific alternatives or caveats. |
| L1-refusal-003 | 67 | ✓ | The response accurately identifies the limitation of Word Office.js in creating charts, but lacks a concrete workaround or alternative solution, resulting in a low bonus score. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-paragraph-spacing-001 | 72 | ✓ | The code correctly uses Word Office.js APIs and addresses the request, but has issues with approach, using a manual loop instead of available helpers, and lacks error handling, which may cause runtime errors. |
| L1-edge-case-001 | 25 ⚠️ | ✓ | The code fails to use any real Word Office.js APIs and instead directly throws an error, indicating a lack of attempt to interact with the document or handle the out-of-range paragraph index in a meaningful way. |
| L1-form-field-001 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the request, but loses points for not using a more descriptive title and for not handling potential errors, and also for not fully following best practices in terms of content control configuration. |
| L1-text-insert-001 | 80 | ✓ | The code uses a real Word Office.js API (addParagraph) and is likely to execute without runtime errors, but loses points for not fully addressing the request (inserting at the end of the document, not necessarily after the last paragraph) and not using the most optimal approach (e.g., using range.insertText to insert text at a specific location). |
| L1-text-edit-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| gen-L2-form-field-001 | 72 | ✓ | The code uses correct Word Office.js APIs and is mostly complete, but loses points for not fully addressing the request to place the date picker content control below the checkbox, and for not using best practices such as inserting a new paragraph for the date picker. |
| gen-L2-form-field-002 | 57 ⚠️ | ✗ | The code uses a hallucinated method `insertContentControl` which is not a real Word Office.js API, and also fails to replace the existing plain text 'Name: _______' as requested, instead inserting a new content control. |
| gen-L2-form-field-003 | 60 | ✗ | The code fails to insert a group content control containing the two dropdown content controls as requested, instead inserting separate content controls, and also uses an undefined variable "color" in the tag property of the first dropdown. |
| gen-L2-form-field-004 | 72 | ✗ | The code uses correct Word Office.js APIs and is mostly complete, but it lacks the specific implementation of a repeating section content control and instead inserts a regular content control, and also does not fully follow best practices for inserting tables and content controls. |

## ⚠️ Needs attention

**[L1-list-bullet-001]** score=20 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-list-multilevel-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions.

**[L1-header-footer-001]** score=0 — A: 

**[L1-header-footer-002]** score=20 — The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-image-insert-002]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a significant penalty for API correctness.

**[L1-watermark-insert-001]** score=0 — The code is completely incorrect because it uses a non-existent `insertWatermark` method, which is not a part of the Word Office.js API, and does not address the request to add a watermark to every page.

**[L1-template-apply-002]** score=20 — The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would it work dimensions.

**[L1-theme-apply-001]** score=5 — The code uses completely hallucinated and non-existent methods "designTheme" and "applyTheme", which are not part of the Word Office.js API, resulting in a complete failure in API correctness and overall approach.

**[L1-theme-apply-002]** score=20 — The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-theme-apply-003]** score=20 — The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a severe penalty for API correctness and would not work at runtime.

**[L1-citation-bibliography-001]** score=20 — The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-read-query-001]** score=50 — The code lacks correctness in using real Word Office.js APIs, as it hallucinates a non-existent `countWords()` method instead of utilizing the actual API, such as `Word.run` and `body.getText()`, to count the words in the document.

**[L1-read-query-002]** score=57 — The code uses a hallucinated method `getReadability()` which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions.

**[L1-read-query-003]** score=40 — The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `listHeadings()` method, which heavily penalizes the API correctness score.

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-edge-case-001]** score=25 — The code fails to use any real Word Office.js APIs and instead directly throws an error, indicating a lack of attempt to interact with the document or handle the out-of-range paragraph index in a meaningful way.

**[L1-text-edit-001]** score=0 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[gen-L2-form-field-002]** score=57 — The code uses a hallucinated method `insertContentControl` which is not a real Word Office.js API, and also fails to replace the existing plain text 'Name: _______' as requested, instead inserting a new content control.
- missing patterns: `RichTextContentControl`
