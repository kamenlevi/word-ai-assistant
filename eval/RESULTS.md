# Word AI — Eval Results
**Last run:** 2026-05-18 11:51:14  
**Overall: 57.0/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 69.5/100 | — | 2 |
| paragraph-format | 41.5/100 | — | 2 |
| list-bullet | 25.0/100 | — | 1 |
| list-multilevel | 0.0/100 | — | 1 |
| style-apply | 85.0/100 | — | 1 |
| table-create | 90.0/100 | — | 1 |
| table-format | 72.0/100 | — | 1 |
| find-replace | 85.0/100 | — | 2 |
| find-replace-regex | 85.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 92.5/100 | — | 2 |
| toc-generate | 67.0/100 | — | 1 |
| section-break | 80.0/100 | — | 1 |
| header-footer | 51.5/100 | — | 2 |
| columns | 72.0/100 | — | 1 |
| image-insert | 72.0/100 | — | 2 |
| watermark-insert | 0.0/100 | — | 1 |
| content-control | 81.5/100 | — | 2 |
| mail-merge | 20.0/100 | — | 1 |
| template-apply | 20.0/100 | — | 2 |
| document-generate | 85.0/100 | — | 1 |
| theme-apply | 18.3/100 | — | 3 |
| citation-bibliography | 45.0/100 | — | 1 |
| equation | 41.5/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 40.0/100 | — | 3 |
| margins-orientation | 82.5/100 | — | 2 |
| web-search-needed | 71.5/100 | — | 2 |
| refusal | 79.3/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 72.0/100 | — | 1 |
| edge-cases | 25.0/100 | — | 1 |
| form-field | 73.6/100 | — | 5 |
| text-insert | 0.0/100 | — | 1 |
| text-edit | 0.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 72 | ✓ | The code uses a non-standard method `addHeading` which is not a real Word Office.js API, but it does address the request and would likely work if the method was correctly implemented. |
| L1-heading-insert-002 | 67 | ✓ | The code uses hallucinated methods addHeading and addParagraph, which are not real Word Office.js APIs, resulting in significant deductions in API correctness and approach dimensions. |
| L1-paragraph-format-001 | 0 ⚠️ | ✗ | The code uses a hallucinated method `addParagraph` which does not exist in the Word Office.js API, and instead of modifying the existing paragraph, it attempts to add a new paragraph with the same text, resulting in a completely incorrect implementation. |
| L1-paragraph-format-002 | 83 | ✓ | The code is mostly correct but loses points for not using the available helper methods, such as applyStyle to set the font size, and for not checking if the first paragraph is indeed the title before applying the style. |
| L1-list-bullet-001 | 25 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-list-multilevel-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions. |
| L1-style-apply-001 | 85 | ✓ | The code correctly applies the Quote style to the third paragraph, but loses points for API correctness and approach due to the use of a generic applyStyle method instead of the more specific paragraph.styleBuiltIn property. |
| L1-table-create-001 | 90 | ✓ | The code accurately uses the Word Office.js API to insert a table, but loses points for not fully addressing the request by not considering the document context and paragraph where the table should be inserted, and for not using the most optimal approach by directly using the insertTable method without checking the current selection or paragraph. |
| L1-table-format-001 | 72 | ✓ | The code uses a non-standard `styleTable` method which is not a real Word Office.js API, and also lacks the use of best practices and available helpers, such as `table.style = "Grid Table 4 - Accent 1"` or `applyStyle` method. |
| L1-find-replace-001 | 80 | ✓ | The code uses a correct API method replaceText, but lacks the surrounding Word.run context and may not fully handle all edge cases, such as multiple occurrences or case sensitivity, and does not follow best practices for error handling and loading the document. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach, such as using `range.font` helpers, and for not handling potential errors that may occur during the execution of the `await context.sync()` calls. |
| L1-find-replace-regex-001 | 85 | ✓ | The code uses the correct Word Office.js API and is likely to work without runtime errors, but loses points for not fully addressing potential edge cases and not using the most optimal approach, such as handling multiple matches or using more specific wildcard patterns. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a correct API method to insert a footnote, but lacks specificity in targeting the current selection and does not follow best practices by directly calling the insertFootnote method without utilizing the provided helpers or ensuring the correct context. |
| L1-comment-insert-001 | 72 | ✓ | The code uses a non-existent insertComment method, which is not a real Word Office.js API, and also lacks proper range or context specification for the comment insertion. |
| L1-track-changes-toggle-001 | 100 ✅ | ✓ | The code perfectly addresses the request to turn on track changes using the correct and available Word Office.js API, specifically the toggleTrackChanges method, with no flaws or unnecessary re-implementations. |
| L1-track-changes-toggle-002 | 85 | ✓ | The code is mostly correct but loses points for not using the Word.run method to encapsulate the API calls, which is a best practice, and also for not checking if there are any tracked revisions before accepting them, which could lead to unnecessary API calls. |
| L1-toc-generate-001 | 67 | ✓ | The code uses a hallucinated method insertTableOfContents, which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-section-break-001 | 80 | ✓ | The code is mostly correct but lacks consideration for the current section and paragraph context, and does not fully follow best practices by directly using low-level APIs like insertSectionBreak and setPageOrientation without utilizing higher-level helpers. |
| L1-header-footer-001 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the best practice approach of directly manipulating the header object, instead relying on a generic setHeader method that may not be optimal. |
| L1-header-footer-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-columns-001 | 72 | ✓ | The code correctly uses real Word Office.js APIs and would likely execute without runtime errors, but it lacks completeness in addressing the user's request to switch the "rest of the document" into a two-column layout, as it only inserts a new section break and applies the two-column layout to the new section, without considering the existing content. |
| L1-image-insert-001 | 72 | ✓ | The code uses a non-standard `insertImage` method which is not a real Word Office.js API, and also lacks proper error handling and context setup, but it attempts to address the user's request and has a correct method signature. |
| L1-image-insert-002 | 72 | ✓ | The code uses a correct API method `insertImage` but lacks context about where in the document the image should be inserted and does not fully follow best practices for image insertion in a Word document. |
| L1-watermark-insert-001 | 0 ⚠️ | ✓ | The code is completely incorrect as it uses a non-existent `insertWatermark` method, which is not a part of the Word Office.js API, and does not address the request to add a watermark to every page. |
| L1-content-control-001 | 83 | ✓ | The code is mostly correct but loses points for not using the Word.run method to execute the content control insertion, and for not fully following best practices in using the available helpers. |
| L1-content-control-002 | 80 | ✓ | The code correctly uses the Office.js API to insert a content control, but loses points for not specifying the exact location of the insertion, which could lead to unexpected behavior, and for not fully following best practices in terms of handling potential errors or edge cases. |
| L1-mail-merge-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "mailMergeReplace" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-template-apply-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a severe penalty for API correctness. |
| L1-template-apply-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would it work dimensions. |
| L1-document-generate-001 | 85 | ✓ | The code is mostly correct and complete, but loses points for not using the best practices, such as not checking if the context is loaded before calling insertTableOfContents, and not handling potential errors, which deducts from the approach and completeness scores. |
| L1-theme-apply-001 | 10 ⚠️ | ✓ | The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-theme-apply-002 | 25 ⚠️ | ✓ | The code heavily hallucinates the "designTheme" and "applyTheme" methods, which are not real Word Office.js APIs, resulting in a significant penalty for API correctness. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, and also does not use the correct API "designTheme" and "applyTheme" to modify the theme. |
| L1-citation-bibliography-001 | 45 ⚠️ | ✓ | The code hallucinates non-existent APIs, specifically `insertCitation` and `insertBibliography`, which are not part of the real Word Office.js APIs, resulting in a significant penalty for API correctness. |
| L1-equation-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "insertEquation" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the user's request, but loses points for not using a more robust approach, such as checking the current selection or range before inserting the equation, and not handling potential errors. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selection for clarity, is mostly clear and relevant, but loses points for not considering the coach panel recommendation as an alternative and not providing additional workarounds or caveats. |
| L1-read-query-001 | 40 ⚠️ | ✓ | The code fails to use real Word Office.js APIs, instead hallucinating a non-existent "countWords" method, which heavily penalizes the API correctness score. |
| L1-read-query-002 | 40 ⚠️ | ✓ | The code fails to utilize actual Word Office.js APIs, instead relying on a hallucinated `getReadability()` method, which is not a real API, resulting in a significant penalty for API correctness. |
| L1-read-query-003 | 40 ⚠️ | ✓ | The code fails to use real Word Office.js APIs, instead hallucinating a non-existent "listHeadings" method, which heavily penalizes the API correctness score. |
| L1-margins-orientation-001 | 85 | ✓ | The code correctly sets page margins using the Word Office.js API, but loses points for not using the most straightforward or best practice approach, such as utilizing a built-in method for setting margins if available, and for potential issues with context.sync() ordering. |
| L1-margins-orientation-002 | 80 | ✓ | The code is mostly correct but lacks the Word.run() context and may not fully address the request for the entire document, also it doesn't use the best practice of using Word.run() to queue the changes. |
| L1-web-search-needed-001 | 67 | ✓ | The response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks clarity on how this information will be used to add a section to the document and does not provide any workarounds or Word-specific guidance. |
| L1-web-search-needed-002 | 76 | ✓ | The response accurately and clearly explains the Pythagorean theorem, directly addressing the question, but loses points for not providing any meaningful bonus information or workarounds, and for including unnecessary code that was not requested. |
| L1-refusal-001 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not exploring alternative solutions or providing more detailed information about the "Save As" feature. |
| L1-refusal-002 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not offering more detailed alternatives or settings options, and for a slightly abrupt tone in the refusal. |
| L1-refusal-003 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js in creating Excel-style charts, but loses points for not providing a more detailed workaround, such as using the Excel REST API or a library to generate the chart, and for inserting a generic image instead of attempting to create a chart with the provided data. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-paragraph-spacing-001 | 72 | ✓ | The code uses real Word Office.js APIs and addresses the request, but has issues with approach, using a manual loop instead of available helpers like `body.paragraphs.each` and applying a style unnecessarily, and lacks error handling, which affects its completeness and potential for runtime errors. |
| L1-edge-case-001 | 25 ⚠️ | ✓ | The code fails to use the Word Office.js API to delete a paragraph and instead throws an error, lacking a proper attempt to access and delete the paragraph using the API, such as `Word.run` and `body.paragraphs.getItem`. |
| L1-form-field-001 | 83 | ✓ | The code correctly uses the Word Office.js API to insert a dropdown content control, but loses points for not fully addressing potential errors or edge cases, and not using the most optimal or best practice approach to inserting the content control. |
| L1-text-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addParagraph" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-text-edit-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| gen-L2-form-field-001 | 72 | ✓ | The code uses correct Word Office.js APIs and is likely to work, but it does not fully address the request as it inserts a date picker with the title "Sign Date" instead of placing it below the checkbox content control with the correct label, and lacks best practices in content control placement and styling. |
| gen-L2-form-field-002 | 47 ⚠️ | ✗ | The code incorrectly inserts two rich text content controls, one for "Name" and one for "Address", instead of replacing the existing "Name: _______" plain text with a single rich text content control and adding a building block gallery control as requested. |
| gen-L2-form-field-003 | 83 | ✗ | The code is mostly correct and complete, but loses points for not using the best practices, such as not checking if the group content control already exists before inserting a new one, and not handling potential errors that may occur during the execution of the code. |
| gen-L2-form-field-004 | 83 | ✗ | The code is mostly correct but loses points for not explicitly using the Word.run method to execute the content control and table insertion, and for not fully addressing the request of adding a repeating section content control, instead inserting a regular content control. |

## ⚠️ Needs attention

**[L1-paragraph-format-001]** score=0 — The code uses a hallucinated method `addParagraph` which does not exist in the Word Office.js API, and instead of modifying the existing paragraph, it attempts to add a new paragraph with the same text, resulting in a completely incorrect implementation.
- missing patterns: `font.bold`, `font.size`

**[L1-list-bullet-001]** score=25 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-list-multilevel-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions.

**[L1-header-footer-002]** score=20 — The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-watermark-insert-001]** score=0 — The code is completely incorrect as it uses a non-existent `insertWatermark` method, which is not a part of the Word Office.js API, and does not address the request to add a watermark to every page.

**[L1-mail-merge-001]** score=20 — The code uses a hallucinated method "mailMergeReplace" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-template-apply-001]** score=20 — The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a severe penalty for API correctness.

**[L1-template-apply-002]** score=20 — The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would it work dimensions.

**[L1-theme-apply-001]** score=10 — The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-theme-apply-002]** score=25 — The code heavily hallucinates the "designTheme" and "applyTheme" methods, which are not real Word Office.js APIs, resulting in a significant penalty for API correctness.

**[L1-theme-apply-003]** score=20 — The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, and also does not use the correct API "designTheme" and "applyTheme" to modify the theme.

**[L1-citation-bibliography-001]** score=45 — The code hallucinates non-existent APIs, specifically `insertCitation` and `insertBibliography`, which are not part of the real Word Office.js APIs, resulting in a significant penalty for API correctness.

**[L1-equation-001]** score=0 — The code uses a hallucinated method "insertEquation" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-read-query-001]** score=40 — The code fails to use real Word Office.js APIs, instead hallucinating a non-existent "countWords" method, which heavily penalizes the API correctness score.

**[L1-read-query-002]** score=40 — The code fails to utilize actual Word Office.js APIs, instead relying on a hallucinated `getReadability()` method, which is not a real API, resulting in a significant penalty for API correctness.

**[L1-read-query-003]** score=40 — The code fails to use real Word Office.js APIs, instead hallucinating a non-existent "listHeadings" method, which heavily penalizes the API correctness score.

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-edge-case-001]** score=25 — The code fails to use the Word Office.js API to delete a paragraph and instead throws an error, lacking a proper attempt to access and delete the paragraph using the API, such as `Word.run` and `body.paragraphs.getItem`.

**[L1-text-insert-001]** score=0 — The code uses a hallucinated method "addParagraph" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-text-edit-001]** score=0 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[gen-L2-form-field-002]** score=47 — The code incorrectly inserts two rich text content controls, one for "Name" and one for "Address", instead of replacing the existing "Name: _______" plain text with a single rich text content control and adding a building block gallery control as requested.
- missing patterns: `RichTextContentControl`
