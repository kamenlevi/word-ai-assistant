# Word AI — Eval Results
**Last run:** 2026-05-25 11:59:25  
**Overall: 58.5/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 36.0/100 | — | 2 |
| paragraph-format | 90.0/100 | — | 2 |
| list-bullet | 0.0/100 | — | 1 |
| list-multilevel | 0.0/100 | — | 1 |
| style-apply | 80.0/100 | — | 1 |
| table-create | 89.0/100 | — | 1 |
| table-format | 72.0/100 | — | 1 |
| find-replace | 85.0/100 | — | 2 |
| find-replace-regex | 83.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 89.0/100 | — | 2 |
| toc-generate | 83.0/100 | — | 1 |
| section-break | 72.0/100 | — | 1 |
| header-footer | 51.5/100 | — | 2 |
| columns | 83.0/100 | — | 1 |
| image-insert | 20.0/100 | — | 2 |
| watermark-insert | 0.0/100 | — | 1 |
| content-control | 78.0/100 | — | 2 |
| mail-merge | 55.0/100 | — | 1 |
| template-apply | 80.0/100 | — | 2 |
| document-generate | 83.0/100 | — | 1 |
| theme-apply | 20.0/100 | — | 3 |
| citation-bibliography | 20.0/100 | — | 1 |
| equation | 79.0/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 26.7/100 | — | 3 |
| margins-orientation | 78.5/100 | — | 2 |
| web-search-needed | 77.0/100 | — | 2 |
| refusal | 75.7/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 20.0/100 | — | 1 |
| edge-cases | 25.0/100 | — | 1 |
| form-field | 75.0/100 | — | 5 |
| text-insert | 72.0/100 | — | 1 |
| text-edit | 0.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addHeading" which is not a real Word Office.js API, resulting in a complete failure across all dimensions. |
| L1-heading-insert-002 | 72 | ✓ | The code uses a non-standard method signature with `addHeading` and `addParagraph` which are not actual Word Office.js APIs, resulting in significant deductions in API correctness and approach dimensions. |
| L1-paragraph-format-001 | 88 | ✓ | The code accurately uses real Word Office.js APIs and fully addresses the request, but loses points for not using the most efficient approach, such as utilizing the `range.font` properties directly on the paragraph object without needing to load and sync the `font` property separately. |
| L1-paragraph-format-002 | 92 | ✓ | The code accurately uses Word Office.js APIs, fully addresses the request, and would likely execute without runtime errors, but loses points for not using the available helpers, such as applyStyle or replaceText, to center-align and resize the title. |
| L1-list-bullet-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-list-multilevel-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions. |
| L1-style-apply-001 | 80 | ✓ | The code correctly applies the Quote style to a paragraph using the applyStyle method, but loses points for not fully addressing potential edge cases and not using the most robust best practices, such as error handling and explicit paragraph indexing. |
| L1-table-create-001 | 89 | ✓ | The code accurately uses the Word Office.js API to insert a table, but loses points for not fully addressing the request by not handling potential errors and not using the most efficient approach by directly using the `body.insertTable` method with the provided parameters. |
| L1-table-format-001 | 72 | ✓ | The code uses a non-standard method `styleTable` which is not a real Word Office.js API, and lacks the use of best practices and available helpers, such as `table.styleBuiltIn`, resulting in a partial loss of points across all dimensions. |
| L1-find-replace-001 | 80 | ✓ | The code uses a correct API method replaceText, but lacks the Word.run wrapper and context setup, and does not fully follow best practices by not utilizing the Word.run method to execute the replacement. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient or best practice approach, such as utilizing the `range.font` properties directly after searching, and not handling potential errors or edge cases. |
| L1-find-replace-regex-001 | 83 | ✓ | The code correctly uses the Word Office.js API and addresses the request, but loses points for not handling potential errors and not using the most efficient approach, such as using a more specific wildcard pattern to match email addresses. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a real Office.js API (insertFootnote) but lacks specificity in targeting the current selection (paragraph 2) and does not fully follow best practices for inserting a footnote at a specific location. |
| L1-comment-insert-001 | 72 | ✓ | The code uses a correct API method `insertComment` but lacks proper context and range handling, and does not fully address the request as it does not specify the exact location of the comment, resulting in potential inaccuracies and deductions in completeness and approach dimensions. |
| L1-track-changes-toggle-001 | 95 ✅ | ✓ | The code perfectly addresses the request using the correct API, but loses points for approach because it manually calls toggleTrackChanges instead of using a potentially more robust built-in helper, if available, although in this case, toggleTrackChanges is a suitable method. |
| L1-track-changes-toggle-002 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the Word.run method to wrap the API calls, and for not checking if the document has tracked revisions before accepting them, which could lead to unnecessary sync calls. |
| L1-toc-generate-001 | 83 | ✓ | The code correctly uses the insertTableOfContents API and addresses the request, but loses points for not using the most straightforward approach and not fully considering potential edge cases, such as the user's desired table of contents title or level settings. |
| L1-section-break-001 | 72 | ✓ | The code correctly uses real Word Office.js APIs and addresses the request, but lacks best practices, has potential sync issues, and does not fully handle the section break and landscape orientation for the next section. |
| L1-header-footer-001 | 83 | ✓ | The code is mostly correct but loses points for not using the Word.run method to ensure the context is loaded before setting the header, and for not checking if the primary header already exists before setting it. |
| L1-header-footer-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-columns-001 | 83 | ✓ | The code is mostly correct but loses points for not using the Word.run method to ensure the context is loaded and for not checking the current section's layout before switching to a two-column layout, and also for not using best practices such as checking the current selection or range. |
| L1-image-insert-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, and does not use the correct API such as Word.run or body.insertInlinePicture to insert an image. |
| L1-image-insert-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-watermark-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "insertWatermark" which is not a real Word Office.js API, resulting in a complete failure to address the request correctly. |
| L1-content-control-001 | 83 | ✓ | The code correctly uses the Office.js API to insert a content control, but loses points for not handling potential errors and not using the most descriptive variable names, and also for not fully addressing the request in terms of providing a complete date picker implementation. |
| L1-content-control-002 | 73 | ✓ | The code is mostly correct but lacks context about the current selection or range, and does not fully address best practices for inserting content controls in the Word Office.js API. |
| L1-mail-merge-001 | 55 ⚠️ | ✓ | The code uses a hallucinated method `mailMergeReplace` which is not a real Word Office.js API, resulting in a complete loss of points for API correctness. |
| L1-template-apply-001 | 80 | ✓ | The code correctly applies the resume template with the provided details, but loses points for not using the most efficient or best practice approaches, such as handling potential errors or using more specific template application methods. |
| L1-template-apply-002 | 80 | ✓ | The code correctly applies the memo template with the specified fields, but loses points for not using Word.run to ensure the context is loaded and for not handling potential errors, and also for not demonstrating best practices in template application. |
| L1-document-generate-001 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient or best-practice approaches, such as potentially re-implementing table of contents logic instead of relying on built-in functionality. |
| L1-theme-apply-001 | 5 ⚠️ | ✓ | The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-theme-apply-002 | 35 ⚠️ | ✓ | The code hallucinates the "designTheme" and "applyTheme" methods, which are not real Word Office.js APIs, resulting in a significant penalty for API correctness. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-citation-bibliography-001 | 20 ⚠️ | ✓ | The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-equation-001 | 75 | ✓ | The code uses a correct and existing API method `insertEquation`, but loses points for not using the `Word.run` context and not specifying the exact range where the equation should be inserted, and also for not following best practices by not using a more specific and robust method to insert the equation at the end of the paragraph. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the user's request, but loses points for not checking the current cursor position and potentially inserting the equation at the wrong location, and for not using a more robust approach to handle potential errors or exceptions. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selection for clarity, but loses points for not considering the special note about recommending the coach panel as an alternative, and for not providing additional context or workarounds, resulting in a slightly lower score for clarity and bonus. |
| L1-read-query-001 | 40 ⚠️ | ✓ | The code hallucinates a non-existent `countWords()` method, which is not a real Word Office.js API, resulting in a significant penalty for API correctness. |
| L1-read-query-002 | 40 ⚠️ | ✓ | The code fails to use real Word Office.js APIs, instead relying on a hallucinated `getReadability()` method, which is not a valid API call in the Office.js library. |
| L1-read-query-003 | 0 ⚠️ | ✓ | The code uses a hallucinated method "listHeadings()" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-margins-orientation-001 | 85 | ✓ | The code correctly sets page margins using the Office.js API, but lacks best practices and does not utilize available helpers, and also assumes the context is already loaded without explicitly showing the Word.run() method to ensure the context is properly synchronized. |
| L1-margins-orientation-002 | 72 | ✓ | The code uses a correct API but lacks best practices, such as using Word.run to execute the operation, and does not fully address the request as it doesn't specify the section or document scope for the page orientation change. |
| L1-web-search-needed-001 | 67 | ✓ | The response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks clarity on how this information will be used to add a section to the document and does not provide any workarounds or Word-specific guidance. |
| L1-web-search-needed-002 | 87 | ✓ | The response accurately and clearly explains the Pythagorean theorem, directly addressing the question, but loses points for not providing additional workarounds or Word-specific caveats, and for including unnecessary code in the response. |
| L1-refusal-001 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not offering more detailed guidance on using the "Save As" feature or alternative solutions. |
| L1-refusal-002 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but lacks a more detailed explanation or alternative solutions, resulting in moderate scores across the dimensions. |
| L1-refusal-003 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js in creating Excel-style charts, but loses points for not providing a more detailed or relevant workaround, such as suggesting how to actually create the chart image or insert the provided data into the chart. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions. |
| L1-paragraph-spacing-001 | 20 ⚠️ | ✓ | The code fails to address the user's request to set line spacing to 1.5 on existing paragraphs, instead attempting to add new paragraphs with the specified line spacing, and also incorrectly loads the "font" property of each paragraph. |
| L1-edge-case-001 | 25 ⚠️ | ✓ | The code fails to use the Word Office.js API to attempt deletion of the paragraph and instead immediately throws an error, lacking a proper approach to handle the out-of-range paragraph index. |
| L1-form-field-001 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the request, but loses points for not using a more specific title or tag that directly relates to the user's request for a "dropdown content control with options Yes / No / Maybe", and for not demonstrating best practices in terms of error handling or content control configuration. |
| L1-text-insert-001 | 72 | ✓ | The code uses a real Office.js API helper method `addParagraph`, but loses points for not explicitly addressing the "at the end of the document" requirement and for not using the most precise method to insert text, such as `body.insertParagraph` or `paragraph.insertText`. |
| L1-text-edit-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure across all dimensions. |
| gen-L2-form-field-001 | 72 | ✓ | The code uses correct Office.js APIs and is likely to execute without runtime errors, but it does not fully address the request as it inserts the date picker content control with a title "Date" instead of placing it below the checkbox content control with the correct label, and it lacks best practices in content control insertion. |
| gen-L2-form-field-002 | 65 | ✗ | The code partially addresses the request but lacks proper placement of the rich text content control, incorrectly uses a comboBox content control instead of a building block gallery control, and does not handle potential runtime errors. |
| gen-L2-form-field-003 | 72 | ✗ | The code uses correct Office.js APIs and mostly addresses the request, but it fails to properly nest the dropdown content controls within the group content control, and does not follow best practices for inserting content controls. |
| gen-L2-form-field-004 | 83 | ✗ | The code is mostly correct but loses points for not explicitly setting the content control as a repeating section and for not fully utilizing best practices such as directly applying the content control to the table. |

## ⚠️ Needs attention

**[L1-heading-insert-001]** score=0 — The code uses a hallucinated method "addHeading" which is not a real Word Office.js API, resulting in a complete failure across all dimensions.

**[L1-list-bullet-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-list-multilevel-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions.

**[L1-header-footer-002]** score=20 — The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-image-insert-001]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, and does not use the correct API such as Word.run or body.insertInlinePicture to insert an image.

**[L1-image-insert-002]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-watermark-insert-001]** score=0 — The code uses a hallucinated method "insertWatermark" which is not a real Word Office.js API, resulting in a complete failure to address the request correctly.

**[L1-mail-merge-001]** score=55 — The code uses a hallucinated method `mailMergeReplace` which is not a real Word Office.js API, resulting in a complete loss of points for API correctness.

**[L1-theme-apply-001]** score=5 — The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-theme-apply-002]** score=35 — The code hallucinates the "designTheme" and "applyTheme" methods, which are not real Word Office.js APIs, resulting in a significant penalty for API correctness.

**[L1-theme-apply-003]** score=20 — The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-citation-bibliography-001]** score=20 — The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-read-query-001]** score=40 — The code hallucinates a non-existent `countWords()` method, which is not a real Word Office.js API, resulting in a significant penalty for API correctness.

**[L1-read-query-002]** score=40 — The code fails to use real Word Office.js APIs, instead relying on a hallucinated `getReadability()` method, which is not a valid API call in the Office.js library.

**[L1-read-query-003]** score=0 — The code uses a hallucinated method "listHeadings()" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions.

**[L1-paragraph-spacing-001]** score=20 — The code fails to address the user's request to set line spacing to 1.5 on existing paragraphs, instead attempting to add new paragraphs with the specified line spacing, and also incorrectly loads the "font" property of each paragraph.

**[L1-edge-case-001]** score=25 — The code fails to use the Word Office.js API to attempt deletion of the paragraph and instead immediately throws an error, lacking a proper approach to handle the out-of-range paragraph index.

**[L1-text-edit-001]** score=0 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure across all dimensions.
