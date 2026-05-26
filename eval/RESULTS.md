# Word AI — Eval Results
**Last run:** 2026-05-26 18:33:21  
**Overall: 63.7/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 67.0/100 | — | 2 |
| paragraph-format | 85.5/100 | — | 2 |
| list-bullet | 20.0/100 | — | 1 |
| list-multilevel | 72.0/100 | — | 1 |
| style-apply | 80.0/100 | — | 1 |
| table-create | 83.0/100 | — | 1 |
| table-format | 72.0/100 | — | 1 |
| find-replace | 81.0/100 | — | 2 |
| find-replace-regex | 83.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 90.0/100 | — | 2 |
| toc-generate | 72.0/100 | — | 1 |
| section-break | 67.0/100 | — | 1 |
| header-footer | 43.5/100 | — | 2 |
| columns | 72.0/100 | — | 1 |
| image-insert | 46.0/100 | — | 2 |
| watermark-insert | 50.0/100 | — | 1 |
| content-control | 83.0/100 | — | 2 |
| mail-merge | 72.0/100 | — | 1 |
| template-apply | 84.0/100 | — | 2 |
| document-generate | 90.0/100 | — | 1 |
| theme-apply | 17.3/100 | — | 3 |
| citation-bibliography | 20.0/100 | — | 1 |
| equation | 79.0/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 45.7/100 | — | 3 |
| margins-orientation | 82.5/100 | — | 2 |
| web-search-needed | 74.0/100 | — | 2 |
| refusal | 79.3/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 0.0/100 | — | 1 |
| edge-cases | 50.0/100 | — | 1 |
| form-field | 71.2/100 | — | 5 |
| text-insert | 85.0/100 | — | 1 |
| text-edit | 0.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 67 | ✓ | The code uses a non-standard method `addHeading` which is not a real Word Office.js API, instead of using the available `paragraph` and `styleBuiltIn` properties to achieve the same result. |
| L1-heading-insert-002 | 67 | ✓ | The code uses non-existent helper functions like addHeading and addParagraph, which are not part of the Word Office.js API, instead of using the actual API methods like paragraph.insertText and paragraph.styleBuiltIn. |
| L1-paragraph-format-001 | 88 | ✓ | The code is mostly correct and complete, but loses points for not using the best practice of applying a built-in style or using the applyStyle helper, and also for potential issues with the load and sync order of the context. |
| L1-paragraph-format-002 | 83 | ✓ | The code is mostly correct but loses points for not using the best practice approach of utilizing available helpers like applyStyle for font size, and for potential issues with variable scope and load order. |
| L1-list-bullet-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would it work dimensions. |
| L1-list-multilevel-001 | 72 | ✓ | The code uses a non-standard `addList` method which is not a real Word Office.js API, and lacks proper integration with the `Word.run` context, but otherwise addresses the request and would likely work with minimal modifications. |
| L1-style-apply-001 | 80 | ✓ | The code correctly applies the Quote style to a paragraph using the applyStyle method, but loses points for not fully addressing the request by assuming the paragraph index is 2 without verifying the paragraph's position, and for not using the most robust and flexible approach by directly accessing the third paragraph. |
| L1-table-create-001 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the `Word.run` method to execute the insertion of the table, and for not fully following best practices, such as not checking if the paragraph exists before inserting the table. |
| L1-table-format-001 | 72 | ✓ | The code uses a non-standard `styleTable` method which is not a real Word Office.js API, and also lacks the use of best practices and available helpers, such as `table.style` or `table.styleBuiltIn`. |
| L1-find-replace-001 | 72 | ✓ | The code uses a non-existent `replaceText` method, which is not a real Word Office.js API, and lacks the proper `Word.run` context, resulting in significant deductions in API correctness and approach dimensions. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach, such as utilizing the `range.font` properties directly on the search results without loading the entire `font` property, and not handling potential errors that may occur during the execution of the code. |
| L1-find-replace-regex-001 | 83 | ✓ | The code uses the correct Word Office.js API and mostly addresses the request, but loses points for not handling potential errors and not using the most efficient approach, such as using Word.run to encapsulate the replacement operation. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a correct API to insert a footnote but lacks consideration for the current selection and does not fully address the request, also it does not use the best practice approach as it manually implements the insertion of a footnote without utilizing the available helpers. |
| L1-comment-insert-001 | 72 | ✓ | The code uses a correct API method `insertComment` but lacks specificity in targeting the first occurrence of "$4.2M" and does not fully address the request by not ensuring the comment is added to the correct location within the document. |
| L1-track-changes-toggle-001 | 97 ✅ | ✓ | The code perfectly addresses the request using the correct API, but loses a few points for not explicitly handling potential errors or edge cases, although the provided code snippet is otherwise flawless. |
| L1-track-changes-toggle-002 | 83 | ✓ | The code is mostly correct but loses points for not using the Word.run method to encapsulate the API calls and for not checking if there are any tracked revisions before accepting them, which could lead to unnecessary sync operations. |
| L1-toc-generate-001 | 72 | ✓ | The code uses a correct API method, `insertTableOfContents`, but loses points for not fully addressing the request, as it doesn't specify the exact location of the Table of Contents, and also for not using best practices, as it directly calls the API method without considering the context and potential errors. |
| L1-section-break-001 | 67 | ✓ | The code uses correct APIs and is likely to work, but it lacks completeness in addressing the request to switch the next section to landscape, and does not follow best practices by directly using low-level methods instead of available helpers. |
| L1-header-footer-001 | 67 | ✓ | The code uses a non-existent setHeader method, which is not a real Word Office.js API, resulting in significant penalties in API correctness and approach dimensions. |
| L1-header-footer-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would it work dimensions. |
| L1-columns-001 | 72 | ✓ | The code correctly uses real Word Office.js APIs and would likely execute without runtime errors, but it lacks completeness in addressing the user's request to switch the rest of the document into a two-column layout, as it doesn't specify the range of the document to apply the column change to, and doesn't follow best practices by not using available helpers. |
| L1-image-insert-001 | 72 | ✓ | The code uses a non-existent `insertImage` method, which is not a real Word Office.js API, and also lacks proper error handling and context setup, but it attempts to address the user's request and uses a correct width parameter. |
| L1-image-insert-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a severe penalty for API correctness and would not work at runtime. |
| L1-watermark-insert-001 | 50 ⚠️ | ✓ | The code uses a non-existent `insertWatermark` method, which is not a real Word Office.js API, and also does not specify the scope of the watermark to be applied to every page, resulting in significant deductions in API correctness and completeness. |
| L1-content-control-001 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the request, but loses points for not using the Word.run method to ensure the code is executed in the correct context and for not handling potential errors, and also for not using the most optimal approach by directly inserting the content control without considering the current selection or cursor position. |
| L1-content-control-002 | 83 | ✓ | The code correctly uses the Office.js API to insert a checkbox content control, but loses points for not fully addressing potential context and sync issues, and for not using the most optimal approach by directly utilizing the insertContentControl helper without considering the cursor position at the end of the document. |
| L1-mail-merge-001 | 72 | ✓ | The code uses a non-standard `mailMergeReplace` method which is not a real Word Office.js API, losing significant points for API correctness, but otherwise addresses the request and would likely execute without major issues. |
| L1-template-apply-001 | 83 | ✓ | The code correctly applies the resume template with the provided information, but loses points for not using the Word.run method to ensure the context is loaded before applying the template, and for not following best practices in terms of error handling and code organization. |
| L1-template-apply-002 | 85 | ✓ | The code accurately applies the memo template with the specified fields, but loses points for API correctness and approach due to the use of a potentially non-existent applyTemplate method, which is not a standard Word Office.js API. |
| L1-document-generate-001 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach, such as potentially reusing the same paragraph or heading objects instead of calling addHeading and addParagraph repeatedly. |
| L1-theme-apply-001 | 12 ⚠️ | ✓ | The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-theme-apply-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not a valid method, resulting in a complete loss of points for API correctness and would it work dimensions. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, and also does not use the correct API "designTheme" and "applyTheme" to modify the theme. |
| L1-citation-bibliography-001 | 20 ⚠️ | ✓ | The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-equation-001 | 75 | ✓ | The code uses a correct and existing API method `insertEquation`, but loses points for not using the `Word.run` context and not specifying the exact range where the equation should be inserted, and also for not following best practices by not using a more specific and robust method to insert the equation at the end of the paragraph. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the user's request, but loses points for not using a more robust method to handle equation insertion, such as checking the current selection or range, and for not following best practices in terms of error handling and code organization. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selection for clarity, is easy to understand, and directly addresses the question, but loses points for not considering alternative approaches, such as recommending the coach panel or providing additional context for the rewrite. |
| L1-read-query-001 | 40 ⚠️ | ✓ | The code fails to use real Word Office.js APIs, instead hallucinating a non-existent "countWords" method, which heavily penalizes its API correctness score. |
| L1-read-query-002 | 57 ⚠️ | ✓ | The code uses a hallucinated method `getReadability()` which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-read-query-003 | 40 ⚠️ | ✓ | The code fails to use real Word Office.js APIs, instead hallucinating a non-existent "listHeadings" method, which heavily penalizes the API correctness score. |
| L1-margins-orientation-001 | 85 | ✓ | The code correctly sets page margins using the Office.js API, but loses points for not using a more explicit or helper-based approach, and for potential issues with context synchronization. |
| L1-margins-orientation-002 | 80 | ✓ | The code is mostly correct but lacks consideration of the document's current state and potential existing content, and does not use the most idiomatic or efficient Word Office.js APIs, such as Word.run, to perform the page orientation switch. |
| L1-web-search-needed-001 | 67 | ✓ | The AI response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks clarity on how this information will be used to add a section to the Word document and does not offer any workarounds or Word-specific guidance. |
| L1-web-search-needed-002 | 81 | ✓ | The response accurately and clearly explains the Pythagorean theorem, directly addressing the question, but loses bonus points for not providing additional workarounds or Word-specific caveats, such as suggesting how to format the equation or insert it at a specific location in the document. |
| L1-refusal-001 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not exploring alternative solutions or providing more detailed information about the "Save As" feature. |
| L1-refusal-002 | 83 | ✓ | The response accurately conveys the limitation of Word Office.js and provides a clear workaround, but loses points for not exploring alternative solutions or providing more detailed information about the print menu option. |
| L1-refusal-003 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js in creating charts, but loses points for not providing a more detailed or creative workaround, and for inserting a placeholder image description instead of a more meaningful alternative. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-paragraph-spacing-001 | 0 ⚠️ | ✓ | The code hallucinates a non-existent `addParagraph` method, which is not a real Word Office.js API, and instead of modifying the existing paragraphs, it attempts to add new ones, resulting in a complete failure to address the user's request. |
| L1-edge-case-001 | 50 ⚠️ | ✓ | The code fails to attempt the deletion of the paragraph and instead immediately throws an error, not utilizing the Word Office.js APIs to handle the out-of-range scenario, and not following best practices for error handling and paragraph deletion. |
| L1-form-field-001 | 83 | ✓ | The code correctly uses the Word Office.js API to insert a dropdown content control, but loses points for not fully addressing potential errors or edge cases, and not using the most optimal approach by directly utilizing the `insertContentControl` method without considering the context of the cursor position in the document body. |
| L1-text-insert-001 | 85 | ✓ | The code uses a real Word Office.js API helper method `addParagraph` to insert the text, but loses points for not explicitly checking if the insertion is at the end of the document, and for not using the most precise method to ensure the text is inserted at the very end. |
| L1-text-edit-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| gen-L2-form-field-001 | 72 | ✓ | The code partially addresses the request but fails to fully meet the requirements, such as not inserting the date picker content control below the checkbox content control and using a different title for the date picker. |
| gen-L2-form-field-002 | 72 | ✗ | The code uses a non-existent "comboBox" content control type and lacks proper replacement of the existing plain text, resulting in incomplete and imperfect implementation. |
| gen-L2-form-field-003 | 57 ⚠️ | ✗ | The code fails to insert a group content control containing the two dropdown content controls, instead inserting three separate content controls, and does not utilize the available helpers for inserting content controls. |
| gen-L2-form-field-004 | 72 | ✗ | The code partially addresses the request but fails to fully implement a repeating section content control, instead using a rich text content control, and also lacks best practices in using available helpers. |

## ⚠️ Needs attention

**[L1-list-bullet-001]** score=20 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would it work dimensions.

**[L1-header-footer-002]** score=20 — The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would it work dimensions.

**[L1-image-insert-002]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a severe penalty for API correctness and would not work at runtime.

**[L1-watermark-insert-001]** score=50 — The code uses a non-existent `insertWatermark` method, which is not a real Word Office.js API, and also does not specify the scope of the watermark to be applied to every page, resulting in significant deductions in API correctness and completeness.

**[L1-theme-apply-001]** score=12 — The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-theme-apply-002]** score=20 — The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not a valid method, resulting in a complete loss of points for API correctness and would it work dimensions.

**[L1-theme-apply-003]** score=20 — The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, and also does not use the correct API "designTheme" and "applyTheme" to modify the theme.

**[L1-citation-bibliography-001]** score=20 — The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-read-query-001]** score=40 — The code fails to use real Word Office.js APIs, instead hallucinating a non-existent "countWords" method, which heavily penalizes its API correctness score.

**[L1-read-query-002]** score=57 — The code uses a hallucinated method `getReadability()` which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions.

**[L1-read-query-003]** score=40 — The code fails to use real Word Office.js APIs, instead hallucinating a non-existent "listHeadings" method, which heavily penalizes the API correctness score.

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-paragraph-spacing-001]** score=0 — The code hallucinates a non-existent `addParagraph` method, which is not a real Word Office.js API, and instead of modifying the existing paragraphs, it attempts to add new ones, resulting in a complete failure to address the user's request.

**[L1-edge-case-001]** score=50 — The code fails to attempt the deletion of the paragraph and instead immediately throws an error, not utilizing the Word Office.js APIs to handle the out-of-range scenario, and not following best practices for error handling and paragraph deletion.

**[L1-text-edit-001]** score=0 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[gen-L2-form-field-003]** score=57 — The code fails to insert a group content control containing the two dropdown content controls, instead inserting three separate content controls, and does not utilize the available helpers for inserting content controls.
- missing patterns: `GroupContentControl`
