# Word AI — Eval Results
**Last run:** 2026-05-19 18:11:23  
**Overall: 60.6/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 78.5/100 | — | 2 |
| paragraph-format | 41.5/100 | — | 2 |
| list-bullet | 25.0/100 | — | 1 |
| list-multilevel | 67.0/100 | — | 1 |
| style-apply | 83.0/100 | — | 1 |
| table-create | 89.0/100 | — | 1 |
| table-format | 72.0/100 | — | 1 |
| find-replace | 78.5/100 | — | 2 |
| find-replace-regex | 85.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 91.5/100 | — | 2 |
| toc-generate | 85.0/100 | — | 1 |
| section-break | 72.0/100 | — | 1 |
| header-footer | 41.5/100 | — | 2 |
| columns | 80.0/100 | — | 1 |
| image-insert | 20.0/100 | — | 2 |
| watermark-insert | 60.0/100 | — | 1 |
| content-control | 81.5/100 | — | 2 |
| mail-merge | 40.0/100 | — | 1 |
| template-apply | 81.5/100 | — | 2 |
| document-generate | 72.0/100 | — | 1 |
| theme-apply | 10.0/100 | — | 3 |
| citation-bibliography | 40.0/100 | — | 1 |
| equation | 83.0/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 26.7/100 | — | 3 |
| margins-orientation | 82.5/100 | — | 2 |
| web-search-needed | 70.0/100 | — | 2 |
| refusal | 75.7/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 92.0/100 | — | 1 |
| edge-cases | 45.0/100 | — | 1 |
| form-field | 61.2/100 | — | 5 |
| text-insert | 80.0/100 | — | 1 |
| text-edit | 0.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 85 | ✓ | The code is mostly correct but loses points for using a potentially non-standard method `addHeading` which is not a built-in Word Office.js API, although it seems to be a helper function that might be defined elsewhere in the codebase. |
| L1-heading-insert-002 | 72 | ✓ | The code uses correct APIs and mostly addresses the request, but lacks best practices and uses generic methods instead of available helpers, and also does not ensure the new paragraph is inserted below the heading. |
| L1-paragraph-format-001 | 0 ⚠️ | ✗ | The code uses a hallucinated method `addParagraph` which does not exist in the Word Office.js API, and instead of modifying the existing paragraph, it attempts to add a new paragraph with the same text, resulting in a completely incorrect solution. |
| L1-paragraph-format-002 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the best practice of applying a built-in style for the title and instead manually setting font size and alignment, and also for not checking if the paragraph is already loaded before accessing its properties. |
| L1-list-bullet-001 | 25 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a severe penalty for API correctness. |
| L1-list-multilevel-001 | 67 | ✓ | The code uses a non-standard method "addList" which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-style-apply-001 | 83 | ✓ | The code correctly applies the Quote style to the third paragraph, but loses points for not using the more specific and recommended `paragraph.styleBuiltIn` property instead of the generic `applyStyle` method, and also for not handling potential errors or edge cases. |
| L1-table-create-001 | 89 | ✓ | The code correctly uses the Word Office.js API to insert a table, but loses points for not fully addressing the request by not handling potential errors and not using the most efficient approach by directly using the `body.insertTable` method with the provided parameters. |
| L1-table-format-001 | 72 | ✓ | The code uses a non-existent `styleTable` method, which is not a real Word Office.js API, resulting in significant penalties in API correctness and approach dimensions. |
| L1-find-replace-001 | 67 | ✓ | The code uses a non-existent `replaceText` method, which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient or best practice approach, such as utilizing the `range.font` properties directly on the search results without explicitly loading them. |
| L1-find-replace-regex-001 | 85 | ✓ | The code uses the correct Word Office.js API and is likely to work, but loses points for not fully addressing the request with a more precise wildcard pattern and not using the most optimal approach by directly utilizing the `replaceText` method with a more specific pattern. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a correct API method to insert a footnote, but lacks specificity in targeting the current selection and does not fully follow best practices for implementing the request. |
| L1-comment-insert-001 | 72 | ✓ | The code uses a correct API method (insertComment) but lacks proper error handling, context loading, and range specification, which might lead to issues with finding the exact occurrence of the text "$4.2M" and deducts from completeness and best practices. |
| L1-track-changes-toggle-001 | 100 ✅ | ✓ | The code perfectly addresses the request to turn on track changes using the correct and available Word Office.js API, specifically the toggleTrackChanges method, with no flaws or unnecessary re-implementations. |
| L1-track-changes-toggle-002 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using a Word.run wrapper and potentially not handling errors, and also for not explicitly checking if there are tracked revisions before accepting them, which could lead to unnecessary sync calls. |
| L1-toc-generate-001 | 85 | ✓ | The code correctly uses the insertTableOfContents API and addresses the request, but loses points for not using the most straightforward approach and not fully specifying the API call with all necessary parameters, such as the location of the Table of Contents. |
| L1-section-break-001 | 72 | ✓ | The code uses correct APIs and is likely to work, but it lacks completeness in fully addressing the request, as it doesn't specify the section to apply the landscape orientation to, and the approach could be improved by utilizing available helpers for more robust and maintainable code. |
| L1-header-footer-001 | 83 | ✓ | The code is mostly correct but loses points for not using the best practice approach of directly manipulating the header object, instead relying on a generic setHeader method which may not be optimal for all scenarios. |
| L1-header-footer-002 | 0 ⚠️ | ✓ | The code uses hallucinated methods "setFooter" and "addPageNumbers" which are not real Word Office.js APIs, resulting in a complete failure across all dimensions. |
| L1-columns-001 | 80 | ✓ | The code correctly uses real Word Office.js APIs and should execute without runtime errors, but it lacks a precise way to define "the rest of the document" and does not utilize the most efficient or best practice approach for applying the two-column layout to the entire remaining document content. |
| L1-image-insert-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-image-insert-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-watermark-insert-001 | 60 | ✓ | The code partially addresses the request but lacks completeness in applying the watermark to every page, and does not utilize the Word Office.js API correctly to achieve the desired outcome, such as using the `section` or `header/footer` properties to insert the watermark on each page. |
| L1-content-control-001 | 83 | ✓ | The code correctly uses the Office.js API to insert a content control, but loses points for not handling potential errors and not using the most robust way to ensure the content control is inserted at the correct location, and also for not fully following best practices in terms of code organization and error handling. |
| L1-content-control-002 | 80 | ✓ | The code correctly uses the insertContentControl API and addresses the request, but loses points for not specifying the location of the content control, which might lead to unexpected behavior, and not following best practices for content control insertion. |
| L1-mail-merge-001 | 40 ⚠️ | ✓ | The code uses a hallucinated method `mailMergeReplace` which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and approach. |
| L1-template-apply-001 | 80 | ✓ | The code correctly applies the resume template with the provided details, but lacks error handling and does not utilize the full range of available Word Office.js APIs and best practices, such as using `Word.run` to execute the template application. |
| L1-template-apply-002 | 83 | ✓ | The code correctly applies the memo template with the specified fields, but loses points for not using the Word.run method to ensure the context is loaded before applying the template, and for not demonstrating best practices in using the available helpers. |
| L1-document-generate-001 | 72 | ✓ | The code correctly uses real Word Office.js APIs and would likely execute without runtime errors, but it lacks completeness in addressing the request, as it only generates headings without adding any paragraphs or content to the sections, and does not follow best practices by not utilizing available helpers like addParagraph to add content to each section. |
| L1-theme-apply-001 | 10 ⚠️ | ✓ | The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-theme-apply-002 | 0 ⚠️ | ✓ | The code uses a completely hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not a valid method, resulting in a total failure across all dimensions. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code hallucinates a non-existent `tweakTheme` method and fails to use the correct `designTheme` and `applyTheme` APIs, resulting in a completely incorrect implementation. |
| L1-citation-bibliography-001 | 40 ⚠️ | ✓ | The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a significant penalty for API correctness. |
| L1-equation-001 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the request, but loses points for not using a more robust approach, such as checking the current selection or range before inserting the equation, and not handling potential errors or exceptions. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the user's request, but loses points for not using a more robust method to handle equation insertion, such as checking the current selection or range, and for not demonstrating best practices in terms of error handling or input validation. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selection for clarity, but loses points for not considering alternative approaches, such as recommending the coach panel, and for using a somewhat generic phrase "Rewriting the selection for clarity" that could be more descriptive. |
| L1-read-query-001 | 40 ⚠️ | ✓ | The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `countWords()` method, which heavily penalizes its API correctness score. |
| L1-read-query-002 | 40 ⚠️ | ✓ | The code fails to utilize real Word Office.js APIs, instead relying on a hallucinated `getReadability()` method, which is not a valid API call, resulting in a significant penalty for API correctness. |
| L1-read-query-003 | 0 ⚠️ | ✓ | The code uses a hallucinated method "listHeadings()" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-margins-orientation-001 | 85 | ✓ | The code correctly sets page margins using the Office.js API, but loses points for not using a more explicit or best-practice approach, such as utilizing a built-in method for margin settings if available, and for potential issues with context synchronization. |
| L1-margins-orientation-002 | 80 | ✓ | The code is mostly correct but lacks the Word.run wrapper and context initialization, which are necessary for the Office.js API to function properly, and also does not fully follow best practices. |
| L1-web-search-needed-001 | 67 | ✓ | The response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks a clear explanation of how this information can be integrated into a Word document using Office.js, resulting in a low bonus score. |
| L1-web-search-needed-002 | 73 | ✓ | The response accurately and clearly explains the Pythagorean theorem, directly addressing the question, but loses points for providing unnecessary code and not offering any workarounds or Word-specific caveats, despite being a question that doesn't require code. |
| L1-refusal-001 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not offering more detailed guidance on the "Save As" feature or alternative solutions, and for a slightly abrupt tone. |
| L1-refusal-002 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but lacks a more detailed explanation or alternative solutions, resulting in moderate scores across the dimensions. |
| L1-refusal-003 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js in creating charts, but loses points for not providing a more detailed workaround or alternative solution that utilizes the provided data to generate a chart image. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code hallucinates a non-existent `addQuote` method, which is not a real Word Office.js API, and does not provide any actual implementation to insert a block quote. |
| L1-paragraph-spacing-001 | 92 | ✓ | The code is mostly correct and complete, but loses points for not using the more efficient and recommended `body.paragraphs.each` method instead of a manual loop, and for not handling potential errors that may occur during the `context.sync()` call. |
| L1-edge-case-001 | 45 ⚠️ | ✓ | The code fails to use the Word Office.js API correctly, instead directly throwing an error without attempting to access or manipulate the document's paragraphs using the available APIs. |
| L1-form-field-001 | 97 ✅ | ✓ | The code is nearly flawless, using the correct Office.js API to insert a dropdown content control, but loses a few points for not demonstrating a best practice of handling potential errors or exceptions that may occur during the execution of the insertContentControl method. |
| L1-text-insert-001 | 80 | ✓ | The code uses a real Word Office.js API and is likely to work, but loses points for not fully addressing the request to insert text at the end of the document, as addParagraph may not necessarily add it at the very end, and for not using the most appropriate helper methods. |
| L1-text-edit-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure across all dimensions. |
| gen-L2-form-field-001 | 72 | ✓ | The code uses correct Office.js APIs and is likely to work, but it does not fully address the request as it inserts a date picker with the title "Sign Date" instead of placing it below the checkbox content control with the label "I agree" as requested. |
| gen-L2-form-field-002 | 50 ⚠️ | ✗ | The code incorrectly inserts two rich text content controls instead of replacing the existing plain text "Name: _______" with a single rich text content control and adding a building block gallery control as requested. |
| gen-L2-form-field-003 | 15 ⚠️ | ✗ | The code uses a hallucinated method "insertContentControl" which is not a real Word Office.js API, and also fails to insert the content controls inside a group content control as requested. |
| gen-L2-form-field-004 | 72 | ✗ | The code uses real Word Office.js APIs and mostly addresses the request, but lacks a repeating section content control and uses a generic insertContentControl method, also the table insertion is correct but the content control type should be specifically set to "repeatingSection" for full completeness. |

## ⚠️ Needs attention

**[L1-paragraph-format-001]** score=0 — The code uses a hallucinated method `addParagraph` which does not exist in the Word Office.js API, and instead of modifying the existing paragraph, it attempts to add a new paragraph with the same text, resulting in a completely incorrect solution.
- missing patterns: `font.bold`, `font.size`

**[L1-list-bullet-001]** score=25 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a severe penalty for API correctness.

**[L1-header-footer-002]** score=0 — The code uses hallucinated methods "setFooter" and "addPageNumbers" which are not real Word Office.js APIs, resulting in a complete failure across all dimensions.

**[L1-image-insert-001]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-image-insert-002]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-mail-merge-001]** score=40 — The code uses a hallucinated method `mailMergeReplace` which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and approach.

**[L1-theme-apply-001]** score=10 — The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-theme-apply-002]** score=0 — The code uses a completely hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not a valid method, resulting in a total failure across all dimensions.

**[L1-theme-apply-003]** score=20 — The code hallucinates a non-existent `tweakTheme` method and fails to use the correct `designTheme` and `applyTheme` APIs, resulting in a completely incorrect implementation.

**[L1-citation-bibliography-001]** score=40 — The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a significant penalty for API correctness.

**[L1-read-query-001]** score=40 — The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `countWords()` method, which heavily penalizes its API correctness score.

**[L1-read-query-002]** score=40 — The code fails to utilize real Word Office.js APIs, instead relying on a hallucinated `getReadability()` method, which is not a valid API call, resulting in a significant penalty for API correctness.

**[L1-read-query-003]** score=0 — The code uses a hallucinated method "listHeadings()" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-quote-insert-001]** score=0 — The code hallucinates a non-existent `addQuote` method, which is not a real Word Office.js API, and does not provide any actual implementation to insert a block quote.

**[L1-edge-case-001]** score=45 — The code fails to use the Word Office.js API correctly, instead directly throwing an error without attempting to access or manipulate the document's paragraphs using the available APIs.

**[L1-text-edit-001]** score=0 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure across all dimensions.

**[gen-L2-form-field-002]** score=50 — The code incorrectly inserts two rich text content controls instead of replacing the existing plain text "Name: _______" with a single rich text content control and adding a building block gallery control as requested.
- missing patterns: `RichTextContentControl`

**[gen-L2-form-field-003]** score=15 — The code uses a hallucinated method "insertContentControl" which is not a real Word Office.js API, and also fails to insert the content controls inside a group content control as requested.
- missing patterns: `GroupContentControl`
