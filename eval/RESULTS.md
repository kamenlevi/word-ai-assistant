# Word AI — Eval Results
**Last run:** 2026-05-18 03:46:42  
**Overall: 53.3/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 0.0/100 | — | 2 |
| paragraph-format | 41.5/100 | — | 2 |
| list-bullet | 0.0/100 | — | 1 |
| list-multilevel | 0.0/100 | — | 1 |
| style-apply | 20.0/100 | — | 1 |
| table-create | 83.0/100 | — | 1 |
| table-format | 72.0/100 | — | 1 |
| find-replace | 85.0/100 | — | 2 |
| find-replace-regex | 83.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 73.0/100 | — | 1 |
| track-changes-toggle | 90.0/100 | — | 2 |
| toc-generate | 83.0/100 | — | 1 |
| section-break | 80.0/100 | — | 1 |
| header-footer | 78.5/100 | — | 2 |
| columns | 83.0/100 | — | 1 |
| image-insert | 46.0/100 | — | 2 |
| watermark-insert | 60.0/100 | — | 1 |
| content-control | 81.5/100 | — | 2 |
| mail-merge | 20.0/100 | — | 1 |
| template-apply | 20.0/100 | — | 2 |
| document-generate | 85.0/100 | — | 1 |
| theme-apply | 15.0/100 | — | 3 |
| citation-bibliography | 20.0/100 | — | 1 |
| equation | 83.0/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 20.0/100 | — | 3 |
| margins-orientation | 75.0/100 | — | 2 |
| web-search-needed | 66.0/100 | — | 2 |
| refusal | 59.0/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 72.0/100 | — | 1 |
| edge-cases | 25.0/100 | — | 1 |
| form-field | 66.0/100 | — | 5 |
| text-insert | 80.0/100 | — | 1 |
| text-edit | 0.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addHeading" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions. |
| L1-heading-insert-002 | 0 ⚠️ | ✓ | The code uses hallucinated methods "addHeading" and "addParagraph" which are not real Word Office.js APIs, resulting in a complete loss of points across all dimensions. |
| L1-paragraph-format-001 | 0 ⚠️ | ✗ | The code uses a hallucinated method `addParagraph` which is not a real Word Office.js API, and instead of modifying the existing paragraph, it attempts to add a new paragraph with the same text, resulting in a completely incorrect solution. |
| L1-paragraph-format-002 | 83 | ✓ | The code is mostly correct but loses points for not using the built-in `applyStyle` method to set the font size and for not checking if the first paragraph is indeed the title before applying the changes. |
| L1-list-bullet-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-list-multilevel-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions. |
| L1-style-apply-001 | 20 ⚠️ | ✓ | A: 20
B: |
| L1-table-create-001 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the Word.run method to execute the insertion of the table, and for not fully following best practices by directly specifying the table style instead of using available helpers. |
| L1-table-format-001 | 72 | ✓ | The code uses a non-existent `styleTable` method, which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-find-replace-001 | 80 | ✓ | The code correctly uses the replaceText API and addresses the user's request, but loses points for not using Word.run to initiate the operation and not handling potential errors, and for not fully following best practices by not utilizing the available helpers in the most optimal way. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient or best-practice approach, such as utilizing the `range.font` properties directly on the search results without explicitly loading the `font` property. |
| L1-find-replace-regex-001 | 83 | ✓ | The code correctly uses the Word Office.js API and addresses the request, but loses points for not handling potential errors and not using the most efficient approach, such as checking if the replacement is necessary before calling replaceText. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a correct API method to insert a footnote, but lacks context about the current selection and does not fully address the request by not specifying the range where the footnote should be inserted, and does not follow best practices by not using the provided helpers and context information. |
| L1-comment-insert-001 | 73 | ✓ | The code uses a correct API method `insertComment` but lacks proper range specification, which might lead to incorrect comment placement, and does not fully follow best practices for searching and commenting on specific text within the document. |
| L1-track-changes-toggle-001 | 100 ✅ | ✓ | The code perfectly addresses the request to turn on track changes using the correct and available Word Office.js API, specifically the toggleTrackChanges method, with no flaws or areas for improvement. |
| L1-track-changes-toggle-002 | 80 | ✓ | The code is mostly correct but lacks the Word.run() context and may not fully handle potential errors, also it does not use the best practice of wrapping the code in a Word.run() block for a more robust solution. |
| L1-toc-generate-001 | 83 | ✓ | The code correctly uses the insertTableOfContents API and addresses the request, but loses points for not using the most specific and efficient approach, such as specifying the exact heading levels present in the document, and for not handling potential errors or edge cases. |
| L1-section-break-001 | 80 | ✓ | The code is mostly correct but lacks consideration for the current section and paragraph context, and does not fully follow best practices by directly using low-level APIs like insertSectionBreak and setPageOrientation without utilizing higher-level helpers. |
| L1-header-footer-001 | 85 | ✓ | The code is mostly correct but loses points for not using the best practice approach of utilizing Word Office.js APIs like Word.run to encapsulate the operation and ensure proper context loading and synchronization. |
| L1-header-footer-002 | 72 | ✗ | The code uses a non-existent `setFooter` method, which is not a real Word Office.js API, resulting in significant penalties in API correctness and approach dimensions. |
| L1-columns-001 | 83 | ✓ | The code is mostly correct and would work, but it lacks consideration for the current selection or range, and does not fully address potential edge cases, such as handling existing section breaks or column settings, and does not use the most efficient or best practice approaches. |
| L1-image-insert-001 | 72 | ✓ | The code uses a non-existent `insertImage` method, which is not a real Word Office.js API, and also lacks consideration for the cursor position at paragraph 3, but otherwise attempts to address the request with a reasonable approach. |
| L1-image-insert-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-watermark-insert-001 | 60 | ✓ | The code partially addresses the request but lacks specificity in applying the watermark to every page, and does not utilize the Word Office.js API correctly to achieve the desired outcome, such as using the `section` or `header/footer` properties to apply the watermark to all pages. |
| L1-content-control-001 | 83 | ✓ | The code correctly uses the Office.js API to insert a content control, but loses points for not handling potential errors and not using the most robust way to ensure the content control is inserted at the correct location, and also for not fully following best practices in terms of error handling and code organization. |
| L1-content-control-002 | 80 | ✓ | The code is mostly correct but loses points for not fully addressing potential context and loading issues, and not using the most idiomatic or best-practice approach to inserting a content control in the Office JavaScript API. |
| L1-mail-merge-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "mailMergeReplace" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-template-apply-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-template-apply-002 | 20 ⚠️ | ✓ | The code hallucinates a non-existent `applyTemplate` method, which is not a real Word Office.js API, resulting in a severe penalty for API correctness. |
| L1-document-generate-001 | 85 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach, such as potentially re-implementing table of contents logic instead of relying on built-in functionality, and not fully utilizing available helpers like applyStyle or replaceText. |
| L1-theme-apply-001 | 5 ⚠️ | ✓ | The code hallucinates non-existent `designTheme` and `applyTheme` methods, which are not part of the real Word Office.js APIs, resulting in a complete failure in API correctness and overall approach. |
| L1-theme-apply-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not a valid method, resulting in a complete loss of points for API correctness and would it work dimensions. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, and also does not use the correct API "designTheme" and "applyTheme" to modify the theme. |
| L1-citation-bibliography-001 | 20 ⚠️ | ✓ | The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would not work at runtime. |
| L1-equation-001 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the request, but loses points for not using a more robust approach, such as checking the current selection or range before inserting the equation, and not handling potential errors or exceptions that may occur during execution. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the user's request, but loses points for not using a more robust method to handle equation insertion, such as checking the current selection or range, and not following best practices for error handling and code organization. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selected text for clarity, but loses points for not considering alternative approaches, such as recommending the coach panel, and for using a somewhat generic phrase "Rewriting the current selection for clarity" that doesn't add much value. |
| L1-read-query-001 | 20 ⚠️ | ✓ | The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `countWords()` method, which heavily penalizes the API correctness score. |
| L1-read-query-002 | 40 ⚠️ | ✓ | The code heavily relies on a hallucinated method "getReadability()" which is not a real Word Office.js API, resulting in a significant penalty for API correctness. |
| L1-read-query-003 | 0 ⚠️ | ✓ | The code hallucinates a non-existent `listHeadings()` method, which is not a real Word Office.js API, and does not provide any actual implementation to list the headings in the document. |
| L1-margins-orientation-001 | 83 | ✓ | The code correctly uses the Word Office.js API and addresses the request, but loses points for not using a best practice approach, as the setMargins method is not a standard Word API, and the code does not handle potential errors or loading issues. |
| L1-margins-orientation-002 | 67 | ✓ | The code uses a correct API but lacks best practices and does not fully utilize available helpers, also the method setPageOrientation is not a standard Word Office.js API, it seems to be a custom or wrapper method. |
| L1-web-search-needed-001 | 67 | ✓ | The response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks clarity on how this information will be used to add a section to the Word document and does not provide any workarounds or Word-specific guidance. |
| L1-web-search-needed-002 | 65 | ✓ | The response accurately and clearly explains the Pythagorean theorem, but loses points for inserting unnecessary code and not providing any workarounds or Word-specific caveats, making the bonus score zero. |
| L1-refusal-001 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not offering more detailed guidance on using the "Save As" feature or alternative solutions. |
| L1-refusal-002 | 22 ⚠️ | ✓ | A: 22
B: |
| L1-refusal-003 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js in creating Excel-style bar charts, but loses points for not providing a more detailed or relevant workaround, such as suggesting how to generate the chart image from the provided data. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-paragraph-spacing-001 | 72 | ✓ | The code correctly uses the Word Office.js API and addresses the user's request, but it lacks best practices, such as using the `applyStyle` method more efficiently, and has potential issues with the load and sync order, which may cause runtime errors. |
| L1-edge-case-001 | 25 ⚠️ | ✓ | The code fails to use the Word Office.js API to attempt deletion of the paragraph and instead immediately throws an error, lacking a proper approach to handle the out-of-range paragraph index. |
| L1-form-field-001 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the user's request, but loses points for not using a more specific and descriptive title for the content control, and for not fully following best practices in terms of code organization and variable naming. |
| L1-text-insert-001 | 80 | ✓ | The code uses a real Word Office.js API (addParagraph) but loses points for not fully addressing the request to insert text at the end of the document, as it doesn't ensure the paragraph is added after the existing content, and also for not using the most optimal approach by directly manipulating the body's paragraphs. |
| L1-text-edit-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions. |
| gen-L2-form-field-001 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the request, but loses points for not using the best practices, such as not specifying the location of the content controls and not using a more robust method to ensure the date picker is inserted below the checkbox. |
| gen-L2-form-field-002 | 72 | ✗ | The code partially addresses the request but lacks specificity in replacing the existing plain text and incorrectly uses a comboBox content control instead of a building block gallery control. |
| gen-L2-form-field-003 | 20 ⚠️ | ✗ | A: 20
B: |
| gen-L2-form-field-004 | 72 | ✗ | The code partially addresses the request but lacks the implementation of a repeating section content control, instead inserting a rich text content control, and does not fully follow best practices for inserting tables and content controls in Word using Office.js APIs. |

## ⚠️ Needs attention

**[L1-heading-insert-001]** score=0 — The code uses a hallucinated method "addHeading" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions.

**[L1-heading-insert-002]** score=0 — The code uses hallucinated methods "addHeading" and "addParagraph" which are not real Word Office.js APIs, resulting in a complete loss of points across all dimensions.

**[L1-paragraph-format-001]** score=0 — The code uses a hallucinated method `addParagraph` which is not a real Word Office.js API, and instead of modifying the existing paragraph, it attempts to add a new paragraph with the same text, resulting in a completely incorrect solution.
- missing patterns: `font.bold`, `font.size`

**[L1-list-bullet-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-list-multilevel-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions.

**[L1-style-apply-001]** score=20 — A: 20
B:

**[L1-image-insert-002]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-mail-merge-001]** score=20 — The code uses a hallucinated method "mailMergeReplace" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-template-apply-001]** score=20 — The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-template-apply-002]** score=20 — The code hallucinates a non-existent `applyTemplate` method, which is not a real Word Office.js API, resulting in a severe penalty for API correctness.

**[L1-theme-apply-001]** score=5 — The code hallucinates non-existent `designTheme` and `applyTheme` methods, which are not part of the real Word Office.js APIs, resulting in a complete failure in API correctness and overall approach.

**[L1-theme-apply-002]** score=20 — The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not a valid method, resulting in a complete loss of points for API correctness and would it work dimensions.

**[L1-theme-apply-003]** score=20 — The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, and also does not use the correct API "designTheme" and "applyTheme" to modify the theme.

**[L1-citation-bibliography-001]** score=20 — The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would not work at runtime.

**[L1-read-query-001]** score=20 — The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `countWords()` method, which heavily penalizes the API correctness score.

**[L1-read-query-002]** score=40 — The code heavily relies on a hallucinated method "getReadability()" which is not a real Word Office.js API, resulting in a significant penalty for API correctness.

**[L1-read-query-003]** score=0 — The code hallucinates a non-existent `listHeadings()` method, which is not a real Word Office.js API, and does not provide any actual implementation to list the headings in the document.

**[L1-refusal-002]** score=22 — A: 22
B:

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-edge-case-001]** score=25 — The code fails to use the Word Office.js API to attempt deletion of the paragraph and instead immediately throws an error, lacking a proper approach to handle the out-of-range paragraph index.

**[L1-text-edit-001]** score=0 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions.

**[gen-L2-form-field-003]** score=20 — A: 20
B:
- missing patterns: `GroupContentControl`
