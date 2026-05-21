# Word AI — Eval Results
**Last run:** 2026-05-21 18:06:31  
**Overall: 56.3/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 42.5/100 | — | 2 |
| paragraph-format | 44.0/100 | — | 2 |
| list-bullet | 25.0/100 | — | 1 |
| list-multilevel | 0.0/100 | — | 1 |
| style-apply | 80.0/100 | — | 1 |
| table-create | 92.0/100 | — | 1 |
| table-format | 67.0/100 | — | 1 |
| find-replace | 81.0/100 | — | 2 |
| find-replace-regex | 80.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 91.5/100 | — | 2 |
| toc-generate | 83.0/100 | — | 1 |
| section-break | 72.0/100 | — | 1 |
| header-footer | 40.0/100 | — | 2 |
| columns | 72.0/100 | — | 1 |
| image-insert | 46.0/100 | — | 2 |
| watermark-insert | 0.0/100 | — | 1 |
| content-control | 83.0/100 | — | 2 |
| mail-merge | 83.0/100 | — | 1 |
| template-apply | 82.5/100 | — | 2 |
| document-generate | 83.0/100 | — | 1 |
| theme-apply | 10.0/100 | — | 3 |
| citation-bibliography | 20.0/100 | — | 1 |
| equation | 41.5/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 20.0/100 | — | 3 |
| margins-orientation | 81.5/100 | — | 2 |
| web-search-needed | 68.0/100 | — | 2 |
| refusal | 75.7/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 0.0/100 | — | 1 |
| edge-cases | 25.0/100 | — | 1 |
| form-field | 80.0/100 | — | 5 |
| text-insert | 80.0/100 | — | 1 |
| text-edit | 0.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 85 | ✓ | The code is mostly correct but loses points for using a non-standard method signature `addHeading` which is not a real Word Office.js API, instead of using the correct `paragraph.styleBuiltIn` and `paragraph.insertText` methods. |
| L1-heading-insert-002 | 0 ⚠️ | ✓ | The code uses hallucinated methods "addHeading" and "addParagraph" which are not real Word Office.js APIs, resulting in a complete failure across all dimensions. |
| L1-paragraph-format-001 | 0 ⚠️ | ✗ | The code uses a hallucinated method `addParagraph` which does not exist in the Word Office.js API, and instead of modifying the existing paragraph, it attempts to add a new paragraph with the same text, resulting in a completely incorrect solution. |
| L1-paragraph-format-002 | 88 | ✓ | The code accurately addresses the user's request but loses points for not using the more idiomatic and efficient `paragraph.styleBuiltIn` or `applyStyle` methods to set the font size and alignment, instead directly modifying the `font` and `alignment` properties. |
| L1-list-bullet-001 | 25 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice. |
| L1-list-multilevel-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the API correctness and best practices criteria. |
| L1-style-apply-001 | 80 | ✓ | The code correctly applies the Quote style to the third paragraph using the applyStyle method, but loses points for not using the Word.run method to execute the operation and for not checking if the paragraph exists before applying the style. |
| L1-table-create-001 | 92 | ✓ | The code is mostly correct and uses the real Word Office.js API, but loses points for completeness as it doesn't fully address the request by not handling the "Notes" column properly and not considering the document context, and also for approach as it manually specifies the table style instead of using the available helpers. |
| L1-table-format-001 | 67 | ✓ | The code uses a non-standard and hallucinated method `styleTable` instead of utilizing the available Word Office.js APIs, such as `table.styleBuiltIn`, to apply the desired table style. |
| L1-find-replace-001 | 72 | ✓ | The code uses a non-existent `replaceText` method, which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient or recommended approach, such as using `range.font` directly on the search results instead of loading and syncing the font property separately. |
| L1-find-replace-regex-001 | 80 | ✓ | The code uses the correct Word Office.js API replaceText with matchWildcards, but loses points for not fully addressing potential edge cases and not using the most robust pattern for matching email addresses, and also for not being wrapped in a Word.run context. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a correct API method to insert a footnote, but lacks context about the current selection and does not fully address the request by not specifying the range where the footnote should be inserted, and does not follow best practices by not using the provided helpers. |
| L1-comment-insert-001 | 72 | ✓ | The code uses a non-existent `insertComment` method with two parameters, which is not a real Word Office.js API, and lacks proper error handling and context setup, but attempts to address the user's request and follows some best practices. |
| L1-track-changes-toggle-001 | 100 ✅ | ✓ | The code perfectly addresses the request to turn on track changes using the correct and efficient Office.js API method toggleTrackChanges, with no flaws or areas for improvement. |
| L1-track-changes-toggle-002 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using a more specific and best-practice approach, such as using Word.run to execute the acceptAllRevisions method, and not providing any error handling or loading the document before syncing. |
| L1-toc-generate-001 | 83 | ✓ | The code correctly uses the insertTableOfContents API and addresses the user's request, but loses points for not fully utilizing best practices and available helpers, and for not perfectly handling potential edge cases in the generated code. |
| L1-section-break-001 | 72 | ✓ | The code uses correct APIs and is likely to work, but it lacks completeness in addressing the request to switch the next section to landscape, as it doesn't specify the section to apply the orientation to, and the approach could be improved by utilizing Word's built-in helpers for section management. |
| L1-header-footer-001 | 80 | ✓ | The code is mostly correct but loses points for not using the Word.run method to execute the API call and for not fully addressing potential edge cases, such as handling multiple sections or existing headers. |
| L1-header-footer-002 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-columns-001 | 72 | ✓ | The code correctly uses real Word Office.js APIs and would likely execute without runtime errors, but it lacks completeness in addressing the user's request to switch the rest of the document into a two-column layout, and does not use best practices or available helpers. |
| L1-image-insert-001 | 72 | ✓ | The code uses the correct insertImage API but loses points for not specifying the exact paragraph where the image should be inserted, and for not using best practices such as checking if the paragraph exists before inserting the image. |
| L1-image-insert-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-watermark-insert-001 | 0 ⚠️ | ✓ | The code hallucinates a non-existent "insertWatermark" method, which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-content-control-001 | 83 | ✓ | The code is mostly correct but loses points for not using the Word.run method to execute the content control insertion, and for not fully following best practices by directly calling insertContentControl instead of using a helper function if available. |
| L1-content-control-002 | 83 | ✓ | The code correctly uses the Office.js API to insert a checkbox content control, but loses points for not handling potential errors and not using a more robust approach to ensure the content control is inserted at the correct location. |
| L1-mail-merge-001 | 83 | ✓ | The code correctly uses the mailMergeReplace API and addresses the user's request, but loses points for not using the most efficient or best-practice approach, such as handling potential errors or using more specific typing for the replace operation. |
| L1-template-apply-001 | 80 | ✓ | The code correctly applies the resume template with the provided details, but loses points for not using the Word.run method to ensure the context is loaded before applying the template, and for not following best practices in terms of error handling and code organization. |
| L1-template-apply-002 | 85 | ✓ | The code correctly applies the memo template with the specified fields, but loses points for API correctness due to the use of a potentially hallucinated `applyTemplate` method, and for approach due to not explicitly using the `Word.run` API to ensure context synchronization. |
| L1-document-generate-001 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient or best-practice approaches, such as potentially re-implementing table of contents logic instead of relying on built-in functionality, and not demonstrating a clear understanding of when to use context.sync(). |
| L1-theme-apply-001 | 10 ⚠️ | ✓ | The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a significant penalty in API correctness and overall score. |
| L1-theme-apply-002 | 0 ⚠️ | ✓ | The code uses a completely hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not a valid method, resulting in a total failure across all dimensions. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, and also does not use the correct API "designTheme" and "applyTheme" to modify the theme. |
| L1-citation-bibliography-001 | 20 ⚠️ | ✓ | The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-equation-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "insertEquation" which is not a real Word Office.js API, resulting in a complete failure across all dimensions. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the user's request, but loses points for not using a more robust method to handle equation insertion, such as checking the current selection or range, and for not following best practices in terms of error handling and code organization. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selection for clarity, but loses points for not considering the coach panel recommendation as an alternative and for minor clarity issues in the presentation of the code. |
| L1-read-query-001 | 40 ⚠️ | ✓ | The code fails to utilize real Word Office.js APIs, instead relying on a hallucinated `countWords()` method, which is not a valid API call. |
| L1-read-query-002 | 20 ⚠️ | ✓ | The code heavily hallucinates a non-existent `getReadability()` method, which is not a real Word Office.js API, resulting in a significant penalty for API correctness. |
| L1-read-query-003 | 0 ⚠️ | ✓ | The code does not use any real Word Office.js APIs and instead calls a non-existent function `listHeadings()`, which is a significant flaw in API correctness, completeness, and approach. |
| L1-margins-orientation-001 | 83 | ✓ | The code is mostly correct but loses points for not using the best practice of explicitly defining the unit of measurement and for not handling potential errors that may occur during the execution of the setMargins method. |
| L1-margins-orientation-002 | 80 | ✓ | The code is mostly correct but lacks consideration of the document's current state and potential existing content, and does not use the most idiomatic or efficient Word Office.js APIs, such as Word.run, to perform the page orientation switch. |
| L1-web-search-needed-001 | 67 | ✓ | The response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks clarity on how this information will be used to add a section to the document and does not provide any workarounds or suggestions. |
| L1-web-search-needed-002 | 69 | ✓ | The response accurately and clearly explains the Pythagorean theorem, but the provided code is unnecessary and unrelated to the user's question, which was a request for information rather than a request to insert an equation into a Word document. |
| L1-refusal-001 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but lacks a more detailed explanation of alternative automation methods and does not directly offer code or more specific guidance on using VBA. |
| L1-refusal-002 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not explicitly stating that it's a refusal due to API limitations and not offering more detailed alternatives or settings options. |
| L1-refusal-003 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js in creating charts, but loses points for not providing a more detailed or relevant workaround, such as suggesting how to generate the chart image from the given data. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-paragraph-spacing-001 | 0 ⚠️ | ✓ | The code hallucinates a non-existent `addParagraph` method and attempts to modify existing paragraphs by re-adding them, which is not a valid approach and would result in significant errors. |
| L1-edge-case-001 | 25 ⚠️ | ✓ | The code does not use the Word Office.js API to attempt to delete the paragraph, instead immediately throwing an error, and does not utilize any helper functions or guards against out-of-range paragraph indices. |
| L1-form-field-001 | 90 | ✓ | The code correctly uses the Word Office.js API to insert a dropdown content control, but loses points for not using a more specific title or tag that directly relates to the user's request for a "Yes / No / Maybe" dropdown. |
| L1-text-insert-001 | 80 | ✓ | The code uses a real Office.js API helper method `addParagraph` but loses points for not explicitly addressing the "at the end of the document" requirement and for not using the most precise method to ensure the text is inserted at the very end. |
| L1-text-edit-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure across all dimensions. |
| gen-L2-form-field-001 | 72 | ✓ | The code uses correct Office.js APIs and is likely to work, but it does not fully address the request as it inserts a date picker with the title "Sign Date" instead of placing it below the checkbox content control with the correct label, and it lacks best practices in terms of content control placement and styling. |
| gen-L2-form-field-002 | 72 | ✗ | The code uses the correct insertContentControl API but fails to fully address the request by not replacing the existing plain text "Name: _______" and not using the best practices for inserting content controls and building block gallery controls. |
| gen-L2-form-field-003 | 83 | ✗ | The code is mostly correct and complete, but loses points for not using the Word.run method to wrap the content control insertion and for not handling potential errors, and also for not using the best practices such as using a more specific method to insert the content control. |
| gen-L2-form-field-004 | 83 | ✗ | The code is mostly correct but loses points for not explicitly setting the content control as a repeating section and not fully addressing the request for a repeating section content control with a table. |

## ⚠️ Needs attention

**[L1-heading-insert-002]** score=0 — The code uses hallucinated methods "addHeading" and "addParagraph" which are not real Word Office.js APIs, resulting in a complete failure across all dimensions.

**[L1-paragraph-format-001]** score=0 — The code uses a hallucinated method `addParagraph` which does not exist in the Word Office.js API, and instead of modifying the existing paragraph, it attempts to add a new paragraph with the same text, resulting in a completely incorrect solution.
- missing patterns: `font.bold`, `font.size`

**[L1-list-bullet-001]** score=25 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice.

**[L1-list-multilevel-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the API correctness and best practices criteria.

**[L1-header-footer-002]** score=0 — The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-image-insert-002]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-watermark-insert-001]** score=0 — The code hallucinates a non-existent "insertWatermark" method, which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-theme-apply-001]** score=10 — The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a significant penalty in API correctness and overall score.

**[L1-theme-apply-002]** score=0 — The code uses a completely hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not a valid method, resulting in a total failure across all dimensions.

**[L1-theme-apply-003]** score=20 — The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, and also does not use the correct API "designTheme" and "applyTheme" to modify the theme.

**[L1-citation-bibliography-001]** score=20 — The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-equation-001]** score=0 — The code uses a hallucinated method "insertEquation" which is not a real Word Office.js API, resulting in a complete failure across all dimensions.

**[L1-read-query-001]** score=40 — The code fails to utilize real Word Office.js APIs, instead relying on a hallucinated `countWords()` method, which is not a valid API call.

**[L1-read-query-002]** score=20 — The code heavily hallucinates a non-existent `getReadability()` method, which is not a real Word Office.js API, resulting in a significant penalty for API correctness.

**[L1-read-query-003]** score=0 — The code does not use any real Word Office.js APIs and instead calls a non-existent function `listHeadings()`, which is a significant flaw in API correctness, completeness, and approach.

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-paragraph-spacing-001]** score=0 — The code hallucinates a non-existent `addParagraph` method and attempts to modify existing paragraphs by re-adding them, which is not a valid approach and would result in significant errors.

**[L1-edge-case-001]** score=25 — The code does not use the Word Office.js API to attempt to delete the paragraph, instead immediately throwing an error, and does not utilize any helper functions or guards against out-of-range paragraph indices.

**[L1-text-edit-001]** score=0 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure across all dimensions.
