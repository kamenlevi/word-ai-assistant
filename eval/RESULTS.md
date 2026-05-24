# Word AI — Eval Results
**Last run:** 2026-05-24 17:07:40  
**Overall: 58.8/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 78.5/100 | — | 2 |
| paragraph-format | 85.5/100 | — | 2 |
| list-bullet | 0.0/100 | — | 1 |
| list-multilevel | 72.0/100 | — | 1 |
| style-apply | 80.0/100 | — | 1 |
| table-create ✓ | 95.0/100 | — | 1 |
| table-format | 72.0/100 | — | 1 |
| find-replace | 81.0/100 | — | 2 |
| find-replace-regex | 83.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 88.5/100 | — | 2 |
| toc-generate | 72.0/100 | — | 1 |
| section-break | 80.0/100 | — | 1 |
| header-footer | 72.0/100 | — | 2 |
| columns | 72.0/100 | — | 1 |
| image-insert | 46.0/100 | — | 2 |
| watermark-insert | 0.0/100 | — | 1 |
| content-control | 41.5/100 | — | 2 |
| mail-merge | 72.0/100 | — | 1 |
| template-apply | 84.0/100 | — | 2 |
| document-generate | 90.0/100 | — | 1 |
| theme-apply | 15.0/100 | — | 3 |
| citation-bibliography | 20.0/100 | — | 1 |
| equation | 77.5/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 20.0/100 | — | 3 |
| margins-orientation | 44.0/100 | — | 2 |
| web-search-needed | 70.0/100 | — | 2 |
| refusal | 75.7/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 0.0/100 | — | 1 |
| edge-cases | 20.0/100 | — | 1 |
| form-field | 71.2/100 | — | 5 |
| text-insert | 72.0/100 | — | 1 |
| text-edit | 0.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 85 | ✓ | The code is mostly correct but loses points for using a hypothetical `addHeading` method which is not a real Word Office.js API, instead of using the actual `paragraph.styleBuiltIn` and `paragraph.insertText` methods. |
| L1-heading-insert-002 | 72 | ✓ | The code uses a hallucinated method addHeading and addParagraph which are not real Word Office.js APIs, losing significant points in API correctness and approach dimensions. |
| L1-paragraph-format-001 | 88 | ✓ | The code is mostly correct and complete, but loses points for not using the best practice of wrapping the code in a Word.run function and for not handling potential errors, and also for not using a more specific font size setting method if available. |
| L1-paragraph-format-002 | 83 | ✓ | The code is mostly correct but loses points for not using the built-in `applyStyle` method to set the font size and for not checking if the first paragraph is indeed the title before applying the changes. |
| L1-list-bullet-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-list-multilevel-001 | 72 | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-style-apply-001 | 80 | ✓ | The code uses a real Word Office.js API, applies the style correctly, and would likely execute without runtime errors, but loses points for not using the most straightforward helper method and assuming a 0-based index for paragraphs without explicit validation. |
| L1-table-create-001 | 95 ✅ | ✓ | The code accurately uses the Word Office.js API to insert a table, but loses points for completeness due to not fully addressing potential paragraph or table styling requests, and for approach due to not utilizing available helpers for potential further table or paragraph formatting. |
| L1-table-format-001 | 72 | ✓ | The code uses a non-standard method `styleTable` which is not a real Word Office.js API, and also lacks proper error handling and context loading, but it attempts to address the user's request and might work in some scenarios. |
| L1-find-replace-001 | 72 | ✓ | The code uses a non-existent `replaceText` method, which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach, such as utilizing the `body.search` method with wildcards or regular expressions more effectively, and not handling potential errors that may occur during the execution of the `await context.sync()` calls. |
| L1-find-replace-regex-001 | 83 | ✓ | The code uses the correct Word Office.js API and mostly addresses the request, but loses points for not handling potential errors and not using the most efficient approach, such as checking if the replacement is necessary before calling replaceText. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a correct API method to insert a footnote, but lacks consideration for the current selection and paragraph context, and does not fully follow best practices for using available helpers. |
| L1-comment-insert-001 | 72 | ✓ | The code uses a correct API method `insertComment` but lacks proper context and range handling, and does not fully address the request as it does not specify the exact location of the comment, which might lead to incorrect placement. |
| L1-track-changes-toggle-001 | 97 ✅ | ✓ | The code perfectly addresses the request using the correct API, but loses a few points for not explicitly handling potential errors that might occur during the execution of the toggleTrackChanges method. |
| L1-track-changes-toggle-002 | 80 | ✓ | The code correctly uses the Word Office.js API and addresses the request, but lacks error handling and does not utilize the `Word.run` method to encapsulate the `acceptAllRevisions` call, which is a best practice. |
| L1-toc-generate-001 | 72 | ✓ | The code uses a correct API but lacks best practices, such as using the provided helpers, and the insertTableOfContents method is not a standard Word Office.js API, which reduces the score. |
| L1-section-break-001 | 80 | ✓ | The code is mostly correct but lacks consideration for the current section and paragraph context, and does not fully follow best practices for inserting section breaks and setting page orientation in a Word document using Office.js APIs. |
| L1-header-footer-001 | 72 | ✓ | The code uses a non-existent `setHeader` method, which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-header-footer-002 | 72 | ✓ | The code uses a non-standard method `addPageNumbers` which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-columns-001 | 72 | ✓ | The code correctly uses real Word Office.js APIs and would likely execute without runtime errors, but it lacks completeness in addressing the user's request to switch the rest of the document into a two-column layout, and does not use best practices or available helpers. |
| L1-image-insert-001 | 72 | ✓ | The code uses the correct insertImage API but loses points for not specifying the exact paragraph where the image should be inserted, and for not using best practices such as checking if the paragraph exists before inserting the image. |
| L1-image-insert-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-watermark-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "insertWatermark" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions. |
| L1-content-control-001 | 0 ⚠️ | ✓ | A:  |
| L1-content-control-002 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the user's request, but loses points for not using the Word.run method to ensure the code is executed in the correct context and for not providing any error handling or loading checks. |
| L1-mail-merge-001 | 72 | ✓ | The code uses a hallucinated method mailMergeReplace, which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-template-apply-001 | 83 | ✓ | The code correctly applies the resume template with the provided information, but loses points for not using the Word.run method to ensure the context is loaded before applying the template, and for not following best practices in terms of error handling and variable scope. |
| L1-template-apply-002 | 85 | ✓ | The code correctly applies the memo template with the requested fields, but loses points for API correctness due to the use of a potentially non-existent `applyTemplate` method, and for approach due to not using the official Word.run method to execute the template application. |
| L1-document-generate-001 | 90 | ✓ | The code accurately uses real Word Office.js APIs, fully addresses the request, and would execute without runtime errors, but loses points for not using the most optimal approach, such as potentially using a loop to generate the sections instead of repeating similar code. |
| L1-theme-apply-001 | 5 ⚠️ | ✓ | The code heavily hallucinates non-existent methods "designTheme" and "applyTheme", which are not part of the real Word Office.js APIs, resulting in a significant penalty. |
| L1-theme-apply-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would_it_work dimensions. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, and also does not use the correct API "designTheme" and "applyTheme" to modify the theme. |
| L1-citation-bibliography-001 | 20 ⚠️ | ✓ | The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-equation-001 | 72 | ✓ | The code uses a correct API method to insert an equation, but lacks consideration for the cursor position and paragraph context, and does not utilize the best practice helper methods provided by the Office.js API. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the request, but loses points for not using a more robust method to handle equation insertion, such as utilizing the range object to specify the exact location of the equation. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selection for clarity, is mostly clear and relevant, but loses points for not considering alternative approaches, such as recommending the coach panel, and for not providing additional context or workarounds. |
| L1-read-query-001 | 20 ⚠️ | ✓ | The code fails to utilize any real Word Office.js APIs, instead relying on a hallucinated `countWords()` method, which is not a valid API call. |
| L1-read-query-002 | 40 ⚠️ | ✓ | The code heavily relies on a hallucinated method "getReadability()" which is not a real Word Office.js API, resulting in a significant penalty in API correctness. |
| L1-read-query-003 | 0 ⚠️ | ✓ | The code uses a hallucinated method "listHeadings()" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-margins-orientation-001 | 88 | ✓ | The code correctly uses the Word Office.js API and fully addresses the request, but loses points for approach due to using a manual method to set margins instead of utilizing available helpers or built-in margin settings. |
| L1-margins-orientation-002 | 0 ⚠️ | ✓ | The code uses a hallucinated method "setPageOrientation" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-web-search-needed-001 | 67 | ✓ | The response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks clarity on how this information will be used to add a section to the Word document and does not provide any workarounds or suggestions. |
| L1-web-search-needed-002 | 73 | ✓ | The response accurately and clearly explains the Pythagorean theorem, directly addressing the question, but loses points for including unnecessary code and not providing any workarounds or Word-specific caveats, resulting in a low bonus score. |
| L1-refusal-001 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a workaround, but lacks a clear explanation of alternative APIs or tools that could be used for automation, deducting points from accuracy and bonus scores. |
| L1-refusal-002 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not offering more detailed alternatives or caveats, and for a slightly abrupt tone in the "NO CODE WAS GENERATED" statement. |
| L1-refusal-003 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a workaround, but the workaround is incomplete and doesn't directly address the user's request for a real Excel-style bar chart from the provided data. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-paragraph-spacing-001 | 0 ⚠️ | ✓ | The code hallucinates a non-existent `addParagraph` method, which is not a real Word Office.js API, and instead of modifying the existing paragraphs, it attempts to add new ones with the same text, resulting in a completely incorrect implementation. |
| L1-edge-case-001 | 20 ⚠️ | ✓ | The code does not use any real Word Office.js APIs, and instead simply includes a comment stating that no action will be taken, which results in a significant deduction in API correctness and would not execute without runtime errors. |
| L1-form-field-001 | 90 | ✓ | The code correctly uses the Word Office.js API to insert a dropdown content control, but loses points for not fully following best practices, such as not checking if the cursor is in a valid location to insert the content control. |
| L1-text-insert-001 | 72 | ✓ | The code uses a simplified `addParagraph` method which is not a standard Word Office.js API, and also lacks consideration for inserting the text at the very end of the document, which might require navigating to the last paragraph and inserting the text there, rather than simply adding a new paragraph. |
| L1-text-edit-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions. |
| gen-L2-form-field-001 | 72 | ✓ | The code uses a non-existent `insertContentControl` method with incorrect parameters, such as `title` instead of `label`, and does not fully address the request by not specifying the correct label for the date picker content control. |
| gen-L2-form-field-002 | 50 ⚠️ | ✗ | The code incorrectly inserts two rich text content controls instead of replacing the existing plain text "Name: _______" with a single rich text content control and adding a building block gallery control as requested. |
| gen-L2-form-field-003 | 72 | ✗ | The code uses a non-existent `insertContentControl` method with incorrect parameters, and also fails to actually group the two dropdown content controls as requested, losing points for API correctness, completeness, and approach. |
| gen-L2-form-field-004 | 72 | ✗ | The code partially addresses the request but fails to fully implement a repeating section content control and instead uses a rich text content control, resulting in lost points across multiple dimensions. |

## ⚠️ Needs attention

**[L1-list-bullet-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-image-insert-002]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-watermark-insert-001]** score=0 — The code uses a hallucinated method "insertWatermark" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions.

**[L1-content-control-001]** score=0 — A: 

**[L1-theme-apply-001]** score=5 — The code heavily hallucinates non-existent methods "designTheme" and "applyTheme", which are not part of the real Word Office.js APIs, resulting in a significant penalty.

**[L1-theme-apply-002]** score=20 — The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would_it_work dimensions.

**[L1-theme-apply-003]** score=20 — The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, and also does not use the correct API "designTheme" and "applyTheme" to modify the theme.

**[L1-citation-bibliography-001]** score=20 — The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-read-query-001]** score=20 — The code fails to utilize any real Word Office.js APIs, instead relying on a hallucinated `countWords()` method, which is not a valid API call.

**[L1-read-query-002]** score=40 — The code heavily relies on a hallucinated method "getReadability()" which is not a real Word Office.js API, resulting in a significant penalty in API correctness.

**[L1-read-query-003]** score=0 — The code uses a hallucinated method "listHeadings()" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-margins-orientation-002]** score=0 — The code uses a hallucinated method "setPageOrientation" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-paragraph-spacing-001]** score=0 — The code hallucinates a non-existent `addParagraph` method, which is not a real Word Office.js API, and instead of modifying the existing paragraphs, it attempts to add new ones with the same text, resulting in a completely incorrect implementation.

**[L1-edge-case-001]** score=20 — The code does not use any real Word Office.js APIs, and instead simply includes a comment stating that no action will be taken, which results in a significant deduction in API correctness and would not execute without runtime errors.

**[L1-text-edit-001]** score=0 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions.

**[gen-L2-form-field-002]** score=50 — The code incorrectly inserts two rich text content controls instead of replacing the existing plain text "Name: _______" with a single rich text content control and adding a building block gallery control as requested.
- missing patterns: `RichTextContentControl`
