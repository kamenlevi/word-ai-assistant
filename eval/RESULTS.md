# Word AI — Eval Results
**Last run:** 2026-05-21 11:29:50  
**Overall: 57.9/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 81.5/100 | — | 2 |
| paragraph-format | 44.0/100 | — | 2 |
| list-bullet | 0.0/100 | — | 1 |
| list-multilevel | 0.0/100 | — | 1 |
| style-apply | 72.0/100 | — | 1 |
| table-create | 83.0/100 | — | 1 |
| table-format | 72.0/100 | — | 1 |
| find-replace | 81.0/100 | — | 2 |
| find-replace-regex | 83.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 90.0/100 | — | 2 |
| toc-generate | 80.0/100 | — | 1 |
| section-break | 72.0/100 | — | 1 |
| header-footer | 36.0/100 | — | 2 |
| columns | 83.0/100 | — | 1 |
| image-insert | 53.5/100 | — | 2 |
| watermark-insert | 50.0/100 | — | 1 |
| content-control | 81.5/100 | — | 2 |
| mail-merge | 55.0/100 | — | 1 |
| template-apply | 10.0/100 | — | 2 |
| document-generate | 90.0/100 | — | 1 |
| theme-apply | 32.3/100 | — | 3 |
| citation-bibliography | 40.0/100 | — | 1 |
| equation | 41.5/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 40.0/100 | — | 3 |
| margins-orientation | 82.5/100 | — | 2 |
| web-search-needed | 77.0/100 | — | 2 |
| refusal | 75.7/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 17.0/100 | — | 1 |
| edge-cases | 5.0/100 | — | 1 |
| form-field | 77.8/100 | — | 5 |
| text-insert | 80.0/100 | — | 1 |
| text-edit | 0.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 83 | ✓ | The code uses a correct and complete approach to add a Heading 1, but lacks explicit use of Word.run and may not handle potential async execution issues, also the addHeading function is not a built-in Word Office.js API, it seems to be a custom helper function. |
| L1-heading-insert-002 | 80 | ✓ | The code is mostly correct but loses points for not checking the current document context and the position where the new heading and paragraph will be inserted, and for not using a more robust method to ensure the insertion point is correct. |
| L1-paragraph-format-001 | 88 | ✓ | The code is mostly correct and complete, but loses points for not using the available helpers and for potential issues with the load and sync order, which might cause runtime errors. |
| L1-paragraph-format-002 | 0 ⚠️ | ✓ | A:  |
| L1-list-bullet-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-list-multilevel-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the API correctness and best practices criteria. |
| L1-style-apply-001 | 72 | ✓ | The code uses a real Word Office.js API, applies the style to the correct paragraph, and would mostly work, but lacks best practices, such as using the `paragraph.styleBuiltIn` property directly, and assumes a 0-based index for paragraphs which may not always be the case. |
| L1-table-create-001 | 83 | ✓ | The code correctly uses the Word Office.js API to insert a table, but loses points for not fully addressing the request by not considering the document context and paragraph placement, and for not using the most efficient approach by directly using the `body.insertTable` method without specifying the insertion point. |
| L1-table-format-001 | 72 | ✓ | The code uses a non-standard method `styleTable` which is not a real Word Office.js API, and also lacks the use of best practices and available helpers, such as `table.styleBuiltIn`, which would be a more appropriate and standard way to apply a table style. |
| L1-find-replace-001 | 72 | ✓ | The code uses a non-existent `replaceText` method, which is not a real Word Office.js API, resulting in a significant deduction in API correctness and approach scores. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most optimal approach, such as utilizing the `range.font` properties directly after searching, and not handling potential errors that may occur during the execution of the `await context.sync()` calls. |
| L1-find-replace-regex-001 | 83 | ✓ | The code uses the correct Word Office.js API replaceText with matchWildcards, but loses points for not fully addressing potential edge cases and not using the most optimal approach for searching and replacing text in the document body. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a correct API method to insert a footnote, but lacks consideration for the current selection context, and does not follow best practices by directly using the insertFootnote method without ensuring the range is properly set to the current selection. |
| L1-comment-insert-001 | 72 | ✓ | The code uses a real Word Office.js API (insertComment) but lacks specificity in targeting the first occurrence of the text "$4.2M" and does not fully address the request by not ensuring the comment is added to the correct location within the document. |
| L1-track-changes-toggle-001 | 97 ✅ | ✓ | The code perfectly addresses the request using the correct API, but loses a few points for not being explicitly wrapped in a Word.run call, which is a best practice for ensuring the code executes correctly in the Word context. |
| L1-track-changes-toggle-002 | 83 | ✓ | The code is mostly correct but loses points for not using the Word.run() method to execute the acceptAllRevisions() action, which is a best practice, and also for not checking if there are any tracked revisions before accepting them. |
| L1-toc-generate-001 | 80 | ✓ | The code correctly uses the insertTableOfContents API and addresses the request, but loses points for not fully specifying the API call with the correct parameters for the given document context and not using best practices for handling the generated table of contents. |
| L1-section-break-001 | 72 | ✓ | The code uses correct APIs and is likely to work, but it lacks completeness in fully addressing the user's request to switch the next section to landscape, and it does not follow best practices by using available helpers, instead using generic methods like insertSectionBreak and setPageOrientation. |
| L1-header-footer-001 | 72 | ✓ | The code uses a non-existent `setHeader` method, which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-header-footer-002 | 0 ⚠️ | ✓ | The code uses a hallucinated method `addPageNumbers` which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-columns-001 | 83 | ✓ | The code is mostly correct but loses points for not using the Word.run context and for not checking if the document is already in a two-column layout before making the change, and also for not using best practices such as wrapping the code in a Word.run async context to ensure proper synchronization. |
| L1-image-insert-001 | 72 | ✓ | The code uses a non-standard method signature for inserting an image, lacking the context of the cursor position and paragraph, and does not utilize the Word.run() method to ensure asynchronous execution. |
| L1-image-insert-002 | 35 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness. |
| L1-watermark-insert-001 | 50 ⚠️ | ✓ | The code lacks completeness and does not fully address the request, as it only inserts a watermark without specifying the scope to every page, and also uses a non-standard method signature without utilizing the Word.run context. |
| L1-content-control-001 | 83 | ✓ | The code correctly uses the Office.js API to insert a content control, but loses points for not handling potential errors and not using the most descriptive variable names, and also for not fully addressing the request in terms of providing a complete date picker experience. |
| L1-content-control-002 | 80 | ✓ | The code is mostly correct but loses points for not specifying the exact location of the content control insertion, which could lead to unexpected behavior, and for not using the Word.run method to ensure the code is executed in the correct context. |
| L1-mail-merge-001 | 55 ⚠️ | ✓ | The code uses a hallucinated method "mailMergeReplace" which is not a real Word Office.js API, resulting in a significant penalty for API correctness. |
| L1-template-apply-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a severe penalty for API correctness and would not work at runtime. |
| L1-template-apply-002 | 0 ⚠️ | ✓ | The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions. |
| L1-document-generate-001 | 90 | ✓ | The code accurately uses real Word Office.js APIs, fully addresses the request, and would execute without runtime errors, but loses points for not using the most efficient approach, such as potentially using a loop to add sections instead of repeating similar code. |
| L1-theme-apply-001 | 5 ⚠️ | ✓ | The code heavily hallucinates non-existent methods "designTheme" and "applyTheme", which are not part of the real Word Office.js APIs, resulting in a significant penalty. |
| L1-theme-apply-002 | 72 | ✓ | The code uses the correct designTheme and applyTheme APIs, but lacks completeness in addressing the request, has some potential issues with execution, and does not fully follow best practices, particularly in terms of using available helpers for styling and content manipulation. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-citation-bibliography-001 | 40 ⚠️ | ✓ | The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a significant penalty for API correctness. |
| L1-equation-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "insertEquation" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the user's request, but loses points for not using a more robust method to handle equation insertion and potential errors, and for not fully following best practices. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selected text for clarity, but loses points for not considering the coach panel as an alternative and for using a somewhat generic phrase in the initial response. |
| L1-read-query-001 | 40 ⚠️ | ✓ | The code fails to utilize real Word Office.js APIs, instead relying on a hallucinated method called "countWords()" which does not exist in the Office.js API. |
| L1-read-query-002 | 40 ⚠️ | ✓ | The code hallucinates a non-existent `getReadability()` method, which is not a real Word Office.js API, resulting in a severe penalty for API correctness. |
| L1-read-query-003 | 40 ⚠️ | ✓ | The code fails to use real Word Office.js APIs, instead hallucinating a non-existent "listHeadings" method, which heavily penalizes the API correctness score. |
| L1-margins-orientation-001 | 85 | ✓ | The code correctly uses the Word Office.js API and fully addresses the request, but loses points for not using the most efficient approach and potentially having issues with the context.sync() call. |
| L1-margins-orientation-002 | 80 | ✓ | The code is mostly correct but lacks a Word.run context and does not handle potential errors, also it does not use the best practice of using Word built-in methods for setting page orientation, instead it uses a generic setPageOrientation method. |
| L1-web-search-needed-001 | 67 | ✓ | The response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks clarity on how this information will be used to add a section to the document, and does not offer any workarounds or Word-specific guidance. |
| L1-web-search-needed-002 | 87 | ✓ | The response accurately and clearly explains the Pythagorean theorem, directly addressing the question, but loses points for not providing additional workarounds or Word-specific caveats, and for including unnecessary code in the response. |
| L1-refusal-001 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not exploring alternative solutions or providing more detailed instructions for the "Save As" feature. |
| L1-refusal-002 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but lacks a more detailed explanation or alternative solutions, resulting in a moderate score. |
| L1-refusal-003 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js in creating dynamic charts, but loses points for not providing a more detailed workaround or alternative solution that directly addresses the user's request for an Excel-style bar chart with specific data. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions. |
| L1-paragraph-spacing-001 | 17 ⚠️ | ✓ | The code fails to address the user's request to set line spacing to 1.5 on every paragraph, instead attempting to add new paragraphs and insert text, resulting in a completely incorrect solution. |
| L1-edge-case-001 | 5 ⚠️ | ✓ | The code does not use any real Word Office.js APIs and does not attempt to delete the paragraph, instead simply stating that no action is needed, resulting in a complete lack of functionality. |
| L1-form-field-001 | 97 ✅ | ✓ | The code is nearly perfect, using the correct Office.js API to insert a dropdown content control, but loses a few points for not using a more descriptive title or tag, and not handling potential errors that may occur during execution. |
| L1-text-insert-001 | 80 | ✓ | The code uses a real Office.js API and is likely to work, but it does not fully address the request as it inserts a new paragraph instead of appending text to the end of the document, and it does not use the best practice of checking the current selection or range before inserting content. |
| L1-text-edit-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| gen-L2-form-field-001 | 72 | ✓ | The code uses correct Office.js APIs and is likely to work, but it does not fully address the request as it inserts a "Sign Date" label instead of no label for the date picker and does not ensure the date picker is inserted below the checkbox. |
| gen-L2-form-field-002 | 65 | ✗ | The code incorrectly inserts two rich text content controls instead of replacing the existing plain text "Name: _______" with a single rich text content control and adding a building block gallery control as requested. |
| gen-L2-form-field-003 | 83 | ✗ | The code is mostly correct and complete, but loses points for not using the best practices, such as not checking if the content control already exists before inserting a new one, and not handling potential errors that may occur during the execution of the code. |
| gen-L2-form-field-004 | 72 | ✗ | The code uses some correct APIs but lacks a repeating section content control and instead inserts a rich text content control, also it does not fully follow best practices and available helpers. |

## ⚠️ Needs attention

**[L1-paragraph-format-002]** score=0 — A: 

**[L1-list-bullet-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-list-multilevel-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the API correctness and best practices criteria.

**[L1-header-footer-002]** score=0 — The code uses a hallucinated method `addPageNumbers` which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-image-insert-002]** score=35 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness.

**[L1-watermark-insert-001]** score=50 — The code lacks completeness and does not fully address the request, as it only inserts a watermark without specifying the scope to every page, and also uses a non-standard method signature without utilizing the Word.run context.

**[L1-mail-merge-001]** score=55 — The code uses a hallucinated method "mailMergeReplace" which is not a real Word Office.js API, resulting in a significant penalty for API correctness.

**[L1-template-apply-001]** score=20 — The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a severe penalty for API correctness and would not work at runtime.

**[L1-template-apply-002]** score=0 — The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions.

**[L1-theme-apply-001]** score=5 — The code heavily hallucinates non-existent methods "designTheme" and "applyTheme", which are not part of the real Word Office.js APIs, resulting in a significant penalty.

**[L1-theme-apply-003]** score=20 — The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-citation-bibliography-001]** score=40 — The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a significant penalty for API correctness.

**[L1-equation-001]** score=0 — The code uses a hallucinated method "insertEquation" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions.

**[L1-read-query-001]** score=40 — The code fails to utilize real Word Office.js APIs, instead relying on a hallucinated method called "countWords()" which does not exist in the Office.js API.

**[L1-read-query-002]** score=40 — The code hallucinates a non-existent `getReadability()` method, which is not a real Word Office.js API, resulting in a severe penalty for API correctness.

**[L1-read-query-003]** score=40 — The code fails to use real Word Office.js APIs, instead hallucinating a non-existent "listHeadings" method, which heavily penalizes the API correctness score.

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions.

**[L1-paragraph-spacing-001]** score=17 — The code fails to address the user's request to set line spacing to 1.5 on every paragraph, instead attempting to add new paragraphs and insert text, resulting in a completely incorrect solution.

**[L1-edge-case-001]** score=5 — The code does not use any real Word Office.js APIs and does not attempt to delete the paragraph, instead simply stating that no action is needed, resulting in a complete lack of functionality.

**[L1-text-edit-001]** score=0 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.
