# Word AI — Eval Results
**Last run:** 2026-05-16 02:59:03  
**Overall: 57.8/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 78.5/100 | — | 2 |
| paragraph-format | 85.5/100 | — | 2 |
| list-bullet | 25.0/100 | — | 1 |
| list-multilevel | 18.0/100 | — | 1 |
| style-apply | 80.0/100 | — | 1 |
| table-create | 83.0/100 | — | 1 |
| table-format | 20.0/100 | — | 1 |
| find-replace | 54.0/100 | — | 2 |
| find-replace-regex | 83.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 65.0/100 | — | 1 |
| track-changes-toggle | 91.5/100 | — | 2 |
| toc-generate | 72.0/100 | — | 1 |
| section-break | 72.0/100 | — | 1 |
| header-footer | 50.0/100 | — | 2 |
| columns | 72.0/100 | — | 1 |
| image-insert | 46.0/100 | — | 2 |
| watermark-insert | 0.0/100 | — | 1 |
| content-control | 87.5/100 | — | 2 |
| mail-merge | 40.0/100 | — | 1 |
| template-apply | 40.0/100 | — | 2 |
| document-generate | 90.0/100 | — | 1 |
| theme-apply | 32.3/100 | — | 3 |
| citation-bibliography | 40.0/100 | — | 1 |
| equation | 50.5/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 20.0/100 | — | 3 |
| margins-orientation | 77.5/100 | — | 2 |
| web-search-needed | 68.5/100 | — | 2 |
| refusal | 75.7/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 72.0/100 | — | 1 |
| edge-cases | 40.0/100 | — | 1 |
| form-field | 68.8/100 | — | 5 |
| text-insert | 80.0/100 | — | 1 |
| text-edit | 0.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 85 | ✓ | The code correctly uses the addHeading API and addresses the request, but loses points for not using a more robust approach, such as checking if the document is empty before adding a heading, and not handling potential errors that may occur during execution. |
| L1-heading-insert-002 | 72 | ✓ | The code uses correct APIs and mostly addresses the request, but lacks best practices and may not correctly insert the paragraph below the heading due to missing range or context management. |
| L1-paragraph-format-001 | 88 | ✓ | The code accurately uses real Word Office.js APIs and fully addresses the request, but loses points for not using the most efficient approach, such as utilizing the `range.font` helper, and for potential issues with load and sync order. |
| L1-paragraph-format-002 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the best practice of applying a built-in style for the title and instead manually setting font size and alignment, and also for not checking if the first paragraph is indeed the title before modifying it. |
| L1-list-bullet-001 | 25 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-list-multilevel-001 | 18 ⚠️ | ✓ | A: 18
B: |
| L1-style-apply-001 | 80 | ✓ | The code uses a real Word Office.js API, applies the style correctly, and would likely execute without runtime errors, but loses points for not using the most straightforward helper method and assuming a 0-based index for paragraphs, which may not always be the case. |
| L1-table-create-001 | 83 | ✓ | The code correctly uses the Word Office.js API to insert a table, but loses points for not fully addressing the request by not handling potential errors and not using the most efficient approach by directly using the `body.insertTable` method with the provided parameters. |
| L1-table-format-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method `styleTable` which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice. |
| L1-find-replace-001 | 18 ⚠️ | ✓ | A: 18
B: |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient or best-practice approach, such as utilizing the `range.font` properties directly on the search results without explicitly loading the font property. |
| L1-find-replace-regex-001 | 83 | ✓ | The code correctly uses the Word Office.js API and addresses the request, but loses points for not handling potential errors and not using the most efficient approach, such as checking if the replacement is necessary before calling replaceText. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a correct API to insert a footnote, but lacks consideration for the current selection and paragraph context, and does not follow best practices for using available helpers. |
| L1-comment-insert-001 | 65 | ✓ | The code uses a correct API method (insertComment) but lacks proper context and range handling, which may lead to incorrect comment insertion or runtime errors, and does not fully utilize best practices for searching and inserting comments in the document. |
| L1-track-changes-toggle-001 | 100 ✅ | ✓ | The code perfectly addresses the request to turn on track changes using the correct and existing Word Office.js API method toggleTrackChanges, and follows best practices. |
| L1-track-changes-toggle-002 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using a more specific best practice approach, such as using Word.run to encapsulate the acceptAllRevisions call, and not providing any error handling or context validation. |
| L1-toc-generate-001 | 72 | ✓ | The code uses a correct API but lacks best |
| L1-section-break-001 | 72 | ✓ | The code uses correct APIs and is likely to work, but it lacks completeness in addressing the user's request to switch the next section to landscape, and it does not follow best practices by using available helpers, instead using generic methods like insertSectionBreak and setPageOrientation. |
| L1-header-footer-001 | 80 | ✓ | The code is mostly correct but lacks consideration for the specific section or page where the header should be applied, and it does not use the most appropriate built-in methods for header manipulation, which would be more in line with best practices. |
| L1-header-footer-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-columns-001 | 72 | ✓ | The code correctly uses real Word Office.js APIs and would likely execute without runtime errors, but it lacks completeness in addressing the user's request to switch the "rest of the document" to a two-column layout, and does not follow best practices by not utilizing available helpers for layout adjustments. |
| L1-image-insert-001 | 72 | ✓ | The code uses the correct insertImage API but loses points for not specifying the exact paragraph where the image should be inserted, and for not using best practices such as checking the context and loading the image before inserting it. |
| L1-image-insert-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would_it_work dimensions. |
| L1-watermark-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "insertWatermark" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-content-control-001 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the request, but loses points for not using the Word.run method to ensure the code is executed in the correct context and for not handling potential errors, and also for not providing a more robust solution by utilizing the full range of available helpers. |
| L1-content-control-002 | 92 | ✓ | The code is mostly correct and complete, but loses points for not specifying the location of the content control insertion, which might lead to unexpected behavior, and for not using the provided helpers to handle potential edge cases. |
| L1-mail-merge-001 | 40 ⚠️ | ✓ | The code uses a hallucinated method "mailMergeReplace" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and approach. |
| L1-template-apply-001 | 80 | ✓ | The code correctly applies the resume template with the provided information, but loses points for not using the Word.run method to execute the API calls and for not handling potential errors, and also for not using the most efficient approach by directly using the applyTemplate helper. |
| L1-template-apply-002 | 0 ⚠️ | ✓ | The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-document-generate-001 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach, such as potentially reusing the `addHeading` and `addParagraph` methods in a loop to generate the 5 sections, instead of writing them out individually. |
| L1-theme-apply-001 | 5 ⚠️ | ✓ | The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would not work in practice. |
| L1-theme-apply-002 | 72 | ✓ | The code uses the correct designTheme and applyTheme APIs, but lacks completeness in addressing the request, as it doesn't specify how the theme will be applied to the existing body content, and the approach could be improved by utilizing more specific helpers for styling the document. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code hallucinates a non-existent `tweakTheme` method, which is not a real Word Office.js API, resulting in a severe penalty for API correctness. |
| L1-citation-bibliography-001 | 40 ⚠️ | ✓ | The code fails to use real Word Office.js APIs, instead hallucinating non-existent methods `insertCitation` and `insertBibliography`, which heavily penalizes the API correctness score. |
| L1-equation-001 | 18 ⚠️ | ✓ | A: 18
B: |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the user's request, but loses points for not handling potential errors and not using more robust best practices, such as checking the current selection or cursor position before inserting the equation. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selection for clarity and directly addresses the question, but loses points for not considering alternative approaches, such as recommending the coach panel, and for minor clarity issues in the code comments. |
| L1-read-query-001 | 20 ⚠️ | ✓ | The code hallucinates a non-existent `countWords()` method, which is not a real Word Office.js API, and does not use the correct API to count words, such as `body.getText()` or `range.getText()`. |
| L1-read-query-002 | 40 ⚠️ | ✓ | The code heavily relies on a hallucinated method "getReadability()" which is not a real Word Office.js API, resulting in a significant penalty for API correctness. |
| L1-read-query-003 | 0 ⚠️ | ✓ | The code hallucinates a non-existent `listHeadings()` method, which is not a real Word Office.js API, and does not provide any actual implementation to list the headings in the document. |
| L1-margins-orientation-001 | 83 | ✓ | The code correctly sets page margins using the Office.js API, but loses points for not using a more explicit best practice approach and for assuming the context is already loaded without showing the surrounding Word.run or context initialization code. |
| L1-margins-orientation-002 | 72 | ✓ | The code uses a non-existent `setPageOrientation` method, which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-web-search-needed-001 | 60 | ✓ | The AI response correctly identifies the need for a web search to get the current Apple stock price, but lacks clarity in explaining why no code was generated and does not offer any workarounds or alternatives. |
| L1-web-search-needed-002 | 77 | ✓ | The response accurately and clearly explains the Pythagorean theorem, directly addressing the question, but loses points for providing a simplistic equation insertion as code, which, although relevant, doesn't fully utilize the potential for educational or explanatory code snippets. |
| L1-refusal-001 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not exploring alternative solutions or providing more detailed information about the "Save As" feature. |
| L1-refusal-002 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but lacks a more detailed explanation or alternative solutions, resulting in moderate scores across the board. |
| L1-refusal-003 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js in creating Excel-style charts, but loses points for not providing a more detailed or relevant workaround, such as suggesting how to generate the chart image or alternative methods to represent the data. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions. |
| L1-paragraph-spacing-001 | 72 | ✓ | The code correctly uses Office.js APIs and addresses the user's request, but it lacks best practices, such as using a more efficient loop or built-in methods, and applies an unnecessary style to each paragraph before setting the line spacing. |
| L1-edge-case-001 | 40 ⚠️ | ✓ | The code does not use any real Word Office.js APIs, instead directly throwing an error, which loses all points for API correctness and approach, but partially addresses the request and would execute without runtime errors. |
| L1-form-field-001 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the user's request, but loses points for not using the most optimal approach and not handling potential errors or edge cases. |
| L1-text-insert-001 | 80 | ✓ | The code uses a real Word Office.js API (addParagraph) but loses points for not explicitly addressing the "at the end of the document" requirement and for not using the most precise method to insert text, which would be to get the last paragraph and insert text at its end. |
| L1-text-edit-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| gen-L2-form-field-001 | 72 | ✓ | The code correctly uses the insertContentControl API, but loses points for not fully addressing the request by including the label "I agree" as the title of the checkbox and not ensuring the date picker is inserted below the checkbox, and also for not using best practices in terms of content control placement and styling. |
| gen-L2-form-field-002 | 72 | ✗ | The code uses correct Office.js APIs and is mostly complete, but it lacks proper handling of the existing plain text replacement and does not fully follow best practices, such as using the `insertContentControl` method correctly and not addressing the removal of the original text. |
| gen-L2-form-field-003 | 45 ⚠️ | ✗ | The code fails to use the real Word Office.js API, specifically the `insertContentControl` method is not a valid API, and instead, it should use `range.insertContentControl` or `body.insertContentControl` with the correct parameters to insert a group content control containing two dropdown content controls. |
| gen-L2-form-field-004 | 72 | ✗ | The code partially addresses the request but lacks a repeating section content control and instead inserts a rich text content control, also it does not fully utilize the available helpers and best practices. |

## ⚠️ Needs attention

**[L1-list-bullet-001]** score=25 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-list-multilevel-001]** score=18 — A: 18
B:

**[L1-table-format-001]** score=20 — The code uses a hallucinated method `styleTable` which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice.

**[L1-find-replace-001]** score=18 — A: 18
B:

**[L1-header-footer-002]** score=20 — The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-image-insert-002]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would_it_work dimensions.

**[L1-watermark-insert-001]** score=0 — The code uses a hallucinated method "insertWatermark" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-mail-merge-001]** score=40 — The code uses a hallucinated method "mailMergeReplace" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and approach.

**[L1-template-apply-002]** score=0 — The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-theme-apply-001]** score=5 — The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would not work in practice.

**[L1-theme-apply-003]** score=20 — The code hallucinates a non-existent `tweakTheme` method, which is not a real Word Office.js API, resulting in a severe penalty for API correctness.

**[L1-citation-bibliography-001]** score=40 — The code fails to use real Word Office.js APIs, instead hallucinating non-existent methods `insertCitation` and `insertBibliography`, which heavily penalizes the API correctness score.

**[L1-equation-001]** score=18 — A: 18
B:

**[L1-read-query-001]** score=20 — The code hallucinates a non-existent `countWords()` method, which is not a real Word Office.js API, and does not use the correct API to count words, such as `body.getText()` or `range.getText()`.

**[L1-read-query-002]** score=40 — The code heavily relies on a hallucinated method "getReadability()" which is not a real Word Office.js API, resulting in a significant penalty for API correctness.

**[L1-read-query-003]** score=0 — The code hallucinates a non-existent `listHeadings()` method, which is not a real Word Office.js API, and does not provide any actual implementation to list the headings in the document.

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions.

**[L1-edge-case-001]** score=40 — The code does not use any real Word Office.js APIs, instead directly throwing an error, which loses all points for API correctness and approach, but partially addresses the request and would execute without runtime errors.

**[L1-text-edit-001]** score=0 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[gen-L2-form-field-003]** score=45 — The code fails to use the real Word Office.js API, specifically the `insertContentControl` method is not a valid API, and instead, it should use `range.insertContentControl` or `body.insertContentControl` with the correct parameters to insert a group content control containing two dropdown content controls.
- missing patterns: `GroupContentControl`
