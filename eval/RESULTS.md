# Word AI — Eval Results
**Last run:** 2026-05-18 18:07:04  
**Overall: 56.4/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 0.0/100 | — | 2 |
| paragraph-format | 41.5/100 | — | 2 |
| list-bullet | 25.0/100 | — | 1 |
| list-multilevel | 0.0/100 | — | 1 |
| style-apply | 80.0/100 | — | 1 |
| table-create | 89.0/100 | — | 1 |
| table-format | 20.0/100 | — | 1 |
| find-replace | 85.0/100 | — | 2 |
| find-replace-regex | 83.0/100 | — | 1 |
| footnote-insert | 54.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 92.5/100 | — | 2 |
| toc-generate | 72.0/100 | — | 1 |
| section-break | 72.0/100 | — | 1 |
| header-footer | 77.5/100 | — | 2 |
| columns | 80.0/100 | — | 1 |
| image-insert | 46.0/100 | — | 2 |
| watermark-insert | 40.0/100 | — | 1 |
| content-control | 83.0/100 | — | 2 |
| mail-merge | 55.0/100 | — | 1 |
| template-apply | 83.0/100 | — | 2 |
| document-generate | 85.0/100 | — | 1 |
| theme-apply | 16.7/100 | — | 3 |
| citation-bibliography | 20.0/100 | — | 1 |
| equation | 41.5/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 40.0/100 | — | 3 |
| margins-orientation | 81.5/100 | — | 2 |
| web-search-needed | 38.5/100 | — | 2 |
| refusal | 73.3/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 72.0/100 | — | 1 |
| edge-cases | 35.0/100 | — | 1 |
| form-field | 73.4/100 | — | 5 |
| text-insert | 80.0/100 | — | 1 |
| text-edit | 0.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addHeading" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions. |
| L1-heading-insert-002 | 0 ⚠️ | ✓ | The code uses hallucinated methods "addHeading" and "addParagraph" which are not real Word Office.js APIs, resulting in a complete failure across all dimensions. |
| L1-paragraph-format-001 | 0 ⚠️ | ✗ | The code attempts to add a new paragraph with the specified style instead of modifying the existing first paragraph, and uses a non-existent `addParagraph` method, indicating a severe lack of understanding of the Word Office.js API. |
| L1-paragraph-format-002 | 83 | ✓ | The code is mostly correct but loses points for not using the built-in `applyStyle` method to set the font size and for not checking if the first paragraph is indeed the title before applying the style. |
| L1-list-bullet-001 | 25 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice. |
| L1-list-multilevel-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-style-apply-001 | 80 | ✓ | The code correctly applies the Quote style to the third paragraph, but uses a generic applyStyle method which may not be the most efficient or Word-specific approach, and assumes the paragraph index is 2, which may not always be the case. |
| L1-table-create-001 | 89 | ✓ | The code correctly uses the Word Office.js API to insert a table, but loses points for not fully addressing the request by not handling potential errors and not using the most efficient approach by directly using the `body.insertTable` method with the provided parameters. |
| L1-table-format-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method `styleTable` which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-find-replace-001 | 80 | ✓ | The code uses a correct and existing API method replaceText, but lacks the surrounding Word.run context and may not fully handle all edge cases, such as multiple occurrences or case sensitivity, and does not follow best practices for error handling and loading the document. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not handling potential errors and not using the most efficient approach, such as using `body.search` with `matchWildcards` instead of a simple string search, and not checking if the search results are valid before trying to access their font properties. |
| L1-find-replace-regex-001 | 83 | ✓ | The code correctly uses the Word Office.js API and addresses the request, but loses points for not handling potential errors and not using the most efficient approach, such as checking if the replacement is necessary before calling replaceText. |
| L1-footnote-insert-001 | 54 ⚠️ | ✓ | The code uses a correct method to insert a footnote, but API consideration for the current lacks selection and does not utilize helpers, such as specifying the range for the provided. |
| L1-comment-insert-001 | 72 | ✓ | The code uses a correct API method `insertComment` but lacks specificity in targeting the first occurrence of the text "$4.2M" and does not utilize Word's built-in search functionality with wildcards or regular expressions as needed for precise text matching. |
| L1-track-changes-toggle-001 | 100 ✅ | ✓ | The code perfectly addresses the request to turn on track changes using the correct and available Word Office.js API, specifically the toggleTrackChanges method, with no flaws or issues identified. |
| L1-track-changes-toggle-002 | 85 | ✓ | The code is mostly correct but loses points for API correctness and approach due to the use of a generic method name "acceptAllRevisions" which is not a real Word Office.js API, instead it should utilize the "doc.changeTrackingMode" or similar APIs to achieve the desired result. |
| L1-toc-generate-001 | 72 | ✓ | The code uses a correct API but loses points for not fully utilizing best practices, such as using the provided helpers, and for not addressing potential issues with the context and variable scope. |
| L1-section-break-001 | 72 | ✓ | The code uses correct APIs and is likely to work, but it lacks completeness in fully addressing the request, as it doesn't specify the section to apply the landscape orientation to, and the approach could be improved by utilizing available helpers for more robust and maintainable code. |
| L1-header-footer-001 | 83 | ✓ | The code is mostly correct but loses points for not using the Word.run method to execute the operation and for not handling potential errors, and also for not using the best practice of using a more specific method like inserting a header paragraph with the desired text instead of a generic setHeader method. |
| L1-header-footer-002 | 72 | ✓ | The code uses a non-standard method `addPageNumbers` which is not a real Word Office.js API, and also lacks best practices and available helpers, but otherwise addresses the request and would likely execute without runtime errors. |
| L1-columns-001 | 80 | ✓ | The code correctly uses real Word Office.js APIs and addresses the request, but loses points for not fully handling potential edge cases and not using the most optimal approach, such as checking the current section layout before inserting a new section break. |
| L1-image-insert-001 | 72 | ✓ | The code uses a non-existent `insertImage` method, which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-image-insert-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-watermark-insert-001 | 40 ⚠️ | ✓ | The code is incomplete and does not use the correct Word Office.js APIs, as it hallucinates a non-existent `insertWatermark` method instead of using the available APIs to achieve the desired result. |
| L1-content-control-001 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the user's request, but loses points for not using a more specific best practice approach and potentially lacking error handling or loading checks. |
| L1-content-control-002 | 83 | ✓ | The code correctly uses the Office.js API to insert a checkbox content control, but loses points for not handling potential errors and not using the most efficient approach, such as checking if the content control already exists before inserting a new one. |
| L1-mail-merge-001 | 55 ⚠️ | ✓ | The code fails to use real Word Office.js APIs, specifically the mailMergeReplace method is not a valid API, resulting in a significant penalty for API correctness. |
| L1-template-apply-001 | 83 | ✓ | The code correctly applies the resume template with the provided details, but loses points for not using the Word.run method to ensure the context is properly loaded and for not handling potential errors, and also for not fully utilizing best practices and available helpers. |
| L1-template-apply-002 | 83 | ✓ | The code correctly applies the memo template with the specified fields, but loses points for not using the most efficient or best practice approach, such as utilizing Word's built-in template application methods or handling potential errors, and also for not fully addressing potential edge cases or providing additional functionality. |
| L1-document-generate-001 | 85 | ✓ | The code correctly uses real Word Office.js APIs and addresses the request, but loses points for not using the most efficient approach and not handling potential errors or edge cases, such as checking if the headings and paragraphs already exist before inserting them. |
| L1-theme-apply-001 | 10 ⚠️ | ✓ | The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also lacks any actual implementation to address the user's request to apply the theme to the document. |
| L1-theme-apply-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would_it_work dimensions. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code hallucinates a non-existent `tweakTheme` method and fails to use the actual Word Office.js APIs, such as `designTheme` and `applyTheme`, to achieve the desired theme modification. |
| L1-citation-bibliography-001 | 20 ⚠️ | ✓ | The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would not work at runtime. |
| L1-equation-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "Equation" which is not a real Word Office.js APIinsert, resulting in a complete failure to meet. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the user's request, but loses points for not handling potential errors and not using more robust best practices, such as checking the cursor position before inserting the equation. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selected text for clarity, but loses points for not considering the coach panel as an alternative and for using a somewhat generic phrase "Rewriting the current selection for clarity" that doesn't add much value. |
| L1-read-query-001 | 40 ⚠️ | ✓ | The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `countWords()` method, which heavily penalizes its API correctness score. |
| L1-read-query-002 | 40 ⚠️ | ✓ | The code fails to utilize real Word Office.js APIs, instead relying on a hallucinated `getReadability()` method, which is not a valid API call. |
| L1-read-query-003 | 40 ⚠️ | ✓ | The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `listHeadings()` method, which heavily penalizes the API correctness score. |
| L1-margins-orientation-001 | 83 | ✓ | The code correctly uses the Word Office.js API and addresses the user's request, but loses points for not using the most efficient or best practice approach, such as using a helper function if available, and assumes the context is already loaded without explicitly showing the loading process. |
| L1-margins-orientation-002 | 80 | ✓ | The code is mostly correct but lacks consideration for the document's current state and potential errors, and does not use the most idiomatic or efficient Office.js APIs for the task, such as Word.run to ensure the operation is executed as a single transaction. |
| L1-web-search-needed-001 | 0 ⚠️ | ✓ | A:  |
| L1-web-search-needed-002 | 77 | ✓ | The response accurately and clearly explains the Pythagorean theorem, directly addressing the question, but lacks significant bonus points for not providing additional workarounds, settings options, or Word-specific caveats beyond the basic explanation and equation insertion. |
| L1-refusal-001 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not offering more detailed instructions or alternatives, and for a slightly vague explanation of the API's capabilities. |
| L1-refusal-002 | 72 | ✓ | The response accurately conveys the limitation of Word Office.js regarding direct printing, but loses points for not providing a more detailed explanation or alternative solutions beyond the basic workaround. |
| L1-refusal-003 | 65 | ✓ | The response accurately identifies the limitation of Word Office.js in creating charts, but loses points for not providing a more detailed or relevant workaround, such as suggesting how to generate the chart image or using an alternative method like inserting an Excel object. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of the task. |
| L1-paragraph-spacing-001 | 72 | ✓ | The code correctly uses Office.js APIs and addresses the request, but has unnecessary `applyStyle` calls, lacks error handling, and doesn't utilize best practices such as using `body.paragraphs.forEach` instead of manual looping. |
| L1-edge-case-001 | 35 ⚠️ | ✓ | The code does not use any real Word Office.js APIs, instead directly throwing an error, which is not a correct approach to handle the request to delete a paragraph. |
| L1-form-field-001 | 83 | ✓ | The code correctly uses the insertContentControl API and fully addresses the request, but loses points for not using a more descriptive title and for potential issues with the load and sync order, as well as not following best practices for content control insertion. |
| L1-text-insert-001 | 80 | ✓ | The code uses a real Office.js API helper method addParagraph, but loses points for not fully addressing potential edge cases and not using the most precise method to insert text at the end of the document, which would be to get the last paragraph and insert text into its range. |
| L1-text-edit-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| gen-L2-form-field-001 | 72 | ✓ | The code uses correct Office.js APIs and is likely to work, but it does not fully address the request as it inserts a "Sign Date" label instead of no label for the date picker and does not ensure the date picker is inserted below the checkbox. |
| gen-L2-form-field-002 | 72 | ✗ | The code uses correct Office.js APIs and is mostly complete, but it lacks proper replacement of the existing plain text "Name: _______" and does not follow best practices for inserting content controls and building block galleries. |
| gen-L2-form-field-003 | 57 ⚠️ | ✗ | The code fails to insert a group content control containing the two dropdown content controls, instead inserting three separate content controls, and does not utilize the best practices and available helpers provided by the Word Office.js API. |
| gen-L2-form-field-004 | 83 | ✗ | The code is mostly correct but loses points for not explicitly using the Word.run method to execute the content control and table insertion, and for not fully addressing the request of adding a repeating section content control, instead inserting a regular content control. |

## ⚠️ Needs attention

**[L1-heading-insert-001]** score=0 — The code uses a hallucinated method "addHeading" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions.

**[L1-heading-insert-002]** score=0 — The code uses hallucinated methods "addHeading" and "addParagraph" which are not real Word Office.js APIs, resulting in a complete failure across all dimensions.

**[L1-paragraph-format-001]** score=0 — The code attempts to add a new paragraph with the specified style instead of modifying the existing first paragraph, and uses a non-existent `addParagraph` method, indicating a severe lack of understanding of the Word Office.js API.
- missing patterns: `font.bold`, `font.size`

**[L1-list-bullet-001]** score=25 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice.

**[L1-list-multilevel-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-table-format-001]** score=20 — The code uses a hallucinated method `styleTable` which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-footnote-insert-001]** score=54 — The code uses a correct method to insert a footnote, but API consideration for the current lacks selection and does not utilize helpers, such as specifying the range for the provided.

**[L1-image-insert-002]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-watermark-insert-001]** score=40 — The code is incomplete and does not use the correct Word Office.js APIs, as it hallucinates a non-existent `insertWatermark` method instead of using the available APIs to achieve the desired result.

**[L1-mail-merge-001]** score=55 — The code fails to use real Word Office.js APIs, specifically the mailMergeReplace method is not a valid API, resulting in a significant penalty for API correctness.

**[L1-theme-apply-001]** score=10 — The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also lacks any actual implementation to address the user's request to apply the theme to the document.

**[L1-theme-apply-002]** score=20 — The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would_it_work dimensions.

**[L1-theme-apply-003]** score=20 — The code hallucinates a non-existent `tweakTheme` method and fails to use the actual Word Office.js APIs, such as `designTheme` and `applyTheme`, to achieve the desired theme modification.

**[L1-citation-bibliography-001]** score=20 — The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would not work at runtime.

**[L1-equation-001]** score=0 — The code uses a hallucinated method "Equation" which is not a real Word Office.js APIinsert, resulting in a complete failure to meet.

**[L1-read-query-001]** score=40 — The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `countWords()` method, which heavily penalizes its API correctness score.

**[L1-read-query-002]** score=40 — The code fails to utilize real Word Office.js APIs, instead relying on a hallucinated `getReadability()` method, which is not a valid API call.

**[L1-read-query-003]** score=40 — The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `listHeadings()` method, which heavily penalizes the API correctness score.

**[L1-web-search-needed-001]** score=0 — A: 

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of the task.

**[L1-edge-case-001]** score=35 — The code does not use any real Word Office.js APIs, instead directly throwing an error, which is not a correct approach to handle the request to delete a paragraph.

**[L1-text-edit-001]** score=0 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[gen-L2-form-field-003]** score=57 — The code fails to insert a group content control containing the two dropdown content controls, instead inserting three separate content controls, and does not utilize the best practices and available helpers provided by the Word Office.js API.
- missing patterns: `GroupContentControl`
