# Word AI — Eval Results
**Last run:** 2026-05-15 15:53:35  
**Overall: 58.9/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 82.5/100 | — | 2 |
| paragraph-format | 56.5/100 | — | 2 |
| list-bullet | 72.0/100 | — | 1 |
| list-multilevel | 67.0/100 | — | 1 |
| style-apply | 83.0/100 | — | 1 |
| table-create | 20.0/100 | — | 1 |
| table-format | 0.0/100 | — | 1 |
| find-replace | 81.0/100 | — | 2 |
| find-replace-regex | 85.0/100 | — | 1 |
| footnote-insert | 0.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 54.0/100 | — | 2 |
| toc-generate | 67.0/100 | — | 1 |
| section-break | 72.0/100 | — | 1 |
| header-footer | 77.5/100 | — | 2 |
| columns | 85.0/100 | — | 1 |
| image-insert | 46.0/100 | — | 2 |
| watermark-insert | 60.0/100 | — | 1 |
| content-control | 90.0/100 | — | 2 |
| mail-merge | 20.0/100 | — | 1 |
| template-apply | 36.0/100 | — | 2 |
| document-generate | 83.0/100 | — | 1 |
| theme-apply | 32.3/100 | — | 3 |
| citation-bibliography | 40.0/100 | — | 1 |
| equation | 77.5/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 26.7/100 | — | 3 |
| margins-orientation | 82.5/100 | — | 2 |
| web-search-needed | 69.5/100 | — | 2 |
| refusal | 79.3/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 93.0/100 | — | 1 |
| edge-cases | 20.0/100 | — | 1 |
| form-field ✓ | 97.0/100 | — | 1 |
| text-insert | 83.0/100 | — | 1 |
| text-edit | 0.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 85 | ✓ | The code is mostly correct, but it uses a custom `addHeading` function which is not a built-in Word Office.js API, and its implementation is not provided, which raises concerns about its correctness and potential for errors. |
| L1-heading-insert-002 | 80 | ✓ | The code is mostly correct but loses points for not fully addressing the request context, such as handling the existing paragraph and the positioning of the new heading and paragraph, and not using the most optimal approach by potentially re-implementing existing helper functions. |
| L1-paragraph-format-001 | 93 | ✓ | The code is mostly correct and complete, but loses points for approach as it manually sets font properties instead of using the available helpers like applyStyle or range.font methods in a more structured way. |
| L1-paragraph-format-002 | 20 ⚠️ | ✓ | A: 20
B: |
| L1-list-bullet-001 | 72 | ✓ | The code uses a non-standard `addList` method which is not a real Word Office.js API, and lacks best practices, but otherwise addresses the request and would likely execute without runtime errors. |
| L1-list-multilevel-001 | 67 | ✓ | The code uses a non-standard method "addList" which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-style-apply-001 | 83 | ✓ | The code correctly applies the Quote style to the third paragraph, but loses points for not using the more specific and recommended paragraph.styleBuiltIn method and for potential indexing issues, as paragraph indexing starts at 0, so the third paragraph should be indexed as 2, which is correct in this case, but may not be robust for all document structures. |
| L1-table-create-001 | 20 ⚠️ | ✓ | A: 20
B: |
| L1-table-format-001 | 0 ⚠️ | ✓ | The code fails to apply the style to the existing table, instead attempting to insert a new table with the specified style, which does not address the user's request. |
| L1-find-replace-001 | 72 | ✓ | The code uses a correct API method replaceText, but loses points for not using the Word.run method to queue the changes and for not fully addressing the request by not checking if the replacement was successful, and also for not using best practices by directly calling replaceText without utilizing the available helpers. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach, such as using `range.font` helpers, and for not handling potential errors that may occur during the execution of the `await context.sync()` calls. |
| L1-find-replace-regex-001 | 85 | ✓ | The code correctly uses the Word Office.js API and would execute without runtime errors, but loses points for not fully addressing the request with a more precise wildcard pattern and for not using the most optimal approach by directly utilizing the replaceText method without considering potential edge cases. |
| L1-footnote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "insertFootnote" which is not a real Word Office.js API, instead it should use the "range.insertFootnote" method to insert a footnote at the current selection. |
| L1-comment-insert-001 | 72 | ✓ | The code uses a correct API method `insertComment` but lacks specificity in targeting the first occurrence of the text "$4.2M" and does not utilize Word's built-in search functionality with wildcards or regular expressions as needed for precise text matching. |
| L1-track-changes-toggle-001 | 25 ⚠️ | ✓ | A: 25
B: |
| L1-track-changes-toggle-002 | 83 | ✓ | The code is mostly correct but loses points for not using the Word.run method to wrap the API calls, and for not checking if there are actually tracked revisions to accept before calling acceptAllRevisions. |
| L1-toc-generate-001 | 67 | ✓ | The code uses a non-existent `insertTableOfContents` method, which is not a real Word Office.js API, resulting in significant penalties in API correctness and approach dimensions. |
| L1-section-break-001 | 72 | ✓ | The code partially addresses the request but lacks best practices and does not fully utilize available helpers, also it does not specify the section to apply the page orientation to, which might lead to incorrect results. |
| L1-header-footer-001 | 83 | ✓ | The code is mostly correct but loses points for not using the Word.run method and for not fully addressing potential edge cases, such as handling multiple sections or existing headers. |
| L1-header-footer-002 | 72 | ✗ | The code uses a non-standard method `setFooter` which is not a real Word Office.js API, and also lacks best practices and available helpers, but otherwise addresses the request and would likely execute without runtime errors. |
| L1-columns-001 | 85 | ✓ | The code is mostly correct and would work, but it lacks the use of best practices, such as utilizing the Word.run method to ensure context and loading, and does not account for potential edge cases, like the current selection or range. |
| L1-image-insert-001 | 72 | ✓ | The code uses a non-standard `insertImage` method which is not a real Word Office.js API, and also lacks proper error handling and context setup, but it attempts to address the request and has a correct width specification. |
| L1-image-insert-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method `insertImage` which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-watermark-insert-001 | 60 | ✓ | The code partially addresses the request but lacks specificity in applying the watermark to every page, and it does not utilize the Word Office.js API correctly for inserting watermarks, which would typically involve working with sections or the document body directly. |
| L1-content-control-001 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the request, but loses points for not using a more specific best practice approach and for potential issues with variable scope and load/sync order. |
| L1-content-control-002 | 97 ✅ | ✓ | The code is nearly flawless, using the correct Office.js API to insert a content control, but loses a few points on approach for not explicitly checking if the insertion was successful or handling potential errors. |
| L1-mail-merge-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "mailMergeReplace" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice. |
| L1-template-apply-001 | 72 | ✓ | The code uses a correct API method applyTemplate but loses points for not using Word.run to execute the API call and for not checking if the template exists, and also for not following best practices in terms of error handling and variable scope. |
| L1-template-apply-002 | 0 ⚠️ | ✓ | The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions. |
| L1-document-generate-001 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach, such as potentially re-implementing functionality that is already available through the provided helpers, and also for not handling potential errors or edge cases that might occur during execution. |
| L1-theme-apply-001 | 5 ⚠️ | ✓ | The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-theme-apply-002 | 72 | ✓ | The code uses the correct designTheme and applyTheme APIs, but lacks completeness in addressing the request, has some potential issues with execution, and does not fully follow best practices, particularly in terms of using available helpers for styling and content manipulation. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work at runtime, although it partially addresses the request. |
| L1-citation-bibliography-001 | 40 ⚠️ | ✓ | The code heavily hallucinates non-existent Office.js APIs, specifically `insertCitation` and `insertBibliography`, which are not real methods in the Word Office.js API. |
| L1-equation-001 | 72 | ✓ | The code uses a correct API but lacks best practices, such as using Word.run to ensure context and load the document, and does not handle potential errors or loading issues, resulting in deductions across all dimensions. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the user's request, but loses points for not using a more robust method to handle the equation insertion, such as checking the current selection or range before inserting the equation. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selection for clarity, is easy to understand, and directly addresses the question, but loses points for not considering the coach panel as an alternative and not providing additional Word-specific caveats or workarounds. |
| L1-read-query-001 | 40 ⚠️ | ✓ | The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `countWords()` method, which heavily penalizes the API correctness score. |
| L1-read-query-002 | 40 ⚠️ | ✓ | The code fails to utilize actual Word Office.js APIs, instead relying on a hallucinated `getReadability()` method, which is not a real API, resulting in a significant penalty for API correctness. |
| L1-read-query-003 | 0 ⚠️ | ✓ | The code uses a hallucinated method "listHeadings()" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-margins-orientation-001 | 85 | ✓ | The code correctly uses the Word Office.js API and fully addresses the request, but loses points for not using the most efficient approach and potentially having issues with the load/sync order. |
| L1-margins-orientation-002 | 80 | ✓ | The code correctly uses the Office.js API and addresses the user's request, but lacks best practices and does not utilize available helpers, and also assumes the existence of a setPageOrientation method which is not a standard Word Office.js API. |
| L1-web-search-needed-001 | 67 | ✓ | The AI response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks clarity on how this information will be used to add a section to the Word document and does not provide any workarounds or suggestions. |
| L1-web-search-needed-002 | 72 | ✓ | The response accurately explains the Pythagorean theorem and directly addresses the question, but loses points for clarity due to assuming prior knowledge of geometric terms and for the bonus category due to not providing additional relevant information or workarounds. |
| L1-refusal-001 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not exploring potential alternative solutions or providing more detailed information about the "Save As" feature. |
| L1-refusal-002 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not offering more detailed or creative workarounds, such as using other APIs or libraries to achieve printing functionality. |
| L1-refusal-003 | 72 | ✓ | The response accurately limitation of Word Office.js in creating identifies the Excel-style bar charts, but loses points for not providing a more detailed or alternative workaround, such as using Excel or suggesting a APIs library that can generate |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-paragraph-spacing-001 | 93 | ✓ | The code accurately sets line spacing to 1.5 for all paragraphs, but loses points for not using a more efficient or best-practice approach, such as utilizing a batch operation or a more targeted range selection, and also for not handling potential errors that may occur during the context sync operation. |
| L1-edge-case-001 | 20 ⚠️ | ✓ | The code to use any fails real Word Office.js APIs, instead throwing an error, which severely penal directlyizes its API correctness and approach scores. |
| L1-form-field-001 | 97 ✅ | ✓ | The code is nearly perfect, using the correct Office.js API to insert a dropdown content control, but loses a few points for not using a more descriptive title or tag, and not handling potential errors that may occur during execution. |
| L1-text-insert-001 | 83 | ✓ | The code is mostly correct but loses points for using the `addParagraph` method with an alignment option, which may not be the most straightforward way to achieve the requested functionality, and for not explicitly addressing the "end of the document" requirement. |
| L1-text-edit-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions. |

## ⚠️ Needs attention

**[L1-paragraph-format-002]** score=20 — A: 20
B:

**[L1-table-create-001]** score=20 — A: 20
B:

**[L1-table-format-001]** score=0 — The code fails to apply the style to the existing table, instead attempting to insert a new table with the specified style, which does not address the user's request.

**[L1-footnote-insert-001]** score=0 — The code uses a hallucinated method "insertFootnote" which is not a real Word Office.js API, instead it should use the "range.insertFootnote" method to insert a footnote at the current selection.

**[L1-track-changes-toggle-001]** score=25 — A: 25
B:

**[L1-image-insert-002]** score=20 — The code uses a hallucinated method `insertImage` which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-mail-merge-001]** score=20 — The code uses a hallucinated method "mailMergeReplace" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice.

**[L1-template-apply-002]** score=0 — The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions.

**[L1-theme-apply-001]** score=5 — The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-theme-apply-003]** score=20 — The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work at runtime, although it partially addresses the request.

**[L1-citation-bibliography-001]** score=40 — The code heavily hallucinates non-existent Office.js APIs, specifically `insertCitation` and `insertBibliography`, which are not real methods in the Word Office.js API.

**[L1-read-query-001]** score=40 — The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `countWords()` method, which heavily penalizes the API correctness score.

**[L1-read-query-002]** score=40 — The code fails to utilize actual Word Office.js APIs, instead relying on a hallucinated `getReadability()` method, which is not a real API, resulting in a significant penalty for API correctness.

**[L1-read-query-003]** score=0 — The code uses a hallucinated method "listHeadings()" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-edge-case-001]** score=20 — The code to use any fails real Word Office.js APIs, instead throwing an error, which severely penal directlyizes its API correctness and approach scores.

**[L1-text-edit-001]** score=0 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions.
