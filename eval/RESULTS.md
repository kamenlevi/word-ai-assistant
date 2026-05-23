# Word AI — Eval Results
**Last run:** 2026-05-23 10:00:15  
**Overall: 59.0/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 42.5/100 | — | 2 |
| paragraph-format | 41.5/100 | — | 2 |
| list-bullet | 0.0/100 | — | 1 |
| list-multilevel | 20.0/100 | — | 1 |
| style-apply | 80.0/100 | — | 1 |
| table-create ✓ | 97.0/100 | — | 1 |
| table-format | 72.0/100 | — | 1 |
| find-replace | 83.0/100 | — | 2 |
| find-replace-regex | 83.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 90.0/100 | — | 2 |
| toc-generate | 67.0/100 | — | 1 |
| section-break | 80.0/100 | — | 1 |
| header-footer | 72.0/100 | — | 2 |
| columns | 72.0/100 | — | 1 |
| image-insert | 46.0/100 | — | 2 |
| watermark-insert | 50.0/100 | — | 1 |
| content-control | 83.0/100 | — | 2 |
| mail-merge | 72.0/100 | — | 1 |
| template-apply | 10.0/100 | — | 2 |
| document-generate | 83.0/100 | — | 1 |
| theme-apply | 15.0/100 | — | 3 |
| citation-bibliography | 20.0/100 | — | 1 |
| equation | 83.0/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 19.0/100 | — | 3 |
| margins-orientation | 84.0/100 | — | 2 |
| web-search-needed | 71.5/100 | — | 2 |
| refusal | 73.3/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 72.0/100 | — | 1 |
| edge-cases | 35.0/100 | — | 1 |
| form-field | 80.6/100 | — | 5 |
| text-insert | 72.0/100 | — | 1 |
| text-edit | 25.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 85 | ✓ | The code is mostly correct but loses points for using a potentially non-standard `addHeading` method, which is not a built-in Word Office.js API, although it seems to be used correctly in this context. |
| L1-heading-insert-002 | 0 ⚠️ | ✓ | The code uses hallucinated methods "addHeading" and "addParagraph" which are not real Word Office.js APIs, resulting in a complete failure to meet the requirements. |
| L1-paragraph-format-001 | 0 ⚠️ | ✗ | The code uses a non-existent `addParagraph` method, which is not a real Word Office.js API, and also modifies the text of the first paragraph instead of applying the requested formatting to the existing paragraph. |
| L1-paragraph-format-002 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the available helper methods, such as applyStyle to set the font size, and for not checking if the first paragraph is indeed the title before applying the changes. |
| L1-list-bullet-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions. |
| L1-list-multilevel-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a significant penalty for API correctness. |
| L1-style-apply-001 | 80 | ✓ | The code uses a real Word Office.js API, applies the style correctly, and would likely execute without runtime errors, but loses points for not fully addressing potential edge cases and not using the most idiomatic or efficient approach to styling a paragraph. |
| L1-table-create-001 | 97 ✅ | ✓ | The code is nearly flawless, using the correct Word Office.js API to insert a table with headers, but loses a few points for not using the most concise and expressive API calls, such as specifying the table style using a more robust method. |
| L1-table-format-001 | 72 | ✓ | The code uses a non-standard `styleTable` method which is not a real Word Office.js API, and also lacks proper error handling and context setup, but it attempts to address the request and might work in some scenarios. |
| L1-find-replace-001 | 83 | ✓ | The code correctly uses the replaceText API and addresses the user's request, but lacks proper error handling and does not utilize Word's built-in search functionality with wildcards, instead relying on a generic replaceText method. |
| L1-find-replace-002 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the `range` object explicitly and for not handling potential errors or edge cases, and also for not using the most efficient or best practice approach to achieve the desired result. |
| L1-find-replace-regex-001 | 83 | ✓ | The code uses the correct Word Office.js API replaceText with matchWildcards, but loses points for not fully addressing potential edge cases and not using the most optimal approach by directly utilizing the provided API without additional error handling or checks. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a correct API method to insert a footnote, but lacks consideration for the current selection and paragraph context, and does not utilize the best practice helpers provided by the Office.js API. |
| L1-comment-insert-001 | 72 | ✓ | The code uses a correct API method `insertComment` but lacks specificity in targeting the first occurrence of the text "$4.2M" and does not fully address the request by not ensuring the comment is added to the correct location within the document. |
| L1-track-changes-toggle-001 | 97 ✅ | ✓ | The code perfectly addresses the request using the correct API, but loses a few points for not explicitly handling potential errors or edge cases, although the provided code snippet is concise and effective. |
| L1-track-changes-toggle-002 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the Word.run method to wrap the API calls, and for not checking if there are actually tracked revisions to accept before calling acceptAllRevisions. |
| L1-toc-generate-001 | 67 | ✓ | The code uses a correct API but lacks best practices, such as utilizing the available helper functions, and the provided function insertTableOfContents is not a standard Word Office.js API, which reduces the score for API correctness and approach. |
| L1-section-break-001 | 80 | ✓ | The code is mostly correct but lacks specificity in targeting the next section for landscape orientation and does not fully follow best practices by directly using setPageOrientation without ensuring the section is properly defined or handled. |
| L1-header-footer-001 | 72 | ✓ | The code uses a non-existent `setHeader` method, which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-header-footer-002 | 72 | ✓ | The code uses a non-existent `addPageNumbers` method, which is not a real Word Office.js API, resulting in significant penalties in API correctness and approach dimensions. |
| L1-columns-001 | 72 | ✓ | The code uses correct APIs and would likely work, but it lacks completeness in addressing the user's request to switch the rest of the document into a two-column layout, as it only inserts a section break and columns without considering the existing content or document structure. |
| L1-image-insert-001 | 72 | ✓ | The code uses a correct API method `insertImage` but loses points for not specifying the exact paragraph where the image should be inserted, and for not using best practices, such as utilizing the `paragraph` object to insert the image at the cursor position. |
| L1-image-insert-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-watermark-insert-001 | 50 ⚠️ | ✓ | The code lacks completeness and does not fully address the request, as it only inserts a watermark without specifying the scope to every page, and also uses a non-standard method signature without utilizing the Word.run context and proper range handling. |
| L1-content-control-001 | 83 | ✓ | The code correctly uses the Office.js API to insert a content control, but loses points for not fully addressing the request by not specifying the exact location of the insertion, and for not using the most optimal approach by not utilizing the provided helpers for inserting content controls at a specific location. |
| L1-content-control-002 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the user's request, but loses points for not using a more specific best practice approach and for potential issues with the context.sync() call. |
| L1-mail-merge-001 | 72 | ✓ | The code uses a correct mail merge approach but lacks proper Word Office.js API usage, such as utilizing the `Word.run` method to execute the mail merge operation, and instead relies on a hallucinated `mailMergeReplace` function. |
| L1-template-apply-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a severe penalty for API correctness and would not work at runtime. |
| L1-template-apply-002 | 0 ⚠️ | ✓ | A:  |
| L1-document-generate-001 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach, such as potentially re-implementing the Table of Contents functionality instead of relying on the built-in insertTableOfContents method, and not handling potential errors or edge cases. |
| L1-theme-apply-001 | 10 ⚠️ | ✓ | The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would_it_work dimensions. |
| L1-theme-apply-002 | 15 ⚠️ | ✓ | The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not a valid method, resulting in a complete loss of points for API correctness and would_it_work dimensions. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code hallucinates a non-existent `tweakTheme` method and fails to use the real Word Office.js APIs, such as `designTheme` and `applyTheme`, to achieve the desired theme modification. |
| L1-citation-bibliography-001 | 20 ⚠️ | ✓ | The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-equation-001 | 83 | ✓ | The code is mostly correct but loses points for not using the Word.run method to execute the insertion of the equation, and for not checking if the cursor is at the end of the paragraph before inserting the equation, which could lead to incorrect placement. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation helper, but loses points for not checking the current cursor position and paragraph context before inserting the equation, and for not using a more robust method to handle potential errors or exceptions. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selected text for clarity, but loses points for not considering alternative approaches, such as recommending the coach panel, and for using a somewhat generic phrase "Rewriting the selected text for clarity" that could be more descriptive. |
| L1-read-query-001 | 0 ⚠️ | ✓ | The code hallucinates a non-existent `countWords()` method, which is not a real Word Office.js API, and does not use any actual Word APIs to count the words in the document. |
| L1-read-query-002 | 57 ⚠️ | ✓ | The code uses a hallucinated method `getReadability()` which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-read-query-003 | 0 ⚠️ | ✓ | The code hallucinates a non-existent `listHeadings()` method, which is not a real Word Office.js API, and does not provide any actual implementation to list the headings in the document. |
| L1-margins-orientation-001 | 88 | ✓ | The code correctly uses the Word Office.js API and fully addresses the request, but loses points for approach due to using a manual method to set margins instead of utilizing available helpers or built-in margin settings. |
| L1-margins-orientation-002 | 80 | ✓ | The code is mostly correct but lacks the surrounding Word.run context and does not handle potential errors, also it does not fully follow best practices by not using a more specific method to set the page orientation for the entire document. |
| L1-web-search-needed-001 | 67 | ✓ | The AI response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks clarity on how this information will be used to add a section to the document and does not provide any workarounds or Word-specific caveats. |
| L1-web-search-needed-002 | 76 | ✓ | The response accurately and clearly explains the Pythagorean theorem, directly addressing the question, but lacks significant bonus points for not providing additional workarounds, settings options, or Word-specific caveats beyond the basic explanation and code insertion. |
| L1-refusal-001 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but lacks a more detailed explanation or alternative solutions to earn full points. |
| L1-refusal-002 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but lacks a more detailed explanation or alternative solutions, resulting in moderate scores across the dimensions. |
| L1-refusal-003 | 65 | ✓ | The response accurately identifies the limitation of Word Office.js in creating charts, but loses points for clarity and relevance due to the vague workaround and lack of direct connection to the provided data, and for bonus points due to not providing a more concrete alternative solution. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-paragraph-spacing-001 | 72 | ✓ | The code correctly uses real Word Office.js APIs and addresses the request, but has issues with approach, using a manual loop instead of available helpers, and lacks optimization in loading and syncing paragraph formats. |
| L1-edge-case-001 | 35 ⚠️ | ✓ | The code does not use any real Word Office.js APIs, instead directly throwing an error, which severely penalizes its API correctness score. |
| L1-form-field-001 | 97 ✅ | ✓ | The code is nearly perfect, using the correct Office.js API to insert a dropdown content control, but loses a few points for not explicitly checking if the cursor is in a valid location to insert the content control, which is a best practice. |
| L1-text-insert-001 | 72 | ✓ | The code uses a non-standard method `addParagraph` which is not a real Word Office.js API, and also lacks consideration for inserting the text at the end of the document, which would require using `body.paragraphs.getLast()` or a similar approach. |
| L1-text-edit-001 | 25 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a severe penalty for API correctness. |
| gen-L2-form-field-001 | 75 | ✓ | The code partially addresses the request but fails to fully meet the requirements, such as inserting the date picker content control below the checkbox and using the exact label for the date picker as specified in the request. |
| gen-L2-form-field-002 | 65 | ✗ | The code partially addresses the request but fails to correctly replace the existing plain text "Name: _______" and incorrectly uses a comboBox content control instead of a building block gallery control. |
| gen-L2-form-field-003 | 83 | ✗ | The code is mostly correct and complete, but loses points for not using the Word.run method to wrap the content control insertion and for not handling potential errors, and also for not using the best practices such as using a more specific method to insert the content control. |
| gen-L2-form-field-004 | 83 | ✗ | The code is mostly correct and complete, but loses points for not explicitly setting the content control as a repeating section and for not using the most idiomatic API methods, such as setting the table style through a separate method call instead of as an option to insertTable. |

## ⚠️ Needs attention

**[L1-heading-insert-002]** score=0 — The code uses hallucinated methods "addHeading" and "addParagraph" which are not real Word Office.js APIs, resulting in a complete failure to meet the requirements.

**[L1-paragraph-format-001]** score=0 — The code uses a non-existent `addParagraph` method, which is not a real Word Office.js API, and also modifies the text of the first paragraph instead of applying the requested formatting to the existing paragraph.
- missing patterns: `font.bold`, `font.size`

**[L1-list-bullet-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions.

**[L1-list-multilevel-001]** score=20 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a significant penalty for API correctness.

**[L1-image-insert-002]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-watermark-insert-001]** score=50 — The code lacks completeness and does not fully address the request, as it only inserts a watermark without specifying the scope to every page, and also uses a non-standard method signature without utilizing the Word.run context and proper range handling.

**[L1-template-apply-001]** score=20 — The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a severe penalty for API correctness and would not work at runtime.

**[L1-template-apply-002]** score=0 — A: 

**[L1-theme-apply-001]** score=10 — The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would_it_work dimensions.

**[L1-theme-apply-002]** score=15 — The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not a valid method, resulting in a complete loss of points for API correctness and would_it_work dimensions.

**[L1-theme-apply-003]** score=20 — The code hallucinates a non-existent `tweakTheme` method and fails to use the real Word Office.js APIs, such as `designTheme` and `applyTheme`, to achieve the desired theme modification.

**[L1-citation-bibliography-001]** score=20 — The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-read-query-001]** score=0 — The code hallucinates a non-existent `countWords()` method, which is not a real Word Office.js API, and does not use any actual Word APIs to count the words in the document.

**[L1-read-query-002]** score=57 — The code uses a hallucinated method `getReadability()` which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions.

**[L1-read-query-003]** score=0 — The code hallucinates a non-existent `listHeadings()` method, which is not a real Word Office.js API, and does not provide any actual implementation to list the headings in the document.

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-edge-case-001]** score=35 — The code does not use any real Word Office.js APIs, instead directly throwing an error, which severely penalizes its API correctness score.

**[L1-text-edit-001]** score=25 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a severe penalty for API correctness.
