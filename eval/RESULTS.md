# Word AI — Eval Results
**Last run:** 2026-05-23 03:02:45  
**Overall: 57.1/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 40.0/100 | — | 2 |
| paragraph-format | 45.0/100 | — | 2 |
| list-bullet | 72.0/100 | — | 1 |
| list-multilevel | 0.0/100 | — | 1 |
| style-apply | 80.0/100 | — | 1 |
| table-create | 90.0/100 | — | 1 |
| table-format | 67.0/100 | — | 1 |
| find-replace | 81.0/100 | — | 2 |
| find-replace-regex | 83.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 90.0/100 | — | 2 |
| toc-generate | 80.0/100 | — | 1 |
| section-break | 72.0/100 | — | 1 |
| header-footer | 73.5/100 | — | 2 |
| columns | 72.0/100 | — | 1 |
| image-insert | 20.0/100 | — | 2 |
| watermark-insert | 30.0/100 | — | 1 |
| content-control | 83.0/100 | — | 2 |
| mail-merge | 55.0/100 | — | 1 |
| template-apply | 40.0/100 | — | 2 |
| document-generate | 90.0/100 | — | 1 |
| theme-apply | 20.0/100 | — | 3 |
| citation-bibliography | 45.0/100 | — | 1 |
| equation | 80.0/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 20.0/100 | — | 3 |
| margins-orientation | 82.5/100 | — | 2 |
| web-search-needed | 33.5/100 | — | 2 |
| refusal | 74.0/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 0.0/100 | — | 1 |
| edge-cases | 50.0/100 | — | 1 |
| form-field | 72.8/100 | — | 5 |
| text-insert | 80.0/100 | — | 1 |
| text-edit | 20.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 80 | ✓ | The code uses a non-standard method signature "addHeading" which is not a real Word Office.js API, but the rest of the code is mostly correct and follows best practices, with the only other issues being potential runtime errors due to the unknown implementation of "addHeading" and the lack of error handling. |
| L1-heading-insert-002 | 0 ⚠️ | ✓ | The code uses hallucinated methods "addHeading" and "addParagraph" which are not real Word Office.js APIs, resulting in a complete loss of points across all dimensions. |
| L1-paragraph-format-001 | 0 ⚠️ | ✗ | The code attempts to add a new paragraph with the specified style instead of modifying the existing first paragraph, and also uses the `addParagraph` method which is not a real Word Office.js API, resulting in a complete failure to address the user's request. |
| L1-paragraph-format-002 | 90 | ✓ | The code accurately addresses the user's request and uses correct Word Office.js APIs, but loses points for not using the more efficient and recommended `paragraph.styleBuiltIn` or `applyStyle` methods to set the font size and alignment, instead directly accessing the `font` and `alignment` properties. |
| L1-list-bullet-001 | 72 | ✓ | The code uses a non-standard `addList` method which is not a real Word Office.js API, and also lacks proper error handling and context setup, but it does attempt to address the user's request and might work in a specific implementation. |
| L1-list-multilevel-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions. |
| L1-style-apply-001 | 80 | ✓ | The code correctly applies the Quote style to the third paragraph using the applyStyle method, but loses points for not using the Word.run method to execute the operation and for not checking if the paragraph exists before applying the style. |
| L1-table-create-001 | 90 | ✓ | The code accurately uses the Word Office.js API to insert a table, but loses points for not fully addressing the request by only populating three of the four rows, and for not using the most optimal approach by directly specifying the table style instead of potentially using a more flexible method. |
| L1-table-format-001 | 67 | ✓ | The code uses a non-standard method "styleTable" which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-find-replace-001 | 72 | ✓ | The code uses a correct API method replaceText but loses points for not using Word.run to initiate the operation and not checking if the replacement was successful, also it does not follow best practices by directly calling replaceText without utilizing the available helpers and context. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most optimal approach, such as utilizing the `range.font` properties directly on the search results, and for potential issues with the load and sync order in more complex scenarios. |
| L1-find-replace-regex-001 | 83 | ✓ | The code correctly uses the Word Office.js API and addresses the request, but loses points for not fully handling potential edge cases and not using the most optimal approach, such as checking the search results before replacing. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a correct API (insertFootnote) but lacks specificity in targeting the current selection (paragraph 2) and does not follow best practices by directly using the insertFootnote method without ensuring the range is properly set to the current selection. |
| L1-comment-insert-001 | 72 | ✓ | The code uses a correct API method `insertComment` but lacks specificity in targeting the first occurrence of "$4.2M" and does not utilize Word's built-in search functionality with wildcards or regular expressions as needed for precise text matching. |
| L1-track-changes-toggle-001 | 95 ✅ | ✓ | The code perfectly addresses the request using the correct API, but loses points for approach due to directly calling `toggleTrackChanges` without utilizing the `Word.run` context, which is a best practice. |
| L1-track-changes-toggle-002 | 85 | ✓ | The code is mostly correct but loses points for API correctness and approach due to the use of a generic method name "acceptAllRevisions" which is not a real Word Office.js API, instead it should use "context.document.changeTrackingMode" or a similar API to achieve the desired result. |
| L1-toc-generate-001 | 80 | ✓ | The code uses a correct and existing API method `insertTableOfContents`, but loses points for not fully addressing the request by not specifying the exact location of the Table of Contents, and not using the best practice of using `Word.run` to execute the operation. |
| L1-section-break-001 | 72 | ✓ | The code partially addresses the request but lacks best practices and does not fully utilize the available Word Office.js APIs, such as properly handling the context and section objects to ensure the landscape orientation is applied to the correct section. |
| L1-header-footer-001 | 80 | ✓ | The code is mostly correct but lacks consideration for the specific section or page where the header should be applied, and it does not use the most appropriate built-in methods for header manipulation, which might lead to issues with header consistency across the document. |
| L1-header-footer-002 | 67 | ✗ | The code uses a non-existent `setFooter` method, which is not a real Word Office.js API, resulting in significant penalties in API correctness and approach dimensions. |
| L1-columns-001 | 72 | ✓ | The code correctly uses real Word Office.js APIs and would likely execute without runtime errors, but it lacks completeness in addressing the user's request to switch the rest of the document into a two-column layout, and does not use best practices or available helpers. |
| L1-image-insert-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-image-insert-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a significant penalty in API correctness and would not work in practice. |
| L1-watermark-insert-001 | 30 ⚠️ | ✓ | The code is incomplete and does not fully address the request, as it only inserts a watermark once without specifying the page range or ensuring it's applied to every page. |
| L1-content-control-001 | 83 | ✓ | The code correctly uses the Office.js API to insert a content control, but loses points for not fully addressing the request by not specifying the exact location of the insertion, and for not using the most optimal approach by not utilizing the available helpers for content control insertion. |
| L1-content-control-002 | 83 | ✓ | The code is mostly correct but loses points for not using the Word.run method to execute the content control insertion and for not fully following best practices in terms of using available helpers. |
| L1-mail-merge-001 | 55 ⚠️ | ✓ | The code uses a hallucinated method "mailMergeReplace" which is not a real Word Office.js API, resulting in a significant penalty in API correctness. |
| L1-template-apply-001 | 80 | ✓ | The code correctly applies the resume template with the provided information, but loses points for not using the Word.run method to ensure the context is loaded before applying the template, and for not fully utilizing best practices and available helpers. |
| L1-template-apply-002 | 0 ⚠️ | ✓ | The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions. |
| L1-document-generate-001 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach, such as potentially reusing the same paragraph or heading objects instead of creating new ones for each section, and not handling potential errors or edge cases. |
| L1-theme-apply-001 | 10 ⚠️ | ✓ | The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-theme-apply-002 | 15 ⚠️ | ✓ | The code heavily hallucinates the "designTheme" and "applyTheme" methods, which are not real Word Office.js APIs, resulting in a significant penalty for API correctness. |
| L1-theme-apply-003 | 35 ⚠️ | ✓ | The code heavily hallucinates the "tweakTheme" method, which is not a real Word Office.js API, resulting in a significant penalty for API correctness. |
| L1-citation-bibliography-001 | 45 ⚠️ | ✓ | The code heavily hallucinates methods, specifically "insertCitation" and "insertBibliography", which are not real Word Office.js APIs, resulting in a significant penalty in API correctness. |
| L1-equation-001 | 75 | ✓ | The code uses a correct and existing API method `insertEquation`, but loses points for not using the `Word.run` context and not specifying the exact location of the insertion, and also for not following best practices in terms of error handling and context management. |
| L1-equation-002 | 85 | ✓ | The code correctly uses the insertEquation helper and addresses the request, but loses points for not fully following best practices, such as not checking the current cursor position or paragraph context before inserting the equation. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selection for clarity, but loses points for not considering alternative approaches, such as recommending the coach panel, and for minor clarity issues in the code comments. |
| L1-read-query-001 | 20 ⚠️ | ✓ | The code hallucinates a non-existent `countWords()` method, which is not a real Word Office.js API, and does not use the correct API to count words, such as `body.getText()` or `range.getText()`. |
| L1-read-query-002 | 40 ⚠️ | ✓ | The code heavily hallucinates a non-existent `getReadability()` method, which is not a real Word Office.js API, resulting in a significant penalty for API correctness. |
| L1-read-query-003 | 0 ⚠️ | ✓ | The code uses a hallucinated method "listHeadings()" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-margins-orientation-001 | 85 | ✓ | The code correctly sets page margins using the Word Office.js API, but loses points for not using a more explicit best practice approach, such as using a predefined constant for the margin value, and for not handling potential errors that may occur during the execution of the code. |
| L1-margins-orientation-002 | 80 | ✓ | The code is mostly correct but lacks the Word.run context and may not fully address the request for the entire document, also it does not use the best practice of using Word built-in methods for page orientation changes. |
| L1-web-search-needed-001 | 67 | ✓ | The AI response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks clarity on how this information will be used to address the user's request to add a section about the stock price to the document. |
| L1-web-search-needed-002 | 0 ⚠️ | ✓ | A:  |
| L1-refusal-001 | 83 | ✓ | The response is mostly accurate and clear, but loses points for not providing more detailed workaround instructions or alternative solutions, and for not explicitly stating any Word-specific caveats related to PDF export. |
| L1-refusal-002 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but lacks a more detailed explanation or alternative solutions, resulting in moderate scores across the dimensions. |
| L1-refusal-003 | 67 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a workaround, but the workaround is incomplete and lacks specific details on how to create a chart image from the provided data. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-paragraph-spacing-001 | 0 ⚠️ | ✓ | The code hallucinates the `addParagraph` method and incorrectly attempts to modify existing paragraphs while also adding new ones, demonstrating a fundamental misunderstanding of the Word Office.js API. |
| L1-edge-case-001 | 50 ⚠️ | ✓ | The code fails to fully address the request by not attempting to delete the paragraph and instead immediately throwing an error, and also lacks the use of best practices and available helpers in the Office.js API. |
| L1-form-field-001 | 83 | ✓ | The code correctly uses the Office.js API to insert a dropdown content control, but loses points for not fully addressing potential edge cases and not using the most idiomatic approach to content control insertion. |
| L1-text-insert-001 | 80 | ✓ | The code uses a correct and existing API method addParagraph, but loses points for not explicitly addressing the "at the end of the document" requirement and not using the more precise body.paragraphs.add() method, and also for not fully utilizing best practices. |
| L1-text-edit-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| gen-L2-form-field-001 | 72 | ✓ | The code partially addresses the request but fails to insert the date picker content control below the checkbox content control and uses a different title for the date picker, also lacking best practices in content control insertion. |
| gen-L2-form-field-002 | 65 | ✗ | The code incorrectly adds a building block gallery control with the title "Address" instead of adding it to the document without specifying a title, and also fails to replace the existing plain text "Name: _______" as requested. |
| gen-L2-form-field-003 | 72 | ✗ | The code fails to insert a group content control containing the two dropdown content controls as requested, instead inserting separate content controls, and does not utilize the best practices and available helpers provided by the Word Office.js API. |
| gen-L2-form-field-004 | 72 | ✗ | The code uses correct APIs and is mostly complete, but it lacks proper integration of the repeating section content control with the table and does not follow best practices for inserting content controls and tables. |

## ⚠️ Needs attention

**[L1-heading-insert-002]** score=0 — The code uses hallucinated methods "addHeading" and "addParagraph" which are not real Word Office.js APIs, resulting in a complete loss of points across all dimensions.

**[L1-paragraph-format-001]** score=0 — The code attempts to add a new paragraph with the specified style instead of modifying the existing first paragraph, and also uses the `addParagraph` method which is not a real Word Office.js API, resulting in a complete failure to address the user's request.
- missing patterns: `font.bold`, `font.size`

**[L1-list-multilevel-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions.

**[L1-image-insert-001]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-image-insert-002]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a significant penalty in API correctness and would not work in practice.

**[L1-watermark-insert-001]** score=30 — The code is incomplete and does not fully address the request, as it only inserts a watermark once without specifying the page range or ensuring it's applied to every page.

**[L1-mail-merge-001]** score=55 — The code uses a hallucinated method "mailMergeReplace" which is not a real Word Office.js API, resulting in a significant penalty in API correctness.

**[L1-template-apply-002]** score=0 — The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions.

**[L1-theme-apply-001]** score=10 — The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-theme-apply-002]** score=15 — The code heavily hallucinates the "designTheme" and "applyTheme" methods, which are not real Word Office.js APIs, resulting in a significant penalty for API correctness.

**[L1-theme-apply-003]** score=35 — The code heavily hallucinates the "tweakTheme" method, which is not a real Word Office.js API, resulting in a significant penalty for API correctness.

**[L1-citation-bibliography-001]** score=45 — The code heavily hallucinates methods, specifically "insertCitation" and "insertBibliography", which are not real Word Office.js APIs, resulting in a significant penalty in API correctness.

**[L1-read-query-001]** score=20 — The code hallucinates a non-existent `countWords()` method, which is not a real Word Office.js API, and does not use the correct API to count words, such as `body.getText()` or `range.getText()`.

**[L1-read-query-002]** score=40 — The code heavily hallucinates a non-existent `getReadability()` method, which is not a real Word Office.js API, resulting in a significant penalty for API correctness.

**[L1-read-query-003]** score=0 — The code uses a hallucinated method "listHeadings()" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-web-search-needed-002]** score=0 — A: 

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-paragraph-spacing-001]** score=0 — The code hallucinates the `addParagraph` method and incorrectly attempts to modify existing paragraphs while also adding new ones, demonstrating a fundamental misunderstanding of the Word Office.js API.

**[L1-edge-case-001]** score=50 — The code fails to fully address the request by not attempting to delete the paragraph and instead immediately throwing an error, and also lacks the use of best practices and available helpers in the Office.js API.

**[L1-text-edit-001]** score=20 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would-it-work dimensions.
