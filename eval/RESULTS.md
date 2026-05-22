# Word AI — Eval Results
**Last run:** 2026-05-22 03:46:46  
**Overall: 56.8/100**  
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
| table-format | 72.0/100 | — | 1 |
| find-replace | 82.0/100 | — | 2 |
| find-replace-regex | 83.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 91.5/100 | — | 2 |
| toc-generate | 72.0/100 | — | 1 |
| section-break | 72.0/100 | — | 1 |
| header-footer | 50.0/100 | — | 2 |
| columns | 72.0/100 | — | 1 |
| image-insert | 27.5/100 | — | 2 |
| watermark-insert | 0.0/100 | — | 1 |
| content-control | 89.0/100 | — | 2 |
| mail-merge | 72.0/100 | — | 1 |
| template-apply | 51.5/100 | — | 2 |
| document-generate | 83.0/100 | — | 1 |
| theme-apply | 20.7/100 | — | 3 |
| citation-bibliography | 40.0/100 | — | 1 |
| equation | 76.0/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 20.0/100 | — | 3 |
| margins-orientation | 87.5/100 | — | 2 |
| web-search-needed | 77.0/100 | — | 2 |
| refusal | 79.3/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 72.0/100 | — | 1 |
| edge-cases | 25.0/100 | — | 1 |
| form-field | 72.8/100 | — | 5 |
| text-insert | 80.0/100 | — | 1 |
| text-edit | 0.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addHeading" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-heading-insert-002 | 0 ⚠️ | ✓ | The code uses hallucinated methods "addHeading" and "addParagraph" which are not real Word Office.js APIs, resulting in a complete failure to meet the requirements. |
| L1-paragraph-format-001 | 0 ⚠️ | ✗ | The code uses a hallucinated method `addParagraph` which is not a real Word Office.js API, and also attempts to modify an existing paragraph's text and style in a way that is not supported by the API. |
| L1-paragraph-format-002 | 83 | ✓ | The code is mostly correct but loses points for not checking if the first paragraph is indeed the title and for not using the more specific and helper-based approach of applying a style with a built-in size. |
| L1-list-bullet-001 | 25 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-list-multilevel-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-style-apply-001 | 80 | ✓ | The code uses a real Word Office.js API, applies the style correctly, and would likely execute without runtime errors, but loses points for not fully addressing the request with a more robust paragraph selection method and not using the most idiomatic helper functions. |
| L1-table-create-001 | 89 | ✓ | The code accurately uses the Word Office.js API to insert a table, but loses points for not fully addressing the request by not handling the document context and paragraph insertion, and for not using the most optimal approach by directly specifying the table style instead of using available helpers. |
| L1-table-format-001 | 72 | ✓ | The code uses a non-existent `styleTable` method, which is not a real Word Office.js API, resulting in significant penalties for API correctness and approach. |
| L1-find-replace-001 | 72 | ✓ | The code uses a correct API method replaceText, but loses points for not using Word.run to initiate the operation and not checking if the replacement was successful, and also for not following best practices in terms of error handling and context synchronization. |
| L1-find-replace-002 | 92 | ✓ | The code is mostly correct and complete, but loses points for not using the most optimal approach, such as utilizing the available helpers, and for a minor potential issue with the load and sync order. |
| L1-find-replace-regex-001 | 83 | ✓ | The code uses the correct Word Office.js API replaceText with matchWildcards, but loses points for not fully handling potential edge cases and not using the most robust best practices for searching and replacing text in the document. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a correct API method to insert a footnote, but lacks specificity in targeting the current selection and does not follow best practices for inserting content at a specific location in the document. |
| L1-comment-insert-001 | 72 | ✓ | The code uses a correct API method to insert a comment, but lacks specificity in targeting the exact occurrence of the text "$4.2M" and does not utilize Word's built-in search functionality with matchWildcards, which would be a more robust approach. |
| L1-track-changes-toggle-001 | 95 ✅ | ✓ | The code perfectly addresses the request using the correct API, but loses points for approach because it manually calls `toggleTrackChanges` instead of using a more explicit method like `doc.changeTrackingMode`, although `toggleTrackChanges` is a valid helper function. |
| L1-track-changes-toggle-002 | 88 | ✓ | The code is mostly correct but loses points for not using the Word.run method to execute the API calls and for not checking if the document has tracked revisions before accepting them, which could lead to potential issues with the approach. |
| L1-toc-generate-001 | 72 | ✓ | The code uses a correct API method to insert a Table of Contents, but lacks best practices and available helpers, and does not fully address the request as it does not specify the exact location of the Table of Contents, which is supposed to be at the top. |
| L1-section-break-001 | 72 | ✓ | The code uses correct APIs and is likely to work, but it lacks completeness in addressing the user's request to switch the next section to landscape, and it does not follow best practices by using available helpers, instead using generic methods like setPageOrientation. |
| L1-header-footer-001 | 80 | ✓ | The code is mostly correct but lacks consideration for the specific section or page where the header should be applied, and it does not utilize the most appropriate built-in methods for header manipulation, which might lead to issues with header consistency across the document. |
| L1-header-footer-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work at runtime. |
| L1-columns-001 | 72 | ✓ | The code uses correct APIs and would likely work, but it lacks completeness in addressing the user's request to switch the rest of the document into a two-column layout, as it doesn't handle the current section break or ensure the new layout applies to the entire document. |
| L1-image-insert-001 | 35 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a severe penalty in API correctness. |
| L1-image-insert-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-watermark-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "insertWatermark" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-content-control-001 | 83 | ✓ | The code correctly uses the Office.js API to insert a content control, but loses points for not fully addressing potential edge cases and not using the most optimal approach, such as handling potential errors or using more specific methods. |
| L1-content-control-002 | 95 ✅ | ✓ | The code perfectly addresses the request using the correct Office.js API, but loses points for not using the most idiomatic approach, as it directly calls `insertContentControl` without considering potential context or error handling. |
| L1-mail-merge-001 | 72 | ✓ | The code uses the correct mailMergeReplace API but lacks proper error handling, does not fully address potential edge cases, and does not follow best practices for handling data and template integration. |
| L1-template-apply-001 | 83 | ✓ | The code correctly applies the resume template with the provided details, but loses points for not using the Word.run method to ensure the context is loaded before applying the template, and for not fully utilizing best practices and available helpers. |
| L1-template-apply-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would_it_work dimensions. |
| L1-document-generate-001 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach, such as potentially re-implementing the Table of Contents logic instead of relying on the built-in insertTableOfContents method, and not fully utilizing the available helpers for adding headings and paragraphs. |
| L1-theme-apply-001 | 10 ⚠️ | ✓ | The code hallucinates non-existent `designTheme` and `applyTheme` methods, which are not part of the real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-theme-apply-002 | 12 ⚠️ | ✓ | The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not a valid method, resulting in a complete loss of points for API correctness and would it work dimensions. |
| L1-theme-apply-003 | 40 ⚠️ | ✓ | The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a significant penalty for API correctness. |
| L1-citation-bibliography-001 | 40 ⚠️ | ✓ | The code uses hallucinated methods `insertCitation` and `insertBibliography` which are not real Word Office.js APIs, resulting in a significant penalty for API correctness. |
| L1-equation-001 | 67 | ✓ | The code uses a non-existent `insertEquation` method, which is not a real Word Office.js API, resulting in significant penalties for API correctness and approach. |
| L1-equation-002 | 85 | ✓ | The code correctly uses the insertEquation API and addresses the user's request, but loses points for not using a more robust approach, such as checking the current cursor position and paragraph context before inserting the equation. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selected text for clarity, but loses points for not considering the coach panel as an alternative and for using a somewhat generic phrase "Rewriting the current selection for clarity" that doesn't add much value. |
| L1-read-query-001 | 20 ⚠️ | ✓ | The code fails to use any real Word Office.js APIs, instead relying on a hallucinated `countWords()` method, which is not a valid API and would result in a runtime error. |
| L1-read-query-002 | 40 ⚠️ | ✓ | The code hallucinates a non-existent `getReadability()` method, which is not a real Word Office.js API, resulting in a severe penalty for API correctness. |
| L1-read-query-003 | 0 ⚠️ | ✓ | The code uses a hallucinated method "listHeadings()" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-margins-orientation-001 | 95 ✅ | ✓ | The code accurately sets page margins to 1 inch on all sides using the correct Office.js API, but does not demonstrate the use of best practices or available helpers, as the task does not require them. |
| L1-margins-orientation-002 | 80 | ✓ | The code is mostly correct but lacks the Word.run context and uses a generic setPageOrientation method which may not be a real Word Office.js API, also it doesn't fully follow best practices. |
| L1-web-search-needed-001 | 67 | ✓ | The response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks a clear explanation of how the information will be used to add a section to the document, and does not provide any workarounds or Word-specific caveats. |
| L1-web-search-needed-002 | 87 | ✓ | The response accurately and clearly explains the Pythagorean theorem, directly addressing the question, but loses points for not providing additional workarounds or Word-specific caveats, and including unnecessary code in the response. |
| L1-refusal-001 | 83 | ✓ | The response is accurate, clear, and relevant, but loses points for not providing a more detailed workaround or alternative solution, such as suggesting a specific API or tool that could be used to automate the export process. |
| L1-refusal-002 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not offering more detailed alternatives or settings options, and for the code sections being unnecessary and empty. |
| L1-refusal-003 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js in creating charts, but loses points for not providing a more detailed or creative workaround, and for inserting a placeholder image description instead of a more meaningful alternative. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-paragraph-spacing-001 | 72 | ✓ | The code uses correct APIs and addresses the request, but has issues with approach, as it manually iterates over paragraphs and applies line spacing instead of using available helpers, and also calls an undefined `applyStyle` method. |
| L1-edge-case-001 | 25 ⚠️ | ✓ | The code does not use any real Word Office.js APIs, instead directly throwing an error without attempting to access or manipulate the document's paragraphs. |
| L1-form-field-001 | 97 ✅ | ✓ | The code is almost perfect, but loses a few points on approach because it manually constructs the content control options instead of potentially using a more Word-API-centric way to define the dropdown items, although the provided code is still correct and idiomatic. |
| L1-text-insert-001 | 80 | ✓ | The code uses a real Office.js API helper method addParagraph, but loses points for not fully addressing potential edge cases and not using the most precise method to insert text at the end of the document, which would be to get the last paragraph and insert text into it. |
| L1-text-edit-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| gen-L2-form-field-001 | 72 | ✓ | The code partially addresses the request but fails to insert the date picker content control below the checkbox content control and uses a different title for the date picker, also lacking best practices in its implementation. |
| gen-L2-form-field-002 | 72 | ✗ | The code uses a non-existent "comboBox" content control type and lacks proper replacement of the existing plain text, resulting in incomplete and non-standard implementation. |
| gen-L2-form-field-003 | 40 ⚠️ | ✗ | The code fails to use the actual Word Office.js API, instead hallucinating a non-existent `insertContentControl` method, and also does not properly group the two dropdown content controls as requested. |
| gen-L2-form-field-004 | 83 | ✗ | The code is mostly correct but loses points for not explicitly setting the content control as a repeating section and for not fully utilizing best practices such as directly applying a style to the inserted table instead of specifying it in the insertTable method. |

## ⚠️ Needs attention

**[L1-heading-insert-001]** score=0 — The code uses a hallucinated method "addHeading" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-heading-insert-002]** score=0 — The code uses hallucinated methods "addHeading" and "addParagraph" which are not real Word Office.js APIs, resulting in a complete failure to meet the requirements.

**[L1-paragraph-format-001]** score=0 — The code uses a hallucinated method `addParagraph` which is not a real Word Office.js API, and also attempts to modify an existing paragraph's text and style in a way that is not supported by the API.
- missing patterns: `font.bold`, `font.size`

**[L1-list-bullet-001]** score=25 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-list-multilevel-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-header-footer-002]** score=20 — The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work at runtime.

**[L1-image-insert-001]** score=35 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a severe penalty in API correctness.

**[L1-image-insert-002]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-watermark-insert-001]** score=0 — The code uses a hallucinated method "insertWatermark" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-template-apply-002]** score=20 — The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would_it_work dimensions.

**[L1-theme-apply-001]** score=10 — The code hallucinates non-existent `designTheme` and `applyTheme` methods, which are not part of the real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-theme-apply-002]** score=12 — The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not a valid method, resulting in a complete loss of points for API correctness and would it work dimensions.

**[L1-theme-apply-003]** score=40 — The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a significant penalty for API correctness.

**[L1-citation-bibliography-001]** score=40 — The code uses hallucinated methods `insertCitation` and `insertBibliography` which are not real Word Office.js APIs, resulting in a significant penalty for API correctness.

**[L1-read-query-001]** score=20 — The code fails to use any real Word Office.js APIs, instead relying on a hallucinated `countWords()` method, which is not a valid API and would result in a runtime error.

**[L1-read-query-002]** score=40 — The code hallucinates a non-existent `getReadability()` method, which is not a real Word Office.js API, resulting in a severe penalty for API correctness.

**[L1-read-query-003]** score=0 — The code uses a hallucinated method "listHeadings()" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-edge-case-001]** score=25 — The code does not use any real Word Office.js APIs, instead directly throwing an error without attempting to access or manipulate the document's paragraphs.

**[L1-text-edit-001]** score=0 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[gen-L2-form-field-003]** score=40 — The code fails to use the actual Word Office.js API, instead hallucinating a non-existent `insertContentControl` method, and also does not properly group the two dropdown content controls as requested.
- missing patterns: `GroupContentControl`
