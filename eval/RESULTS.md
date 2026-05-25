# Word AI — Eval Results
**Last run:** 2026-05-25 03:57:21  
**Overall: 62.5/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 78.5/100 | — | 2 |
| paragraph-format | 86.5/100 | — | 2 |
| list-bullet | 25.0/100 | — | 1 |
| list-multilevel | 72.0/100 | — | 1 |
| style-apply | 80.0/100 | — | 1 |
| table-create | 83.0/100 | — | 1 |
| table-format | 72.0/100 | — | 1 |
| find-replace | 36.0/100 | — | 2 |
| find-replace-regex | 83.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 91.0/100 | — | 2 |
| toc-generate | 72.0/100 | — | 1 |
| section-break | 72.0/100 | — | 1 |
| header-footer | 40.0/100 | — | 2 |
| columns | 72.0/100 | — | 1 |
| image-insert | 46.0/100 | — | 2 |
| watermark-insert | 0.0/100 | — | 1 |
| content-control | 89.0/100 | — | 2 |
| mail-merge | 79.0/100 | — | 1 |
| template-apply | 82.5/100 | — | 2 |
| document-generate | 83.0/100 | — | 1 |
| theme-apply | 15.0/100 | — | 3 |
| citation-bibliography | 40.0/100 | — | 1 |
| equation | 84.0/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 26.7/100 | — | 3 |
| margins-orientation | 87.5/100 | — | 2 |
| web-search-needed | 70.0/100 | — | 2 |
| refusal | 79.3/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 72.0/100 | — | 1 |
| edge-cases | 25.0/100 | — | 1 |
| form-field | 76.4/100 | — | 5 |
| text-insert | 80.0/100 | — | 1 |
| text-edit | 0.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 85 | ✓ | The code is mostly correct but loses points for using a non-standard method `addHeading` which is not a real Word Office.js API, instead of using the available helpers or implementing it manually with `body.paragraphs.add` and `paragraph.styleBuiltIn`. |
| L1-heading-insert-002 | 72 | ✓ | The code uses a hallucinated method addHeading and addParagraph, which are not real Word Office.js APIs, resulting in significant deductions in API correctness and approach dimensions. |
| L1-paragraph-format-001 | 90 | ✓ | The code accurately uses real Word Office.js APIs and fully addresses the request, but loses points for not using the available helpers, such as applyStyle, and for potential issues with the load/sync order, although the provided code snippet seems mostly correct. |
| L1-paragraph-format-002 | 83 | ✓ | The code is mostly correct, but it uses manual property setting for font size instead of using the available helpers, and it assumes the first paragraph is the title without checking, which might not always be the case. |
| L1-list-bullet-001 | 25 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a severe penalty for API correctness. |
| L1-list-multilevel-001 | 72 | ✓ | The code uses a non-standard method "addList" which is not a real |
| L1-style-apply-001 | 80 | ✓ | The code correctly applies the Quote style to the third paragraph using the applyStyle method, but loses points for not using the Word.run method to ensure the context is properly loaded and for not checking if the paragraph exists before applying the style. |
| L1-table-create-001 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the Word.run method to execute the API calls and for not fully following best practices, such as not checking if the table can be inserted at the current location. |
| L1-table-format-001 | 72 | ✓ | The code uses a non-standard method `styleTable` which is not a real Word Office.js API, and also lacks the use of best practices and available helpers, such as `table.style` or `table.styleBuiltIn`. |
| L1-find-replace-001 | 72 | ✓ | The code uses a correct API method replaceText, but lacks the surrounding Word.run context and may not fully address the request if there are multiple instances of "Acme" in different parts of the document, and does not follow best practices by not utilizing the Word.run method. |
| L1-find-replace-002 | 0 ⚠️ | ✓ | A:  |
| L1-find-replace-regex-001 | 83 | ✓ | The code uses the correct Word Office.js API replaceText with matchWildcards, but loses points for not explicitly checking if the replacement was successful and for not using a more precise wildcard pattern to match email addresses. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a real Office.js API (insertFootnote) but lacks specificity in targeting the current selection (paragraph 2) and does not fully address the request by not ensuring the footnote is added to the correct location, and also does not follow best practices by not using the provided helpers for more complex operations. |
| L1-comment-insert-001 | 72 | ✓ | The code uses a non-existent `insertComment` method with incorrect parameters, instead of utilizing the `range.insertComment` method provided by the Word Office.js API. |
| L1-track-changes-toggle-001 | 97 ✅ | ✓ | The code perfectly addresses the request using the correct API, but loses a few points on approach for not explicitly mentioning that it's using a best practice helper method, although in this case, `toggleTrackChanges` is indeed the correct and most straightforward helper method to use. |
| L1-track-changes-toggle-002 | 85 | ✓ | The code is mostly correct and complete, but loses points for not using a more specific best practice approach, such as using Word.run to encapsulate the acceptAllRevisions method, and for not handling potential errors that may occur during the execution of the code. |
| L1-toc-generate-001 | 72 | ✓ | The code uses a correct API but lacks best practices, such as not checking if a Table of Contents already exists, and unnecessarily inserts a page break, also the `insertTableOfContents` method is not a standard Word Office.js API, it seems to be a custom method. |
| L1-section-break-001 | 72 | ✓ | The code is mostly correct but lacks consideration for the current section and paragraph context, and does not use the best practices provided by the Word Office.js API, such as checking the current section orientation before switching to landscape. |
| L1-header-footer-001 | 80 | ✓ | The code uses a non-standard `setHeader` method which is not a real Word Office.js API, losing points for API correctness, and also lacks consideration of best practices for header manipulation, but is otherwise complete and would likely work with minimal adjustments. |
| L1-header-footer-002 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions. |
| L1-columns-001 | 72 | ✓ | The code uses correct APIs and would likely work, but it only partially addresses the request by not specifying the range for the two-column layout, and it does not follow best practices by not using a more targeted approach to apply the layout change. |
| L1-image-insert-001 | 72 | ✓ | The code uses the correct insertImage API but loses points for not specifying the paragraph where the image should be inserted, and for not fully following best practices by not utilizing the available helpers to handle the insertion at the correct cursor position. |
| L1-image-insert-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-watermark-insert-001 | 0 ⚠️ | ✓ | The code hallucinates a non-existent "insertWatermark" method, which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-content-control-001 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the request, but loses points for not using a more specific best practice approach and potentially lacking error handling or loading checks. |
| L1-content-control-002 | 95 ✅ | ✓ | The code is almost perfect, using the correct Office.js API to insert a content control, but loses points for approach because it directly calls `insertContentControl` without considering whether the cursor is at a valid location, such as the end of a paragraph, which could be handled more robustly using provided helpers. |
| L1-mail-merge-001 | 79 | ✓ | The code uses the correct mailMergeReplace API but loses points for not fully addressing potential edge cases, such as handling errors or checking the template context, and for not explicitly demonstrating best practices like error handling or loading the document context. |
| L1-template-apply-001 | 80 | ✓ | The code correctly applies the resume template with the provided details, but loses points for not using the Word.run method to ensure the context is loaded before applying the template, and for not following best practices in terms of error handling and potential improvements in code structure. |
| L1-template-apply-002 | 85 | ✓ | The code correctly applies the memo template with the specified fields, but loses points for API correctness due to the use of a potentially non-existent `applyTemplate` method, and for approach due to not using the official Word.run method to execute the template application. |
| L1-document-generate-001 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach, such as potentially re-implementing table of contents logic instead of relying on built-in helpers, and not handling potential errors or edge cases. |
| L1-theme-apply-001 | 5 ⚠️ | ✓ | The code heavily hallucinates non-existent APIs, specifically "designTheme" and "applyTheme", which are not real Word Office.js APIs, resulting in a significant penalty. |
| L1-theme-apply-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not a valid method, resulting in a complete loss of points for API correctness and would it work dimensions. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would_it_work dimensions. |
| L1-citation-bibliography-001 | 40 ⚠️ | ✓ | The code hallucinates non-existent `insertCitation` and `insertBibliography` methods, which are not part of the real Word Office.js APIs, resulting in a severe penalty for API correctness. |
| L1-equation-001 | 83 | ✓ | The code is mostly correct but loses points for not using the Word.run method to ensure the context is loaded before inserting the equation, and for not checking if the insertEquation method is a valid helper function in the Office.js API. |
| L1-equation-002 | 85 | ✓ | The code correctly uses the insertEquation helper and executes without runtime errors, but loses points for not fully addressing potential user needs, such as allowing customization of the equation, and not using the most optimal approach by directly utilizing the provided helper without additional error checking or handling. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selection for clarity and provides relevant code, but loses points for not considering alternative approaches, such as recommending the coach panel, and for minor clarity issues in the code comments. |
| L1-read-query-001 | 40 ⚠️ | ✓ | The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `countWords()` method, which heavily penalizes the API correctness score. |
| L1-read-query-002 | 40 ⚠️ | ✓ | The code hallucinates a non-existent `getReadability()` method, which is not a real Word Office.js API, resulting in a severe penalty for API correctness. |
| L1-read-query-003 | 0 ⚠️ | ✓ | The code hallucinates a non-existent `listHeadings()` method, which is not a real Word Office.js API, and does not provide any actual implementation to list the headings. |
| L1-margins-orientation-001 | 95 ✅ | ✓ | The code perfectly addresses the request using correct APIs and best practices, but loses points for not using a more explicit or helper-based approach, such as using a predefined unit or margin object, although the provided solution is straightforward and effective. |
| L1-margins-orientation-002 | 80 | ✓ | The code correctly uses the Office.js API and addresses the user's request, but lacks best practices and does not utilize available helpers, and also assumes the existence of a setPageOrientation method which is not a standard Word Office.js API. |
| L1-web-search-needed-001 | 67 | ✓ | The response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks a clear explanation or workaround for integrating this information into a Microsoft Word document using Office JavaScript API. |
| L1-web-search-needed-002 | 73 | ✓ | The response accurately and clearly explains the Pythagorean theorem, directly addressing the question, but lacks any notable bonus features, such as providing the actual formula (a^2 + b^2 = c^2) or offering additional geometric insights. |
| L1-refusal-001 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not exploring alternative solutions or providing more detailed information about the "Save As" feature. |
| L1-refusal-002 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not offering more detailed alternatives or caveats, and for the code sections being unnecessary and empty. |
| L1-refusal-003 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js in creating charts, but loses points for not providing a more detailed or creative workaround, and for inserting a placeholder image description instead of a more meaningful alternative. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-paragraph-spacing-001 | 72 | ✓ | The code uses correct Word Office.js APIs and addresses the request, but lacks best practices, such as using a loop based on the actual number of paragraphs in the document instead of hardcoding the number, and incorrectly applies a style with `applyStyle(i, "paragraph")` which is not necessary for setting line spacing. |
| L1-edge-case-001 | 25 ⚠️ | ✓ | The code fails to use the Word Office.js API correctly, instead directly throwing an error without attempting to access or manipulate the document's paragraphs using the available APIs. |
| L1-form-field-001 | 83 | ✓ | The code correctly uses the Office.js API to insert a dropdown content control, but loses points for not fully addressing potential edge cases and not using the most optimal approach, such as handling potential errors or using more descriptive variable names. |
| L1-text-insert-001 | 80 | ✓ | The code uses a real Word Office.js API (addParagraph) and is likely to work, but loses points for not fully addressing the request (inserting at the end of the document, not necessarily after the last paragraph) and not using the most optimal approach (inserting a paragraph may not always be the best way to add text to the end of a document). |
| L1-text-edit-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure across all dimensions. |
| gen-L2-form-field-001 | 72 | ✓ | The code uses correct Office.js APIs and is likely to work, but it does not fully address the request as it inserts a date picker with the title "Sign Date" instead of placing it below the checkbox content control with no specific label requested for the date picker, and the approach could be improved by utilizing more specific helpers or addressing layout. |
| gen-L2-form-field-002 | 72 | ✗ | The code uses a hallucinated method "insertContentControl" with incorrect parameters, and does not properly address the request to replace existing plain text, resulting in a significant deduction in API correctness and approach scores. |
| gen-L2-form-field-003 | 72 | ✗ | The code uses a non-existent `insertContentControl` method, which is not a real Word Office.js API, and also fails to properly nest the dropdown content controls within the group content control, resulting in a loss of points across multiple dimensions. |
| gen-L2-form-field-004 | 83 | ✗ | The code is mostly correct and complete, but loses points for not explicitly setting the content control as a repeating section, and for not using the most idiomatic API methods, such as setting the table style through a separate method call instead of as an option to insertTable. |

## ⚠️ Needs attention

**[L1-list-bullet-001]** score=25 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a severe penalty for API correctness.

**[L1-find-replace-002]** score=0 — A: 

**[L1-header-footer-002]** score=0 — The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions.

**[L1-image-insert-002]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-watermark-insert-001]** score=0 — The code hallucinates a non-existent "insertWatermark" method, which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-theme-apply-001]** score=5 — The code heavily hallucinates non-existent APIs, specifically "designTheme" and "applyTheme", which are not real Word Office.js APIs, resulting in a significant penalty.

**[L1-theme-apply-002]** score=20 — The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not a valid method, resulting in a complete loss of points for API correctness and would it work dimensions.

**[L1-theme-apply-003]** score=20 — The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would_it_work dimensions.

**[L1-citation-bibliography-001]** score=40 — The code hallucinates non-existent `insertCitation` and `insertBibliography` methods, which are not part of the real Word Office.js APIs, resulting in a severe penalty for API correctness.

**[L1-read-query-001]** score=40 — The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `countWords()` method, which heavily penalizes the API correctness score.

**[L1-read-query-002]** score=40 — The code hallucinates a non-existent `getReadability()` method, which is not a real Word Office.js API, resulting in a severe penalty for API correctness.

**[L1-read-query-003]** score=0 — The code hallucinates a non-existent `listHeadings()` method, which is not a real Word Office.js API, and does not provide any actual implementation to list the headings.

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-edge-case-001]** score=25 — The code fails to use the Word Office.js API correctly, instead directly throwing an error without attempting to access or manipulate the document's paragraphs using the available APIs.

**[L1-text-edit-001]** score=0 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure across all dimensions.
