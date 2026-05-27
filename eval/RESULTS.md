# Word AI — Eval Results
**Last run:** 2026-05-27 03:56:33  
**Overall: 59.9/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 69.5/100 | — | 2 |
| paragraph-format | 44.0/100 | — | 2 |
| list-bullet | 25.0/100 | — | 1 |
| list-multilevel | 0.0/100 | — | 1 |
| style-apply | 80.0/100 | — | 1 |
| table-create | 83.0/100 | — | 1 |
| table-format | 72.0/100 | — | 1 |
| find-replace | 81.0/100 | — | 2 |
| find-replace-regex | 83.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 91.5/100 | — | 2 |
| toc-generate | 83.0/100 | — | 1 |
| section-break | 72.0/100 | — | 1 |
| header-footer | 40.0/100 | — | 2 |
| columns | 72.0/100 | — | 1 |
| image-insert | 46.0/100 | — | 2 |
| watermark-insert | 0.0/100 | — | 1 |
| content-control | 81.5/100 | — | 2 |
| mail-merge | 55.0/100 | — | 1 |
| template-apply | 78.5/100 | — | 2 |
| document-generate | 83.0/100 | — | 1 |
| theme-apply | 21.7/100 | — | 3 |
| citation-bibliography | 20.0/100 | — | 1 |
| equation | 81.5/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 32.3/100 | — | 3 |
| margins-orientation | 86.5/100 | — | 2 |
| web-search-needed | 70.0/100 | — | 2 |
| refusal | 79.3/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 0.0/100 | — | 1 |
| edge-cases | 25.0/100 | — | 1 |
| form-field | 81.4/100 | — | 5 |
| text-insert | 85.0/100 | — | 1 |
| text-edit | 0.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 72 | ✓ | The code uses a non-standard method `addHeading` which is not a real Word Office.js API, and lacks consideration for best practices and available helpers, but otherwise addresses the request and would likely execute without major runtime errors. |
| L1-heading-insert-002 | 67 | ✓ | The code uses non-existent helper functions like addHeading and addParagraph, which are not part of the Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-paragraph-format-001 | 0 ⚠️ | ✗ | The code uses a hallucinated method `addParagraph` which is not a real Word Office.js API, and also attempts to modify a paragraph by adding a new paragraph with the same text, instead of modifying the existing paragraph's font properties. |
| L1-paragraph-format-002 | 88 | ✓ | The code correctly uses real Word Office.js APIs and fully addresses the request, but loses points for not using the best practice of applying a style to set the font size and alignment, and for potential issues with the load/sync order. |
| L1-list-bullet-001 | 25 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a severe penalty for API correctness. |
| L1-list-multilevel-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions. |
| L1-style-apply-001 | 80 | ✓ | The code is mostly correct but loses points for not fully addressing the request with perfect accuracy, as it applies the style to the paragraph at index 2, which is correct given that indices are 0-based, but the approach could be more robust and the completeness is slightly lacking due to the simplicity of the solution. |
| L1-table-create-001 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the `Word.run` method to execute the insertion of the table, and for not fully addressing the request by not handling potential errors or edge cases, and for not using the most efficient approach by directly inserting the table without utilizing available helpers. |
| L1-table-format-001 | 72 | ✓ | The code uses a non-standard method `styleTable` which is not a real Word Office.js API, and lacks the use of best practices and available helpers, such as `table.styleBuiltIn`, which would be a more correct and efficient approach. |
| L1-find-replace-001 | 72 | ✓ | The code uses a non-existent `replaceText` method, which is not a real Word Office.js API, and lacks the proper `Word.run` context, resulting in significant deductions in API correctness and approach dimensions. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient or best practice approach, such as using a more specific search method or handling potential errors, and also for not considering the possibility of multiple search results needing to be handled individually. |
| L1-find-replace-regex-001 | 83 | ✓ | The code uses the correct Word Office.js API replaceText with matchWildcards, but loses points for not fully addressing potential edge cases and not using the most optimal approach, such as handling multiple search results or using a more specific wildcard pattern. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a real Office.js API (insertFootnote) but lacks context about the current selection and does not fully address best practices for implementing the request, such as using Word.run to ensure the code executes in the correct context. |
| L1-comment-insert-001 | 72 | ✓ | The code uses a correct API method `insertComment` but lacks specificity in targeting the first occurrence of the text "$4.2M" and does not fully address the request by not ensuring the comment is added to the correct location within the document. |
| L1-track-changes-toggle-001 | 100 ✅ | ✓ | The code perfectly addresses the request to turn on track changes using the correct and available Word Office.js API, specifically the toggleTrackChanges method, with no flaws or unnecessary re-implementations. |
| L1-track-changes-toggle-002 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the Word.run method to execute the acceptAllRevisions action, and for not following best practices by directly calling acceptAllRevisions without wrapping it in a Word.run context. |
| L1-toc-generate-001 | 83 | ✓ | The code correctly uses the insertTableOfContents API and addresses the request, but loses points for not using the most efficient approach and for inserting a page break without considering the document's existing content. |
| L1-section-break-001 | 72 | ✓ | The code uses correct APIs and is likely to work, but it lacks completeness in addressing the request to switch the next section to landscape, as it doesn't specify the section to apply the orientation to, and the approach could be improved by utilizing Word's built-in helpers for section management. |
| L1-header-footer-001 | 80 | ✓ | The code uses a non-standard `setHeader` method which is not a real Word Office.js API, and also lacks consideration for the header's section and page layout, but is otherwise well-structured and likely to execute without errors. |
| L1-header-footer-002 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions. |
| L1-columns-001 | 72 | ✓ | The code correctly uses real Word Office.js APIs and would likely execute without runtime errors, but it lacks completeness in addressing the user's request to switch the "rest of the document" into a two-column layout, and does not follow best practices by not using available helpers to handle the section break and column insertion. |
| L1-image-insert-001 | 72 | ✓ | The code uses a non-standard `insertImage` method which is not a real Word Office.js API, and also lacks proper error handling and context setup, but it attempts to address the user's request and has a correct sync order. |
| L1-image-insert-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-watermark-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "insertWatermark" which is not a real Word Office.js API, and does not address the request to add a watermark to every page of a 4-page document. |
| L1-content-control-001 | 83 | ✓ | The code is mostly correct but loses points for not using the Word.run method to execute the code and for not checking if the content control can be inserted at the current cursor position, and also for not fully following best practices by directly calling insertContentControl instead of using a more structured approach. |
| L1-content-control-002 | 80 | ✓ | The code is mostly correct and complete, but loses points for not using the Word.run() method to execute the content control insertion, and for not fully following best practices by directly calling insertContentControl without utilizing the provided helpers. |
| L1-mail-merge-001 | 55 ⚠️ | ✓ | The code uses a hallucinated method "mailMergeReplace" which is not a real Word Office.js API, resulting in a significant penalty in the API_CORRECTNESS dimension. |
| L1-template-apply-001 | 72 | ✓ | The code uses a correct API method applyTemplate but lacks best practices, such as using Word.run to execute the template application and potentially missing error handling, resulting in deductions across all dimensions. |
| L1-template-apply-002 | 85 | ✓ | The code correctly applies the memo template with the specified fields, but loses points for API correctness due to the use of a potentially non-existent `applyTemplate` method, and for approach due to not using the official Word JavaScript API methods, such as `Word.run` and `body.insertText`. |
| L1-document-generate-001 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the best practices, such as not checking if the context is loaded before calling insertTableOfContents, and not handling potential errors, which deducts from the approach and would_it_work dimensions. |
| L1-theme-apply-001 | 25 ⚠️ | ✓ | The code heavily hallucinates methods, using non-existent `designTheme` and `applyTheme` functions, which are not part of the real Word Office.js APIs. |
| L1-theme-apply-002 | 20 ⚠️ | ✓ | The code heavily hallucinates the "designTheme" and "applyTheme" methods, which are not real Word Office.js APIs, resulting in a significant penalty for API correctness. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code heavily hallucinates the "tweakTheme" method, which is not a real Word Office.js API, resulting in a significant penalty for API correctness. |
| L1-citation-bibliography-001 | 20 ⚠️ | ✓ | The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-equation-001 | 80 | ✓ | The code uses a correct and existing helper method insertEquation, but loses points for not using Word.run to ensure the code is executed in the correct context and for not handling potential errors or loading issues. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the user's request, but loses points for not fully utilizing best practices and available helpers, and for potential issues with variable scope and method signatures. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selected text for clarity, but loses points for not considering alternative approaches, such as recommending the coach panel, and for using a somewhat generic phrase "Rewriting the current selection for clarity" that doesn't add much value. |
| L1-read-query-001 | 40 ⚠️ | ✓ | The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `countWords()` method, which heavily penalizes the API correctness score. |
| L1-read-query-002 | 57 ⚠️ | ✓ | The code uses a hallucinated method "getReadability()" which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-read-query-003 | 0 ⚠️ | ✓ | The code uses a hallucinated method "listHeadings()" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-margins-orientation-001 | 88 | ✓ | The code correctly sets the page margins to 1 inch on all sides using the Word Office.js API, but loses points for approach due to using a generic setMargins method instead of utilizing the available helpers or directly manipulating the section properties. |
| L1-margins-orientation-002 | 85 | ✓ | The code is mostly correct but lacks the use of Word.run to queue the command and does not check the current page orientation before switching, also it does not use the best practice of using Word built-in methods for setting page orientation. |
| L1-web-search-needed-001 | 67 | ✓ | The response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks clarity on how this information will be used to add a section to the document, and does not provide any workarounds or Word-specific caveats. |
| L1-web-search-needed-002 | 73 | ✓ | The response accurately and clearly explains the Pythagorean theorem, directly addressing the question, but lacks a meaningful bonus contribution, such as providing additional examples or noting any Word-specific caveats related to equation insertion. |
| L1-refusal-001 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but lacks a detailed explanation of the "Save As" feature and its potential caveats. |
| L1-refusal-002 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not exploring alternative solutions or providing more detailed information about the workaround. |
| L1-refusal-003 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js in creating charts, but loses points for not providing a more detailed workaround or alternative solution that directly addresses the user's request for an Excel-style bar chart with specific data. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-paragraph-spacing-001 | 0 ⚠️ | ✓ | The code hallucinates the `addParagraph` and `applyStyle` methods, which are not real Word Office.js APIs, and also attempts to modify existing paragraphs by adding new ones instead of updating the existing ones. |
| L1-edge-case-001 | 25 ⚠️ | ✓ | The code fails to use the Word Office.js API to delete a paragraph and instead throws an error, lacking a proper attempt to handle the out-of-range paragraph index using the available APIs. |
| L1-form-field-001 | 97 ✅ | ✓ | The code perfectly addresses the request using the correct Office.js API, but loses a few points on approach for not explicitly checking if the cursor is in a valid location to insert a content control, although the provided code is otherwise well-structured and uses best practices. |
| L1-text-insert-001 | 85 | ✓ | The code uses a real Word Office.js API (addParagraph) and is likely to work without runtime errors, but loses points for not explicitly addressing the "at the end of the document" requirement and not using the most precise method to insert text (paragraph.insertText). |
| L1-text-edit-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the API correctness and other dimensions. |
| gen-L2-form-field-001 | 72 | ✓ | The code uses real Word Office.js APIs and is mostly complete, but loses points for not fully addressing the request, such as not specifying the exact position of the content controls, and not using best practices like inserting a paragraph or setting the label of the date picker content control to match the user's request. |
| gen-L2-form-field-002 | 72 | ✗ | The code uses correct Office.js APIs and is mostly complete, but it lacks proper replacement of the existing plain text and does not follow best practices for inserting content controls and building block gallery controls. |
| gen-L2-form-field-003 | 83 | ✗ | The code accurately uses real Word Office.js APIs and addresses the request, but loses points for not using the most efficient or best practice approaches, such as directly utilizing the `insertContentControl` method with the `group` type and not fully handling potential errors or edge cases during execution. |
| gen-L2-form-field-004 | 83 | ✗ | The code is mostly correct and complete, but loses points for not explicitly setting the content control as a repeating section and not fully utilizing best practices for inserting tables and content controls. |

## ⚠️ Needs attention

**[L1-paragraph-format-001]** score=0 — The code uses a hallucinated method `addParagraph` which is not a real Word Office.js API, and also attempts to modify a paragraph by adding a new paragraph with the same text, instead of modifying the existing paragraph's font properties.
- missing patterns: `font.bold`, `font.size`

**[L1-list-bullet-001]** score=25 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a severe penalty for API correctness.

**[L1-list-multilevel-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions.

**[L1-header-footer-002]** score=0 — The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions.

**[L1-image-insert-002]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-watermark-insert-001]** score=0 — The code uses a hallucinated method "insertWatermark" which is not a real Word Office.js API, and does not address the request to add a watermark to every page of a 4-page document.

**[L1-mail-merge-001]** score=55 — The code uses a hallucinated method "mailMergeReplace" which is not a real Word Office.js API, resulting in a significant penalty in the API_CORRECTNESS dimension.

**[L1-theme-apply-001]** score=25 — The code heavily hallucinates methods, using non-existent `designTheme` and `applyTheme` functions, which are not part of the real Word Office.js APIs.

**[L1-theme-apply-002]** score=20 — The code heavily hallucinates the "designTheme" and "applyTheme" methods, which are not real Word Office.js APIs, resulting in a significant penalty for API correctness.

**[L1-theme-apply-003]** score=20 — The code heavily hallucinates the "tweakTheme" method, which is not a real Word Office.js API, resulting in a significant penalty for API correctness.

**[L1-citation-bibliography-001]** score=20 — The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-read-query-001]** score=40 — The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `countWords()` method, which heavily penalizes the API correctness score.

**[L1-read-query-002]** score=57 — The code uses a hallucinated method "getReadability()" which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions.

**[L1-read-query-003]** score=0 — The code uses a hallucinated method "listHeadings()" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-paragraph-spacing-001]** score=0 — The code hallucinates the `addParagraph` and `applyStyle` methods, which are not real Word Office.js APIs, and also attempts to modify existing paragraphs by adding new ones instead of updating the existing ones.

**[L1-edge-case-001]** score=25 — The code fails to use the Word Office.js API to delete a paragraph and instead throws an error, lacking a proper attempt to handle the out-of-range paragraph index using the available APIs.

**[L1-text-edit-001]** score=0 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the API correctness and other dimensions.
