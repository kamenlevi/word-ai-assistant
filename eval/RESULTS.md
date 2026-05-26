# Word AI — Eval Results
**Last run:** 2026-05-26 03:39:58  
**Overall: 58.7/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 69.5/100 | — | 2 |
| paragraph-format | 44.0/100 | — | 2 |
| list-bullet | 0.0/100 | — | 1 |
| list-multilevel | 0.0/100 | — | 1 |
| style-apply | 80.0/100 | — | 1 |
| table-create | 83.0/100 | — | 1 |
| table-format | 93.0/100 | — | 1 |
| find-replace | 85.0/100 | — | 2 |
| find-replace-regex | 83.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 73.0/100 | — | 1 |
| track-changes-toggle | 89.0/100 | — | 2 |
| toc-generate | 83.0/100 | — | 1 |
| section-break | 72.0/100 | — | 1 |
| header-footer | 51.5/100 | — | 2 |
| columns | 72.0/100 | — | 1 |
| image-insert | 20.0/100 | — | 2 |
| watermark-insert | 40.0/100 | — | 1 |
| content-control | 83.0/100 | — | 2 |
| mail-merge | 45.0/100 | — | 1 |
| template-apply | 40.0/100 | — | 2 |
| document-generate | 90.0/100 | — | 1 |
| theme-apply | 15.0/100 | — | 3 |
| citation-bibliography | 20.0/100 | — | 1 |
| equation | 84.0/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 26.7/100 | — | 3 |
| margins-orientation | 81.5/100 | — | 2 |
| web-search-needed | 73.5/100 | — | 2 |
| refusal | 79.3/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 67.0/100 | — | 1 |
| edge-cases | 15.0/100 | — | 1 |
| form-field | 72.8/100 | — | 5 |
| text-insert | 80.0/100 | — | 1 |
| text-edit | 25.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 67 | ✓ | The code uses a non-standard method `addHeading` which is not a real Word Office.js API, indicating a lack of best practices and reliance on potentially unsupported helpers. |
| L1-heading-insert-002 | 72 | ✓ | The code uses a non-standard method signature for adding a heading and paragraph, and lacks the use of best practices and available helpers, such as using the `Word.run` method to execute the operations. |
| L1-paragraph-format-001 | 0 ⚠️ | ✗ | The code uses a hallucinated method `addParagraph` which is not a real Word Office.js API, and also attempts to modify a paragraph's text and style in a way that is not supported by the API. |
| L1-paragraph-format-002 | 88 | ✓ | The code correctly uses Word Office.js APIs and fully addresses the request, but loses points for not using the best practice of applying a style to set the font size and alignment, and for potentially causing issues with the load and sync order. |
| L1-list-bullet-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions. |
| L1-list-multilevel-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-style-apply-001 | 80 | ✓ | The code correctly applies the Quote style to the third paragraph, but uses a generic applyStyle method which may not be the most efficient or Word-specific approach, and assumes the paragraph index is 2, which may not always be the case. |
| L1-table-create-001 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the `Word.run` context and not handling potential errors, and also for not fully following best practices by not checking if the table can be inserted at the current location. |
| L1-table-format-001 | 93 | ✓ | The code accurately applies the requested table style, but loses points for not using the available helper methods, such as applyStyle, and for potential issues with the load and sync order in a more complex document context. |
| L1-find-replace-001 | 80 | ✓ | The code uses a correct API method replaceText, but lacks the Word.run wrapper and context setup, and does not fully follow best practices for completeness and approach. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach, such as utilizing the `range.font` properties directly on the search results, and for not handling potential errors that may occur during the execution of the `await context.sync()` calls. |
| L1-find-replace-regex-001 | 83 | ✓ | The code uses the correct Word Office.js API replaceText with matchWildcards, but loses points for not fully addressing potential edge cases and not using the most optimal approach, such as handling multiple search results or using a more specific wildcard pattern. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a correct API method to insert a footnote, but lacks specificity in targeting the current selection and does not follow best practices by directly calling the insertFootnote method without utilizing the provided helpers or ensuring the correct context. |
| L1-comment-insert-001 | 73 | ✓ | The code uses a correct API (insertComment) but lacks context about the range or location where the comment should be inserted, potentially leading to incorrect placement, and does not fully follow best practices for searching and inserting comments in a document. |
| L1-track-changes-toggle-001 | 95 ✅ | ✓ | The code perfectly addresses the request using the correct API, but loses points for approach because it manually calls `toggleTrackChanges` instead of using a more explicit method like `doc.changeTrackingMode`, although `toggleTrackChanges` is a valid helper function. |
| L1-track-changes-toggle-002 | 83 | ✓ | The code is mostly correct and complete, but it lacks the use of the Word.run method to execute the acceptAllRevisions action, and the approach could be improved by utilizing the available helpers, such as toggleTrackChanges, instead of a direct method call. |
| L1-toc-generate-001 | 83 | ✓ | The code correctly uses the insertTableOfContents API and addresses the user's request, but loses points for not fully utilizing best practices and available helpers, and for potential issues with variable scope and method signatures. |
| L1-section-break-001 | 72 | ✓ | The code uses correct APIs and is mostly complete, but lacks best practices and uses generic methods instead of available helpers, and also assumes the section break and page orientation methods are correctly implemented without considering the actual Word Office.js API specifics. |
| L1-header-footer-001 | 83 | ✓ | The code is mostly correct but loses points for not using the best practice approach, as it manually sets the header instead of potentially using a more specific helper method if available, and also assumes the existence of a setHeader function without clarifying its origin or ensuring it aligns with Word Office.js APIs. |
| L1-header-footer-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-columns-001 | 72 | ✓ | The code uses correct Word Office.js APIs and would likely execute without runtime errors, but it lacks completeness in addressing the user's request to switch the rest of the document into a two-column layout, and does not use best practices or available helpers. |
| L1-image-insert-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-image-insert-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-watermark-insert-001 | 40 ⚠️ | ✓ | The code uses a non-existent `insertWatermark` method, which is not a real Word Office.js API, and does not fully address the request to add a watermark to every page. |
| L1-content-control-001 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the request, but loses points for not using the Word.run method to ensure the code is executed in a context that supports asynchronous operations and for not fully following best practices. |
| L1-content-control-002 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the user's request, but loses points for not using the most efficient or best-practice approach, such as handling potential errors or using more specific methods. |
| L1-mail-merge-001 | 45 ⚠️ | ✓ | The code uses a hallucinated method "mailMergeReplace" which is not a real Word Office.js API, resulting in a significant penalty for API correctness. |
| L1-template-apply-001 | 80 | ✓ | The code correctly applies the resume template with the provided details, but loses points for not using the Word.run method to ensure the context is loaded before applying the template, and for not following best practices in terms of error handling and variable scope. |
| L1-template-apply-002 | 0 ⚠️ | ✓ | The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions. |
| L1-document-generate-001 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach, such as potentially using a loop to generate the sections instead of repeating the addHeading and addParagraph calls. |
| L1-theme-apply-001 | 5 ⚠️ | ✓ | The code heavily hallucinates non-existent APIs, specifically "designTheme" and "applyTheme", which are not real Word Office.js APIs, resulting in a significant penalty. |
| L1-theme-apply-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, resulting in a significant penalty for API correctness. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work at runtime. |
| L1-citation-bibliography-001 | 20 ⚠️ | ✓ | The code heavily hallucinates methods, specifically "insertCitation" and "insertBibliography", which are not real Word Office.js APIs, resulting in a significant penalty. |
| L1-equation-001 | 85 | ✓ | The code correctly uses the insertEquation API and would execute without runtime errors, but loses points for not fully addressing potential context or paragraph specifics and not using the most robust or best-practice approach to inserting equations in various document scenarios. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the user's request, but loses points for not using a more robust method to handle equation insertion, such as utilizing the range object to specify the exact location of the equation. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selection for clarity and provides relevant code, but loses points for not considering alternative approaches, such as recommending the coach panel, and for minor clarity issues in the code presentation. |
| L1-read-query-001 | 40 ⚠️ | ✓ | The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `countWords()` method, which heavily penalizes the API correctness score. |
| L1-read-query-002 | 40 ⚠️ | ✓ | The code hallucinates a non-existent `getReadability()` method, which is not a real Word Office.js API, resulting in a severe penalty for API correctness. |
| L1-read-query-003 | 0 ⚠️ | ✓ | The code hallucinates a non-existent API method "listHeadings()" which is not a real Word Office.js API, resulting in a complete failure across all dimensions. |
| L1-margins-orientation-001 | 83 | ✓ | The code correctly uses the Office.js API and addresses the request, but loses points for not using best practices and available helpers, and for potential issues with variable scope and method signatures. |
| L1-margins-orientation-002 | 80 | ✓ | The code is mostly correct but lacks the Word.run() context and may not fully address the request for the entire document, also it does not use the best practice of using Word built-in methods for page orientation changes. |
| L1-web-search-needed-001 | 60 | ✓ | The response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks clarity on how this information will be used to add a section to the Word document and does not provide any workarounds or suggestions. |
| L1-web-search-needed-002 | 87 | ✓ | The response accurately and clearly explains the Pythagorean theorem, directly addressing the question, but loses points for not providing additional workarounds or Word-specific caveats, and for including unnecessary code in the response. |
| L1-refusal-001 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not exploring alternative solutions or providing more detailed information about the "Save As" feature. |
| L1-refusal-002 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not offering more detailed alternatives or caveats, and for a slightly vague explanation. |
| L1-refusal-003 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js in creating Excel-style charts, but loses points for not providing a more detailed workaround or alternative solution that directly utilizes the provided data. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-paragraph-spacing-001 | 67 | ✓ | The code uses a manual loop and directly accesses paragraphFormat, but fails to utilize the provided helpers like applyStyle effectively and does not handle potential errors or dynamic paragraph counts. |
| L1-edge-case-001 | 15 ⚠️ | ✓ | The code fails to use the Word Office.js API to delete a paragraph and instead throws an error, also it does not attempt to guard against out-of-range paragraph numbers or use the available helpers. |
| L1-form-field-001 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the request, but loses points for not using a more descriptive title and tag, and for not handling potential errors or edge cases, which affects its completeness and adherence to best practices. |
| L1-text-insert-001 | 80 | ✓ | The code uses a real Office.js API (addParagraph) but loses points for not explicitly checking if the paragraph is being added at the end of the document, and for not using the most precise method to achieve the user's request, which would be to use range.insertText to insert the text at the end of the document body. |
| L1-text-edit-001 | 25 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a severe penalty for API correctness. |
| gen-L2-form-field-001 | 72 | ✓ | The code uses correct Office.js APIs and is mostly complete, but loses points for not fully addressing the request by including the label "I agree" as the title of the checkbox and not placing the date picker content control explicitly below the checkbox, and also for not using best practices in terms of layout and content control placement. |
| gen-L2-form-field-002 | 65 | ✗ | The code partially addresses the request but fails to accurately replace the existing plain text "Name: _______" and incorrectly uses a "comboBox" content control instead of a "buildingBlockGallery" content control, while also lacking proper error handling and context consideration. |
| gen-L2-form-field-003 | 72 | ✗ | The code uses real Word Office.js APIs and mostly addresses the request, but it fails to insert the content controls within a group content control as requested, and does not follow best practices for inserting content controls. |
| gen-L2-form-field-004 | 72 | ✗ | The code uses a non-existent `insertContentControl` method and `insertTable` method with incorrect parameters, but it attempts to address the request and would mostly work if the correct APIs were used. |

## ⚠️ Needs attention

**[L1-paragraph-format-001]** score=0 — The code uses a hallucinated method `addParagraph` which is not a real Word Office.js API, and also attempts to modify a paragraph's text and style in a way that is not supported by the API.
- missing patterns: `font.bold`, `font.size`

**[L1-list-bullet-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions.

**[L1-list-multilevel-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-header-footer-002]** score=20 — The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-image-insert-001]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-image-insert-002]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-watermark-insert-001]** score=40 — The code uses a non-existent `insertWatermark` method, which is not a real Word Office.js API, and does not fully address the request to add a watermark to every page.

**[L1-mail-merge-001]** score=45 — The code uses a hallucinated method "mailMergeReplace" which is not a real Word Office.js API, resulting in a significant penalty for API correctness.

**[L1-template-apply-002]** score=0 — The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions.

**[L1-theme-apply-001]** score=5 — The code heavily hallucinates non-existent APIs, specifically "designTheme" and "applyTheme", which are not real Word Office.js APIs, resulting in a significant penalty.

**[L1-theme-apply-002]** score=20 — The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, resulting in a significant penalty for API correctness.

**[L1-theme-apply-003]** score=20 — The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work at runtime.

**[L1-citation-bibliography-001]** score=20 — The code heavily hallucinates methods, specifically "insertCitation" and "insertBibliography", which are not real Word Office.js APIs, resulting in a significant penalty.

**[L1-read-query-001]** score=40 — The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `countWords()` method, which heavily penalizes the API correctness score.

**[L1-read-query-002]** score=40 — The code hallucinates a non-existent `getReadability()` method, which is not a real Word Office.js API, resulting in a severe penalty for API correctness.

**[L1-read-query-003]** score=0 — The code hallucinates a non-existent API method "listHeadings()" which is not a real Word Office.js API, resulting in a complete failure across all dimensions.

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-edge-case-001]** score=15 — The code fails to use the Word Office.js API to delete a paragraph and instead throws an error, also it does not attempt to guard against out-of-range paragraph numbers or use the available helpers.

**[L1-text-edit-001]** score=25 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a severe penalty for API correctness.
