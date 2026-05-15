# Word AI — Eval Results
**Last run:** 2026-05-15 17:32:32  
**Overall: 54.1/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 42.5/100 | — | 2 |
| paragraph-format | 44.5/100 | — | 2 |
| list-bullet | 25.0/100 | — | 1 |
| list-multilevel | 0.0/100 | — | 1 |
| style-apply | 80.0/100 | — | 1 |
| table-create | 83.0/100 | — | 1 |
| table-format | 0.0/100 | — | 1 |
| find-replace | 81.0/100 | — | 2 |
| find-replace-regex | 83.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 54.0/100 | — | 2 |
| toc-generate | 72.0/100 | — | 1 |
| section-break | 72.0/100 | — | 1 |
| header-footer | 77.5/100 | — | 2 |
| columns | 72.0/100 | — | 1 |
| image-insert | 51.5/100 | — | 2 |
| watermark-insert | 0.0/100 | — | 1 |
| content-control | 83.0/100 | — | 2 |
| mail-merge | 18.0/100 | — | 1 |
| template-apply | 76.0/100 | — | 2 |
| document-generate | 83.0/100 | — | 1 |
| theme-apply | 15.0/100 | — | 3 |
| citation-bibliography | 40.0/100 | — | 1 |
| equation | 77.5/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 26.7/100 | — | 3 |
| margins-orientation | 81.5/100 | — | 2 |
| web-search-needed | 71.5/100 | — | 2 |
| refusal | 75.7/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 88.0/100 | — | 1 |
| edge-cases | 5.0/100 | — | 1 |
| form-field | 49.0/100 | — | 5 |
| text-insert | 60.0/100 | — | 1 |
| text-edit | 0.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 85 | ✓ | The code is mostly correct but loses points for using a hypothetical `addHeading` method which is not a real Word Office.js API, instead of using the actual `paragraph.styleBuiltIn` and `paragraph.insertText` methods. |
| L1-heading-insert-002 | 0 ⚠️ | ✓ | The code uses hallucinated methods "addHeading" and "addParagraph" which are not real Word Office.js APIs, resulting in a complete failure to meet the requirements. |
| L1-paragraph-format-001 | 0 ⚠️ | ✗ | The code uses a hallucinated method `addParagraph` which does not exist in the Word Office.js API, and also attempts to modify a paragraph's text and style in a way that is not supported by the API. |
| L1-paragraph-format-002 | 89 | ✓ | The code is mostly correct and would work, but it lacks the use of best practices, such as using the `paragraph.styleBuiltIn` or `paragraph.font` helpers, and assumes the first run in the paragraph is the title, which might not always be the case. |
| L1-list-bullet-001 | 25 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a severe penalty for API correctness. |
| L1-list-multilevel-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-style-apply-001 | 80 | ✓ | The code correctly applies the Quote style to the third paragraph using the applyStyle method, but loses points for not using the Word.run method to ensure context and for not handling potential errors, and also for not using the most idiomatic approach to paragraph indexing. |
| L1-table-create-001 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the `Word.run` method to execute the insertion of the table, and for not fully addressing the request by not handling potential errors or edge cases, and for not using the most optimal approach by directly inserting the table without considering the document context. |
| L1-table-format-001 | 0 ⚠️ | ✓ | The code fails to apply the style to the existing table, instead attempting to insert a new table with the specified style, which does not address the user's request. |
| L1-find-replace-001 | 72 | ✓ | The code uses a correct API method replaceText, but lacks the context of a Word.run call and does not handle potential errors, also it does not follow best practices as it does not utilize the Word.run method which is the entry point for most Word JavaScript API operations. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient or best practice approach, such as utilizing the `range.font` properties directly on the search results without explicitly loading the "font" property. |
| L1-find-replace-regex-001 | 83 | ✓ | The code correctly uses the Word Office.js API and addresses the request, but loses points for not handling potential errors and not using the most efficient approach, such as checking if the replacement is necessary before calling replaceText. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a correct API method (insertFootnote) but lacks specificity in targeting the current selection (paragraph 2) and does not fully follow best practices by not utilizing the range object to specify the insertion point. |
| L1-comment-insert-001 | 72 | ✓ | The code uses a correct API method `insertComment` but lacks specificity in targeting the first occurrence of "$4.2M" and does not fully address the request by not ensuring the comment is added to the correct location within the document context. |
| L1-track-changes-toggle-001 | 25 ⚠️ | ✓ | A: 25
B: |
| L1-track-changes-toggle-002 | 83 | ✓ | The code is mostly correct and complete, but it lacks the use of the Word.run method to execute the acceptAllRevisions action, which is a best practice in Office.js APIs. |
| L1-toc-generate-001 | 72 | ✓ | The code uses a correct API method to insert a Table of Contents, but lacks best practices and does not fully utilize available helpers, and the `minLevel` and `maxLevel` parameters may not perfectly match the user's request. |
| L1-section-break-001 | 72 | ✓ | The code uses correct APIs and is mostly complete, but lacks best practices and does not fully address the request as it doesn't specify the section to apply the landscape orientation to, and also doesn't handle potential errors or edge cases. |
| L1-header-footer-001 | 83 | ✓ | The code is mostly correct, but it uses a generic `setHeader` method which is not a real Word Office.js API, instead of using the `section.getHeader` and `header.insertText` methods, and also lacks error handling and context loading checks. |
| L1-header-footer-002 | 72 | ✗ | The code uses a non-existent setFooter method, which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-columns-001 | 72 | ✓ | The code uses correct APIs and would likely work, but it only partially addresses the request by not specifying the scope of the section break and column insertion, and does not follow best practices by not using a more targeted approach to apply the two-column layout to the rest of the document. |
| L1-image-insert-001 | 83 | ✓ | The code correctly uses the insertImage API and addresses the user's request, but loses points for not explicitly specifying the paragraph where the image should be inserted, and for not using the most robust approach to handle potential errors or edge cases. |
| L1-image-insert-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice. |
| L1-watermark-insert-001 | 0 ⚠️ | ✓ | The code is completely incorrect as it uses a non-existent `insertWatermark` method, which is not a part of the Word Office.js API, and does not address the request to add a watermark to every page. |
| L1-content-control-001 | 83 | ✓ | The code correctly uses the Office.js API to insert a content control, but loses points for not fully addressing potential context and scope issues, and not using the most optimal approach by directly utilizing the provided helpers without additional error handling or checks. |
| L1-content-control-002 | 83 | ✓ | The code correctly uses the Office.js API to insert a checkbox content control, but loses points for not handling potential errors and not using the most efficient approach, such as checking if the content control already exists before inserting a new one. |
| L1-mail-merge-001 | 18 ⚠️ | ✓ | A: 18
B: |
| L1-template-apply-001 | 72 | ✓ | The code uses a correct API method applyTemplate but loses points for not using Word.run to execute the API call and for not checking if the template exists, and also for not following best practices in terms of error handling and variable scope. |
| L1-template-apply-002 | 80 | ✓ | The code correctly applies the memo template with the provided details, but lacks error handling and does not utilize the Word.run method, which is a best practice for executing asynchronous operations in Word Office.js APIs. |
| L1-document-generate-001 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach, such as potentially re-implementing the table of contents logic instead of relying on the built-in insertTableOfContents method, and not fully utilizing the available helpers for adding headings and paragraphs. |
| L1-theme-apply-001 | 5 ⚠️ | ✓ | The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would_it_work dimensions. |
| L1-theme-apply-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not a valid method, resulting in a complete loss of points for API correctness and would_it_work dimensions. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work at runtime. |
| L1-citation-bibliography-001 | 40 ⚠️ | ✓ | The code uses hallucinated methods `insertCitation` and `insertBibliography` which are not real Word Office.js APIs, resulting in a significant penalty for API correctness. |
| L1-equation-001 | 72 | ✓ | The code uses a non-existent `insertEquation` method, which is not a real Word Office.js API, resulting in significant penalties in API correctness and approach dimensions. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation method, a real Word Office.js API, but loses points for not fully addressing potential context or user-specific requests and not using the most robust or flexible approach to inserting equations. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selected text for clarity and provides relevant code, but loses points for not offering alternative solutions or workarounds, such as recommending the coach panel, and for minor clarity issues in the code comments. |
| L1-read-query-001 | 40 ⚠️ | ✓ | The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `countWords()` method, which heavily penalizes the API correctness score. |
| L1-read-query-002 | 40 ⚠️ | ✓ | The code fails to utilize real Word Office.js APIs, instead relying on a hallucinated `getReadability()` method, which is not a valid API call in the Office.js library. |
| L1-read-query-003 | 0 ⚠️ | ✓ | The code hallucinates a non-existent `listHeadings()` method, which is not a real Word Office.js API, and does not provide any actual implementation to list the headings. |
| L1-margins-orientation-001 | 83 | ✓ | The code correctly uses the Word Office.js API and addresses the request, but loses points for not using a more explicit best practice approach, such as using a predefined unit type for margin settings, and for potential issues with variable scope and method signatures. |
| L1-margins-orientation-002 | 80 | ✓ | The code correctly uses the Office.js API and addresses the user's request, but lacks best practices and does not utilize available helpers, and also assumes the existence of a setPageOrientation method which is not a standard Word Office.js API. |
| L1-web-search-needed-001 | 67 | ✓ | The AI response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks clarity on how this information will be used to add a section to the Word document and does not provide any workarounds or suggestions. |
| L1-web-search-needed-002 | 76 | ✓ | The response accurately and clearly explains the Pythagorean theorem, directly addressing the question, but lacks significant bonus points for not providing additional workarounds, settings options, or Word-specific caveats beyond the basic equation insertion. |
| L1-refusal-001 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not offering a more detailed or alternative solution, such as using other APIs or add-ins to achieve the PDF export. |
| L1-refusal-002 | 72 | ✓ | The response accurately conveys the limitation of Word Office.js and provides a clear workaround, but loses points for not offering more detailed alternatives or settings options, and for including redundant information in the code snippet. |
| L1-refusal-003 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js in creating charts, but loses points for not providing a more detailed workaround that directly addresses the user's request for an Excel-style bar chart from specific data. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, and does not provide any actual implementation to insert a block quote. |
| L1-paragraph-spacing-001 | 88 | ✓ | The code is mostly correct and complete, but loses points for not using a more dynamic approach to get the number of paragraphs, and for not handling potential errors or edge cases, such as an empty document or paragraphs with no text. |
| L1-edge-case-001 | 5 ⚠️ | ✓ | The code does not use any real Word Office.js APIs and does not attempt to delete the paragraph, instead simply stating that no action will be taken, which is a significant flaw in addressing the user's request. |
| L1-form-field-001 | 93 | ✓ | The code accurately uses the Word Office.js API to insert a dropdown content control, but may not handle potential asynchronous execution issues and does not explicitly check the cursor position in the body. |
| L1-text-insert-001 | 60 | ✓ | The code uses a real Office.js API (addParagraph) and isA to work, but loses likely points for not fully addressing the request (inserting at not just the end of the document, adding a new paragraph) and not using the most optimal approach (e.g., using body.paragraphs.getFirst() and then inserting after it). |
| L1-text-edit-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| gen-L2-form-field-001 | 72 | ✓ | The code correctly uses the insertContentControl API and executes without runtime errors, but loses points for not fully addressing the request by inserting the content controls below the existing text and not using best practices to position the controls. |
| gen-L2-form-field-002 | 30 ⚠️ | ✗ | The code fails to use real Word Office.js APIs, instead hallucinating non-existent methods like insertContentControl, and does not fully address the request by not replacing the existing plain text and not using the correct API methods. |
| gen-L2-form-field-003 | 40 ⚠️ | ✗ | The code fails to use the correct API to insert a group content control containing two dropdown content controls, instead inserting separate content controls for rich text, color, and size, and does not utilize the available helpers for inserting content controls. |
| gen-L2-form-field-004 | 10 ⚠️ | ✗ | The code hallucinates non-existent `insertContentControl` and `insertTable` methods, which are not part of the real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions. |

## ⚠️ Needs attention

**[L1-heading-insert-002]** score=0 — The code uses hallucinated methods "addHeading" and "addParagraph" which are not real Word Office.js APIs, resulting in a complete failure to meet the requirements.

**[L1-paragraph-format-001]** score=0 — The code uses a hallucinated method `addParagraph` which does not exist in the Word Office.js API, and also attempts to modify a paragraph's text and style in a way that is not supported by the API.
- missing patterns: `font.bold`, `font.size`

**[L1-list-bullet-001]** score=25 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a severe penalty for API correctness.

**[L1-list-multilevel-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-table-format-001]** score=0 — The code fails to apply the style to the existing table, instead attempting to insert a new table with the specified style, which does not address the user's request.

**[L1-track-changes-toggle-001]** score=25 — A: 25
B:

**[L1-image-insert-002]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice.

**[L1-watermark-insert-001]** score=0 — The code is completely incorrect as it uses a non-existent `insertWatermark` method, which is not a part of the Word Office.js API, and does not address the request to add a watermark to every page.

**[L1-mail-merge-001]** score=18 — A: 18
B:

**[L1-theme-apply-001]** score=5 — The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would_it_work dimensions.

**[L1-theme-apply-002]** score=20 — The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not a valid method, resulting in a complete loss of points for API correctness and would_it_work dimensions.

**[L1-theme-apply-003]** score=20 — The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work at runtime.

**[L1-citation-bibliography-001]** score=40 — The code uses hallucinated methods `insertCitation` and `insertBibliography` which are not real Word Office.js APIs, resulting in a significant penalty for API correctness.

**[L1-read-query-001]** score=40 — The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `countWords()` method, which heavily penalizes the API correctness score.

**[L1-read-query-002]** score=40 — The code fails to utilize real Word Office.js APIs, instead relying on a hallucinated `getReadability()` method, which is not a valid API call in the Office.js library.

**[L1-read-query-003]** score=0 — The code hallucinates a non-existent `listHeadings()` method, which is not a real Word Office.js API, and does not provide any actual implementation to list the headings.

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, and does not provide any actual implementation to insert a block quote.

**[L1-edge-case-001]** score=5 — The code does not use any real Word Office.js APIs and does not attempt to delete the paragraph, instead simply stating that no action will be taken, which is a significant flaw in addressing the user's request.

**[L1-text-edit-001]** score=0 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[gen-L2-form-field-002]** score=30 — The code fails to use real Word Office.js APIs, instead hallucinating non-existent methods like insertContentControl, and does not fully address the request by not replacing the existing plain text and not using the correct API methods.
- missing patterns: `RichTextContentControl`

**[gen-L2-form-field-003]** score=40 — The code fails to use the correct API to insert a group content control containing two dropdown content controls, instead inserting separate content controls for rich text, color, and size, and does not utilize the available helpers for inserting content controls.
- missing patterns: `GroupContentControl`

**[gen-L2-form-field-004]** score=10 — The code hallucinates non-existent `insertContentControl` and `insertTable` methods, which are not part of the real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions.
- missing patterns: `RepeatingSectionContentControl`
