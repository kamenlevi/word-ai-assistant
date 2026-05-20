# Word AI — Eval Results
**Last run:** 2026-05-20 03:40:09  
**Overall: 55.4/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 78.5/100 | — | 2 |
| paragraph-format | 0.0/100 | — | 2 |
| list-bullet | 20.0/100 | — | 1 |
| list-multilevel | 0.0/100 | — | 1 |
| style-apply | 83.0/100 | — | 1 |
| table-create | 83.0/100 | — | 1 |
| table-format | 72.0/100 | — | 1 |
| find-replace | 85.0/100 | — | 2 |
| find-replace-regex | 83.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 91.5/100 | — | 2 |
| toc-generate | 83.0/100 | — | 1 |
| section-break | 72.0/100 | — | 1 |
| header-footer | 40.0/100 | — | 2 |
| columns | 72.0/100 | — | 1 |
| image-insert | 50.0/100 | — | 2 |
| watermark-insert | 0.0/100 | — | 1 |
| content-control | 84.0/100 | — | 2 |
| mail-merge | 67.0/100 | — | 1 |
| template-apply | 54.0/100 | — | 2 |
| document-generate | 85.0/100 | — | 1 |
| theme-apply | 17.3/100 | — | 3 |
| citation-bibliography | 20.0/100 | — | 1 |
| equation | 77.5/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 32.3/100 | — | 3 |
| margins-orientation | 41.5/100 | — | 2 |
| web-search-needed | 66.0/100 | — | 2 |
| refusal | 74.0/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 67.0/100 | — | 1 |
| edge-cases | 50.0/100 | — | 1 |
| form-field | 58.4/100 | — | 5 |
| text-insert | 72.0/100 | — | 1 |
| text-edit | 0.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 85 | ✓ | The code is mostly correct but loses points for using a non-standard method signature for adding a heading, which may not be a real Word Office.js API, and not utilizing the best practice approach of using the `paragraph.styleBuiltIn` property to apply the Heading 1 style. |
| L1-heading-insert-002 | 72 | ✓ | The code uses non-existent helper functions like addHeading and addParagraph, which are not part of the Word Office.js API, and lacks proper context and range handling. |
| L1-paragraph-format-001 | 0 ⚠️ | ✗ | The code uses a hallucinated method `addParagraph` which is not a real Word Office.js API, and also modifies the first paragraph's text instead of its style, resulting in a complete failure to address the user's request. |
| L1-paragraph-format-002 | 0 ⚠️ | ✓ | A:  |
| L1-list-bullet-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice. |
| L1-list-multilevel-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-style-apply-001 | 83 | ✓ | The code correctly applies the Quote style to the third paragraph, but loses points for not using the more specific and recommended `paragraph.styleBuiltIn` property instead of the generic `applyStyle` method, and also for not handling potential errors or edge cases. |
| L1-table-create-001 | 83 | ✓ | The code is mostly correct but loses points for not fully addressing the request by only populating three of the four rows in the table as specified by the user. |
| L1-table-format-001 | 72 | ✓ | The code uses a non-existent `styleTable` method, which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-find-replace-001 | 80 | ✓ | The code uses a correct API method replaceText, but lacks the surrounding Word.run context and may not fully handle all edge cases, such as multiple occurrences of 'Acme' in a single paragraph or in different paragraphs. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient or best-practice approach, such as utilizing the `range.font` properties directly on the search results without explicitly loading the font property. |
| L1-find-replace-regex-001 | 83 | ✓ | The code correctly uses the Word Office.js API and addresses the request, but loses points for not handling potential errors and not using the most efficient approach, such as checking if the replacement is necessary before calling replaceText. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a correct API method insertFootnote but lacks context about the current selection and does not fully follow best practices, as it does not explicitly use the range object to insert the footnote at the correct location. |
| L1-comment-insert-001 | 72 | ✓ | The code uses a correct API method to insert a comment, but lacks specificity in targeting the exact occurrence of the text and does not fully utilize best practices for searching and inserting comments in the document. |
| L1-track-changes-toggle-001 | 100 ✅ | ✓ | The code perfectly addresses the request to turn on track changes using the correct and existing Word Office.js API method toggleTrackChanges, making it a complete, correct, and well-structured solution. |
| L1-track-changes-toggle-002 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the Word.run method to wrap the API calls, and for not demonstrating best practices in using available helpers, such as toggleTrackChanges to accept all revisions. |
| L1-toc-generate-001 | 83 | ✓ | The code correctly uses the insertTableOfContents API and addresses the request, but loses points for not using the most specific and efficient approach, such as specifying the exact heading levels present in the document, and for not handling potential errors or edge cases. |
| L1-section-break-001 | 72 | ✓ | The code uses correct APIs and is likely to work, but it lacks completeness in addressing the user's request to switch the next section to landscape, and does not follow best practices in using available helpers. |
| L1-header-footer-001 | 80 | ✓ | The code is mostly correct but lacks consideration for the specific section or page range where the header should be applied, and uses a generic `setHeader` method which might not be a real Word Office.js API, deducting points for API correctness and completeness. |
| L1-header-footer-002 | 0 ⚠️ | ✗ | The code uses a hallucinated method "setFooter" which is not a real Word Office.js API, resulting in a complete failure across all dimensions. |
| L1-columns-001 | 72 | ✓ | The code uses correct APIs and would likely work, but it lacks completeness in addressing the user's request to switch the rest of the document into a two-column layout, as it doesn't handle the current section's layout, and the approach could be improved by using more specific methods to target the rest of the document. |
| L1-image-insert-001 | 80 | ✓ | The code correctly uses the insertImage API and addresses the request, but loses points for not specifying the exact paragraph where the image should be inserted and potential issues with the context.sync() order. |
| L1-image-insert-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-watermark-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "insertWatermark" which is not a real Word Office.js API, and also does not address the request to add the watermark to every page. |
| L1-content-control-001 | 85 | ✓ | The code correctly uses the Office.js API to insert a content control, but loses points for not using the most specific and efficient methods, such as specifying the location of the insertion, and not handling potential errors that may occur during the execution of the code. |
| L1-content-control-002 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the Word.run method to execute the content control insertion, and for not handling potential errors or loading the document context before inserting the content control. |
| L1-mail-merge-001 | 67 | ✓ | The code uses a hallucinated method mailMergeReplace, which is not a real Word Office.js API, resulting in significant penalties in API correctness and approach dimensions. |
| L1-template-apply-001 | 83 | ✓ | The code is mostly correct but loses points for not using the Word.run method to execute the API calls and for not handling potential errors, and also for not using the best practice of using Word.run to apply the template. |
| L1-template-apply-002 | 25 ⚠️ | ✓ | The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-document-generate-001 | 85 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach, such as utilizing the `addSection` or `insertBreak` methods to create distinct sections, and not fully leveraging the available helpers to simplify the code. |
| L1-theme-apply-001 | 12 ⚠️ | ✓ | The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would_it_work dimensions. |
| L1-theme-apply-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not a valid method, resulting in a complete loss of points for API correctness and would it work dimensions. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code hallucinates a non-existent `tweakTheme` method, which is not a real Word Office.js API, resulting in a severe penalty for API correctness. |
| L1-citation-bibliography-001 | 20 ⚠️ | ✓ | The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-equation-001 | 72 | ✓ | The code uses a non-existent `insertEquation` method, which is not a real Word Office.js API, resulting in significant penalties in API correctness and approach dimensions. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation helper, but loses points for not explicitly checking the cursor position before inserting the equation, and not using a more robust method to handle potential errors or edge cases. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selection for clarity, but loses points for not considering alternative approaches, such as recommending the coach panel, and for minor clarity issues in the code comments. |
| L1-read-query-001 | 40 ⚠️ | ✓ | The code fails to utilize real Word Office.js APIs, instead relying on a hallucinated `countWords()` method, which is not a valid API call. |
| L1-read-query-002 | 57 ⚠️ | ✓ | The code uses a hallucinated method `getReadability()` which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-read-query-003 | 0 ⚠️ | ✓ | The code hallucinates a non-existent API method "listHeadings()" which is not a real Word Office.js API, resulting in a complete failure across all dimensions. |
| L1-margins-orientation-001 | 83 | ✓ | The code correctly uses the Office.js API and addresses the request, but loses points for not using a more explicit best practice approach, such as using a predefined unit for the margin size, and for not handling potential errors or edge cases. |
| L1-margins-orientation-002 | 0 ⚠️ | ✓ | The code uses a hallucinated method "setPageOrientation" which is not a real Word Office.js API, resulting in a complete failure across all dimensions. |
| L1-web-search-needed-001 | 67 | ✓ | The response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks clarity on how this information will be used to add a section to the Word document and does not provide any workarounds or Word-specific caveats. |
| L1-web-search-needed-002 | 65 | ✓ | The response accurately and clearly explains the Pythagorean theorem, but loses points for inserting unnecessary code and not providing any workarounds or Word-specific caveats, making the bonus score zero. |
| L1-refusal-001 | 72 | ✓ | The response is mostly accurate and clear, but loses points for relevance due to the brief mention of VBA and external libraries without elaboration, and for bonus points due to the lack of a more detailed workaround or alternative solution. |
| L1-refusal-002 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not offering more detailed or alternative solutions, such as using other APIs or third-party libraries. |
| L1-refusal-003 | 67 | ✓ | The response accurately identifies the limitation of Word Office.js in creating charts, but loses points for not providing a more detailed workaround or alternative solution that directly utilizes the provided data to generate a chart image. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-paragraph-spacing-001 | 67 | ✓ | The code uses a correct API but lacks best practices, such as using a loop based on the actual number of paragraphs in the document instead of hardcoding the number, and incorrectly applies a style before setting the line spacing, which is not necessary and may alter the paragraph's original style. |
| L1-edge-case-001 | 50 ⚠️ | ✓ | The code fails to use any real Word Office.js APIs, instead simply stating that no changes are needed, which results in a severe penalty for API correctness. |
| L1-form-field-001 | 83 | ✓ | The code correctly uses the Office.js API to insert a dropdown content control, but loses points for not fully addressing potential edge cases and not using the most idiomatic approach to inserting content controls. |
| L1-text-insert-001 | 72 | ✓ | The code uses a non-standard method `addParagraph` which is not a real Word Office.js API, and instead of using `body.paragraphs.add` or `range.insertText`, it relies on a helper function that may not be defined or may not work as expected. |
| L1-text-edit-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure across all dimensions. |
| gen-L2-form-field-001 | 72 | ✓ | The code uses correct Word Office.js APIs and is likely to execute without runtime errors, but it does not fully address the request as it inserts a date picker with the title "Sign Date" instead of placing it below the checkbox content control with the correct label, and also lacks best practices in content control insertion. |
| gen-L2-form-field-002 | 72 | ✗ | The code partially addresses the request but lacks proper integration of the rich text content control with the replacement of the existing plain text, and it incorrectly uses a comboBox content control instead of a building block gallery control. |
| gen-L2-form-field-003 | 65 | ✗ | The code fails to fully address the request by not inserting the dropdown content controls within a group content control, resulting in a partial solution that lacks the required structure. |
| gen-L2-form-field-004 | 0 ⚠️ | ✗ | A:  |

## ⚠️ Needs attention

**[L1-paragraph-format-001]** score=0 — The code uses a hallucinated method `addParagraph` which is not a real Word Office.js API, and also modifies the first paragraph's text instead of its style, resulting in a complete failure to address the user's request.
- missing patterns: `font.bold`, `font.size`

**[L1-paragraph-format-002]** score=0 — A: 

**[L1-list-bullet-001]** score=20 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in practice.

**[L1-list-multilevel-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-header-footer-002]** score=0 — The code uses a hallucinated method "setFooter" which is not a real Word Office.js API, resulting in a complete failure across all dimensions.
- missing patterns: `addPageNumbers`

**[L1-image-insert-002]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-watermark-insert-001]** score=0 — The code uses a hallucinated method "insertWatermark" which is not a real Word Office.js API, and also does not address the request to add the watermark to every page.

**[L1-template-apply-002]** score=25 — The code uses a hallucinated method "applyTemplate" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-theme-apply-001]** score=12 — The code uses hallucinated methods "designTheme" and "applyTheme" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would_it_work dimensions.

**[L1-theme-apply-002]** score=20 — The code uses a hallucinated method "designTheme" which is not a real Word Office.js API, and also "applyTheme" is not a valid method, resulting in a complete loss of points for API correctness and would it work dimensions.

**[L1-theme-apply-003]** score=20 — The code hallucinates a non-existent `tweakTheme` method, which is not a real Word Office.js API, resulting in a severe penalty for API correctness.

**[L1-citation-bibliography-001]** score=20 — The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-read-query-001]** score=40 — The code fails to utilize real Word Office.js APIs, instead relying on a hallucinated `countWords()` method, which is not a valid API call.

**[L1-read-query-002]** score=57 — The code uses a hallucinated method `getReadability()` which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions.

**[L1-read-query-003]** score=0 — The code hallucinates a non-existent API method "listHeadings()" which is not a real Word Office.js API, resulting in a complete failure across all dimensions.

**[L1-margins-orientation-002]** score=0 — The code uses a hallucinated method "setPageOrientation" which is not a real Word Office.js API, resulting in a complete failure across all dimensions.

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-edge-case-001]** score=50 — The code fails to use any real Word Office.js APIs, instead simply stating that no changes are needed, which results in a severe penalty for API correctness.

**[L1-text-edit-001]** score=0 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure across all dimensions.

**[gen-L2-form-field-004]** score=0 — A: 
- missing patterns: `RepeatingSectionContentControl`
