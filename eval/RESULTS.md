# Word AI — Eval Results
**Last run:** 2026-05-25 17:47:24  
**Overall: 60.2/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 78.5/100 | — | 2 |
| paragraph-format | 41.5/100 | — | 2 |
| list-bullet | 25.0/100 | — | 1 |
| list-multilevel | 0.0/100 | — | 1 |
| style-apply | 80.0/100 | — | 1 |
| table-create | 83.0/100 | — | 1 |
| table-format | 67.0/100 | — | 1 |
| find-replace | 85.0/100 | — | 2 |
| find-replace-regex | 0.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 91.5/100 | — | 2 |
| toc-generate | 83.0/100 | — | 1 |
| section-break | 72.0/100 | — | 1 |
| header-footer | 69.5/100 | — | 2 |
| columns | 80.0/100 | — | 1 |
| image-insert | 76.0/100 | — | 2 |
| watermark-insert | 50.0/100 | — | 1 |
| content-control | 81.5/100 | — | 2 |
| mail-merge | 55.0/100 | — | 1 |
| template-apply | 83.0/100 | — | 2 |
| document-generate | 90.0/100 | — | 1 |
| theme-apply | 20.0/100 | — | 3 |
| citation-bibliography | 40.0/100 | — | 1 |
| equation | 41.5/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 32.3/100 | — | 3 |
| margins-orientation | 84.0/100 | — | 2 |
| web-search-needed | 74.0/100 | — | 2 |
| refusal | 74.0/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 0.0/100 | — | 1 |
| edge-cases | 25.0/100 | — | 1 |
| form-field | 76.4/100 | — | 5 |
| text-insert | 80.0/100 | — | 1 |
| text-edit | 0.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 85 | ✓ | The code is mostly correct but loses points for using a non-standard method `addHeading` which is not a real Word Office.js API, instead of using the available helpers or implementing it manually using `body.paragraphs.add` and `paragraph.styleBuiltIn`. |
| L1-heading-insert-002 | 72 | ✓ | The code uses correct APIs and is likely to work, but it lacks best practices and does not utilize available helpers, instead using generic methods like addHeading and addParagraph that may not be part of the official Word Office.js API. |
| L1-paragraph-format-001 | 0 ⚠️ | ✗ | The code uses a hallucinated method "addParagraph" which is not a real Word Office.js API, and also attempts to add a new paragraph instead of modifying the existing first paragraph. |
| L1-paragraph-format-002 | 83 | ✓ | The code is mostly correct but loses points for not using the applyStyle helper to set the font size to 28pt and instead directly modifying the font size property, and also for not checking if the first paragraph is indeed the title before applying the style. |
| L1-list-bullet-001 | 25 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a severe penalty for API correctness. |
| L1-list-multilevel-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-style-apply-001 | 80 | ✓ | The code correctly applies the Quote style to the third paragraph, but loses points for not using the most straightforward helper method and potential indexing issues, as paragraph indexing starts at 0, so the third paragraph should be indexed as 2, which is correct in this case, but may not be robust for all document structures. |
| L1-table-create-001 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the exact number of rows requested (4 rows with headers would be 3 rows of data plus 1 header row, so it should be a 4x3 table with only 3 data rows), and the approach could be improved by using more Word API helpers, such as setting the table style in a separate step. |
| L1-table-format-001 | 67 | ✓ | The code uses a non-standard `styleTable` method which is not a real Word Office.js API, and instead of using the available `applyStyle` helper or `paragraph.styleBuiltIn` property, it relies on a custom method that may not work as expected. |
| L1-find-replace-001 | 80 | ✓ | The code uses a correct API method replaceText, but lacks the surrounding Word.run context and may not fully address the request if there are multiple occurrences of "Acme" in different parts of the document, and does not follow best practices by not using Word.run to queue the change. |
| L1-find-replace-002 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient or best-practice approach, such as utilizing the `range.font` properties directly on the search results without explicitly loading the font property. |
| L1-find-replace-regex-001 | 0 ⚠️ | ✗ | The AI response failed to generate any code, making it impossible to evaluate its correctness, completeness, executability, or approach to solving the user's request. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a correct API method to insert a footnote, but lacks consideration for the current selection and paragraph context, and does not utilize the best practice helpers provided by the Office.js API. |
| L1-comment-insert-001 | 72 | ✓ | The code uses a correct API but lacks best practices and does not fully address the request as it does not specify the search context, potentially leading to incorrect comment insertion. |
| L1-track-changes-toggle-001 | 100 ✅ | ✓ | The code perfectly addresses the request to turn on track changes using the correct and available Word Office.js API, specifically the toggleTrackChanges method, with no flaws or unnecessary re-implementations. |
| L1-track-changes-toggle-002 | 83 | ✓ | The code is mostly correct and complete, but loses points for not using the Word.run method to execute the acceptAllRevisions action, and for not following best practices by directly calling acceptAllRevisions without wrapping it in a Word.run context. |
| L1-toc-generate-001 | 83 | ✓ | The code correctly uses the insertTableOfContents API and addresses the request, but loses points for not using the most efficient approach and potentially inserting a page break unnecessarily. |
| L1-section-break-001 | 72 | ✓ | The code uses correct APIs and addresses the request, but lacks best practices and does not fully handle the section break and page orientation change as it should apply the page orientation to the newly inserted section. |
| L1-header-footer-001 | 67 | ✓ | The code uses a hallucinated method "setHeader" which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-header-footer-002 | 72 | ✓ | The code uses a non-existent `addPageNumbers` method, which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-columns-001 | 80 | ✓ | The code is mostly correct but loses points for not fully addressing the request, as it only inserts a section break and switches to a two-column layout without ensuring the rest of the document is actually affected by this change. |
| L1-image-insert-001 | 72 | ✓ | The code uses a non-standard `insertImage` method which is not a real Word Office.js API, and also lacks proper error handling and context setup, but it attempts to address the user's request and has a correct sync order. |
| L1-image-insert-002 | 80 | ✓ | The code uses the correct insertImage API and addresses the request, but loses points for not using a more robust approach to handling image insertion, such as checking if the image is successfully inserted before syncing the context. |
| L1-watermark-insert-001 | 50 ⚠️ | ✓ | The code is incomplete and does not fully address the request as it only inserts a watermark without specifying the scope to every page, and also lacks proper API usage for watermark insertion. |
| L1-content-control-001 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the user's request, but loses points for not using a more specific best practice approach and potentially lacking error handling or context checks. |
| L1-content-control-002 | 80 | ✓ | The code correctly uses the insertContentControl API and addresses the request, but loses points for not using a more robust approach, such as checking the current selection or range before inserting the content control, and for not handling potential errors or exceptions. |
| L1-mail-merge-001 | 55 ⚠️ | ✓ | The code fails to use real Word Office.js APIs, specifically the mailMergeReplace method is not a valid API, resulting in a significant penalty for API correctness. |
| L1-template-apply-001 | 83 | ✓ | The code correctly applies the resume template with the provided details, but loses points for not using the most efficient or best practice approach, such as utilizing available helpers for inserting content, and also assumes the existence of a "resume" template without handling potential errors. |
| L1-template-apply-002 | 83 | ✓ | The code correctly applies the memo template with the specified fields, but loses points for not using Word.run to initiate the operation and for not handling potential errors, and also for not demonstrating the best approach by directly using the applyTemplate helper without considering other potential template application scenarios. |
| L1-document-generate-001 | 90 | ✓ | The code is mostly correct and complete, but loses points for not using the most efficient approach, such as potentially reusing the same paragraph or heading styles instead of re-creating them for each section, and not handling potential errors or edge cases. |
| L1-theme-apply-001 | 15 ⚠️ | ✓ | The code hallucinates non-existent `designTheme` and `applyTheme` methods, which are not part of the real Word Office.js APIs, resulting in a significant penalty for API correctness and would-it-work dimensions. |
| L1-theme-apply-002 | 25 ⚠️ | ✓ | The code heavily hallucinates the "designTheme" and "applyTheme" methods, which are not real Word Office.js APIs, resulting in a significant penalty for API correctness. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work at runtime. |
| L1-citation-bibliography-001 | 40 ⚠️ | ✓ | The code heavily hallucinates methods, such as insertCitation and insertBibliography, which are not real Word Office.js APIs, resulting in a significant penalty for API correctness. |
| L1-equation-001 | 0 ⚠️ | ✓ | The code hallucinates a non-existent `insertEquation` method, which is not a real Word Office.js API, resulting in a complete failure across all dimensions. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the user's request, but loses points for not using a more robust method to handle equation insertion, such as checking the current selection or range, and for not providing any error handling or feedback to the user. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selected text for clarity, but loses points for not considering the coach panel recommendation as an alternative and for using a generic phrase "Rewriting the selected text for clarity" that doesn't add much value. |
| L1-read-query-001 | 0 ⚠️ | ✓ | The code does not use any real Word Office.js APIs, instead calling a non-existent function `countWords()` which is not a valid method in the Word JavaScript API. |
| L1-read-query-002 | 57 ⚠️ | ✓ | The code uses a hallucinated method `getReadability()` which is not a real Word Office.js API, heavily penalizing API correctness, and also lacks best practices, such as using `Word.run` and `body` properties to access document content. |
| L1-read-query-003 | 40 ⚠️ | ✓ | The code fails to use real Word Office.js APIs, instead hallucinating a non-existent "listHeadings" method, which heavily penalizes its API correctness score. |
| L1-margins-orientation-001 | 85 | ✓ | The code correctly sets page margins using the Word Office.js API, but lacks best practices and does not utilize available helpers, and also assumes the context is already loaded without explicitly showing the Word.run call. |
| L1-margins-orientation-002 | 83 | ✓ | The code is mostly correct but lacks the Word.run context and does not fully follow best practices, such as using a more specific method to set the page orientation for the entire document. |
| L1-web-search-needed-001 | 67 | ✓ | The response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks clarity on how this information will be used to add a section to the Word document and does not provide any workarounds or Word-specific caveats. |
| L1-web-search-needed-002 | 81 | ✓ | The response accurately explains the Pythagorean theorem and provides relevant code, but lacks additional workarounds or Word-specific caveats to warrant full bonus points. |
| L1-refusal-001 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not exploring alternative solutions or providing more detailed information about the "Save As" feature. |
| L1-refusal-002 | 72 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but lacks a more detailed explanation or alternative solutions, resulting in a moderate score. |
| L1-refusal-003 | 67 | ✓ | The response accurately identifies the limitation of Word Office.js in creating charts, but lacks a clear and detailed workaround, resulting in a low bonus score. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-paragraph-spacing-001 | 0 ⚠️ | ✓ | The code hallucinates a non-existent `addParagraph` method and attempts to modify existing paragraphs by re-adding them, which is not a valid approach in the Word Office.js API. |
| L1-edge-case-001 | 25 ⚠️ | ✓ | The code fails to use the Word Office.js API correctly, instead throwing a generic error, and does not attempt to delete the paragraph or handle the out-of-range scenario using available API methods. |
| L1-form-field-001 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the user's request, but loses points for not using a more descriptive title and not handling potential errors, and also for not fully following best practices in terms of code structure and variable handling. |
| L1-text-insert-001 | 80 | ✓ | The code uses a real Word Office.js API (addParagraph) but loses points for not fully addressing the request of inserting text at the end of the document, as it doesn't explicitly ensure the paragraph is added at the end, and also for not using the most optimal approach by directly utilizing the paragraph.insertText method. |
| L1-text-edit-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| gen-L2-form-field-001 | 72 | ✓ | The code partially addresses the request but loses points for not exactly matching the requested label for the date picker content control and not utilizing best practices for inserting content controls in a specific location, such as below the checkbox. |
| gen-L2-form-field-002 | 72 | ✗ | The code uses the correct insertContentControl API but fails to fully address the request by not replacing the existing plain text "Name: _______" and incorrectly uses a comboBox content control instead of a building block gallery control. |
| gen-L2-form-field-003 | 83 | ✗ | The code is mostly correct and complete, but loses points for not using the Word.run method to wrap the content control insertion and for not handling potential errors, and also for not using the most efficient approach by directly using the insertContentControl method without utilizing available helpers. |
| gen-L2-form-field-004 | 72 | ✗ | The code partially addresses the request but fails to fully implement a repeating section content control, instead using a rich text content control, and lacks best practices in using available helpers. |

## ⚠️ Needs attention

**[L1-paragraph-format-001]** score=0 — The code uses a hallucinated method "addParagraph" which is not a real Word Office.js API, and also attempts to add a new paragraph instead of modifying the existing first paragraph.
- missing patterns: `font.bold`, `font.size`

**[L1-list-bullet-001]** score=25 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a severe penalty for API correctness.

**[L1-list-multilevel-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-find-replace-regex-001]** score=0 — The AI response failed to generate any code, making it impossible to evaluate its correctness, completeness, executability, or approach to solving the user's request.
- missing patterns: `matchWildcards`, `redacted`

**[L1-watermark-insert-001]** score=50 — The code is incomplete and does not fully address the request as it only inserts a watermark without specifying the scope to every page, and also lacks proper API usage for watermark insertion.

**[L1-mail-merge-001]** score=55 — The code fails to use real Word Office.js APIs, specifically the mailMergeReplace method is not a valid API, resulting in a significant penalty for API correctness.

**[L1-theme-apply-001]** score=15 — The code hallucinates non-existent `designTheme` and `applyTheme` methods, which are not part of the real Word Office.js APIs, resulting in a significant penalty for API correctness and would-it-work dimensions.

**[L1-theme-apply-002]** score=25 — The code heavily hallucinates the "designTheme" and "applyTheme" methods, which are not real Word Office.js APIs, resulting in a significant penalty for API correctness.

**[L1-theme-apply-003]** score=20 — The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work at runtime.

**[L1-citation-bibliography-001]** score=40 — The code heavily hallucinates methods, such as insertCitation and insertBibliography, which are not real Word Office.js APIs, resulting in a significant penalty for API correctness.

**[L1-equation-001]** score=0 — The code hallucinates a non-existent `insertEquation` method, which is not a real Word Office.js API, resulting in a complete failure across all dimensions.

**[L1-read-query-001]** score=0 — The code does not use any real Word Office.js APIs, instead calling a non-existent function `countWords()` which is not a valid method in the Word JavaScript API.

**[L1-read-query-002]** score=57 — The code uses a hallucinated method `getReadability()` which is not a real Word Office.js API, heavily penalizing API correctness, and also lacks best practices, such as using `Word.run` and `body` properties to access document content.

**[L1-read-query-003]** score=40 — The code fails to use real Word Office.js APIs, instead hallucinating a non-existent "listHeadings" method, which heavily penalizes its API correctness score.

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-paragraph-spacing-001]** score=0 — The code hallucinates a non-existent `addParagraph` method and attempts to modify existing paragraphs by re-adding them, which is not a valid approach in the Word Office.js API.

**[L1-edge-case-001]** score=25 — The code fails to use the Word Office.js API correctly, instead throwing a generic error, and does not attempt to delete the paragraph or handle the out-of-range scenario using available API methods.

**[L1-text-edit-001]** score=0 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.
