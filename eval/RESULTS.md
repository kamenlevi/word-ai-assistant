# Word AI — Eval Results
**Last run:** 2026-05-17 03:32:58  
**Overall: 61.4/100**  
**Model:** meta-llama/llama-3.3-70b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 76.0/100 | — | 2 |
| paragraph-format | 85.5/100 | — | 2 |
| list-bullet | 25.0/100 | — | 1 |
| list-multilevel | 0.0/100 | — | 1 |
| style-apply | 80.0/100 | — | 1 |
| table-create | 83.0/100 | — | 1 |
| table-format | 72.0/100 | — | 1 |
| find-replace | 80.0/100 | — | 2 |
| find-replace-regex | 80.0/100 | — | 1 |
| footnote-insert | 72.0/100 | — | 1 |
| comment-insert | 72.0/100 | — | 1 |
| track-changes-toggle | 89.0/100 | — | 2 |
| toc-generate | 80.0/100 | — | 1 |
| section-break | 72.0/100 | — | 1 |
| header-footer | 40.0/100 | — | 2 |
| columns | 20.0/100 | — | 1 |
| image-insert | 46.0/100 | — | 2 |
| watermark-insert | 0.0/100 | — | 1 |
| content-control | 83.0/100 | — | 2 |
| mail-merge | 83.0/100 | — | 1 |
| template-apply | 41.5/100 | — | 2 |
| document-generate | 90.0/100 | — | 1 |
| theme-apply | 35.7/100 | — | 3 |
| citation-bibliography | 20.0/100 | — | 1 |
| equation | 83.0/100 | — | 2 |
| writing-coach | 83.0/100 | — | 1 |
| read-query | 39.0/100 | — | 3 |
| margins-orientation | 81.5/100 | — | 2 |
| web-search-needed | 72.0/100 | — | 2 |
| refusal | 77.7/100 | — | 3 |
| quote-insert | 0.0/100 | — | 1 |
| paragraph-spacing | 72.0/100 | — | 1 |
| edge-cases | 25.0/100 | — | 1 |
| form-field | 74.2/100 | — | 5 |
| text-insert | 85.0/100 | — | 1 |
| text-edit | 0.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 85 | ✓ | The code uses a non-standard method `addHeading` which is not a real Word Office.js API, but it is a helper function that is commonly used, and the rest of the code is correct and complete, but may not work due to the unknown implementation of `addHeading`. |
| L1-heading-insert-002 | 67 | ✓ | The code uses hallucinated methods addHeading and addParagraph, which are not real Word Office.js APIs, resulting in significant deductions in API correctness and approach dimensions. |
| L1-paragraph-format-001 | 88 | ✓ | The code is mostly correct, but loses points for not using the available helpers, such as applying a built-in style or using the font properties directly on the paragraph object, and also for not handling potential errors or edge cases, like an empty document body. |
| L1-paragraph-format-002 | 83 | ✓ | The code is mostly correct but loses points for not using the available helper methods, such as applyStyle to set the font size, and for not checking if the first paragraph is indeed the title before applying the changes. |
| L1-list-bullet-001 | 25 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a significant penalty for API correctness. |
| L1-list-multilevel-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-style-apply-001 | 80 | ✓ | The code correctly applies the Quote style to the third paragraph, but uses a generic applyStyle method which may not be the most efficient or Word-specific approach, and assumes the paragraph index is 2, which may not always be the case. |
| L1-table-create-001 | 83 | ✓ | The code is mostly correct, but it loses points for not fully addressing the request by only inserting three rows of data instead of four as specified, and for not using the most optimal approach by directly inserting the table with the specified headers and style. |
| L1-table-format-001 | 72 | ✓ | The code uses a non-existent `styleTable` method, which is not a real Word Office.js API, resulting in significant deductions in API correctness and approach dimensions. |
| L1-find-replace-001 | 72 | ✓ | The code uses a correct API method replaceText, but lacks the context of a Word.run function to ensure the code is executed within a valid runtime context, and also does not utilize the best practice approach of using Word.run and queueing changes. |
| L1-find-replace-002 | 88 | ✓ | The code is mostly correct, but loses points for not using the most efficient or best practice approach, such as utilizing the `range.font` properties directly on the search results without needing to load and sync the font property separately. |
| L1-find-replace-regex-001 | 80 | ✓ | The code correctly uses the Word Office.js API and should work without runtime errors, but loses points for not fully addressing the request with more precise wildcard patterns and not using the most optimal approach by directly utilizing the available helpers without considering potential edge cases. |
| L1-footnote-insert-001 | 72 | ✓ | The code uses a correct API to insert a footnote but lacks specificity in targeting the current selection and does not fully follow best practices for handling the insertion within the context of the user's request. |
| L1-comment-insert-001 | 72 | ✓ | The code uses a correct API method `insertComment` but lacks specificity in targeting the first occurrence of the text "$4.2M" and does not utilize Word's built-in search functionality with `body.search` to locate the exact text before inserting the comment. |
| L1-track-changes-toggle-001 | 95 ✅ | ✓ | The code perfectly addresses the request using the correct API, but loses points for approach because it manually calls `toggleTrackChanges` instead of using a potentially more robust built-in helper if available, although in this case, `toggleTrackChanges` is a suitable method. |
| L1-track-changes-toggle-002 | 83 | ✓ | The code is mostly correct but loses points for not using the Word.run helper to encapsulate the API calls and for not checking the state of the document before accepting revisions. |
| L1-toc-generate-001 | 80 | ✓ | The code uses a correct and existing Word Office.js API method `insertTableOfContents`, but loses points for not fully addressing the request as it doesn't specify the exact location of the Table of Contents, and also doesn't follow best practices by not utilizing the available helpers for more precise control over the Table of Contents. |
| L1-section-break-001 | 72 | ✓ | The code uses correct APIs and is likely to work, but it lacks completeness in fully addressing the request, as it doesn't specify the section to apply the landscape orientation to, and the approach could be improved by utilizing available helpers for more robust and maintainable code. |
| L1-header-footer-001 | 80 | ✓ | The code is mostly correct but loses points for not using the Word.run method to execute the API call and for not fully addressing potential edge cases, such as handling multiple sections or existing headers. |
| L1-header-footer-002 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions. |
| L1-columns-001 | 20 ⚠️ | ✓ | A: 20
B: |
| L1-image-insert-001 | 72 | ✓ | The code uses the correct insertImage API but loses points for not specifying the insertion point at the cursor in paragraph 3, and for not following best practices in terms of handling the image insertion and context synchronization. |
| L1-image-insert-002 | 20 ⚠️ | ✓ | The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would_it_work dimensions. |
| L1-watermark-insert-001 | 0 ⚠️ | ✓ | The code hallucinates a non-existent method "insertWatermark" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-content-control-001 | 83 | ✓ | The code correctly uses the insertContentControl API and addresses the request, but loses points for not using the Word.run method to ensure the code is executed in the correct context and for not handling potential errors, and also for not being a complete solution as it doesn't specify the position of the content control. |
| L1-content-control-002 | 83 | ✓ | The code correctly uses the Office.js API to insert a checkbox content control, but loses points for not handling potential errors and not using a more robust approach to ensure the content control is inserted at the correct location. |
| L1-mail-merge-001 | 83 | ✓ | The code correctly uses the mailMergeReplace API and addresses the user's request, but loses points for not using the most efficient or best-practice approach, such as handling potential errors or using more specific typing for the replace operation. |
| L1-template-apply-001 | 83 | ✓ | The code correctly applies the resume template with the provided information, but loses points for not using the Word.run method to ensure context synchronization and for not fully utilizing best practices and available helpers, such as using Word.run to wrap the applyTemplate call. |
| L1-template-apply-002 | 0 ⚠️ | ✓ | The code hallucinates a non-existent method "applyTemplate" which is not a real Word Office.js API, resulting in a complete failure across all dimensions. |
| L1-document-generate-001 | 90 | ✓ | The code accurately uses Office.js APIs and fully addresses the request, but loses points for not using the most efficient approach, such as potentially utilizing a loop to generate the sections instead of repeating similar code. |
| L1-theme-apply-001 | 15 ⚠️ | ✓ | The code heavily hallucinates methods, using non-existent "designTheme" and "applyTheme" functions, which are not part of the real Word Office.js APIs. |
| L1-theme-apply-002 | 72 | ✓ | The code uses the correct designTheme and applyTheme APIs, but lacks completeness in addressing the request, as it doesn't specify how the theme will be applied to the existing body content, and the approach could be improved by utilizing more specific helpers for styling the document. |
| L1-theme-apply-003 | 20 ⚠️ | ✓ | The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors. |
| L1-citation-bibliography-001 | 20 ⚠️ | ✓ | The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions. |
| L1-equation-001 | 83 | ✓ | The code is mostly correct, but it uses a generic function name `insertEquation` which is not a real Word Office.js API, instead of using the available helper `range.insertOoxml` to insert the equation, losing points for API correctness and approach. |
| L1-equation-002 | 83 | ✓ | The code correctly uses the insertEquation API and addresses the user's request, but loses points for not explicitly specifying the range where the equation should be inserted, which could lead to potential issues with the equation's placement in the document. |
| L1-writing-coach-001 | 83 | ✗ | The response accurately rewrites the selected text for clarity, but loses points for not considering alternative approaches, such as recommending the coach panel, and for using a somewhat generic phrase "Rewriting the selected text for clarity" that could be more descriptive. |
| L1-read-query-001 | 20 ⚠️ | ✓ | The code uses a hallucinated method "countWords()" which is not a real Word Office.js API, resulting in a severe penalty for API correctness and would not work in Word without runtime errors. |
| L1-read-query-002 | 57 ⚠️ | ✓ | The code uses a hallucinated method `getReadability()` which is not a real Word Office.js API, resulting in significant penalties across multiple dimensions. |
| L1-read-query-003 | 40 ⚠️ | ✓ | The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `listHeadings()` method, which heavily penalizes the API correctness score. |
| L1-margins-orientation-001 | 83 | ✓ | The code correctly uses the Office.js API and addresses the user's request, but loses points for not using a more explicit best practice approach, such as using a built-in method to set margins, and for not handling potential errors or edge cases. |
| L1-margins-orientation-002 | 80 | ✓ | The code correctly uses the Office.js API and addresses the request, but lacks context about the document and section objects, which are necessary for setting page orientation, resulting in potential runtime errors and incomplete implementation of best practices. |
| L1-web-search-needed-001 | 67 | ✓ | The AI response correctly identifies the need for a web search to obtain the current Apple stock price, but lacks clarity on how this information will be used to add a section to the Word document and does not offer any workarounds or suggestions. |
| L1-web-search-needed-002 | 77 | ✓ | The response accurately and clearly explains the Pythagorean theorem, directly addressing the question, but the bonus score is low because it doesn't provide additional workarounds or Word-specific caveats beyond inserting the equation into the document. |
| L1-refusal-001 | 83 | ✓ | The response is mostly accurate and clear, but loses points for not providing more detailed workaround steps or alternative solutions, and for not explicitly stating the limitations of the Word JavaScript API in a more concise manner. |
| L1-refusal-002 | 83 | ✓ | The response accurately identifies the limitation of Word Office.js and provides a clear workaround, but loses points for not exploring potential alternative solutions or providing more detailed information about the print functionality. |
| L1-refusal-003 | 67 | ✓ | The response accurately identifies the limitation of Word Office.js in creating Excel-style charts, but lacks a clear and relevant workaround that directly addresses the user's request for a chart based on specific data. |
| L1-quote-insert-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements. |
| L1-paragraph-spacing-001 | 72 | ✓ | The code uses correct APIs and mostly addresses the request, but it unnecessarily applies a style to each paragraph before setting line spacing and lacks the use of best practices such as directly accessing paragraphs without looping through each one with an index. |
| L1-edge-case-001 | 25 ⚠️ | ✓ | The code fails to use the Word Office.js API correctly, instead throwing an error without attempting to access or delete the paragraph, and not utilizing any of the available API methods or helpers. |
| L1-form-field-001 | 90 | ✓ | The code correctly uses the Office.js API to insert a dropdown content control, but loses points for not fully following best practices, such as not checking if the cursor is in a valid location to insert the content control. |
| L1-text-insert-001 | 85 | ✓ | The code uses a real Word Office.js API helper method `addParagraph`, but loses points for not explicitly checking if the insertion is at the end of the document, and for not using the most precise method `body.paragraphs.add` with a specific location. |
| L1-text-edit-001 | 0 ⚠️ | ✓ | The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions. |
| gen-L2-form-field-001 | 72 | ✓ | The code uses correct Office.js APIs and is likely to work, but it does not fully address the request as it inserts a date picker with the title "Sign Date" instead of placing it below the checkbox content control with no specific title, and lacks best practices in content control placement and styling. |
| gen-L2-form-field-002 | 72 | ✗ | The code partially addresses the request but fails to correctly replace the existing plain text with a rich text content control and incorrectly uses a comboBox content control instead of a building block gallery control. |
| gen-L2-form-field-003 | 65 | ✗ | The code partially addresses the request but fails to insert the dropdown content controls within a group content control, and also uses a non-existent `insertContentControl` method with incorrect parameters, instead of using the correct `range.insertContentControl` method. |
| gen-L2-form-field-004 | 72 | ✗ | The code partially addresses the request but fails to fully implement a repeating section content control, instead inserting a rich text content control and a table, and also lacks best practices in using available helpers. |

## ⚠️ Needs attention

**[L1-list-bullet-001]** score=25 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a significant penalty for API correctness.

**[L1-list-multilevel-001]** score=0 — The code uses a hallucinated method "addList" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-header-footer-002]** score=0 — The code uses a hallucinated method "addPageNumbers" which is not a real Word Office.js API, resulting in a complete loss of points across all dimensions.

**[L1-columns-001]** score=20 — A: 20
B:

**[L1-image-insert-002]** score=20 — The code uses a hallucinated method "insertImage" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would_it_work dimensions.

**[L1-watermark-insert-001]** score=0 — The code hallucinates a non-existent method "insertWatermark" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-template-apply-002]** score=0 — The code hallucinates a non-existent method "applyTemplate" which is not a real Word Office.js API, resulting in a complete failure across all dimensions.

**[L1-theme-apply-001]** score=15 — The code heavily hallucinates methods, using non-existent "designTheme" and "applyTheme" functions, which are not part of the real Word Office.js APIs.

**[L1-theme-apply-003]** score=20 — The code uses a hallucinated method "tweakTheme" which is not a real Word Office.js API, resulting in a complete loss of points for API correctness and would not work in Word without runtime errors.

**[L1-citation-bibliography-001]** score=20 — The code uses hallucinated methods "insertCitation" and "insertBibliography" which are not real Word Office.js APIs, resulting in a complete loss of points for API correctness and would-it-work dimensions.

**[L1-read-query-001]** score=20 — The code uses a hallucinated method "countWords()" which is not a real Word Office.js API, resulting in a severe penalty for API correctness and would not work in Word without runtime errors.

**[L1-read-query-002]** score=57 — The code uses a hallucinated method `getReadability()` which is not a real Word Office.js API, resulting in significant penalties across multiple dimensions.

**[L1-read-query-003]** score=40 — The code fails to use real Word Office.js APIs, instead hallucinating a non-existent `listHeadings()` method, which heavily penalizes the API correctness score.

**[L1-quote-insert-001]** score=0 — The code uses a hallucinated method "addQuote" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements.

**[L1-edge-case-001]** score=25 — The code fails to use the Word Office.js API correctly, instead throwing an error without attempting to access or delete the paragraph, and not utilizing any of the available API methods or helpers.

**[L1-text-edit-001]** score=0 — The code uses a hallucinated method "addTitle" which is not a real Word Office.js API, resulting in a complete failure to meet the requirements of all four dimensions.
