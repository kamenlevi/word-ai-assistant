# Word AI — Eval Results
**Last run:** 2026-06-29 04:01:34  
**Overall: 65.1/100**  
**Model:** meta-llama/llama-3.1-8b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 91.5/100 | — | 2 |
| paragraph-format | 87.0/100 | — | 2 |
| list-bullet | 89.0/100 | — | 1 |
| list-multilevel | 85.0/100 | — | 1 |
| style-apply | 94.0/100 | — | 1 |
| table-create | 89.0/100 | — | 1 |
| table-format | 85.0/100 | — | 1 |
| find-replace | 88.0/100 | — | 2 |
| find-replace-regex | 60.0/100 | — | 1 |
| footnote-insert | 60.0/100 | — | 1 |
| comment-insert | 85.0/100 | — | 1 |
| track-changes-toggle | 87.0/100 | — | 2 |
| toc-generate | 89.0/100 | — | 1 |
| section-break | 60.0/100 | — | 1 |
| header-footer | 69.0/100 | — | 2 |
| columns | 47.0/100 | — | 1 |
| image-insert | 70.5/100 | — | 2 |
| watermark-insert | 0.0/100 | — | 1 |
| content-control | 78.5/100 | — | 2 |
| mail-merge | 85.0/100 | — | 1 |
| template-apply | 81.0/100 | — | 2 |
| document-generate | 84.0/100 | — | 1 |
| theme-apply | 53.0/100 | — | 3 |
| citation-bibliography | 67.0/100 | — | 1 |
| equation | 46.0/100 | — | 2 |
| writing-coach | 47.0/100 | — | 1 |
| read-query | 38.7/100 | — | 3 |
| margins-orientation | 61.0/100 | — | 2 |
| web-search-needed | 39.0/100 | — | 2 |
| refusal | 47.0/100 | — | 3 |
| quote-insert | 45.0/100 | — | 1 |
| paragraph-spacing | 92.0/100 | — | 1 |
| edge-cases | 47.0/100 | — | 1 |
| form-field | 68.4/100 | — | 5 |
| text-insert | 0.0/100 | — | 1 |
| text-edit | 47.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 94 | ✓ | The code is mostly correct, but it lacks a Word.run block to execute the addHeading method, which is a crucial part of the Office.js API. |
| L1-heading-insert-002 | 89 | ✓ | The code uses real Word Office.js APIs, fully addresses the request, and would likely execute without runtime errors, but it doesn't use the available helpers for inserting content and doesn't handle potential errors when adding the paragraph. |
| L1-paragraph-format-001 | 89 | ✗ | The code uses the correct Word Office.js API, fully addresses the request, and would execute without runtime errors, but it doesn't use the available helpers and instead re-implements the functionality manually, which is a suboptimal approach. |
| L1-paragraph-format-002 | 85 | ✓ | The code uses real Word Office.js APIs, fully addresses the request, and would execute without runtime errors, but it doesn't use the available helpers and re-implements the addParagraph function manually. |
| L1-list-bullet-001 | 89 | ✓ | The code uses the correct Word Office.js API, but lacks a necessary Word.run context and uses a deprecated method signature for addList. |
| L1-list-multilevel-001 | 85 | ✓ | The code is mostly correct, but lacks a Word.run block and uses an undocumented method "addList" which is not a standard Office.js API. |
| L1-style-apply-001 | 94 | ✓ | The code uses the correct Word Office.js API, fully addresses the request, and would execute without runtime errors, but it could benefit from using the `applyStyle` method with a more robust way to access the paragraph, such as `paragraphs.item(2).styleBuiltIn("Quote")`. |
| L1-table-create-001 | 89 | ✓ | A: 22
The code uses the correct Word Office.js API, `insertTable`, but it's missing the `Word.run` context setup.

B: 25 |
| L1-table-format-001 | 85 | ✓ | The code uses the correct Word API, but lacks a crucial check for the table's existence before attempting to style it, and also doesn't follow best practices by not using the available helpers for styling tables. |
| L1-find-replace-001 | 85 | ✓ | The code uses the correct Word API, but lacks a loop to replace all instances of "Acme" in the document, and uses a direct string replacement without considering the context or potential edge cases. |
| L1-find-replace-002 | 91 | ✓ | The code correctly uses the Office.js API to search for and modify the text, but deducts points for not using the `matchWildcards` option in the `search` method, and for not addressing the color change in a more robust way. |
| L1-find-replace-regex-001 | 60 | ✓ | The code is partially correct, but it does not fully address the request as it only replaces text, not the actual email addresses in the document, and it uses a hardcoded pattern instead of a more robust solution. |
| L1-footnote-insert-001 | 60 | ✓ | A: 22
The code uses the correct Word Office.js API, `insertFootnote`, but lacks the necessary Word.run context.

B: 0
Th |
| L1-comment-insert-001 | 85 | ✓ | A: 20
The code uses the correct Word Office.js APIs, but it unnecessarily calls findText twice, which is inefficient.

B |
| L1-track-changes-toggle-001 | 89 | ✓ | A: 22
The code uses the correct Word Office.js API, `toggleTrackChanges`, but does not follow the recommended `Word.run` |
| L1-track-changes-toggle-002 | 85 | ✓ | The code uses the correct Word Office.js API, but deducts points for not addressing the request fully, potential runtime errors, and not using best practices for theme application. |
| L1-toc-generate-001 | 89 | ✓ | The code uses the correct Office.js API, but it doesn't check if the document already has a Table of Contents, and it doesn't handle potential errors when inserting the TOC. |
| L1-section-break-001 | 60 | ✓ | A: 22
The code uses the correct Word Office.js APIs, but it's missing the part that switches the next section to landsca |
| L1-header-footer-001 | 91 | ✓ | The code correctly uses the Word Office.js API, fully addresses the request, and would execute without runtime errors, but it does not use the available helpers for setting the header, instead re-implementing it manually. |
| L1-header-footer-002 | 47 ⚠️ | ✓ | A: 22
The code uses the correct Word Office.js API, `addPageNumbers`, but lacks the necessary context to determine if it |
| L1-columns-001 | 47 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, `insertSectionBreak` and `insertColumns`, but it's missing the cruci |
| L1-image-insert-001 | 94 | ✓ | The code uses the correct Office.js API, fully addresses the request, and would execute without runtime errors, but it does not use the recommended helper function for inserting images, instead using a custom method. |
| L1-image-insert-002 | 47 ⚠️ | ✓ | The code uses the correct Word Office.js API, but it does not fully address the request, and the approach is not optimal as it does not use the available helpers for inserting images. |
| L1-watermark-insert-001 | 0 ⚠️ | ✓ | The code uses a non-existent method "insertWatermark" and does not address the request of adding a watermark to every page, nor does it consider the correct load/sync order or variable scope. |
| L1-content-control-001 | 72 | ✓ | A: 22
The code uses the correct Word Office.js API, `insertContentControl`, but lacks proper error handling and assumes  |
| L1-content-control-002 | 85 | ✓ | A: 22
The code uses the correct Office.js API, `insertContentControl`, but lacks the necessary Word.run context.

B: 25
 |
| L1-mail-merge-001 | 85 | ✓ | A: 22
The code uses the real Word Office.js API `mailMergeReplace`, but it's missing the `Word.run` context setup.

B: 2 |
| L1-template-apply-001 | 85 | ✓ | The code is mostly correct, but lacks error handling and uses |
| L1-template-apply-002 | 77 | ✓ | A: 22
The code uses the correct Word API, applyTemplate, but does not check if the template exists or if the document is |
| L1-document-generate-001 | 84 | ✓ | A: 22
The code uses real Word Office.js APIs, but it could be improved by using the `Word.run` method to ensure correct  |
| L1-theme-apply-001 | 37 ⚠️ | ✓ | The code uses real Word Office.js APIs, but it lacks any actual implementation, and the approach is overly simplistic and doesn't address the request's requirements. |
| L1-theme-apply-002 | 47 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, designTheme and applyTheme, but lacks proper error handling and assu |
| L1-theme-apply-003 | 75 | ✓ | A: 18
The code uses the correct Word API, `applyTheme`, but incorrectly uses a string argument instead of an object with |
| L1-citation-bibliography-001 | 67 | ✓ | A: 18
The code uses the correct Word Office.js API, `Word.run`, but it's missing the necessary `context` and `async`/`aw |
| L1-equation-001 | 45 ⚠️ | ✓ | A: 22
The code uses the correct Word Office.js API, insertEquation, but lacks proper error handling and does not check i |
| L1-equation-002 | 47 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, `insertEquation`, but lacks proper error handling and assumes the fo |
| L1-writing-coach-001 | 47 ⚠️ | ✗ | The response fails to provide a rewritten version of the text as requested, instead only offering a replacement code. |
| L1-read-query-001 | 22 ⚠️ | ✓ | The code relies on a non-existent function and does not address the request, making it a poor solution. |
| L1-read-query-002 | 72 | ✓ | A: 18
The code uses the `getReadability()` method, which is not a real Word Office.js API. The correct API to use is `Wo |
| L1-read-query-003 | 22 ⚠️ | ✗ | The code relies on a non-existent method, which is a fundamental flaw. |
| L1-margins-orientation-001 | 75 | ✓ | A: 22
The code uses the correct Word Office.js API, setMargins, but it's not the most straightforward way to set margins |
| L1-margins-orientation-002 | 47 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, `setPageOrientation`, but it's missing the context of the document,  |
| L1-web-search-needed-001 | 18 ⚠️ | ✓ | The response is a refusal to provide code, but it fails to clearly state that it cannot fulfill the request due to lack of web access, and instead asks the user to enable web search, which is not a solution. |
| L1-web-search-needed-002 | 60 | ✓ | The response is mostly accurate and clear, but it could be improved by being more concise and relevant to the question. |
| L1-refusal-001 | 47 ⚠️ | ✓ | A: 18
The response is factually correct that Word Office.js doesn't expose direct PDF export, but it fails to mention th |
| L1-refusal-002 | 47 ⚠️ | ✗ | The response is clear and directly addresses the question, but fails to provide any relevant information or code, and does not offer any workarounds or suggestions. |
| L1-refusal-003 | 47 ⚠️ | ✓ | A: 18
The response is factually correct that Word Office.js doesn't expose direct chart creation, but it doesn't fully a |
| L1-quote-insert-001 | 45 ⚠️ | ✓ | A: 18
The code uses the correct Word API, but it's missing the necessary context to execute correctly.

B: 0
The code do |
| L1-paragraph-spacing-001 | 92 | ✓ | A: 23
The code uses the correct Word Office.js API, Word.run, and the correct method, paragraph.lineSpacing.

B: 25
The  |
| L1-edge-case-001 | 47 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, but it doesn't attempt to delete paragraph 99, which is the original |
| L1-form-field-001 | 91 | ✓ | The code correctly uses the Office.js API to insert a dropdown content control, but lacks a check for the cursor position and assumes it's in the body, which might not always be the case. |
| L1-text-insert-001 | 0 ⚠️ | ✓ | The code uses a non-existent API |
| L1-text-edit-001 | 47 ⚠️ | ✓ | A: 22
The code uses the correct Word Office.js API, `addTitle`, but does not account for the fact that the document's ti |
| gen-L2-form-field-001 | 67 | ✓ | A: 20
The code uses the correct Word Office.js API, `insertContentControl`, but lacks the necessary context to properly  |
| gen-L2-form-field-002 | 47 ⚠️ | ✗ | A: 18
The code uses the correct Word Office.js API, `insertContentControl`, but lacks the crucial step of replacing the  |
| gen-L2-form-field-003 | 70 | ✗ | A: 18
The code uses the correct Word Office.js API, `insertContentControl`, but incorrectly assumes the `title` property |
| gen-L2-form-field-004 | 67 | ✗ | A: 20
The code uses the correct Word Office.js APIs, but it's missing the crucial `Word.run` function to execute the cod |

## ⚠️ Needs attention

**[L1-header-footer-002]** score=47 — A: 22
The code uses the correct Word Office.js API, `addPageNumbers`, but lacks the necessary context to determine if it

**[L1-columns-001]** score=47 — A: 18
The code uses the correct Word Office.js API, `insertSectionBreak` and `insertColumns`, but it's missing the cruci

**[L1-image-insert-002]** score=47 — The code uses the correct Word Office.js API, but it does not fully address the request, and the approach is not optimal as it does not use the available helpers for inserting images.

**[L1-watermark-insert-001]** score=0 — The code uses a non-existent method "insertWatermark" and does not address the request of adding a watermark to every page, nor does it consider the correct load/sync order or variable scope.

**[L1-theme-apply-001]** score=37 — The code uses real Word Office.js APIs, but it lacks any actual implementation, and the approach is overly simplistic and doesn't address the request's requirements.

**[L1-theme-apply-002]** score=47 — A: 18
The code uses the correct Word Office.js API, designTheme and applyTheme, but lacks proper error handling and assu

**[L1-equation-001]** score=45 — A: 22
The code uses the correct Word Office.js API, insertEquation, but lacks proper error handling and does not check i

**[L1-equation-002]** score=47 — A: 18
The code uses the correct Word Office.js API, `insertEquation`, but lacks proper error handling and assumes the fo

**[L1-writing-coach-001]** score=47 — The response fails to provide a rewritten version of the text as requested, instead only offering a replacement code.
- missing patterns: `coach`

**[L1-read-query-001]** score=22 — The code relies on a non-existent function and does not address the request, making it a poor solution.

**[L1-read-query-003]** score=22 — The code relies on a non-existent method, which is a fundamental flaw.
- missing patterns: `return`

**[L1-margins-orientation-002]** score=47 — A: 18
The code uses the correct Word Office.js API, `setPageOrientation`, but it's missing the context of the document, 

**[L1-web-search-needed-001]** score=18 — The response is a refusal to provide code, but it fails to clearly state that it cannot fulfill the request due to lack of web access, and instead asks the user to enable web search, which is not a solution.

**[L1-refusal-001]** score=47 — A: 18
The response is factually correct that Word Office.js doesn't expose direct PDF export, but it fails to mention th

**[L1-refusal-002]** score=47 — The response is clear and directly addresses the question, but fails to provide any relevant information or code, and does not offer any workarounds or suggestions.
- missing patterns: `doesn't expose`

**[L1-refusal-003]** score=47 — A: 18
The response is factually correct that Word Office.js doesn't expose direct chart creation, but it doesn't fully a

**[L1-quote-insert-001]** score=45 — A: 18
The code uses the correct Word API, but it's missing the necessary context to execute correctly.

B: 0
The code do

**[L1-edge-case-001]** score=47 — A: 18
The code uses the correct Word Office.js API, but it doesn't attempt to delete paragraph 99, which is the original

**[L1-text-insert-001]** score=0 — The code uses a non-existent API

**[L1-text-edit-001]** score=47 — A: 22
The code uses the correct Word Office.js API, `addTitle`, but does not account for the fact that the document's ti

**[gen-L2-form-field-002]** score=47 — A: 18
The code uses the correct Word Office.js API, `insertContentControl`, but lacks the crucial step of replacing the 
- missing patterns: `RichTextContentControl`
