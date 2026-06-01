# Word AI — Eval Results
**Last run:** 2026-06-01 04:12:02  
**Overall: 58.5/100**  
**Model:** meta-llama/llama-3.1-8b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 47.0/100 | — | 2 |
| paragraph-format | 51.5/100 | — | 2 |
| list-bullet | 47.0/100 | — | 1 |
| list-multilevel | 94.0/100 | — | 1 |
| style-apply | 85.0/100 | — | 1 |
| table-create | 75.0/100 | — | 1 |
| table-format | 89.0/100 | — | 1 |
| find-replace | 53.5/100 | — | 2 |
| find-replace-regex | 40.0/100 | — | 1 |
| footnote-insert | 47.0/100 | — | 1 |
| comment-insert | 47.0/100 | — | 1 |
| track-changes-toggle | 74.5/100 | — | 2 |
| toc-generate | 89.0/100 | — | 1 |
| section-break | 64.0/100 | — | 1 |
| header-footer | 75.5/100 | — | 2 |
| columns | 47.0/100 | — | 1 |
| image-insert | 46.0/100 | — | 2 |
| watermark-insert | 0.0/100 | — | 1 |
| content-control | 85.0/100 | — | 2 |
| mail-merge | 67.0/100 | — | 1 |
| template-apply | 85.0/100 | — | 2 |
| document-generate | 78.0/100 | — | 1 |
| theme-apply | 62.0/100 | — | 3 |
| citation-bibliography | 67.0/100 | — | 1 |
| equation | 46.0/100 | — | 2 |
| writing-coach | 47.0/100 | — | 1 |
| read-query | 41.3/100 | — | 3 |
| margins-orientation | 52.0/100 | — | 2 |
| web-search-needed | 41.0/100 | — | 2 |
| refusal | 47.0/100 | — | 3 |
| quote-insert | 47.0/100 | — | 1 |
| paragraph-spacing | 85.0/100 | — | 1 |
| edge-cases | 0.0/100 | — | 1 |
| form-field | 64.8/100 | — | 5 |
| text-insert | 47.0/100 | — | 1 |
| text-edit | 85.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 47 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, addHeading, but lacks proper error handling and context awareness.

 |
| L1-heading-insert-002 | 47 ⚠️ | ✓ | A: 22
The code uses the correct Word Office.js API, Word.run, and the correct method signatures, but it lacks the necess |
| L1-paragraph-format-001 | 18 ⚠️ | ✗ | The code does not address the request at all, as it only applies a heading style to the entire document, not just the first paragraph, and does not change the font size to 14pt. |
| L1-paragraph-format-002 | 85 | ✓ | The code uses real Word Office.js APIs, fully addresses the request, and would likely execute without runtime errors, but it doesn't use the available helpers and doesn't check if the paragraph already exists before trying to center-align it. |
| L1-list-bullet-001 | 47 ⚠️ | ✓ | A: 22
The code uses the correct Word Office.js API, addList, but lacks the necessary Word.run context.

B: 0
The code do |
| L1-list-multilevel-001 | 94 | ✓ | The code uses the correct Word Office.js API, fully addresses the request, would execute without runtime errors, and uses best practices, but does not take advantage of the available helpers for adding a numbered list. |
| L1-style-apply-001 | 85 | ✓ | A: 22
The code uses the correct Word Office.js API, `applyStyle`, but does not account for potential errors or exception |
| L1-table-create-001 | 75 | ✓ | A: 22
The code uses the correct Word Office.js API, `insertTable`, but lacks proper error handling and assumes the table |
| L1-table-format-001 | 89 | ✓ | A: 22
The code uses the correct Word Office.js API, `styleTable`, to apply the style to the existing table.

B: 25
The c |
| L1-find-replace-001 | 47 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, `replaceText`, but lacks context and assumes the text to be replaced |
| L1-find-replace-002 | 60 | ✓ | A: 22
The code uses the correct Word Office.js API, but it's missing the `await context.sync()` call after setting the f |
| L1-find-replace-regex-001 | 40 ⚠️ | ✓ | The code is incomplete as it only replaces text, but does not address the request to make email addresses bold and red, and it uses a real Word API but does not follow best practices. |
| L1-footnote-insert-001 | 47 ⚠️ | ✓ | A: 22
The code uses the correct Word Office.js API, `insertFootnote`, but lacks the necessary Word.run context.

B: 0
Th |
| L1-comment-insert-001 | 47 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, `insertComment`, but lacks the necessary context to identify the fir |
| L1-track-changes-toggle-001 | 64 | ✓ | A: 22
The code uses the correct Word Office.js API, toggleTrackChanges, but does not account for the document context, a |
| L1-track-changes-toggle-002 | 85 | ✓ | The code uses the correct Word API method, but lacks best practices and helper functions, and does not address the request fully in the context of tracked revisions. |
| L1-toc-generate-001 | 89 | ✓ | The code uses the correct Office.js API, fully addresses the request, and would execute without runtime errors, but it doesn't use the available helpers for inserting a Table of Contents, which is a complex operation that requires more than just a single API call. |
| L1-section-break-001 | 64 | ✓ | A: 22
The code uses the correct Word Office.js APIs, but it's missing the crucial step of switching the next section to  |
| L1-header-footer-001 | 60 | ✓ | A: 20
The code uses the correct Word Office.js API, `setHeader`, but lacks the `primary` and `alignment` properties, whi |
| L1-header-footer-002 | 91 | ✓ | The code uses the correct Word Office.js API, fully addresses the request, and would execute without runtime errors, but it doesn't utilize the available helpers and doesn't check if the document already has a footer. |
| L1-columns-001 | 47 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, `insertColumns`, but it lacks the necessary context to switch the re |
| L1-image-insert-001 | 45 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, `insertImage`, but lacks the `Word.run` context.

B: 0
The code does |
| L1-image-insert-002 | 47 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, `insertImage`, but lacks context and parameters for the image.

B: 0 |
| L1-watermark-insert-001 | 0 ⚠️ | ✓ | A: 0
The code uses a non-existent method "insertWatermark" which is not a real Word Office.js API.

B: 0
The code does n |
| L1-content-control-001 | 85 | ✓ | A: 22
The code uses the correct Office.js API, `insertContentControl`, but lacks the necessary `Word.run` context.

B: 2 |
| L1-content-control-002 | 85 | ✓ | A: 22
The code uses the correct Office.js API, `insertContentControl`, but lacks the necessary Word.run context.

B: 25
 |
| L1-mail-merge-001 | 67 | ✓ | A: 20
The code uses the correct Word Office.js API, mailMergeReplace, but lacks proper error handling and context.

B: 1 |
| L1-template-apply-001 | 85 | ✓ | A: 22
The code uses the correct Word Office.js API, applyTemplate, but does not handle potential errors or exceptions.

 |
| L1-template-apply-002 | 85 | ✓ | A: 22
The code uses the correct Word Office.js API, applyTemplate, but does not check if the template exists or if the d |
| L1-document-generate-001 | 78 | ✓ | The code uses real Word Office.js APIs, but lacks a clear Table of Contents structure and uses incorrect method signatures, and does not follow best practices for inserting images and themes. |
| L1-theme-apply-001 | 45 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, `designTheme` and `applyTheme`, but lacks any implementation details |
| L1-theme-apply-002 | 67 | ✓ | A: 18
The code uses the correct Word Office.js API, designTheme and applyTheme, but lacks proper error handling and assu |
| L1-theme-apply-003 | 74 | ✓ | A: 20
The code uses the correct Word Office.js API, `applyTheme`, but incorrectly assumes the `tweakTheme` method is a r |
| L1-citation-bibliography-001 | 67 | ✓ | A: 22
The code uses the correct Word Office.js API, `insertCitation` and `insertBibliography`, but it's missing the `Wor |
| L1-equation-001 | 47 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, `insertEquation`, but lacks proper error handling and context awaren |
| L1-equation-002 | 45 ⚠️ | ✓ | A: 20
The code uses the correct Word Office.js API, `insertEquation`, but lacks proper error handling and context awaren |
| L1-writing-coach-001 | 47 ⚠️ | ✗ | A: 22
The code accurately replaces the selection text with a clearer alternative, but it doesn't address the "rewrite" a |
| L1-read-query-001 | 30 ⚠️ | ✓ | The code uses a non-existent method "countWords()" and does not address the request at all, making it incomplete and incorrect. |
| L1-read-query-002 | 72 | ✓ | A: 18
The code uses the `getReadability()` method, which is not a real Word Office.js API, and instead relies on a hypot |
| L1-read-query-003 | 22 ⚠️ | ✗ | The code does not use any real Word Office.js APIs, and the approach is completely manual and incorrect. |
| L1-margins-orientation-001 | 64 | ✓ | A: 22
The code uses the correct Word Office.js API, `setMargins`, but does not account for the context of default margin |
| L1-margins-orientation-002 | 40 ⚠️ | ✓ | A: 22
The code uses the correct Word Office.js API, `setPageOrientation`, but lacks the necessary Word.run context.

B:  |
| L1-web-search-needed-001 | 22 ⚠️ | ✗ | The response fails to provide any relevant information or code to address the |
| L1-web-search-needed-002 | 60 | ✓ | The response is a good start, but it could be improved with more concise language and a clearer presentation of the formula. |
| L1-refusal-001 | 47 ⚠️ | ✓ | A: 18
The response is factually correct that Word Office.js doesn't expose direct PDF export, but it doesn't provide a c |
| L1-refusal-002 | 47 ⚠️ | ✗ | The response is clear and directly addresses the question, but it fails to provide any relevant information or code, and does not offer any workarounds or suggestions. |
| L1-refusal-003 | 47 ⚠️ | ✓ | A: 18
The response is factually correct that Word Office.js doesn't expose direct chart creation, but it doesn't provide |
| L1-quote-insert-001 | 47 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, `addQuote`, but lacks context and proper usage.

B: 0
The code does  |
| L1-paragraph-spacing-001 | 85 | ✓ | A: 22
The code uses the correct Word Office.js APIs, but it's missing the initial Word.run context setup.

B: 25
The cod |
| L1-edge-case-001 | 0 ⚠️ | ✗ | The AI response does not generate any code, and the provided explanation is not a solution to the user's request. |
| L1-form-field-001 | 85 | ✓ | A: 22
The code uses the correct Word Office.js API, `insertContentControl`, but lacks the necessary `Word.run` context.
 |
| L1-text-insert-001 | 47 ⚠️ | ✓ | A: 18
The code uses a real Word Office.js API, but it's incomplete and lacks proper error handling.

B: 0
The code does  |
| L1-text-edit-001 | 85 | ✓ | A: 22
The code uses the correct Word Office.js API, `addTitle`, but does not account for the fact that the document's ti |
| gen-L2-form-field-001 | 47 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, `insertContentControl`, but lacks the necessary context to insert th |
| gen-L2-form-field-002 | 60 | ✗ | A: 20
The code uses the correct Word Office.js API, `insertContentControl`, but lacks the crucial step of replacing the  |
| gen-L2-form-field-003 | 85 | ✗ | A: 20
The code uses the correct Word Office.js APIs, but it lacks the `await Word.run()` wrapper, which is essential for |
| gen-L2-form-field-004 | 47 ⚠️ | ✗ | A: 18
The code uses the correct Word Office.js API, `insertContentControl`, but lacks the crucial `repeating` property t |

## ⚠️ Needs attention

**[L1-heading-insert-001]** score=47 — A: 18
The code uses the correct Word Office.js API, addHeading, but lacks proper error handling and context awareness.



**[L1-heading-insert-002]** score=47 — A: 22
The code uses the correct Word Office.js API, Word.run, and the correct method signatures, but it lacks the necess

**[L1-paragraph-format-001]** score=18 — The code does not address the request at all, as it only applies a heading style to the entire document, not just the first paragraph, and does not change the font size to 14pt.
- missing patterns: `font.bold`, `font.size`

**[L1-list-bullet-001]** score=47 — A: 22
The code uses the correct Word Office.js API, addList, but lacks the necessary Word.run context.

B: 0
The code do

**[L1-find-replace-001]** score=47 — A: 18
The code uses the correct Word Office.js API, `replaceText`, but lacks context and assumes the text to be replaced

**[L1-find-replace-regex-001]** score=40 — The code is incomplete as it only replaces text, but does not address the request to make email addresses bold and red, and it uses a real Word API but does not follow best practices.

**[L1-footnote-insert-001]** score=47 — A: 22
The code uses the correct Word Office.js API, `insertFootnote`, but lacks the necessary Word.run context.

B: 0
Th

**[L1-comment-insert-001]** score=47 — A: 18
The code uses the correct Word Office.js API, `insertComment`, but lacks the necessary context to identify the fir

**[L1-columns-001]** score=47 — A: 18
The code uses the correct Word Office.js API, `insertColumns`, but it lacks the necessary context to switch the re

**[L1-image-insert-001]** score=45 — A: 18
The code uses the correct Word Office.js API, `insertImage`, but lacks the `Word.run` context.

B: 0
The code does

**[L1-image-insert-002]** score=47 — A: 18
The code uses the correct Word Office.js API, `insertImage`, but lacks context and parameters for the image.

B: 0

**[L1-watermark-insert-001]** score=0 — A: 0
The code uses a non-existent method "insertWatermark" which is not a real Word Office.js API.

B: 0
The code does n

**[L1-theme-apply-001]** score=45 — A: 18
The code uses the correct Word Office.js API, `designTheme` and `applyTheme`, but lacks any implementation details

**[L1-equation-001]** score=47 — A: 18
The code uses the correct Word Office.js API, `insertEquation`, but lacks proper error handling and context awaren

**[L1-equation-002]** score=45 — A: 20
The code uses the correct Word Office.js API, `insertEquation`, but lacks proper error handling and context awaren

**[L1-writing-coach-001]** score=47 — A: 22
The code accurately replaces the selection text with a clearer alternative, but it doesn't address the "rewrite" a
- missing patterns: `coach`

**[L1-read-query-001]** score=30 — The code uses a non-existent method "countWords()" and does not address the request at all, making it incomplete and incorrect.

**[L1-read-query-003]** score=22 — The code does not use any real Word Office.js APIs, and the approach is completely manual and incorrect.
- missing patterns: `return`

**[L1-margins-orientation-002]** score=40 — A: 22
The code uses the correct Word Office.js API, `setPageOrientation`, but lacks the necessary Word.run context.

B: 

**[L1-web-search-needed-001]** score=22 — The response fails to provide any relevant information or code to address the
- missing patterns: `NEEDS_WEB_SEARCH`

**[L1-refusal-001]** score=47 — A: 18
The response is factually correct that Word Office.js doesn't expose direct PDF export, but it doesn't provide a c

**[L1-refusal-002]** score=47 — The response is clear and directly addresses the question, but it fails to provide any relevant information or code, and does not offer any workarounds or suggestions.
- missing patterns: `doesn't expose`

**[L1-refusal-003]** score=47 — A: 18
The response is factually correct that Word Office.js doesn't expose direct chart creation, but it doesn't provide

**[L1-quote-insert-001]** score=47 — A: 18
The code uses the correct Word Office.js API, `addQuote`, but lacks context and proper usage.

B: 0
The code does 

**[L1-edge-case-001]** score=0 — The AI response does not generate any code, and the provided explanation is not a solution to the user's request.
- missing patterns: `paragraphs`

**[L1-text-insert-001]** score=47 — A: 18
The code uses a real Word Office.js API, but it's incomplete and lacks proper error handling.

B: 0
The code does 

**[gen-L2-form-field-001]** score=47 — A: 18
The code uses the correct Word Office.js API, `insertContentControl`, but lacks the necessary context to insert th

**[gen-L2-form-field-004]** score=47 — A: 18
The code uses the correct Word Office.js API, `insertContentControl`, but lacks the crucial `repeating` property t
- missing patterns: `RepeatingSectionContentControl`
