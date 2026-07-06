# Word AI — Eval Results
**Last run:** 2026-07-06 03:48:17  
**Overall: 60.7/100**  
**Model:** meta-llama/llama-3.1-8b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 87.5/100 | — | 2 |
| paragraph-format | 47.0/100 | — | 2 |
| list-bullet | 47.0/100 | — | 1 |
| list-multilevel | 60.0/100 | — | 1 |
| style-apply | 89.0/100 | — | 1 |
| table-create | 85.0/100 | — | 1 |
| table-format | 85.0/100 | — | 1 |
| find-replace | 67.5/100 | — | 2 |
| find-replace-regex | 60.0/100 | — | 1 |
| footnote-insert | 64.0/100 | — | 1 |
| comment-insert | 47.0/100 | — | 1 |
| track-changes-toggle | 90.5/100 | — | 2 |
| toc-generate | 75.0/100 | — | 1 |
| section-break | 52.0/100 | — | 1 |
| header-footer | 68.0/100 | — | 2 |
| columns | 60.0/100 | — | 1 |
| image-insert | 53.5/100 | — | 2 |
| watermark-insert | 0.0/100 | — | 1 |
| content-control | 85.0/100 | — | 2 |
| mail-merge | 75.0/100 | — | 1 |
| template-apply | 85.0/100 | — | 2 |
| document-generate | 75.0/100 | — | 1 |
| theme-apply | 50.7/100 | — | 3 |
| citation-bibliography | 40.0/100 | — | 1 |
| equation | 67.0/100 | — | 2 |
| writing-coach | 64.0/100 | — | 1 |
| read-query | 23.3/100 | — | 3 |
| margins-orientation | 47.0/100 | — | 2 |
| web-search-needed | 41.0/100 | — | 2 |
| refusal | 53.0/100 | — | 3 |
| quote-insert | 47.0/100 | — | 1 |
| paragraph-spacing | 85.0/100 | — | 1 |
| edge-cases | 0.0/100 | — | 1 |
| form-field | 72.6/100 | — | 5 |
| text-insert | 40.0/100 | — | 1 |
| text-edit | 89.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 85 | ✓ | The code uses the correct Word Office.js API, fully addresses the request, and would execute without runtime errors, but it doesn't use the available helper functions and instead re-implements the addHeading method manually. |
| L1-heading-insert-002 | 90 | ✓ | The code correctly uses the Word.run API and the addHeading and addParagraph functions, but it does not check if the paragraph is inserted below the heading, and it does not handle potential errors that might occur during execution. |
| L1-paragraph-format-001 | 22 ⚠️ | ✗ | The code uses a real Word Office.js API, but it does not address the request to make the first paragraph bold and 14pt, and it does not use the available helpers to achieve this. |
| L1-paragraph-format-002 | 72 | ✓ | The code uses the correct Word Office.js API, fully addresses the request, and would execute without runtime errors, but it fails to use best practices by not utilizing the available helpers, specifically the addParagraph function, which is a built-in helper for adding paragraphs. |
| L1-list-bullet-001 | 47 ⚠️ | ✓ | A: 22
The code uses the correct Word Office.js API, addList, but lacks the necessary Word.run context.

B: 0
The code do |
| L1-list-multilevel-001 | 60 | ✓ | A: 22
The code uses the correct Word API, addList, but lacks a Word.run context.

B: 0
The code does not fully address t |
| L1-style-apply-001 | 89 | ✓ | A: 22
The code uses the correct Word Office.js API, `applyStyle`, but deducts 3 points for not checking if the paragraph |
| L1-table-create-001 | 85 | ✓ | A: 22
The code uses the correct Word Office.js API, `insertTable`, but deducts 3 points for not using the `Word.run` con |
| L1-table-format-001 | 85 | ✓ | A: 22
The code uses the correct Word Office.js API, `styleTable`, but it's missing the `Word.run` context.

B: 25
The co |
| L1-find-replace-001 | 47 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, `replaceText`, but does not account for the context of the document, |
| L1-find-replace-002 | 88 | ✓ | The code correctly uses the Word Office.js API, fully addresses the request, and would execute without runtime errors, but it re-implements the font formatting manually instead of using the available helpers, and does not check if the "WARNING" text is actually found in the document. |
| L1-find-replace-regex-001 | 60 | ✓ | A: 22
The code uses the correct Word API, replaceText, but it's missing the context of how to get the document object.

 |
| L1-footnote-insert-001 | 64 | ✓ | A: 22
The code uses the correct Word Office.js API, `insertFootnote`, but lacks the necessary Word.run context.

B: 0
Th |
| L1-comment-insert-001 | 47 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, Word.run, but lacks proper error handling and context.

B: 0
The cod |
| L1-track-changes-toggle-001 | 96 ✅ | ✓ | The code uses the correct Word Office.js API to toggle track changes, but it could benefit from using the `Word.run` method to ensure proper load/sync order. |
| L1-track-changes-toggle-002 | 85 | ✓ | The code uses the correct Office.js API, but deducts points for not addressing the request fully, potential runtime errors, and not using best practices for theme application. |
| L1-toc-generate-001 | 75 | ✓ | A: 22
The code uses the correct Word Office.js API, `insertTableOfContents`, but I deduct 3 points because it doesn't ha |
| L1-section-break-001 | 52 ⚠️ | ✓ | A: 22
The code uses the correct Word Office.js API, `insertSectionBreak` and `setPageOrientation`, but it's missing the  |
| L1-header-footer-001 | 89 | ✓ | The code uses the correct Word Office.js API, but it doesn't fully address the request as it doesn't specify the header's position or style, and it uses a manual approach instead of the recommended addHeading helper. |
| L1-header-footer-002 | 47 ⚠️ | ✓ | A: 18
The code uses a real Word Office.js API, but it's incomplete and doesn't account for the specific request of addin |
| L1-columns-001 | 60 | ✓ | A: 20
The code uses the correct Word Office.js APIs, but it lacks the necessary context to switch the entire document to |
| L1-image-insert-001 | 67 | ✓ | A: 18
The code uses the correct Word Office.js API, `insertImage`, but lacks the necessary Word.run context.

B: 22
The  |
| L1-image-insert-002 | 40 ⚠️ | ✓ | The code uses a non-existent method "insertImage" and lacks any implementation to actually add an image to the page, and it does not follow best practices by re-implementing the image insertion manually. |
| L1-watermark-insert-001 | 0 ⚠️ | ✓ | A: 0
The code uses a non-existent method `insertWatermark()` which is not a valid Word Office.js API.

B: 0
The code doe |
| L1-content-control-001 | 85 | ✓ | The code uses the correct Office.js API, but it lacks a crucial check to ensure the content control is inserted at the correct location, and it doesn't follow best practices by not using the available helpers for inserting content controls. |
| L1-content-control-002 | 85 | ✓ | A: 20
The code uses the correct Word Office.js API, `insertContentControl`, but lacks the `Word.run` context.

B: 25
The |
| L1-mail-merge-001 | 75 | ✓ | A: 22
The code uses the correct Word Office.js API, mailMergeReplace, but lacks proper error handling and assumes the te |
| L1-template-apply-001 | 85 | ✓ | The code uses the correct Word Office.js API, but it lacks error handling and assumes the template "resume" exists, and it uses a hardcoded template name instead of a dynamic one, and it doesn't check if the template is applied successfully. |
| L1-template-apply-002 | 85 | ✓ | A: 22
The code uses the correct Word API, applyTemplate, but does not check if the template exists or if the document is |
| L1-document-generate-001 | 75 | ✓ | A: 18
The code uses the correct Word Office.js APIs, but it incorrectly uses `addTitle` instead of `addHeading` for the  |
| L1-theme-apply-001 | 40 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js APIs, but it lacks the actual implementation of the designTheme and apply |
| L1-theme-apply-002 | 45 ⚠️ | ✓ | A: 20
The code uses the correct Word API, designTheme() and applyTheme(), but it's missing the necessary import statemen |
| L1-theme-apply-003 | 67 | ✓ | A: 22
The code uses the correct Word Office.js APIs, but it lacks a crucial detail in the tweakTheme method call.

B: 18 |
| L1-citation-bibliography-001 | 40 ⚠️ | ✓ | A: 22
The code uses the correct Word Office.js API, `Word.run`, but incorrectly assumes the `insertCitation` and `insert |
| L1-equation-001 | 40 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, `insertEquation`, but lacks the necessary Word.run context.

B: 0
Th |
| L1-equation-002 | 94 | ✓ | The code uses the correct Word Office.js API, fully addresses the request, would execute without runtime errors, but could be improved by using the available helpers for equation insertion. |
| L1-writing-coach-001 | 64 | ✗ | A: 22
The response is factually correct, but it doesn't address the "for clarity" aspect, which implies a need for a mor |
| L1-read-query-001 | 8 ⚠️ | ✓ | The code is a hallucinated solution that does not use any real Word Office.js APIs, does not address the request, and would not execute in Word without runtime errors due to the non-existent `countWords()` method. |
| L1-read-query-002 | 40 ⚠️ | ✓ | A: 18
The code uses the `getReadability()` function, which is not a real Word Office.js API, and `r.flesch`, `r.grade`,  |
| L1-read-query-003 | 22 ⚠️ | ✗ | The code relies on a non-existent method, which is a fundamental flaw. |
| L1-margins-orientation-001 | 47 ⚠️ | ✓ | A: 22
The code uses the correct Word Office.js API, setMargins, but does not account for the context of default margins. |
| L1-margins-orientation-002 | 47 ⚠️ | ✓ | A: 22
The code uses the correct Word Office.js API, `setPageOrientation`, but lacks the necessary Word.run context.

B:  |
| L1-web-search-needed-001 | 22 ⚠️ | ✗ | The response fails to address the request and provide a solution, instead providing a generic refusal. |
| L1-web-search-needed-002 | 60 | ✓ | The response is accurate and clear, but could be more concise and directly address the question without padding. |
| L1-refusal-001 | 47 ⚠️ | ✓ | A: 18
The response is factually correct that Word Office.js doesn't expose direct PDF export, but it's incomplete as it  |
| L1-refusal-002 | 65 | ✓ | The response is clear and directly addresses the question, but fails to offer any workarounds or additional information. |
| L1-refusal-003 | 47 ⚠️ | ✓ | A: 18
The response is factually correct that Word Office.js doesn't expose direct chart creation, but it doesn't provide |
| L1-quote-insert-001 | 47 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, `addQuote`, but lacks proper context and handling.

B: 0
The code do |
| L1-paragraph-spacing-001 | 85 | ✓ | A: 22
The code uses the correct Word Office.js API, but it's missing the `await context.sync()` call after the `forEach` |
| L1-edge-case-001 | 0 ⚠️ | ✓ | The AI response did not generate any code, failing to address the request, execute correctly, or demonstrate best practices. |
| L1-form-field-001 | 85 | ✓ | A: 22
The code uses the correct Word Office.js API, `insertContentControl`, but lacks the necessary `Word.run` context.
 |
| L1-text-insert-001 | 40 ⚠️ | ✓ | A: 12
The code uses the `addParagraph` method, which is a real Word Office.js API, but it's not the recommended way to i |
| L1-text-edit-001 | 89 | ✓ | A: 22
The code uses the correct Word API, addTitle, but does not handle the case where the title already exists.

B: 25
 |
| gen-L2-form-field-001 | 63 | ✓ | A: 18
The code uses the correct Word Office.js API, `insertContentControl`, but it's missing the crucial `Word.run` cont |
| gen-L2-form-field-002 | 60 | ✗ | A: 20
The code uses the correct Word Office.js API, `insertContentControl`, but incorrectly assumes it can insert a rich |
| gen-L2-form-field-003 | 85 | ✗ | A: 20
The code uses the correct Word Office.js APIs, but it's missing the `await Word.run()` wrapper, which is essential |
| gen-L2-form-field-004 | 70 | ✗ | A: 18
The code uses the correct Word Office.js APIs, but it lacks the crucial `Word.run` function to execute the code in |

## ⚠️ Needs attention

**[L1-paragraph-format-001]** score=22 — The code uses a real Word Office.js API, but it does not address the request to make the first paragraph bold and 14pt, and it does not use the available helpers to achieve this.
- missing patterns: `font.bold`, `font.size`

**[L1-list-bullet-001]** score=47 — A: 22
The code uses the correct Word Office.js API, addList, but lacks the necessary Word.run context.

B: 0
The code do

**[L1-find-replace-001]** score=47 — A: 18
The code uses the correct Word Office.js API, `replaceText`, but does not account for the context of the document,

**[L1-comment-insert-001]** score=47 — A: 18
The code uses the correct Word Office.js API, Word.run, but lacks proper error handling and context.

B: 0
The cod

**[L1-section-break-001]** score=52 — A: 22
The code uses the correct Word Office.js API, `insertSectionBreak` and `setPageOrientation`, but it's missing the 

**[L1-header-footer-002]** score=47 — A: 18
The code uses a real Word Office.js API, but it's incomplete and doesn't account for the specific request of addin

**[L1-image-insert-002]** score=40 — The code uses a non-existent method "insertImage" and lacks any implementation to actually add an image to the page, and it does not follow best practices by re-implementing the image insertion manually.

**[L1-watermark-insert-001]** score=0 — A: 0
The code uses a non-existent method `insertWatermark()` which is not a valid Word Office.js API.

B: 0
The code doe

**[L1-theme-apply-001]** score=40 — A: 18
The code uses the correct Word Office.js APIs, but it lacks the actual implementation of the designTheme and apply

**[L1-theme-apply-002]** score=45 — A: 20
The code uses the correct Word API, designTheme() and applyTheme(), but it's missing the necessary import statemen

**[L1-citation-bibliography-001]** score=40 — A: 22
The code uses the correct Word Office.js API, `Word.run`, but incorrectly assumes the `insertCitation` and `insert

**[L1-equation-001]** score=40 — A: 18
The code uses the correct Word Office.js API, `insertEquation`, but lacks the necessary Word.run context.

B: 0
Th

**[L1-read-query-001]** score=8 — The code is a hallucinated solution that does not use any real Word Office.js APIs, does not address the request, and would not execute in Word without runtime errors due to the non-existent `countWords()` method.

**[L1-read-query-002]** score=40 — A: 18
The code uses the `getReadability()` function, which is not a real Word Office.js API, and `r.flesch`, `r.grade`, 

**[L1-read-query-003]** score=22 — The code relies on a non-existent method, which is a fundamental flaw.
- missing patterns: `return`

**[L1-margins-orientation-001]** score=47 — A: 22
The code uses the correct Word Office.js API, setMargins, but does not account for the context of default margins.

**[L1-margins-orientation-002]** score=47 — A: 22
The code uses the correct Word Office.js API, `setPageOrientation`, but lacks the necessary Word.run context.

B: 

**[L1-web-search-needed-001]** score=22 — The response fails to address the request and provide a solution, instead providing a generic refusal.
- missing patterns: `NEEDS_WEB_SEARCH`

**[L1-refusal-001]** score=47 — A: 18
The response is factually correct that Word Office.js doesn't expose direct PDF export, but it's incomplete as it 

**[L1-refusal-003]** score=47 — A: 18
The response is factually correct that Word Office.js doesn't expose direct chart creation, but it doesn't provide

**[L1-quote-insert-001]** score=47 — A: 18
The code uses the correct Word Office.js API, `addQuote`, but lacks proper context and handling.

B: 0
The code do

**[L1-edge-case-001]** score=0 — The AI response did not generate any code, failing to address the request, execute correctly, or demonstrate best practices.

**[L1-text-insert-001]** score=40 — A: 12
The code uses the `addParagraph` method, which is a real Word Office.js API, but it's not the recommended way to i
