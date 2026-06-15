# Word AI — Eval Results
**Last run:** 2026-06-15 04:23:47  
**Overall: 62.0/100**  
**Model:** meta-llama/llama-3.1-8b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 56.0/100 | — | 2 |
| paragraph-format | 74.5/100 | — | 2 |
| list-bullet | 47.0/100 | — | 1 |
| list-multilevel | 91.0/100 | — | 1 |
| style-apply | 91.0/100 | — | 1 |
| table-create | 91.0/100 | — | 1 |
| table-format | 94.0/100 | — | 1 |
| find-replace | 62.0/100 | — | 2 |
| find-replace-regex | 72.0/100 | — | 1 |
| footnote-insert | 47.0/100 | — | 1 |
| comment-insert | 94.0/100 | — | 1 |
| track-changes-toggle | 75.0/100 | — | 2 |
| toc-generate | 65.0/100 | — | 1 |
| section-break | 85.0/100 | — | 1 |
| header-footer | 56.0/100 | — | 2 |
| columns | 52.0/100 | — | 1 |
| image-insert | 46.0/100 | — | 2 |
| watermark-insert | 25.0/100 | — | 1 |
| content-control | 87.0/100 | — | 2 |
| mail-merge | 75.0/100 | — | 1 |
| template-apply | 74.5/100 | — | 2 |
| document-generate | 83.0/100 | — | 1 |
| theme-apply | 63.0/100 | — | 3 |
| citation-bibliography | 75.0/100 | — | 1 |
| equation | 56.0/100 | — | 2 |
| writing-coach | 60.0/100 | — | 1 |
| read-query | 53.7/100 | — | 3 |
| margins-orientation | 53.5/100 | — | 2 |
| web-search-needed | 32.0/100 | — | 2 |
| refusal | 49.7/100 | — | 3 |
| quote-insert | 45.0/100 | — | 1 |
| paragraph-spacing | 77.0/100 | — | 1 |
| edge-cases | 0.0/100 | — | 1 |
| form-field | 66.8/100 | — | 5 |
| text-insert | 0.0/100 | — | 1 |
| text-edit | 89.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 45 ⚠️ | ✓ | A: 22
The code uses the correct Word API, `addHeading`, but does not check if the heading already exists, which could le |
| L1-heading-insert-002 | 67 | ✓ | A: 22
The code uses the correct Word Office.js API, Word.run, and the addHeading and addParagraph functions, but it lack |
| L1-paragraph-format-001 | 60 | ✗ | A: 20
The code uses the correct Word Office.js API, but it incorrectly uses the "Heading 1" style, which is not the same |
| L1-paragraph-format-002 | 89 | ✓ | The code uses real Word Office.js APIs, fully addresses the request, and would likely execute without runtime errors, but it doesn't use best practices and available helpers, instead re-implementing the functionality manually. |
| L1-list-bullet-001 | 47 ⚠️ | ✓ | A: 22
The code uses the correct Word Office.js API, `addList`, but lacks proper error handling and context awareness.

B |
| L1-list-multilevel-001 | 91 | ✓ | The code uses the correct Word Office.js API, but deducts points for not addressing the "add" request explicitly, and for not using the available helpers for inserting a list. |
| L1-style-apply-001 | 91 | ✓ | A: 22
The code uses the correct Word Office.js API, `applyStyle`, but I deduct 3 points for not checking if the paragrap |
| L1-table-create-001 | 91 | ✓ | A: 22
The code uses the correct Word Office.js API, `insertTable`, but deducts 3 points for not using the `Word.run` met |
| L1-table-format-001 | 94 | ✓ | The code correctly applies the style to the existing table, but lacks explicit error handling and uses a direct method call without checking if the table exists. |
| L1-find-replace-001 | 40 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, `replaceText`, but does not account for the context of the document, |
| L1-find-replace-002 | 84 | ✓ | A: 22
The code uses the correct Word Office.js APIs, but it does not handle the case where no matches are found in the s |
| L1-find-replace-regex-001 | 72 | ✓ | A: 22
The code uses the correct Word Office.js API, `replaceText`, but does not account for the specific email address f |
| L1-footnote-insert-001 | 47 ⚠️ | ✓ | A: 22
The code uses the correct Word Office.js API, `insertFootnote`, but lacks the necessary context to execute correct |
| L1-comment-insert-001 | 94 | ✓ | The code uses the correct Word Office.js API, fully addresses the request, and would execute without runtime errors, but it could be improved by using the available helpers for inserting comments. |
| L1-track-changes-toggle-001 | 65 | ✓ | A: 22
The code uses the correct Word API, toggleTrackChanges, but does not account for the asynchronous nature of the Of |
| L1-track-changes-toggle-002 | 85 | ✓ | The code is mostly correct, but it lacks a check for the number of revisions to be accepted, and it doesn't use the available helpers for accepting revisions. |
| L1-toc-generate-001 | 65 | ✓ | A: 22
The code uses the correct Word Office.js API, `insertTableOfContents`, but does not specify the document object, w |
| L1-section-break-001 | 85 | ✓ | A: 22
The code uses the correct Word Office.js API, `insertSectionBreak` and `setPageOrientation`, but it's missing the  |
| L1-header-footer-001 | 47 ⚠️ | ✓ | A: 22
The code uses the correct Word Office.js API, `setHeader`, but lacks the necessary Word.run context.

B: 0
The cod |
| L1-header-footer-002 | 65 | ✓ | The code uses a real Word Office.js API, fully addresses the request, and would execute without runtime errors, but it fails to use best practices and helpers, instead directly calling the addPageNumbers function without leveraging the available helpers. |
| L1-columns-001 | 52 ⚠️ | ✓ | A: 22
The code uses real Word Office.js APIs, but it's missing the crucial step of applying the two-column layout to the |
| L1-image-insert-001 | 45 ⚠️ | ✓ | A: 22
The code uses the correct Word Office.js API, insertImage, but lacks the correct syntax for specifying the image t |
| L1-image-insert-002 | 47 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, `insertImage`, but lacks context and parameters for the image.

B: 0 |
| L1-watermark-insert-001 | 25 ⚠️ | ✓ | The code is entirely incorrect, using a non-existent method, and fails to address the request, while also having potential runtime errors due to missing context. |
| L1-content-control-001 | 85 | ✓ | A: 22
The code uses the correct Word Office.js API, `insertContentControl`, but lacks the necessary `Word.run` block to  |
| L1-content-control-002 | 89 | ✓ | A: 22
The code uses the correct Word Office.js API, insertContentControl, but the tag property is not a valid argument f |
| L1-mail-merge-001 | 75 | ✓ | A: 22
The code uses the correct Word Office.js API, mailMergeReplace, but lacks proper error handling and assumes the re |
| L1-template-apply-001 | 77 | ✓ | The main flaw is the incorrect assumption about the |
| L1-template-apply-002 | 72 | ✓ | A: 22
The code uses the correct Word Office.js API, applyTemplate, but does not account for potential errors or exceptio |
| L1-document-generate-001 | 83 | ✓ | A: 20
The code uses real Word Office.js APIs, but there are some minor issues with method signatures and variable scope. |
| L1-theme-apply-001 | 37 ⚠️ | ✓ | The code uses real Word Office.js APIs, but it lacks any actual implementation to address the request, and it uses a hallucinated method name "designTheme" which is not a real Word API. |
| L1-theme-apply-002 | 67 | ✓ | The code uses real Word Office.js APIs, but lacks a crucial detail to fully address the request, and has a flawed approach by not using the recommended helpers for theme design and application. |
| L1-theme-apply-003 | 85 | ✓ | A: 20
The code uses the correct Word Office.js API, tweakTheme and applyTheme, but incorrectly assumes the tweakTheme me |
| L1-citation-bibliography-001 | 75 | ✓ | A: 22
The code uses the correct Word API, `Word.run`, but it's missing the necessary `async/await` wrapper and the `awai |
| L1-equation-001 | 65 | ✓ | A: 18
The code uses the correct Word Office.js API, `insertEquation`, but lacks proper error handling and assumes the eq |
| L1-equation-002 | 47 ⚠️ | ✓ | A: 22
The code uses the correct Word Office.js API, `insertEquation`, but lacks the necessary setup and context to execu |
| L1-writing-coach-001 | 60 | ✗ | A: 22
The response is factually correct and complete, but it doesn't address the 'for clarity' part, which implies a nee |
| L1-read-query-001 | 77 | ✓ | The code uses the correct Word API, but lacks proper error handling and uses a manual approach to counting words, which is not a recommended practice. |
| L1-read-query-002 | 47 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, getReadability(), but does not fully address the request.

B: 0
The  |
| L1-read-query-003 | 37 ⚠️ | ✗ | The code is incomplete |
| L1-margins-orientation-001 | 60 | ✓ | A: 20
The code uses the correct Word Office.js API, `setMargins`, but does not account for the context of default margin |
| L1-margins-orientation-002 | 47 ⚠️ | ✓ | A: 22
The code uses the correct Word Office.js API, `setPageOrientation`, but it's missing the context of the document,  |
| L1-web-search-needed-001 | 22 ⚠️ | ✗ | A: 0
The response is factually incorrect as it implies the Office.js API cannot provide information on the current stock |
| L1-web-search-needed-002 | 42 ⚠️ | ✓ | A: 20
The response accurately explains the Pythagorean theorem, but it doesn't provide any information about how it appl |
| L1-refusal-001 | 65 | ✓ | The main flaw is that the response fails to provide a clear and actionable solution for exporting the document |
| L1-refusal-002 | 47 ⚠️ | ✓ | A: 18
The response is partially accurate, but it doesn't mention that Word Office.js can use the `Office.context.documen |
| L1-refusal-003 | 37 ⚠️ | ✓ | The response fails to provide a solution that meets the user's |
| L1-quote-insert-001 | 45 ⚠️ | ✓ | A: 20
The code uses the correct Word API, addQuote, but it's a very basic method that doesn't provide any options for cu |
| L1-paragraph-spacing-001 | 77 | ✓ | A: 22
The code uses the correct Word Office.js APIs, but it's missing the initial Word.run() call.

B: 25
The code fully |
| L1-edge-case-001 | 0 ⚠️ | ✓ | The code does not attempt to delete paragraph 99, does not use any Word Office.js APIs, and does not address the request at all. |
| L1-form-field-001 | 85 | ✓ | A: 22
The code uses the correct Word Office.js API, `insertContentControl`, but it's missing the `Word.run` context setu |
| L1-text-insert-001 | 0 ⚠️ | ✓ | The code uses a non-existent method `addParagraph()` which is not a valid Word Office.js API. |
| L1-text-edit-001 | 89 | ✓ | A: 22
The code uses the correct Word Office.js API, `addTitle`, but lacks the `Word.run` context.

B: 25
The code fully  |
| gen-L2-form-field-001 | 47 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, `insertContentControl`, but lacks the necessary context to insert th |
| gen-L2-form-field-002 | 60 | ✗ | A: 20
The code uses the correct Word Office.js API, `insertContentControl`, but incorrectly assumes the existing plain t |
| gen-L2-form-field-003 | 75 | ✗ | A: 22
The code uses the correct Word Office.js API, `insertContentControl`, but lacks proper error handling and assumes  |
| gen-L2-form-field-004 | 67 | ✗ | A: 20
The code uses the correct Word Office.js APIs, but it's missing the crucial `await Word.run` block to execute the  |

## ⚠️ Needs attention

**[L1-heading-insert-001]** score=45 — A: 22
The code uses the correct Word API, `addHeading`, but does not check if the heading already exists, which could le

**[L1-list-bullet-001]** score=47 — A: 22
The code uses the correct Word Office.js API, `addList`, but lacks proper error handling and context awareness.

B

**[L1-find-replace-001]** score=40 — A: 18
The code uses the correct Word Office.js API, `replaceText`, but does not account for the context of the document,

**[L1-footnote-insert-001]** score=47 — A: 22
The code uses the correct Word Office.js API, `insertFootnote`, but lacks the necessary context to execute correct

**[L1-header-footer-001]** score=47 — A: 22
The code uses the correct Word Office.js API, `setHeader`, but lacks the necessary Word.run context.

B: 0
The cod

**[L1-columns-001]** score=52 — A: 22
The code uses real Word Office.js APIs, but it's missing the crucial step of applying the two-column layout to the

**[L1-image-insert-001]** score=45 — A: 22
The code uses the correct Word Office.js API, insertImage, but lacks the correct syntax for specifying the image t

**[L1-image-insert-002]** score=47 — A: 18
The code uses the correct Word Office.js API, `insertImage`, but lacks context and parameters for the image.

B: 0

**[L1-watermark-insert-001]** score=25 — The code is entirely incorrect, using a non-existent method, and fails to address the request, while also having potential runtime errors due to missing context.

**[L1-theme-apply-001]** score=37 — The code uses real Word Office.js APIs, but it lacks any actual implementation to address the request, and it uses a hallucinated method name "designTheme" which is not a real Word API.

**[L1-equation-002]** score=47 — A: 22
The code uses the correct Word Office.js API, `insertEquation`, but lacks the necessary setup and context to execu

**[L1-read-query-002]** score=47 — A: 18
The code uses the correct Word Office.js API, getReadability(), but does not fully address the request.

B: 0
The 

**[L1-read-query-003]** score=37 — The code is incomplete
- missing patterns: `return`

**[L1-margins-orientation-002]** score=47 — A: 22
The code uses the correct Word Office.js API, `setPageOrientation`, but it's missing the context of the document, 

**[L1-web-search-needed-001]** score=22 — A: 0
The response is factually incorrect as it implies the Office.js API cannot provide information on the current stock
- missing patterns: `NEEDS_WEB_SEARCH`

**[L1-web-search-needed-002]** score=42 — A: 20
The response accurately explains the Pythagorean theorem, but it doesn't provide any information about how it appl

**[L1-refusal-002]** score=47 — A: 18
The response is partially accurate, but it doesn't mention that Word Office.js can use the `Office.context.documen

**[L1-refusal-003]** score=37 — The response fails to provide a solution that meets the user's

**[L1-quote-insert-001]** score=45 — A: 20
The code uses the correct Word API, addQuote, but it's a very basic method that doesn't provide any options for cu

**[L1-edge-case-001]** score=0 — The code does not attempt to delete paragraph 99, does not use any Word Office.js APIs, and does not address the request at all.

**[L1-text-insert-001]** score=0 — The code uses a non-existent method `addParagraph()` which is not a valid Word Office.js API.

**[gen-L2-form-field-001]** score=47 — A: 18
The code uses the correct Word Office.js API, `insertContentControl`, but lacks the necessary context to insert th
