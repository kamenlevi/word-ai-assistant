# Word AI — Eval Results
**Last run:** 2026-06-22 04:20:45  
**Overall: 64.7/100**  
**Model:** meta-llama/llama-3.1-8b-instruct

## Scores by category

| Category | Score | Trend | Cases |
|---|---|---|---|
| heading-insert | 66.0/100 | — | 2 |
| paragraph-format | 47.0/100 | — | 2 |
| list-bullet | 60.0/100 | — | 1 |
| list-multilevel | 91.0/100 | — | 1 |
| style-apply ✓ | 95.0/100 | — | 1 |
| table-create | 85.0/100 | — | 1 |
| table-format | 89.0/100 | — | 1 |
| find-replace | 66.0/100 | — | 2 |
| find-replace-regex | 89.0/100 | — | 1 |
| footnote-insert | 89.0/100 | — | 1 |
| comment-insert | 60.0/100 | — | 1 |
| track-changes-toggle | 91.5/100 | — | 2 |
| toc-generate | 75.0/100 | — | 1 |
| section-break | 64.0/100 | — | 1 |
| header-footer | 87.0/100 | — | 2 |
| columns | 47.0/100 | — | 1 |
| image-insert | 47.0/100 | — | 2 |
| watermark-insert | 12.0/100 | — | 1 |
| content-control | 87.0/100 | — | 2 |
| mail-merge | 75.0/100 | — | 1 |
| template-apply | 84.0/100 | — | 2 |
| document-generate | 55.0/100 | — | 1 |
| theme-apply | 77.0/100 | — | 3 |
| citation-bibliography | 85.0/100 | — | 1 |
| equation | 47.0/100 | — | 2 |
| writing-coach | 65.0/100 | — | 1 |
| read-query | 30.0/100 | — | 3 |
| margins-orientation | 66.0/100 | — | 2 |
| web-search-needed | 57.0/100 | — | 2 |
| refusal | 47.0/100 | — | 3 |
| quote-insert | 83.0/100 | — | 1 |
| paragraph-spacing | 89.0/100 | — | 1 |
| edge-cases | 12.0/100 | — | 1 |
| form-field | 61.2/100 | — | 5 |
| text-insert | 22.0/100 | — | 1 |
| text-edit | 89.0/100 | — | 1 |

## All test cases

| ID | Score | Pass | Reason |
|---|---|---|---|
| L1-heading-insert-001 | 47 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, addHeading, but lacks proper error handling and context awareness.

 |
| L1-heading-insert-002 | 85 | ✓ | A: 22
The code uses the correct Word Office.js API, Word.run, and the addHeading and addParagraph functions are real API |
| L1-paragraph-format-001 | 22 ⚠️ | ✗ | A: 0
The code uses a non-existent method "applyStyle" which is not a part of the Office.js API.

B: 0
The code does not  |
| L1-paragraph-format-002 | 72 | ✓ | The code uses real Word Office.js APIs, fully addresses the request, and would execute without runtime errors, but it doesn't use the available helpers and instead re-implements the addParagraph function manually. |
| L1-list-bullet-001 | 60 | ✓ | A: 20
The code uses the correct Word Office.js API, `addList`, but lacks the necessary `Word.run` context.

B: 0
The cod |
| L1-list-multilevel-001 | 91 | ✓ | The code uses the correct Word Office.js API, fully addresses the request, and would execute without runtime errors, but it doesn't use the available helper for adding a list, and it doesn't handle potential errors that might occur when adding the list. |
| L1-style-apply-001 | 95 ✅ | ✓ | The code uses the correct Word Office.js API, fully addresses the request, and would execute without runtime errors, but it could benefit from using the `applyStyle` method with a more robust way to specify the style, such as `paragraph.styleBuiltIn` instead of a string. |
| L1-table-create-001 | 85 | ✓ | A: 22
The code uses the correct Word Office.js API, `insertTable`, but I deduct 3 points for not using the `Word.run` me |
| L1-table-format-001 | 89 | ✓ | The code uses the correct Word Office.js API, but it lacks the necessary Word.run context and the table ID is not properly validated, and it does not use the available helpers for styling tables. |
| L1-find-replace-001 | 47 ⚠️ | ✓ | A: 22
The code uses the correct Word Office.js API, replaceText, but lacks context and handling for potential errors.

B |
| L1-find-replace-002 | 85 | ✓ | The code correctly uses the Word Office.js API, but it does not address the request fully, as it only makes the text bold and red, but does not change the color to red, and it uses a hardcoded color value instead of a named color. |
| L1-find-replace-regex-001 | 89 | ✓ | The code uses the correct Word Office.js API, but deducts points for not using the available helpers for wildcard replacement and not checking if the replacement is case-sensitive. |
| L1-footnote-insert-001 | 89 | ✓ | The code uses the correct Office.js API, fully addresses the request, and would execute without runtime errors, but it doesn't use the available helpers for inserting a footnote, instead re-implementing it manually. |
| L1-comment-insert-001 | 60 | ✓ | A: 22
The code uses the correct Word Office.js APIs, but it's missing the `await Word.run()` wrapper, which is essential |
| L1-track-changes-toggle-001 | 94 | ✓ | The code uses the correct Word Office.js API to toggle track changes, but it could benefit from using the available helpers for a more idiomatic and efficient solution. |
| L1-track-changes-toggle-002 | 89 | ✓ | The code correctly uses the Office.js API, fully addresses the request, and would execute without runtime errors, but it fails to use best practices by not checking if there are any tracked revisions before accepting them. |
| L1-toc-generate-001 | 75 | ✓ | A: 22
The code uses the correct Word Office.js API, `insertTableOfContents`, but lacks proper error handling and assumes |
| L1-section-break-001 | 64 | ✓ | A: 22
The code uses the correct Word Office.js APIs, but it lacks the necessary context to switch the next section to la |
| L1-header-footer-001 | 85 | ✓ | The code uses the correct Word Office.js API, but it doesn't address the request fully, and it uses a manual approach to setting the header instead of the available helper method addHeading. |
| L1-header-footer-002 | 89 | ✓ | The code uses the correct Word Office.js API, fully addresses the request, and would execute without runtime errors, but it doesn't use the available helpers and instead re-implements the functionality manually. |
| L1-columns-001 | 47 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, `insertSectionBreak` and `insertColumns`, but it lacks the necessary |
| L1-image-insert-001 | 47 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, `insertImage`, but lacks the necessary Word.run context.

B: 0
The c |
| L1-image-insert-002 | 47 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, insertImage, but lacks context and parameters for the image.

B: 0
T |
| L1-watermark-insert-001 | 12 ⚠️ | ✓ | The code is a hallucinated method call with no actual implementation, and does not address the request of adding a watermark to every page, nor does it consider the best practices for using the Office.js API. |
| L1-content-control-001 | 89 | ✓ | The code uses the correct Office.js API, fully addresses the request, and would execute without runtime errors, but it doesn't use the available helpers for inserting content controls, which is a missed opportunity for best practices. |
| L1-content-control-002 | 85 | ✓ | A: 22
The code uses the correct Office.js API, `insertContentControl`, but lacks the necessary Word.run context.

B: 25
 |
| L1-mail-merge-001 | 75 | ✓ | A: 22
The code uses the correct Word Office.js API, mailMergeReplace, but lacks proper error handling and context.

B: 1 |
| L1-template-apply-001 | 83 | ✓ | The code uses the correct Word Office.js API, applyTemplate, but lacks a crucial check for the template's existence, which would cause a runtime error if the template is not found. |
| L1-template-apply-002 | 85 | ✓ | A: 22
The code uses the correct Word Office.js API, applyTemplate, but does not handle potential errors or exceptions.

 |
| L1-document-generate-001 | 55 ⚠️ | ✓ | The code is incomplete and lacks |
| L1-theme-apply-001 | 52 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js APIs, but it lacks crucial details for a complete solution.

B: 5
The cod |
| L1-theme-apply-002 | 85 | ✓ | The code uses real Word Office.js APIs, fully addresses the request, and would likely execute without runtime errors, but it doesn't use best practices and available helpers, specifically re-implementing the theme design and application manually. |
| L1-theme-apply-003 | 94 | ✓ | The code uses real Word Office.js APIs, fully addresses the request, and would execute without runtime errors, but it could be improved by using the designTheme() helper to specify the theme changes. |
| L1-citation-bibliography-001 | 85 | ✓ | A: 20
The code uses the correct Word Office.js API, `Word.run`, but it's missing the essential `context.sync()` call to  |
| L1-equation-001 | 47 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, `insertEquation`, but lacks the necessary Word.run context.

B: 0
Th |
| L1-equation-002 | 47 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, `insertEquation`, but lacks proper error handling and context awaren |
| L1-writing-coach-001 | 65 | ✗ | A: 22
The code accurately replaces the selection text with a clearer alternative, but it doesn't address the "rewrite fo |
| L1-read-query-001 | 18 ⚠️ | ✗ | The code does not use any real Word Office.js APIs, and the approach is completely manual and incorrect, with no consideration for best practices or available helpers. |
| L1-read-query-002 | 72 | ✓ | The code uses the correct Word Office.js API, getReadability(), but lacks best practices and available helpers, specifically for handling themes and images. |
| L1-read-query-003 | 0 ⚠️ | ✗ | The code relies on a non-existent method, which is a fundamental flaw in its design. |
| L1-margins-orientation-001 | 85 | ✓ | A: 20
The code uses the correct Word Office.js API, setMargins, but does not account for the unit of measurement (inches |
| L1-margins-orientation-002 | 47 ⚠️ | ✓ | A: 18
The code uses the correct Word API, `setPageOrientation`, but it's missing the context of the document, which is d |
| L1-web-search-needed-001 | 47 ⚠️ | ✓ | The response fails to directly address the user's request and instead focuses on the limitation of the Office.js API. |
| L1-web-search-needed-002 | 67 | ✓ | The response accurately and clearly explains the Pythagorean theorem, but fails to provide any relevant code or workarounds, and does not acknowledge the context of the Office.js API. |
| L1-refusal-001 | 47 ⚠️ | ✓ | A: 18
The response is factually correct that Word Office.js doesn't expose direct PDF export, but it fails to mention th |
| L1-refusal-002 | 47 ⚠️ | ✗ | The response is clear and directly addresses the question, but fails to provide any relevant information or code, and does not offer any workarounds or suggestions. |
| L1-refusal-003 | 47 ⚠️ | ✓ | A: 18
The response is partially correct, but it doesn't fully address the user's request for a dynamic Excel-style bar c |
| L1-quote-insert-001 | 83 | ✓ | The code uses the correct Word API, fully addresses the request, and would execute without runtime errors, but it doesn't use the available helpers and instead re-implements the addQuote method manually. |
| L1-paragraph-spacing-001 | 89 | ✓ | A: 22
The code uses the correct Word Office.js APIs, but it's missing the context variable declaration.

B: 25
The code  |
| L1-edge-case-001 | 12 ⚠️ | ✓ | The AI response does not address the original request to delete paragraph 99, and the generated code is a partial solution that only deletes the first paragraph. |
| L1-form-field-001 | 67 | ✓ | A: 22
The code uses the correct Office.js API, `insertContentControl`, but lacks proper error handling and assumes the ` |
| L1-text-insert-001 | 22 ⚠️ | ✓ | The code uses a non-existent method `addParagraph` and lacks proper error handling and scope management, which would likely cause runtime errors. |
| L1-text-edit-001 | 89 | ✓ | A: 22
The code uses the correct Word Office.js API, `addTitle`, but does not handle the case where the title is already  |
| gen-L2-form-field-001 | 47 ⚠️ | ✓ | A: 18
The code uses the correct Word Office.js API, `insertContentControl`, but lacks the necessary Word.run context and |
| gen-L2-form-field-002 | 60 | ✗ | A: 20
The code uses the correct Word Office.js API, `insertContentControl`, but lacks the crucial step of replacing the  |
| gen-L2-form-field-003 | 85 | ✗ | The code uses real Word Office.js APIs, but lacks proper handling of the group content control's title and the dropdown content controls' tag attributes, which are not explicitly set in the generated code. |
| gen-L2-form-field-004 | 47 ⚠️ | ✗ | A: 18
The code uses the correct Word Office.js API, `insertContentControl`, but lacks the crucial `Word.run` context.

B |

## ⚠️ Needs attention

**[L1-heading-insert-001]** score=47 — A: 18
The code uses the correct Word Office.js API, addHeading, but lacks proper error handling and context awareness.



**[L1-paragraph-format-001]** score=22 — A: 0
The code uses a non-existent method "applyStyle" which is not a part of the Office.js API.

B: 0
The code does not 
- missing patterns: `font.bold`, `font.size`

**[L1-find-replace-001]** score=47 — A: 22
The code uses the correct Word Office.js API, replaceText, but lacks context and handling for potential errors.

B

**[L1-columns-001]** score=47 — A: 18
The code uses the correct Word Office.js API, `insertSectionBreak` and `insertColumns`, but it lacks the necessary

**[L1-image-insert-001]** score=47 — A: 18
The code uses the correct Word Office.js API, `insertImage`, but lacks the necessary Word.run context.

B: 0
The c

**[L1-image-insert-002]** score=47 — A: 18
The code uses the correct Word Office.js API, insertImage, but lacks context and parameters for the image.

B: 0
T

**[L1-watermark-insert-001]** score=12 — The code is a hallucinated method call with no actual implementation, and does not address the request of adding a watermark to every page, nor does it consider the best practices for using the Office.js API.

**[L1-document-generate-001]** score=55 — The code is incomplete and lacks

**[L1-theme-apply-001]** score=52 — A: 18
The code uses the correct Word Office.js APIs, but it lacks crucial details for a complete solution.

B: 5
The cod

**[L1-equation-001]** score=47 — A: 18
The code uses the correct Word Office.js API, `insertEquation`, but lacks the necessary Word.run context.

B: 0
Th

**[L1-equation-002]** score=47 — A: 18
The code uses the correct Word Office.js API, `insertEquation`, but lacks proper error handling and context awaren

**[L1-read-query-001]** score=18 — The code does not use any real Word Office.js APIs, and the approach is completely manual and incorrect, with no consideration for best practices or available helpers.
- missing patterns: `return`

**[L1-read-query-003]** score=0 — The code relies on a non-existent method, which is a fundamental flaw in its design.
- missing patterns: `return`

**[L1-margins-orientation-002]** score=47 — A: 18
The code uses the correct Word API, `setPageOrientation`, but it's missing the context of the document, which is d

**[L1-web-search-needed-001]** score=47 — The response fails to directly address the user's request and instead focuses on the limitation of the Office.js API.

**[L1-refusal-001]** score=47 — A: 18
The response is factually correct that Word Office.js doesn't expose direct PDF export, but it fails to mention th

**[L1-refusal-002]** score=47 — The response is clear and directly addresses the question, but fails to provide any relevant information or code, and does not offer any workarounds or suggestions.
- missing patterns: `doesn't expose`

**[L1-refusal-003]** score=47 — A: 18
The response is partially correct, but it doesn't fully address the user's request for a dynamic Excel-style bar c

**[L1-edge-case-001]** score=12 — The AI response does not address the original request to delete paragraph 99, and the generated code is a partial solution that only deletes the first paragraph.

**[L1-text-insert-001]** score=22 — The code uses a non-existent method `addParagraph` and lacks proper error handling and scope management, which would likely cause runtime errors.

**[gen-L2-form-field-001]** score=47 — A: 18
The code uses the correct Word Office.js API, `insertContentControl`, but lacks the necessary Word.run context and

**[gen-L2-form-field-004]** score=47 — A: 18
The code uses the correct Word Office.js API, `insertContentControl`, but lacks the crucial `Word.run` context.

B
- missing patterns: `RepeatingSectionContentControl`
