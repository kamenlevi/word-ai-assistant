# Word AI Assistant

A private, local AI assistant for Microsoft Word — built as an Office Task Pane add-in. Ask in plain English to write a draft, format paragraphs, build a table, design a theme, run a mail merge, fix track-changes, generate a full document. The AI emits Office JavaScript that runs directly inside your Word document.

## What it does

- **Talk to your document.** "Make the title bold and 28pt." "Replace all 'Acme' with 'Globex'." "Add a footnote 'See Appendix A' to this paragraph." The AI writes and executes Word.js so changes appear instantly.
- **Generate full documents from a prompt.** "Write a 5-section whitepaper on Postgres migration for an engineering audience." A guided dialog (topic / audience / tone / length / TOC / cover / bibliography) streams a complete document into Word.
- **AI Theme Designer — unlimited.** Describe any look in natural language ("calm trustworthy theme for a healthcare startup", "1970s sci-fi novel", "brutalist black-and-yellow legal contract") and the AI generates a complete palette + typography object, validates fonts against a Word-safe whitelist and colors against WCAG AA contrast, and applies it across the document. Tweak with "same but darker, swap the accent for green". "Surprise me" rolls a random creative theme.
- **14 document templates** — resume, cover letter, business letter, formal report, proposal, contract, essay, academic paper, blog post, meeting notes, project plan, invoice, memo, whitepaper — with field substitution.
- **Writing coach.** Click "Coach me". Pick scope (selection / paragraph / whole doc) × focus (clarity / tone / grammar / brevity / persuasive / reading-level) × action (review / rewrite / translate). Apply the suggested rewrite with one click.
- **Find & Replace** with Match-case, Whole-word, and Word-wildcards (not JS regex — see cheatsheet below).
- **Mail Merge.** Auto-detects `{{Field}}` and `«Field»` placeholders, accepts CSV or JSON records, and produces one section per record.
- **Track changes** — toggle on/off, accept all, reject all.
- **Citations & bibliography** in APA / MLA / Chicago / Harvard.
- **Equations** — paste LaTeX, get OMML inserted.
- **Content controls** — rich text, plain text, checkbox, dropdown, combo box, date picker.
- **Table of Contents** — single helper, auto-refreshes (right-click → Update Field if needed).
- **Local image asset pack.** Download a curated CC0 image library (18 categories — people, technology, healthcare, document-icons, watermarks, letterhead-graphics, page-borders, …). When the AI needs an image, it picks the best match from your pack. **No AI image generation, ever.**
- **Web search on demand.** When the AI needs live data (today's stock prices, current weather, recent news), it tells you and offers a one-time search you can approve. Or set web search to always-on. **Google Custom Search** is the default provider; Brave Search is a drop-in alternative.
- **Live document stats** — running word count + Flesch grade-level badge in the header.

## Features at a glance

- **CODE_JS protocol**: AI emits `CODE_JS::...::END_CODE` blocks; the task pane wraps them in `Word.run(...)` with every helper in scope.
- **Multi-provider AI**: OpenRouter (default — Llama 3.3 70B), Groq, Ollama (local), MLX (Mac).
- **Chat history** with auto-summarization after 10 messages.
- **Model picker** with filter/sort/search; AI-recommended model based on your hardware.
- **Cost limits** per-prompt or per-session with running session-cost display.
- **Advanced options**: Deep Think · Dynamic Depth · Auto Model · Plan First · Section-by-section preview.
- **Auto-retry** when generated code fails; AI rewrites and re-executes.
- **Eval framework** with LLM-as-judge scoring and auto-improvement via GitHub Actions every 8 hours.

## Requirements

- **Node.js 18+**
- **Microsoft Word** (desktop or web — desktop works most reliably)
- An **OpenRouter API key** (free tier available at openrouter.ai)
- Optional: **Groq API key** for ultra-fast inference
- Optional: a **web search provider**:
  - **Google Custom Search** (default) — 100 free queries/day. Set `GOOGLE_KEY` + `GOOGLE_CX`.
  - **Brave Search** — 2,000 free queries/month. Set `BRAVE_KEY`.
  - Pick one with `WEB_SEARCH_PROVIDER=google|brave` in `.env`.

## Setup

1. Clone this repo and install dependencies:
   ```
   npm install
   ```

2. Copy `.env.example` to `.env` and add your API keys:
   ```
   OPENROUTER_KEY=sk-or-v1-...
   GROQ_KEY=gsk_...
   GOOGLE_KEY=...
   GOOGLE_CX=...
   BRAVE_KEY=...
   WEB_SEARCH_PROVIDER=google
   ```

3. Generate the icon PNGs (one time):
   ```
   node create-icons.js
   ```

4. Trust the dev cert (one time, so Word desktop doesn't complain):
   ```
   npx office-addin-dev-certs install
   ```

5. Start the local HTTPS server:
   ```
   npm start
   ```
   It listens on `https://localhost:3000`. The first run will use either the
   trusted office-addin-dev-certs cert or an ephemeral self-signed cert.

6. Sideload the add-in in Word:
   - **Desktop**: Insert → My Add-ins → Upload My Add-in → choose `manifest.xml`
   - **Web (word.office.com)**: Insert → Add-ins → Upload My Add-in → `manifest.xml`

7. Click the "Word AI" button on the Home tab to open the task pane.

## First-time use

- **Image asset pack**: Go to Settings → Data → click "Download asset pack". This pulls the starter set (~33 CC0 images across 18 categories). The maintainer can run `npm run build-pack` with Pexels/Unsplash/Pixabay API keys to expand to a full ~900-image library.
- **Web search**: By default the AI will *ask before searching*. To enable always-on web access, go to Settings → Data → Web Search → "Always on".
- **Theme**: Click the 🎨 button. Describe any look you want, click Generate, preview, then Apply.

## Example prompts

- "Write a 5-section proposal about Postgres migration with TOC"
- "Apply the resume template with Name=Alice Smith, Title=Engineer, Email=alice@example.com"
- "Make the title bold and 28pt centered"
- "Insert a 4×3 table with headers Quarter, Revenue, Notes"
- "Replace all 'Acme' with 'Globex'"
- "Wildcard-replace email addresses like 'foo@bar.com' with [redacted]"
- "Add a CONFIDENTIAL watermark to every page"
- "Insert a footnote 'See Appendix A' on this paragraph"
- "Insert a Table of Contents at the top"
- "Turn on track changes"
- "Switch the next section to landscape"
- "Design a calm trustworthy theme for a healthcare startup and apply it"
- "Now make it darker and swap the accent for green"
- "Make it look like a 1970s sci-fi novel"
- "Mail merge with these records: Name,Amount\nAlice,100\nBob,200"
- "Insert a date-picker content control labeled Sign Date"
- "Insert the quadratic formula"
- "Add an office-workplace image at the cursor, 320pt wide"
- "How many words does this document have?"
- "What's my Flesch reading-ease grade level?"

## Word wildcards — NOT JS regex

`body.search()` accepts Word wildcards (not JS regex). Cheat sheet:

| Pattern | Meaning |
|---|---|
| `?` | single character |
| `*` | any string of characters |
| `[aeiou]` | single char in set |
| `[!aeiou]` | single char NOT in set |
| `[a-z]` | range |
| `<word>` | word boundaries — matches whole word |
| `@` | one or more of preceding (e.g. `ro@m` → room, roooom) |
| `{n}` | exactly n |
| `{n,m}` | n to m |
| `\1`, `\2` | back-references |
| `()` | group |

In JS, **escape your backslashes**: `'\\('` for a literal `(`.

## What the AI can do — even the "hard" stuff

Word.js doesn't expose certain APIs, but the assistant works around them where it can:

| Capability | How |
|---|---|
| **Page-wide watermark** | Injected as OOXML into every section's primary header |
| **Multi-column section** | OOXML `<w:cols w:num="N"/>` in a continuous section break |
| **Page orientation switch** | OOXML `<w:pgSz orient="landscape"/>` section property |
| **Table of Contents** | Word's TOC field code via `insertOoxml` |
| **Citations & bibliography** | In-memory list + style-aware formatter (APA/MLA/Chicago/Harvard) |
| **LaTeX equations** | Tiny LaTeX → OMML converter for common patterns |
| **Mail merge** | Body copy + per-record placeholder replacement with page-break separators |
| **Content controls** (checkbox/dropdown/datePicker) | SDT OOXML for the special kinds Word.js doesn't expose |
| **Cell merging in tables** | Refuses cleanly — Word.js doesn't expose it. Recommends manual merge. |
| **Live web data** | One-time Google or Brave search with user permission |

## What the AI still can't do (and will say so)

- **Export to PDF / image / any export** — no JS API.
- **Print** — no JS API.
- **Chart creation** — no JS API. Workaround: insert a `charts-diagrams` image from the asset pack.
- **SmartArt** — no JS API.
- **AI image generation** — by design.

## Themes — AI-generated, unlimited

There is **no fixed catalog**. The user describes the look in natural language, and the LLM emits a strict JSON theme object:

```json
{
  "name":        "Healthcare Calm",
  "headingFont": "Inter",
  "bodyFont":    "Source Sans Pro",
  "primary":     "#0F4C81",
  "secondary":   "#3A6FA0",
  "accent":      "#5BA697",
  "bg":          "#FFFFFF",
  "text":        "#1F2937",
  "headingSize": 18,
  "tone":        "cool",
  "rationale":   "Soft blue + sage green evokes trust and gentle care, without the clinical chill of pure teal."
}
```

Server-side guards:
- **Font safety**: any font not on the Word-safe whitelist is swapped to the closest substitute. The user is told if a swap happens.
- **WCAG AA contrast**: text-on-background contrast is checked (>=4.5); the LLM gets one retry with the constraint, then text is auto-darkened if still failing.

The 8 quick-start chips ("Corporate clean", "Academic serif", etc.) are just pre-fills — every theme is freshly generated by the LLM on each click, so the same chip can yield different results.

## Project structure

```
word-ai-assistant/
├── server.js                       # Express backend (~1.1k lines)
├── manifest.xml                    # Office add-in manifest (Host=Document)
├── create-icons.js                 # Icon PNG generator
├── package.json
├── .env.example
├── public/
│   ├── index.html                  # Task pane shell
│   ├── app.js                      # All client-side logic, incl. execCode()
│   ├── helpers.js                  # Static data: templates, theme seeds, OOXML snippets, font/contrast guards
│   ├── commands.html
│   └── icon-{16,32,80}.png
├── assets/pack/
│   └── manifest.json               # Image pack index (per-image tags, license, sha256)
├── scripts/
│   └── build-pack.js               # Maintainer script for expanding the image pack
├── eval/
│   ├── run.js                      # Eval runner with LLM-as-judge
│   ├── cases.json                  # ~52 Word test cases across 28 categories
│   ├── generated-cases.json        # Auto-generated harder/easier cases
│   ├── progress.json               # Per-category mastery tracking
│   ├── improvements.txt            # Auto-patched system prompt additions
│   ├── watchdog.js                 # Fallback runner if CI misses its window
│   ├── RESULTS.md                  # Latest eval scoreboard
│   └── results/                    # JSON history of every run
└── .github/workflows/
    └── eval.yml                    # Runs every 8 hours, commits improvements
```

## How it works

1. **You type a request.** The task pane snapshots the current document (paragraph outline with styles, selection text, word count, Flesch grade — up to 40 paragraphs) and sends it plus your message to `/api/chat`.
2. **The AI plans and emits code.** Backed by a ~600-line system prompt that knows the Word Office.js surface, every available helper, and the forbidden APIs. The response includes a `CODE_JS::...::END_CODE` block.
3. **The task pane executes the code.** Wraps it in `Word.run(async ctx => { /* helpers in scope */ /* AI code */ })`. Helpers like `addHeading`, `insertTable`, `insertWatermark`, `mailMergeReplace`, `applyTheme(designTheme(desc))` do the messy work.
4. **Auto-retry on failure.** If the code throws, the error is fed back to the AI which rewrites and re-runs.
5. **Auto-improvement loop.** Every 8 hours, GitHub Actions runs `npm run eval` against ~52 test cases. Failing cases (<70/100) trigger an LLM-generated patch to `eval/improvements.txt` which is appended to the system prompt on next start. Mastered categories (≥95) auto-generate harder cases. Stuck categories auto-generate easier ones.

## Web-search permission model

| Mode (Settings → Data → Web Search) | AI behavior |
|---|---|
| **Off** (default) | AI emits `NEEDS_WEB_SEARCH::{...}::END`. Task pane shows a permission card with the suggested query and an "Allow this search" button. After approval, the AI gets the results and continues. |
| **On** | Searches happen automatically. A small pill appears on the message showing the query and result count; click to expand and see the sources. |
| **Disabled** | AI never searches. If you ask something time-sensitive, it tells you to enable web search in Settings. |

## Image asset pack

The pack lives at `assets/pack/`. The manifest (`assets/pack/manifest.json`) is checked into git; the actual image files are downloaded on demand via Settings → Data → "Download asset pack". All images are CC0 from Pexels, Unsplash, and Pixabay — free for personal and commercial use.

To build a fuller pack, set API keys and run:
```
PEXELS_KEY=xxx UNSPLASH_KEY=yyy PIXABAY_KEY=zzz npm run build-pack
```
This downloads up to ~50 images per category × 18 categories ≈ 900 images, computes SHA-256 hashes, and rewrites the manifest. Then commit the updated manifest and bundle the images as a GitHub Release tarball.

## FAQ

**How does theme generation work?** You type any natural-language description. The server sends it to the LLM with a tightly-scoped system prompt that demands strict JSON output (name, headingFont, bodyFont, primary/secondary/accent/bg/text colors, headingSize, tone, rationale). The server validates fonts against a 60-name Word-safe whitelist (swapping unknown fonts for the closest match) and color contrast against WCAG AA (auto-darkening text if needed). The cleaned theme is returned and saved in `data/themes.json` for the "Recent" strip.

**Can I save a favorite theme?** Yes — every applied theme is auto-saved to `data/themes.json` keyed by hash of its description. The Theme Designer shows up to 6 recent themes you can re-apply with a click.

**Why doesn't the TOC populate?** Word's TOC field needs to be updated. Right-click the field and choose "Update Field", or press F9. The helper inserts the field reference with placeholder text by default.

**Cell merging in tables doesn't work?** Word Office.js doesn't expose cell merging on most builds. The helper refuses cleanly and recommends merging manually after the table is applied, or restructuring as a multi-row header.

**Where is my chat history stored?** Locally in `data/chats.json` on the machine running the server (not in any cloud).

**Cert trust issues in Word desktop?** Run `npx office-addin-dev-certs install` once. On macOS / Windows it'll add the dev cert to your trust store.

**How do I switch the AI model?** Click the model name in the header. The picker has filter/sort, recent-models, hardware-based recommendations, and free-tier highlighting.

**Word wildcards vs JS regex?** `body.search()` only accepts Word wildcards (e.g. `<*@*.*>` for emails). For JS regex you can pass `{ regex: /.../ }` to the `replaceText` helper, which post-processes the body text on the client.

## Configuration

In `server.js`:
- `DEFAULT_MODEL` — defaults to `meta-llama/llama-3.3-70b-instruct`
- `USE_OPENROUTER` / `USE_GROQ` / `USE_MLX` flags
- `MACBOOK_IP` — for MLX/Ollama on a separate Mac
- All API keys are read from `.env` (never logged)

## License

MIT.

## Acknowledgments

Sister projects: `excel-ai-assistant` and `powerpoint-ai-assistant`. Same architecture, same eval loop, same multi-provider model picker.
