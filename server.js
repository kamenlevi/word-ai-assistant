require('dotenv').config();
const express = require('express');
const https   = require('https');
const path    = require('path');
const fs      = require('fs');
const crypto  = require('crypto');
const selfsigned = require('selfsigned');
const fetch   = require('node-fetch');
const { v4: uuidv4 } = require('uuid');

const app = express();
app.use(express.json({ limit: '20mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ── SSL cert — use office-addin-dev-certs (trusted by Word desktop) ──
const DEV_CERTS_DIR = path.join(process.env.USERPROFILE || process.env.HOME, '.office-addin-dev-certs');
let pems;
try {
  pems = {
    cert:    fs.readFileSync(path.join(DEV_CERTS_DIR, 'localhost.crt'), 'utf8'),
    private: fs.readFileSync(path.join(DEV_CERTS_DIR, 'localhost.key'), 'utf8'),
  };
} catch {
  // Fallback: generate a self-signed cert in memory
  const attrs = [{ name: 'commonName', value: 'localhost' }];
  const gen = selfsigned.generate(attrs, { days: 365 });
  pems = { cert: gen.cert, private: gen.private };
  console.warn('[server] office-addin-dev-certs not found — using ephemeral self-signed cert.');
}

// ── Data dir for chat persistence ─────────────────────────────────────────────
const DATA_DIR    = path.join(__dirname, 'data');
const CHATS_FILE  = path.join(DATA_DIR, 'chats.json');
const THEMES_FILE = path.join(DATA_DIR, 'themes.json');
const WEB_LOG     = path.join(DATA_DIR, 'web-search-log.json');
fs.mkdirSync(DATA_DIR, { recursive: true });
function readChats() {
  try { return JSON.parse(fs.readFileSync(CHATS_FILE, 'utf8')); } catch { return {}; }
}
function writeChats(chats) { fs.writeFileSync(CHATS_FILE, JSON.stringify(chats, null, 2), 'utf8'); }

function readThemes() {
  try { return JSON.parse(fs.readFileSync(THEMES_FILE, 'utf8')); } catch { return { recent: [] }; }
}
function writeThemes(t) { fs.writeFileSync(THEMES_FILE, JSON.stringify(t, null, 2), 'utf8'); }

// ── Asset pack dir ────────────────────────────────────────────────────────────
const ASSETS_DIR  = path.join(__dirname, 'assets');
const PACK_DIR    = path.join(ASSETS_DIR, 'pack');
const MANIFEST_FILE = path.join(PACK_DIR, 'manifest.json');
fs.mkdirSync(PACK_DIR, { recursive: true });

// ── CONFIG ────────────────────────────────────────────────────────────────────
const MACBOOK_IP     = '192.168.1.206';
const MLX_PORT       = '8080';
const OLLAMA_PORT    = '11434';

const USE_OPENROUTER = true;
const USE_MLX        = false;
const USE_GROQ       = false;

const OPENROUTER_KEY = process.env.OPENROUTER_KEY;
const GROQ_KEY       = process.env.GROQ_KEY;
const BRAVE_KEY      = process.env.BRAVE_KEY;
const GOOGLE_KEY     = process.env.GOOGLE_KEY;
const GOOGLE_CX      = process.env.GOOGLE_CX;
const PROVIDER_PREF  = (process.env.WEB_SEARCH_PROVIDER || 'google').toLowerCase();

function resolvedProvider() {
  if (PROVIDER_PREF === 'google' && GOOGLE_KEY && GOOGLE_CX) return 'google';
  if (PROVIDER_PREF === 'brave'  && BRAVE_KEY)               return 'brave';
  if (GOOGLE_KEY && GOOGLE_CX)                                return 'google';
  if (BRAVE_KEY)                                              return 'brave';
  return null;
}

const DEFAULT_MODEL = USE_OPENROUTER
  ? 'meta-llama/llama-3.3-70b-instruct'
  : USE_MLX
    ? 'mlx-community/Qwen3-32B-4bit'
    : 'qwen3:32b';

// ── Word-safe font whitelist (server-side mirror of helpers.js) ───────────────
const WORD_SAFE_FONTS = [
  'Calibri','Calibri Light','Cambria','Cambria Math','Georgia','Times New Roman',
  'Garamond','Book Antiqua','Arial','Arial Black','Arial Narrow','Helvetica',
  'Verdana','Tahoma','Trebuchet MS','Lucida Sans Unicode','Courier New','Consolas',
  'Lucida Console','Comic Sans MS','Impact','Palatino Linotype','Century Gothic',
  'Inter','Roboto','Roboto Slab','Open Sans','Source Sans Pro','Source Serif Pro',
  'Source Code Pro','Lora','Merriweather','Playfair Display','Montserrat','Poppins',
  'Quicksand','Raleway','Oswald','Bebas Neue','Lato','Nunito','Nunito Sans','PT Sans',
  'PT Serif','Fira Sans','Fira Code','Work Sans','IBM Plex Sans','IBM Plex Serif',
  'IBM Plex Mono','Crimson Text','EB Garamond','Cormorant Garamond','Libre Baskerville',
  'Bitter','Karla','Manrope','Mulish','Rubik','Ubuntu','Futura','Avenir','Avenir Next',
  'Helvetica Neue','Optima','Baskerville','Didot','Hoefler Text',
];
const FONT_FALLBACKS_SRV = {
  'helvetica neue condensed':'Helvetica Neue','helvetica condensed':'Helvetica',
  'avenir next condensed':'Avenir Next','gotham':'Montserrat','gotham bold':'Montserrat',
  'proxima nova':'Montserrat','museo sans':'Open Sans','din':'Oswald','din alternate':'Oswald',
  'helvetica neue thin':'Helvetica Neue','sf pro display':'Inter','sf pro text':'Inter',
  'system':'Calibri','system-ui':'Calibri','sans-serif':'Calibri','serif':'Cambria',
  'monospace':'Consolas','cursive':'Comic Sans MS','fantasy':'Impact','comic sans':'Comic Sans MS',
  'comic sans replacement':'Comic Sans MS','arial unicode ms':'Arial','times':'Times New Roman',
  'courier':'Courier New','roboto condensed':'Roboto','open sans condensed':'Open Sans',
  'pt sans narrow':'PT Sans','futura bold':'Futura',
};

function sanitizeFontSrv(name) {
  if (!name) return { font: 'Calibri', changed: false };
  const trimmed = String(name).trim();
  if (!trimmed) return { font: 'Calibri', changed: false };
  for (const f of WORD_SAFE_FONTS) {
    if (f.toLowerCase() === trimmed.toLowerCase()) return { font: f, changed: false };
  }
  const low = trimmed.toLowerCase();
  if (FONT_FALLBACKS_SRV[low]) return { font: FONT_FALLBACKS_SRV[low], changed: true, original: trimmed };
  for (const f of WORD_SAFE_FONTS) {
    if (low.startsWith(f.toLowerCase() + ' ')) return { font: f, changed: true, original: trimmed };
  }
  if (/mono|code|courier/.test(low))   return { font: 'Consolas',         changed: true, original: trimmed };
  if (/serif|antiqua|roman/.test(low)) return { font: 'Cambria',           changed: true, original: trimmed };
  if (/script|cursive|hand/.test(low)) return { font: 'Comic Sans MS',     changed: true, original: trimmed };
  return { font: 'Calibri', changed: true, original: trimmed };
}

// WCAG contrast guard
function hexToRgb(hex) {
  const m = String(hex || '').replace('#', '').match(/^([\da-f]{6})$/i);
  if (!m) return null;
  const n = parseInt(m[1], 16);
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255 };
}
function luminance({ r, g, b }) {
  const a = [r, g, b].map(v => { v /= 255; return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4); });
  return 0.2126 * a[0] + 0.7152 * a[1] + 0.0722 * a[2];
}
function contrast(h1, h2) {
  const c1 = hexToRgb(h1), c2 = hexToRgb(h2);
  if (!c1 || !c2) return 1;
  const L1 = luminance(c1), L2 = luminance(c2);
  return (Math.max(L1, L2) + 0.05) / (Math.min(L1, L2) + 0.05);
}
function darken(hex, amount) {
  const c = hexToRgb(hex);
  if (!c) return hex;
  const a = Math.max(0, Math.min(1, amount));
  const r = Math.max(0, Math.round(c.r * (1 - a)));
  const g = Math.max(0, Math.round(c.g * (1 - a)));
  const b = Math.max(0, Math.round(c.b * (1 - a)));
  return '#' + [r, g, b].map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase();
}

// ─────────────────────────────────────────────────────────────────────────────
// SYSTEM PROMPT
// ─────────────────────────────────────────────────────────────────────────────
const SYSTEM_PROMPT = `You are an AI assistant built directly into Microsoft Word. You have complete read/write access to the entire document.

CRITICAL RULE — HOW YOU MAKE CHANGES:
When the user asks you to do ANYTHING to the document, you MUST output a CODE_JS block containing valid Office JavaScript API code. This code runs inside Word.run(async (context) => { ... }). You have access to "context", "doc" (= context.document), the full "Word" namespace, and these injected helpers: addHeading, addParagraph, addTitle, addSubtitle, addQuote, addList, applyStyle, replaceText, findText, insertPageBreak, insertSectionBreak, insertColumns, setHeader, setFooter, addPageNumbers, insertTableOfContents, setMargins, setPageOrientation, insertTable, setTableCell, styleTable, mergeCellsBestEffort, insertImage, insertWatermark, insertFootnote, insertEndnote, insertComment, toggleTrackChanges, acceptAllRevisions, rejectAllRevisions, insertCitation, insertBibliography, insertEquation, insertContentControl, insertFormField, mailMergeReplace, applyTheme, recolorDocument, tweakTheme, countWords, countCharacters, countParagraphs, getReadability, getSelectionText, getSelectionHtml, listSections, listHeadings, getDocumentSummary, applyTemplate, designTheme, TEMPLATES, OOXML.

FORMAT:
CODE_JS::
// your Office JS code here
::END_CODE

RULES FOR YOUR CODE:
- You write the BODY of an async function. \`context\` and \`doc\` and \`Word\` are in scope, along with every helper listed above.
- Always call "await context.sync();" after .load() and after writes.
- NEVER use console.log. Throw errors with throw new Error("...").
- Use Word.js 1.1–1.4 APIs only (Word on the web compatible).

WHAT YOU CAN DO (use helpers — they wrap the messy parts):
- Insert headings, paragraphs, titles, subtitles, quotes, bullet lists, numbered lists.
- Apply built-in or custom paragraph styles.
- Find & replace with optional regex/wildcards.
- Insert page/section breaks, switch to multi-column layout, change orientation.
- Set headers, footers, page numbers, margins.
- Insert tables and style them.
- Insert images from the local CC0 asset pack (NEVER hallucinate images — use insertImage).
- Insert footnotes, endnotes, comments.
- Toggle/accept/reject tracked changes.
- Insert citations and a final bibliography (APA / MLA / Chicago / Harvard).
- Insert LaTeX equations (auto-converted to OMML).
- Insert content controls (rich text, plain text, checkbox, dropdown, combo box, date picker).
- Insert a Table of Contents (uses Word's TOC field — user may need to right-click → Update Field).
- Generate a theme on-demand from any natural-language description and apply it.
- Apply one of 14 document templates with field substitution.
- Count words / paragraphs / pages, calculate Flesch reading-ease grade.

NEVER USE THESE — they DO NOT EXIST in Word's Office.js:
- doc.exportToPdf(), doc.export*() (no export API — refuse cleanly)
- doc.print() (no print API)
- doc.startSlideShow() / startMailMerge() etc. (no slideshow / interactive merge — emit text records yourself)
- "JavaScript regex" inside body.search(). Word search uses WORD WILDCARDS, not JS regex. See cheatsheet below.
- Direct chart APIs on a document (no Word chart API — refuse or insert a charts-diagrams image)

WORD WILDCARD CHEAT-SHEET (used with body.search(query, { matchWildcards:true })):
- ?   single character
- *   any string of characters
- [aeiou]      single char in set
- [!aeiou]     single char NOT in set
- [a-z]        range
- <word>       word boundaries — matches whole word
- @            one or more of preceding (e.g. ro@m → room, roooom)
- {n}          exactly n
- {n,m}        n to m
- \\1, \\2 …    back-references to grouped (...) parts
- ()           group
ALWAYS escape backslashes in JS string literals: '\\\\(' for a literal '('.

HOW TO ADD CONTENT — use the helpers:

  await addHeading(text, level, opts?)            // level 1–9
    Example: await addHeading("Executive Summary", 1);

  await addParagraph(text, opts?)
    opts: { bold, italic, size, color, alignment ("left"|"center"|"right"|"justify"),
            spaceBefore, spaceAfter, lineSpacing, indent }

  await addTitle(text, opts?) / addSubtitle(text, opts?)
  await addQuote(text, opts?)

  await addList(items, type, opts?)               // type: "bullet" | "number"
    items: array of strings. opts.level: 0-8 for nesting.

  await applyStyle(target, styleName)             // target: "selection" | "paragraph" index | "all-headings" etc.

  await replaceText(find, replace, opts?)
    opts: { matchCase, matchWholeWord, matchWildcards, regex (JS regex, post-processed) }

  await findText(query, opts?)                    // returns array of {paragraph, before, match, after}

  await insertPageBreak() / insertSectionBreak(kind?)  // kind: "nextPage"|"continuous"|"evenPage"|"oddPage"

  await insertColumns(count)                      // applies to the following content via a continuous section break

  await setHeader(text, opts?) / setFooter(text, opts?)
    opts: { primary, firstPage, evenPages, alignment, includePageNumber, includeTotalPages, font, size }

  await addPageNumbers(opts?)                     // opts: { location: "header"|"footer", alignment, format: "1"|"i"|"I"|"a"|"A" }

  await insertTableOfContents(opts?)              // opts: { minLevel, maxLevel, title }

  await setMargins({ top, bottom, left, right })  // in points
  await setPageOrientation("portrait" | "landscape")

  await insertTable(rows, cols, values?, opts?)
    values: 2D array of strings.
    opts: { headerRow: true, style: "Grid Table 4 - Accent 1", autofit: true }

  await insertImage(category, keywordOrTags, opts?)
    Picks an image from the local CC0 asset pack.
    category ∈ injected ASSET_PACK_READY categories.
    keywordOrTags: string OR array of tag strings.
    opts: { width, height, alignment: "inline"|"center"|"left"|"right" }
    Example: await insertImage("office-workplace", ["desk","minimal"], { width: 320 });

  await insertWatermark(text, opts?)              // text watermark on every page

  await insertFootnote(text) / insertEndnote(text)
  await insertComment(targetQuery, text)          // finds first occurrence of targetQuery, attaches comment
  await toggleTrackChanges(on)                    // true / false
  await acceptAllRevisions() / rejectAllRevisions()

  await insertCitation({ author, year, title, source }) → returns "(Author, Year)" inline marker
  await insertBibliography(style)                 // "APA" | "MLA" | "Chicago" | "Harvard"
  await insertEquation(latex)                     // e.g. "E = mc^2" or "\\\\frac{a}{b}"

  await insertContentControl(kind, opts?)
    kind ∈ "richText" | "plainText" | "checkbox" | "dropdown" | "comboBox" | "datePicker"
    opts: { title, tag, placeholder, items[], checked, format }

  await insertFormField(kind, opts?)              // legacy form field; mostly delegates to content controls
  await mailMergeReplace(records)                 // records: array of {Field: value} objects → duplicates the doc per record

THEMES — AI-GENERATED, UNLIMITED:
There are NO preset theme names. Every theme is generated on demand from a natural-language description by calling the designTheme helper, which returns a palette+typography object you pass to applyTheme(...).

  const t = await designTheme("calm trustworthy theme for a healthcare startup");
  await applyTheme(t);

Or, in a single line if you want to chain:
  await applyTheme(await designTheme("brutalist black-and-yellow legal contract"));

Tweaking an applied theme:
  await applyTheme(await tweakTheme(currentTheme, "same but darker, with a green accent"));

recolorDocument(palette) applies an explicit palette object you've already built:
  await recolorDocument({ primary:"#0F4C81", secondary:"#3A6FA0", accent:"#D4A017", bg:"#FFFFFF", text:"#1A1A1A", headingFont:"Inter", bodyFont:"Inter" });

TEMPLATES — Fourteen ready-to-use document templates with field substitution:
  resume · cover-letter · business-letter · formal-report · proposal · contract ·
  essay · academic-paper · blog-post · meeting-notes · project-plan · invoice · memo · whitepaper

  await applyTemplate("resume", { Name: "Alice", Title: "Engineer", Email: "a@x.com", ... });

When the user provides field values inline ("Name=Alice Title=Engineer"), parse them and pass them.

EXAMPLES:

User: "add a Heading 1 saying Quarterly Review"
Response:
Adding a Heading 1 'Quarterly Review'.
CODE_JS::
await addHeading("Quarterly Review", 1);
::END_CODE

User: "make every instance of WARNING bold and red"
Response:
Making every instance of "WARNING" bold and red.
CODE_JS::
const results = doc.body.search("WARNING", { matchCase: true });
results.load("font");
await context.sync();
results.items.forEach(r => { r.font.bold = true; r.font.color = "#D32F2F"; });
await context.sync();
::END_CODE

User: "replace all email addresses with [redacted]"
Response:
Wildcard-replacing email addresses with [redacted].
CODE_JS::
await replaceText("<*@*.*>", "[redacted]", { matchWildcards: true });
::END_CODE

User: "insert a 4x3 table with headers Quarter, Revenue, Notes"
Response:
Inserting a 4×3 table.
CODE_JS::
await insertTable(4, 3, [
  ["Quarter", "Revenue", "Notes"],
  ["Q1", "", ""],
  ["Q2", "", ""],
  ["Q3", "", ""],
], { headerRow: true, style: "Grid Table 4 - Accent 1" });
::END_CODE

User: "add a footnote 'See appendix A' to this paragraph"
Response:
Adding a footnote to the current selection.
CODE_JS::
await insertFootnote("See appendix A.");
::END_CODE

User: "insert a Table of Contents"
Response:
Inserting a TOC. After Word reloads, right-click and choose "Update Field" if it doesn't auto-populate.
CODE_JS::
await insertTableOfContents({ minLevel: 1, maxLevel: 3, title: "Table of Contents" });
::END_CODE

User: "turn on track changes"
Response:
Turning track changes on.
CODE_JS::
await toggleTrackChanges(true);
::END_CODE

User: "switch the next section to landscape"
Response:
Switching to landscape via a continuous section break.
CODE_JS::
await insertSectionBreak("continuous");
await setPageOrientation("landscape");
::END_CODE

User: "design a calm trustworthy theme for a healthcare startup and apply it"
Response:
Designing a calm trustworthy healthcare theme and applying it.
CODE_JS::
const t = await designTheme("calm trustworthy theme for a healthcare startup — soft blues and greens, generous spacing, friendly modern sans-serif headings");
await applyTheme(t);
::END_CODE

User: "now make it darker with a green accent"
Response:
Tweaking the active theme.
CODE_JS::
const t = await tweakTheme(activeTheme, "make it darker overall and swap the accent for a saturated forest green");
await applyTheme(t);
::END_CODE

User: "make it look like a 1970s sci-fi novel"
Response:
Generating a 70s sci-fi novel theme.
CODE_JS::
const t = await designTheme("1970s science fiction paperback novel — orange, brown, off-white, retro condensed sans-serif headings, classic serif body");
await applyTheme(t);
::END_CODE

User: "apply the resume template with Name=Alice Smith, Title=Software Engineer, Email=alice@example.com"
Response:
Applying the resume template.
CODE_JS::
await applyTemplate("resume", {
  Name: "Alice Smith",
  Title: "Software Engineer",
  Email: "alice@example.com",
});
::END_CODE

User: "insert a date-picker content control labeled Sign Date"
Response:
Inserting a date-picker content control.
CODE_JS::
await insertContentControl("datePicker", { title: "Sign Date", tag: "signDate", placeholder: "Pick a date…" });
::END_CODE

User: "mail-merge Name,Amount\\nAlice,100\\nBob,200"
Response:
Running the mail merge.
CODE_JS::
await mailMergeReplace([
  { Name: "Alice", Amount: "100" },
  { Name: "Bob",   Amount: "200" },
]);
::END_CODE

User: "insert an equation: E = mc^2"
Response:
Inserting the equation.
CODE_JS::
await insertEquation("E = mc^2");
::END_CODE

User: "add an office-workplace image"
Response:
Inserting an office-workplace photo.
CODE_JS::
await insertImage("office-workplace", ["desk","minimal"], { width: 360 });
::END_CODE

User: "add CONFIDENTIAL watermark to every page"
Response:
Adding a CONFIDENTIAL watermark.
CODE_JS::
await insertWatermark("CONFIDENTIAL");
::END_CODE

READING DATA — only "return" when the user wants to KNOW something:

User: "how many words does this have?"
Response:
Counting words.
CODE_JS::
const n = await countWords();
return "Your document has " + n + " words.";
::END_CODE

User: "what's my Flesch reading-ease score?"
Response:
Calculating Flesch reading-ease.
CODE_JS::
const r = await getReadability();
return "Flesch: " + r.flesch + " (grade " + r.grade + "), " + r.wordCount + " words across " + r.sentenceCount + " sentences.";
::END_CODE

IMAGE INSERTION RULES:
- NEVER generate, hallucinate, or describe AI-generated images.
- NEVER call doc.body.insertInlinePictureFromBase64 with a remote URL — Word.js needs base64 anyway. The ONLY sanctioned image path is insertImage(...).
- The available pack categories are injected at runtime as <ASSET_PACK_READY categories="..."/>.
- If <ASSET_PACK_MISSING/> is present, do NOT call insertImage. Instead insert a paragraph reading "[image: <description>]" and tell the user to download the asset pack from Settings.

WEB ACCESS RULES:
You do NOT have direct internet access. You cannot fetch live prices, weather, news, sports scores, or anything that changes after your training cutoff.

The current web-search mode will be injected as <WEB_SEARCH_MODE value="off|on|disabled"/>:

- mode=off (DEFAULT — ASK each time):
    If a request needs live real-world data, emit ONE block on its own lines exactly:
        NEEDS_WEB_SEARCH::{"query": "<best search query>", "reason": "<one short sentence>"}::END
    Then STOP. Do NOT also emit CODE_JS in the same turn. Wait for the host to provide results.

- mode=on (ALWAYS-ON):
    The host will run searches automatically. Proceed as if you have web access. If you need data, emit the same NEEDS_WEB_SEARCH block — the host will silently intercept and re-prompt with results.

- mode=disabled (NEVER):
    Do NOT emit NEEDS_WEB_SEARCH. Tell the user clearly: "I can't get this without web access, and you have web search disabled in Settings. Enable it in Settings → Web search if you'd like me to look this up."

NEVER make up current values, prices, dates, or news under any circumstance.
When you DO receive <WEB_SEARCH_RESULTS query="..."> in the conversation, use those facts in your reply and continue normally (with CODE_JS if needed).

CITATION FORMAT RULES:
- APA:     (Author, Year)         — e.g. (Smith, 2021)
- MLA:     (Author Page)          — e.g. (Smith 42)
- Chicago: (Author Year, Page)    — e.g. (Smith 2021, 42)
- Harvard: (Author Year: Page)    — e.g. (Smith 2021: 42)
The insertCitation/insertBibliography helpers take a style name and format the inline marker and bibliography entry accordingly.

TRACK CHANGES ETIQUETTE:
- When the user says "review", "suggest edits", "find typos" → toggle track changes ON, make changes, leave them as revisions.
- When the user says "apply", "accept", "commit" → make the changes directly (no track changes).
- Never mix: if track changes is already on, NEVER call acceptAllRevisions unless explicitly asked.

VBA FALLBACK:
For things Word.js cannot do (mass mail-merge with a real data source, complex print configuration), write complete VBA in a triple-backtick \`\`\`vba block and tell the user to press Alt+F11 to open the VBA editor.

OTHER RULES:
- Be concise. One short sentence explaining what you're doing, then the CODE_JS block.
- For CHART CREATION, PDF EXPORT: refuse cleanly — "Word Office.js doesn't expose <X>. As a workaround you can <Y>." For charts, suggest insertImage with category "charts-diagrams".
- Only skip CODE_JS if the user is purely asking a question with no changes needed.`
+ (DEFAULT_MODEL.toLowerCase().includes('qwen') ? '\n/no_think' : '')
+ (() => {
    try {
      const imp = fs.readFileSync(path.join(__dirname, 'eval/improvements.txt'), 'utf8').trim();
      return imp ? '\n\nADDITIONAL RULES FROM EVAL:\n' + imp : '';
    } catch { return ''; }
  })();

// ── Call the AI ───────────────────────────────────────────────────────────────
async function callAI(messages, maxTokens = 4096, model = null, useOllama = false, useGroq = false, apiKey = null, groqKey = null) {
  const effectiveModel = model || DEFAULT_MODEL;
  const orKey  = apiKey  || OPENROUTER_KEY;
  const gKey   = groqKey || GROQ_KEY;

  if (useOllama) {
    const res = await fetch(`http://localhost:${OLLAMA_PORT}/api/chat`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        model: effectiveModel, messages, stream: false,
        options: { num_ctx: 4096, temperature: 0.15, top_p: 0.9, repeat_penalty: 1.1 }
      })
    });
    const data = await res.json();
    return { text: data.message.content, usage: null };
  }

  if (useGroq || USE_GROQ) {
    const groqModel = (useGroq && model) ? model : 'llama-3.3-70b-versatile';
    const res = await fetch('https://api.groq.com/openai/v1/chat/completions', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${gKey}` },
      body: JSON.stringify({
        model: groqModel, messages, max_tokens: maxTokens, temperature: 0.15, top_p: 0.9
      })
    });
    const data = await res.json();
    if (data.error) throw new Error(data.error.message);
    return { text: data.choices[0].message.content, usage: data.usage || null };
  }

  if (USE_OPENROUTER || orKey) {
    const res = await fetch('https://openrouter.ai/api/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${orKey}`,
        'HTTP-Referer': 'https://localhost:3000',
        'X-Title': 'Word AI Assistant'
      },
      body: JSON.stringify({
        model: effectiveModel, messages, max_tokens: maxTokens, temperature: 0.15, top_p: 0.9
      })
    });
    const data = await res.json();
    if (data.error) throw new Error(JSON.stringify(data.error));
    return { text: data.choices[0].message.content, usage: data.usage || null };
  }

  if (USE_MLX) {
    const res = await fetch(`http://${MACBOOK_IP}:${MLX_PORT}/v1/chat/completions`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        model: effectiveModel, messages, max_tokens: maxTokens, temperature: 0.15, top_p: 0.9
      })
    });
    const data = await res.json();
    return { text: data.choices[0].message.content, usage: data.usage || null };
  }

  // Fallback: Ollama on MacBook
  const res = await fetch(`http://${MACBOOK_IP}:${OLLAMA_PORT}/api/chat`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      model: 'qwen3:32b', messages, stream: false,
      options: { num_ctx: 4096, temperature: 0.15, top_p: 0.9, repeat_penalty: 1.1 }
    })
  });
  const data = await res.json();
  return { text: data.message.content, usage: null };
}

// ── Parse code from response ─────────────────────────────────────────────────
function parseResponse(text) {
  let code = null;
  const codeMatch = text.match(/CODE_JS::([\s\S]*?)::END_CODE/);
  if (codeMatch) code = codeMatch[1].trim();
  const cleaned = text
    .replace(/CODE_JS::[\s\S]*?::END_CODE/g, '')
    .replace(/<think>[\s\S]*?<\/think>/g, '')
    .trim();
  return { code, cleaned };
}

function parseWebSearchRequest(text) {
  const m = text.match(/NEEDS_WEB_SEARCH::([\s\S]*?)::END/);
  if (!m) return null;
  try { return JSON.parse(m[1].trim()); } catch { return null; }
}

// ── Detect if model forgot to include code for an action request ─────────────
function isQuestion(userMessage) {
  const s = userMessage.toLowerCase().trim();
  const dataQueryWords = [
    'how many words', 'how many paragraphs', 'how many pages', 'how many sections',
    'how many headings', 'word count', 'page count', 'count', 'list all', 'find all',
    'which heading', 'which section', 'what does', 'what is in', 'what color',
    'what font', 'reading level', 'reading-ease', 'flesch',
  ];
  if (dataQueryWords.some(q => s.includes(q))) return false;
  const questionStarters = [
    'why ', 'explain', 'tell me', 'describe', 'can you explain', 'can you tell',
    'what does', 'what is a ', 'what are the', 'how do i', 'how does', 'is there ',
    'are there ', 'does word', 'what is a content control', 'what is a style',
    'what is mail merge', 'what is a section break',
  ];
  return questionStarters.some(q => s.startsWith(q));
}

function modelForgotCode(responseText, userMessage) {
  if (responseText.includes('CODE_JS::')) return false;
  if (responseText.includes('NEEDS_WEB_SEARCH::')) return false;
  if (/doesn't expose|does not expose|not supported|no JS API/i.test(responseText)) return false;
  return !isQuestion(userMessage);
}

// ── Chat history endpoints ────────────────────────────────────────────────────
app.get('/api/chats', (req, res) => {
  const chats = readChats();
  const list = Object.values(chats)
    .sort((a, b) => new Date(b.updatedAt) - new Date(a.updatedAt))
    .map(({ id, title, updatedAt, messageCount, totalTokens }) => ({ id, title, updatedAt, messageCount, totalTokens }));
  res.json(list);
});

app.post('/api/chats', (req, res) => {
  const chats = readChats();
  const id = uuidv4();
  const now = new Date().toISOString();
  chats[id] = { id, title: req.body.title || 'New Chat', messages: req.body.messages || [], summary: null, totalTokens: 0, messageCount: 0, createdAt: now, updatedAt: now };
  writeChats(chats);
  res.json(chats[id]);
});

app.get('/api/chats/:id', (req, res) => {
  const chats = readChats();
  const chat = chats[req.params.id];
  if (!chat) return res.status(404).json({ error: 'Not found' });
  res.json(chat);
});

app.patch('/api/chats/:id', (req, res) => {
  const chats = readChats();
  const chat = chats[req.params.id];
  if (!chat) return res.status(404).json({ error: 'Not found' });
  const { title, messages, summary, totalTokens } = req.body;
  if (title !== undefined) chat.title = title;
  if (messages !== undefined) { chat.messages = messages; chat.messageCount = messages.length; }
  if (summary !== undefined) chat.summary = summary;
  if (totalTokens !== undefined) chat.totalTokens = totalTokens;
  chat.updatedAt = new Date().toISOString();
  writeChats(chats);
  res.json(chat);
});

app.delete('/api/chats/:id', (req, res) => {
  const chats = readChats();
  delete chats[req.params.id];
  writeChats(chats);
  res.json({ ok: true });
});

// ── OpenRouter model catalog (cached 1 h) ─────────────────────────────────────
let catalogCache = null, catalogCacheAt = 0;
app.get('/api/models/catalog', async (req, res) => {
  if (catalogCache && Date.now() - catalogCacheAt < 3600_000) return res.json(catalogCache);
  try {
    const r = await fetch('https://openrouter.ai/api/v1/models', {
      headers: OPENROUTER_KEY ? { 'Authorization': `Bearer ${OPENROUTER_KEY}` } : {}
    });
    const data = await r.json();
    catalogCache = (data.data || []).map(m => ({
      id:      m.id,
      name:    m.name || m.id,
      in:      m.pricing?.prompt     ? Math.round(parseFloat(m.pricing.prompt)     * 1e6 * 100) / 100 : 0,
      out:     m.pricing?.completion ? Math.round(parseFloat(m.pricing.completion) * 1e6 * 100) / 100 : 0,
      context: m.context_length || null,
    })).sort((a, b) => a.name.localeCompare(b.name));
    catalogCacheAt = Date.now();
    res.json(catalogCache);
  } catch (err) {
    res.status(500).json({ error: 'Catalog fetch failed' });
  }
});

// ── Local model detection (Ollama) ────────────────────────────────────────────
app.get('/api/models/local', async (req, res) => {
  try {
    const r = await fetch(`http://localhost:${OLLAMA_PORT}/api/tags`, { timeout: 2000 });
    const data = await r.json();
    const models = (data.models || []).map(m => ({ id: m.name, name: m.name }));
    res.json(models);
  } catch {
    res.json([]);
  }
});

// ── Model recommendation ─────────────────────────────────────────────────────
app.post('/api/recommend-model', async (req, res) => {
  const { specs, modelUsage, preferences, availableModels, apiKey } = req.body;
  const orKey = apiKey || OPENROUTER_KEY;
  if (!orKey) return res.status(400).json({ error: 'No OpenRouter API key configured.' });

  const modelList = (availableModels || []).map(m =>
    `${m.id} | ${m.name} | $${m.in}/$${m.out} per 1M tok${(m.tags||[]).includes('free') ? ' [FREE]' : ''}`
  ).join('\n');
  const usageStr = Object.entries(modelUsage || {}).sort((a,b)=>b[1]-a[1]).slice(0,5)
    .map(([id,n])=>`${id}: used ${n}x`).join('\n') || 'No history yet';

  try {
    const { text } = await callAI([
      { role: 'system', content: 'You are an expert in AI model selection. Respond ONLY with valid JSON: {"modelId":"exact-id","reason":"2-3 sentence explanation"}. No markdown. No extra text.' },
      { role: 'user', content: `A user runs a Word AI Assistant. Recommend the single best model.

PC SPECS:
- CPU cores: ${specs?.cores || 'unknown'}
- RAM: ${specs?.memory || 'unknown'} GB
- GPU: ${specs?.gpu || 'unknown'}
- GPU Vendor: ${specs?.gpuVendor || 'unknown'}
- Platform: ${specs?.platform || 'unknown'}

USAGE HISTORY (most used first):
${usageStr}

USER PREFERENCES: ${preferences || 'None stated'}

AVAILABLE MODELS (id | name | price in/out per 1M):
${modelList}

Consider: code generation quality for Word tasks (paragraph manipulation, theming, footnotes/citations, table generation, content controls), instruction-following, cost, speed. GPU/RAM only matter for local Ollama models.

Respond with ONLY this JSON: {"modelId":"exact-id-from-list","reason":"2-3 sentences explaining why"}` }
    ], 400, 'meta-llama/llama-3.3-70b-instruct', false, false, orKey, null);

    const clean = text.replace(/<think>[\s\S]*?<\/think>/g, '').trim();
    let parsed = null;
    try { parsed = JSON.parse(clean); } catch {}
    if (!parsed) {
      const match = clean.match(/\{[\s\S]*\}/);
      if (match) try { parsed = JSON.parse(match[0]); } catch {}
    }
    if (parsed?.modelId) {
      console.log('[recommend] →', parsed.modelId);
      return res.json(parsed);
    }
    res.json({ modelId: null, reason: 'Could not parse recommendation.' });
  } catch (err) {
    console.error('[recommend] Error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── Feedback endpoint ────────────────────────────────────────────────────────
app.post('/api/feedback', (req, res) => {
  const entry = req.body;
  const filePath = path.join(__dirname, 'feedback.jsonl');
  fs.appendFileSync(filePath, JSON.stringify(entry) + '\n');
  console.log('[feedback] Saved entry for prompt:', entry.prompt?.slice(0, 60));
  res.json({ ok: true });
});

// ── Title generation endpoint ────────────────────────────────────────────────
app.post('/api/title', async (req, res) => {
  const { userMsg, aiReply, apiKey } = req.body;
  try {
    const { text } = await callAI([
      { role: 'system', content: 'Write a chat title: 4-6 words, no quotes, no punctuation at the end. Summarise what the user asked.' },
      { role: 'user', content: `User: ${String(userMsg).slice(0, 200)}\nAssistant: ${String(aiReply).slice(0, 200)}` }
    ], 30, null, false, false, apiKey || null);
    const title = text.replace(/<think>[\s\S]*?<\/think>/g, '').trim().replace(/^["'`]|["'`]$/g, '').slice(0, 60);
    res.json({ title });
  } catch { res.status(500).json({ error: 'Title generation failed' }); }
});

// ── Summarize endpoint ───────────────────────────────────────────────────────
app.post('/api/summarize', async (req, res) => {
  const { messages, previousSummary, apiKey } = req.body;
  try {
    const prevContext = previousSummary ? `Previous context (already summarised):\n${previousSummary}\n\nNew messages to incorporate:\n` : 'Summarise:\n';
    const { text } = await callAI([
      {
        role: 'system',
        content: 'You summarise Word assistant conversations. Write one short paragraph (max 100 words) covering: how many sections/paragraphs/pages, what styles/themes were applied, what changes were made, key headings or selection contexts. If given a previous summary, merge it with the new messages into one updated summary. Be specific. No fluff.'
      },
      {
        role: 'user',
        content: `${prevContext}\n${messages.map(m => `${m.role}: ${m.content}`).join('\n\n')}`
      }
    ], 250, null, false, false, apiKey || null);
    res.json({ summary: text.replace(/<think>[\s\S]*?<\/think>/g, '').trim() });
  } catch (err) {
    res.status(500).json({ error: 'Summarization failed' });
  }
});

// ───────────────────── ASSET PACK ENDPOINTS ──────────────────────────────────
function readManifest() {
  try { return JSON.parse(fs.readFileSync(MANIFEST_FILE, 'utf8')); } catch { return null; }
}

function packStatus() {
  const manifest = readManifest();
  if (!manifest) return { downloaded: false, sizeBytes: 0, imageCount: 0, version: null, categories: [] };
  let count = 0, size = 0;
  for (const img of manifest.images || []) {
    const p = path.join(PACK_DIR, img.id + (img.ext || '.jpg'));
    if (fs.existsSync(p)) { count++; size += fs.statSync(p).size; }
  }
  const categories = [...new Set((manifest.images || []).map(i => i.category))];
  return {
    downloaded: count > 0 && count >= (manifest.images || []).length * 0.9,
    sizeBytes: size,
    imageCount: count,
    totalImages: (manifest.images || []).length,
    version: manifest.version,
    categories,
  };
}

app.get('/api/assets/manifest', (req, res) => {
  const m = readManifest();
  if (!m) return res.json({ version: null, images: [], categories: [] });
  res.json(m);
});

app.get('/api/assets/status', (req, res) => res.json(packStatus()));

// SSE progress stream
let downloadState = { active: false, total: 0, done: 0, errors: [] };
app.get('/api/assets/download-pack/progress', (req, res) => {
  res.set({
    'Content-Type': 'text/event-stream',
    'Cache-Control': 'no-cache',
    'Connection': 'keep-alive',
  });
  const send = () => res.write(`data: ${JSON.stringify(downloadState)}\n\n`);
  send();
  const t = setInterval(send, 1000);
  req.on('close', () => clearInterval(t));
});

app.post('/api/assets/download-pack', async (req, res) => {
  const manifest = readManifest();
  if (!manifest) return res.status(400).json({ error: 'No pack manifest. Run npm run build-pack or download manifest.json first.' });
  if (downloadState.active) return res.status(409).json({ error: 'Download already in progress.' });

  res.json({ started: true, total: (manifest.images || []).length });

  downloadState = { active: true, total: (manifest.images || []).length, done: 0, errors: [] };
  (async () => {
    for (const img of manifest.images || []) {
      const ext = img.ext || '.jpg';
      const dest = path.join(PACK_DIR, img.id + ext);
      try {
        if (fs.existsSync(dest)) {
          const buf = fs.readFileSync(dest);
          const sha = crypto.createHash('sha256').update(buf).digest('hex');
          if (!img.sha256 || sha === img.sha256) { downloadState.done++; continue; }
        }
        const r = await fetch(img.source);
        if (!r.ok) throw new Error('HTTP ' + r.status);
        const buf = Buffer.from(await r.arrayBuffer());
        if (img.sha256) {
          const sha = crypto.createHash('sha256').update(buf).digest('hex');
          if (sha !== img.sha256) throw new Error('checksum mismatch');
        }
        fs.writeFileSync(dest, buf);
        downloadState.done++;
      } catch (err) {
        downloadState.errors.push({ id: img.id, error: err.message });
        downloadState.done++;
      }
    }
    downloadState.active = false;
  })();
});

app.get('/api/assets/pick', (req, res) => {
  const manifest = readManifest();
  if (!manifest) return res.status(404).json({ error: 'No pack manifest.' });
  const status = packStatus();
  if (!status.downloaded && status.imageCount === 0) {
    return res.status(404).json({ error: 'Pack not downloaded.' });
  }
  const category = String(req.query.category || '').toLowerCase();
  const keywords = String(req.query.keywords || '').toLowerCase().split(',').map(s => s.trim()).filter(Boolean);
  const aspect = parseFloat(req.query.aspect || '0');
  const styleHint = String(req.query.style || '').toLowerCase();

  const candidates = (manifest.images || [])
    .filter(img => !category || img.category === category)
    .filter(img => fs.existsSync(path.join(PACK_DIR, img.id + (img.ext || '.jpg'))));
  if (!candidates.length) return res.status(404).json({ error: 'No matching image found.' });

  const score = (img) => {
    const tagSet = new Set((img.tags || []).map(t => t.toLowerCase()));
    let s = keywords.reduce((acc, k) => acc + (tagSet.has(k) ? 3 : 0), 0);
    if (styleHint && img.style === styleHint) s += 2;
    if (aspect && img.w && img.h) {
      const a = img.w / img.h;
      if (Math.abs(a - aspect) < 0.2) s += 1;
    }
    s += Math.random() * 0.5;
    return s;
  };

  const picked = candidates.slice().sort((a, b) => score(b) - score(a))[0];
  const filePath = path.join(PACK_DIR, picked.id + (picked.ext || '.jpg'));
  const buf = fs.readFileSync(filePath);
  const base64 = buf.toString('base64');

  res.json({
    id: picked.id,
    base64,
    width: picked.w,
    height: picked.h,
    license: picked.license,
    attribution: picked.attribution || null,
    source: picked.source || null,
  });
});

app.get('/api/assets/image/:id', (req, res) => {
  const manifest = readManifest();
  if (!manifest) return res.status(404).end();
  const img = (manifest.images || []).find(i => i.id === req.params.id);
  if (!img) return res.status(404).end();
  const filePath = path.join(PACK_DIR, img.id + (img.ext || '.jpg'));
  if (!fs.existsSync(filePath)) return res.status(404).end();
  res.set('Content-Type', img.ext === '.png' ? 'image/png' : 'image/jpeg');
  res.send(fs.readFileSync(filePath));
});

// ───────────────────── THEME DESIGNER ENDPOINTS ──────────────────────────────
//
// AI-generated themes are unlimited. The user describes any look in natural
// language; the LLM emits a strict JSON palette+typography object; we validate
// fonts against the Word-safe whitelist and color contrast against WCAG AA;
// then return the cleaned theme.

const THEME_SYS_PROMPT = `You design typography + color palettes for Microsoft Word documents.

Return ONLY a single valid JSON object with EXACTLY these keys:
{
  "name":        "<3-5 word descriptive title>",
  "headingFont": "<font name, must be in WORD-SAFE list>",
  "bodyFont":    "<font name, must be in WORD-SAFE list>",
  "primary":     "#RRGGBB",
  "secondary":   "#RRGGBB",
  "accent":      "#RRGGBB",
  "bg":          "#RRGGBB",
  "text":        "#RRGGBB",
  "headingSize": <integer 14-32>,
  "tone":        "warm" | "cool" | "neutral" | "high-contrast" | "muted",
  "rationale":   "<one sentence justifying the design choices>"
}

WORD-SAFE FONTS (pick ONLY from this list):
Calibri, Calibri Light, Cambria, Georgia, Times New Roman, Garamond, Book Antiqua,
Arial, Arial Black, Arial Narrow, Helvetica, Verdana, Tahoma, Trebuchet MS,
Courier New, Consolas, Comic Sans MS, Impact, Palatino Linotype, Century Gothic,
Inter, Roboto, Open Sans, Source Sans Pro, Source Serif Pro, Lora, Merriweather,
Playfair Display, Montserrat, Poppins, Quicksand, Raleway, Oswald, Bebas Neue,
Lato, Nunito, PT Sans, PT Serif, Fira Sans, Work Sans, IBM Plex Sans, IBM Plex Serif,
EB Garamond, Libre Baskerville, Bitter, Karla, Manrope, Rubik, Futura, Avenir,
Avenir Next, Helvetica Neue, Optima, Baskerville, Didot, Hoefler Text.

DESIGN PRINCIPLES:
- Body text on background MUST have WCAG AA contrast (ratio >= 4.5 for body, >= 3 for headings).
- For documents, bg is usually #FFFFFF or a very near-white. Dark mode is allowed only if explicitly requested.
- Heading font and body font CAN be the same, but pairs (one serif + one sans) usually feel more refined.
- primary = heading color, secondary = subheading / accent stroke, accent = link / callout.
- Be FAITHFUL to the user's description (e.g. "matcha" → green palette; "1970s sci-fi" → orange/brown).

NO markdown. NO code fences. NO preamble. Just the JSON object.`;

async function designThemeLLM(description, prev, apiKey) {
  const messages = [
    { role: 'system', content: THEME_SYS_PROMPT },
  ];
  if (prev) {
    messages.push({
      role: 'user',
      content: `Current theme:\n${JSON.stringify(prev, null, 2)}\n\nUser wants to tweak it: "${description}"\n\nReturn the UPDATED theme object only.`,
    });
  } else {
    messages.push({ role: 'user', content: `Design a theme for: ${description}` });
  }
  const { text } = await callAI(messages, 600, null, false, false, apiKey || null);
  const clean = text.replace(/<think>[\s\S]*?<\/think>/g, '').replace(/^```(?:json)?\s*|\s*```$/g, '').trim();
  let obj;
  try { obj = JSON.parse(clean); }
  catch {
    const m = clean.match(/\{[\s\S]*\}/);
    if (m) obj = JSON.parse(m[0]);
    else throw new Error('LLM did not return valid JSON');
  }
  return obj;
}

function validateAndPatchTheme(t) {
  const warnings = [];
  // Defaults
  t.name        = String(t.name || 'Custom Theme').slice(0, 60);
  t.headingFont = String(t.headingFont || 'Calibri Light');
  t.bodyFont    = String(t.bodyFont    || 'Calibri');
  t.primary     = String(t.primary     || '#1A1A1A').toUpperCase();
  t.secondary   = String(t.secondary   || '#4A4A4A').toUpperCase();
  t.accent      = String(t.accent      || '#3F72AF').toUpperCase();
  t.bg          = String(t.bg          || '#FFFFFF').toUpperCase();
  t.text        = String(t.text        || '#1A1A1A').toUpperCase();
  t.headingSize = Math.max(12, Math.min(48, parseInt(t.headingSize) || 18));
  t.tone        = ['warm','cool','neutral','high-contrast','muted'].includes(t.tone) ? t.tone : 'neutral';
  t.rationale   = String(t.rationale   || '').slice(0, 240);

  // Font safety
  const hf = sanitizeFontSrv(t.headingFont);
  if (hf.changed) { warnings.push(`Heading font "${hf.original}" swapped to "${hf.font}" (not Word-safe).`); }
  t.headingFont = hf.font;
  const bf = sanitizeFontSrv(t.bodyFont);
  if (bf.changed) { warnings.push(`Body font "${bf.original}" swapped to "${bf.font}" (not Word-safe).`); }
  t.bodyFont = bf.font;

  // Color format
  ['primary','secondary','accent','bg','text'].forEach(k => {
    if (!/^#[0-9A-F]{6}$/.test(t[k])) {
      warnings.push(`Color ${k}="${t[k]}" was invalid — using default.`);
      t[k] = ({primary:'#1A1A1A',secondary:'#4A4A4A',accent:'#3F72AF',bg:'#FFFFFF',text:'#1A1A1A'})[k];
    }
  });

  // Contrast guard for body text on bg
  if (contrast(t.text, t.bg) < 4.5) {
    const bgLum = luminance(hexToRgb(t.bg));
    t.text = bgLum > 0.5 ? '#1A1A1A' : '#F0F0F0';
    warnings.push(`Body text adjusted to "${t.text}" for WCAG AA contrast on bg "${t.bg}".`);
  }
  if (contrast(t.primary, t.bg) < 3) {
    const bgLum = luminance(hexToRgb(t.bg));
    t.primary = bgLum > 0.5 ? darken(t.primary, 0.35) : '#FFFFFF';
    warnings.push(`Heading color adjusted to "${t.primary}" for WCAG AA contrast on bg.`);
  }

  return { theme: t, warnings };
}

function themeHash(description) {
  return crypto.createHash('sha1').update(String(description).toLowerCase().trim()).digest('hex').slice(0, 12);
}

app.get('/api/themes/recent', (req, res) => {
  const t = readThemes();
  res.json({ recent: t.recent || [] });
});

app.delete('/api/themes/recent/:id', (req, res) => {
  const t = readThemes();
  t.recent = (t.recent || []).filter(r => r.id !== req.params.id);
  writeThemes(t);
  res.json({ ok: true });
});

app.post('/api/design-theme', async (req, res) => {
  const { description, previous, apiKey } = req.body || {};
  if (!description || !String(description).trim()) {
    return res.status(400).json({ error: 'missing_description' });
  }
  try {
    // First attempt
    let raw;
    try { raw = await designThemeLLM(description, previous, apiKey); }
    catch (err) { return res.status(500).json({ error: 'llm_failed', detail: err.message }); }
    let { theme, warnings } = validateAndPatchTheme(raw);

    // If contrast guard kicked in heavily, give the LLM one more shot with a constraint.
    if (warnings.length >= 2 && contrast(theme.text, theme.bg) < 4.5) {
      try {
        const second = await designThemeLLM(
          description + '\n\nIMPORTANT: text-on-background contrast must satisfy WCAG AA (>=4.5).',
          previous, apiKey
        );
        const v2 = validateAndPatchTheme(second);
        if (v2.warnings.length < warnings.length) { theme = v2.theme; warnings = v2.warnings; }
      } catch {}
    }

    // Save in recent
    const data = readThemes();
    if (!Array.isArray(data.recent)) data.recent = [];
    const id = themeHash(description);
    data.recent = [{ id, description, theme, createdAt: new Date().toISOString() }]
      .concat(data.recent.filter(r => r.id !== id))
      .slice(0, 12);
    writeThemes(data);

    res.json({ id, theme, warnings, description });
  } catch (err) {
    res.status(500).json({ error: 'theme_failed', detail: err.message });
  }
});

// "Surprise me" — random creative theme. Generates a fresh palette via the same
// pipeline as /api/design-theme using one of a rotating set of seed descriptions.
app.post('/api/design-theme/surprise', async (req, res) => {
  const seeds = [
    'a deeply unusual editorial theme inspired by a real-world art movement of your choice',
    'a vibrant 1990s zine-style theme with high contrast and an unexpected accent',
    'a serene Japanese minimalist theme inspired by Muji catalogs',
    'a moody noir editorial theme — dark grey, oxblood accent, modern serif',
    'a Wes Anderson pastel theme with Futura headings',
    'a Bauhaus theme using only red, blue, yellow, black, and cream',
    'a botanical conservation report theme — sage, parchment, deep forest accent',
    'a tactile newspaper feature theme — newsprint cream, ink black, single red accent',
    'a fintech investor memo theme — deep navy, electric teal accent, modern sans',
    'a craft coffee menu theme — warm parchment, terracotta accent, soft serif',
  ];
  const pick = seeds[Math.floor(Math.random() * seeds.length)];
  try {
    let raw;
    try { raw = await designThemeLLM(pick, null, req.body?.apiKey); }
    catch (err) { return res.status(500).json({ error: 'llm_failed', detail: err.message }); }
    const { theme, warnings } = validateAndPatchTheme(raw);
    const data = readThemes();
    if (!Array.isArray(data.recent)) data.recent = [];
    const id = themeHash(pick);
    data.recent = [{ id, description: pick, theme, createdAt: new Date().toISOString() }]
      .concat(data.recent.filter(r => r.id !== id))
      .slice(0, 12);
    writeThemes(data);
    res.json({ id, theme, warnings, description: pick });
  } catch (err) {
    res.status(500).json({ error: 'theme_failed', detail: err.message });
  }
});

// ───────────────────── WEB SEARCH ENDPOINT ───────────────────────────────────
let webSearchBucket = { tokens: 10, last: Date.now() };
function takeWebToken() {
  const now = Date.now();
  const refill = (now - webSearchBucket.last) / 6000; // 10/min = 1 per 6s
  webSearchBucket.tokens = Math.min(10, webSearchBucket.tokens + refill);
  webSearchBucket.last = now;
  if (webSearchBucket.tokens < 1) return false;
  webSearchBucket.tokens -= 1;
  return true;
}

app.get('/api/web-search/status', (req, res) => {
  const provider = resolvedProvider();
  res.json({
    provider,
    configured: !!provider,
    pref: PROVIDER_PREF,
    available: {
      google: !!(GOOGLE_KEY && GOOGLE_CX),
      brave:  !!BRAVE_KEY,
    },
  });
});

async function runWebSearch(query, count) {
  const provider = resolvedProvider();
  if (!provider) throw Object.assign(new Error('no_provider_configured'), { status: 503 });

  if (provider === 'google') {
    const url = 'https://www.googleapis.com/customsearch/v1?key=' + encodeURIComponent(GOOGLE_KEY)
              + '&cx=' + encodeURIComponent(GOOGLE_CX)
              + '&q=' + encodeURIComponent(query)
              + '&num=' + count;
    const r = await fetch(url);
    const data = await r.json();
    if (data.error) throw new Error(data.error.message || 'google_search_failed');
    const results = (data.items || []).slice(0, count).map(x => ({
      title: x.title || '',
      url: x.link || '',
      snippet: (x.snippet || '').replace(/\s+/g, ' ').slice(0, 400),
      age: null,
    }));
    return { query, provider: 'google', results, fetchedAt: new Date().toISOString() };
  }

  // brave
  const url = 'https://api.search.brave.com/res/v1/web/search?q=' + encodeURIComponent(query) + '&count=' + count;
  const r = await fetch(url, {
    headers: { 'X-Subscription-Token': BRAVE_KEY, 'Accept': 'application/json' }
  });
  const data = await r.json();
  const results = (data.web?.results || []).slice(0, count).map(x => ({
    title: x.title || '',
    url: x.url || '',
    snippet: (x.description || '').replace(/<[^>]+>/g, '').slice(0, 400),
    age: x.age || null,
  }));
  return { query, provider: 'brave', results, fetchedAt: new Date().toISOString() };
}

app.post('/api/web-search', async (req, res) => {
  if (!resolvedProvider()) return res.status(503).json({ error: 'no_provider_configured' });
  if (!takeWebToken()) return res.status(429).json({ error: 'rate_limited' });

  const query = String(req.body.query || '').trim();
  const count = Math.max(1, Math.min(10, parseInt(req.body.count || 5, 10)));
  if (!query) return res.status(400).json({ error: 'missing_query' });

  try {
    const out = await runWebSearch(query, count);
    try {
      const log = fs.existsSync(WEB_LOG) ? JSON.parse(fs.readFileSync(WEB_LOG, 'utf8')) : [];
      log.push({ query, provider: out.provider, count: out.results.length, fetchedAt: out.fetchedAt });
      fs.writeFileSync(WEB_LOG, JSON.stringify(log.slice(-200), null, 2));
    } catch {}
    res.json(out);
  } catch (err) {
    const status = err.status || 500;
    res.status(status).json({ error: status === 503 ? 'no_provider_configured' : 'search_failed', detail: err.message });
  }
});

// ───────────────────── DOCUMENT GENERATION ENDPOINT ──────────────────────────
app.post('/api/generate-doc', async (req, res) => {
  const {
    topic, audience = 'general', tone = 'formal', length = 'medium',
    includeToc = false, includeCover = false, includeBibliography = false,
    model, useOllama, useGroq, apiKey, groqKey,
  } = req.body || {};
  if (!topic) return res.status(400).json({ error: 'missing_topic' });
  const lengthHint = { short:'400-700', medium:'900-1400', long:'1800-2500', 'very-long':'3500-5000' }[length] || '900-1400';

  const sys = `You design Word document outlines. Output ONLY a JSON array of section specs. NO markdown, NO prose, NO code fences. Each spec: {"type": "title"|"subtitle"|"heading1"|"heading2"|"heading3"|"paragraph"|"bullet"|"number"|"quote"|"pageBreak"|"toc"|"bibliography", "text": "..."}.

Rules:
- The FIRST item should be {"type":"title","text":"<doc title>"}; optionally a {"type":"subtitle","text":"..."} after it.
- Use heading1 for top-level sections, heading2/3 for subsections.
- "text" can include line breaks (\\n) within a paragraph but NOT for bullets — emit multiple {"type":"bullet"} items instead.
- Body must be SPECIFIC to the topic with concrete facts, examples, and reasoning. No filler.
- Tone "${tone}", audience "${audience}". Target total word count: ${lengthHint}.
${includeToc ? '- Insert {"type":"toc","text":"Table of Contents"} as the second or third item.' : ''}
${includeCover ? '- After the title/subtitle, insert {"type":"pageBreak"} so the title acts as a cover page.' : ''}
${includeBibliography ? '- Last item should be {"type":"bibliography","text":"APA"} (APA style).' : ''}`;

  const usr = `Generate a ${tone} ${length} Word document about: ${topic}
Audience: ${audience}

Return ONLY the JSON array. No backticks, no preamble.`;

  try {
    const { text, usage } = await callAI([
      { role: 'system', content: sys },
      { role: 'user', content: usr }
    ], 5500, model || null, useOllama || false, useGroq || false, apiKey || null, groqKey || null);
    const clean = text.replace(/<think>[\s\S]*?<\/think>/g, '').replace(/^```(?:json)?\s*|\s*```$/g, '').trim();
    let outline;
    try { outline = JSON.parse(clean); } catch {
      const m = clean.match(/\[[\s\S]*\]/);
      if (m) outline = JSON.parse(m[0]);
    }
    if (!Array.isArray(outline)) throw new Error('Outline was not an array.');
    const allowed = ['title','subtitle','heading1','heading2','heading3','paragraph','bullet','number','quote','pageBreak','toc','bibliography'];
    outline = outline.map(s => ({
      type: allowed.includes(s.type) ? s.type : 'paragraph',
      text: String(s.text || '').slice(0, 4000),
    })).filter(s => s.type === 'pageBreak' || s.type === 'toc' || s.text.trim());
    res.json({ outline, usage });
  } catch (err) {
    console.error('[generate-doc] Error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ───────────────────── WRITING COACH ENDPOINT ────────────────────────────────
app.post('/api/coach', async (req, res) => {
  const {
    text = '', scope = 'selection', focus = 'clarity', action = 'review',
    model, useOllama, useGroq, apiKey, groqKey,
  } = req.body || {};

  // Two flavors:
  //   action="review"   → free-form feedback (no rewrite); structured response.
  //   action="rewrite"  → returns a single rewritten string the UI can paste over selection.
  const sys = action === 'rewrite'
    ? `You are a writing assistant. Rewrite the text BELOW with focus="${focus}".

Rules:
- "clarity":   simplify wording, prefer concrete nouns, kill jargon, keep meaning.
- "tone":      match tone="${req.body.toneHint || 'professional'}" (e.g. friendly, persuasive, academic, casual).
- "grammar":   fix grammar and punctuation; preserve voice and rhythm.
- "brevity":   cut by ~30% without losing key information.
- "persuasive":sharpen the argument, add a specific call to action.
- "reading-level": rewrite so it's understandable at grade level ${req.body.gradeHint || '8'} or lower.
- "translate": translate into ${req.body.targetLang || 'Spanish'}; preserve formatting.
Return ONLY the rewritten text. No commentary, no headers, no quotes around it. Plain prose with paragraph breaks.`
    : `You are a writing coach. Review the text below with focus="${focus}".

Output EXACTLY this structure:
SUMMARY: 1-2 sentences naming what works and the single biggest weakness.
CLARITY: 1-2 sentences with a specific suggestion.
TONE: 1 sentence on tone match.
GRAMMAR: bullet 0-3 specific errors with corrections.
SUGGESTED REWRITE: 2-4 sentences the user can paste to replace a key passage (no quotes around it).
READING LEVEL: report the Flesch-Kincaid grade estimate and one suggestion to lower/raise it.

Total under 280 words. Never output CODE_JS.`;

  const usr = `SCOPE: ${scope}
FOCUS: ${focus}
${req.body.toneHint ? 'TONE_HINT: ' + req.body.toneHint + '\n' : ''}${req.body.targetLang ? 'TARGET_LANG: ' + req.body.targetLang + '\n' : ''}
TEXT:
${String(text).slice(0, 8000) || '(empty)'}`;

  try {
    const { text: out, usage } = await callAI([
      { role: 'system', content: sys },
      { role: 'user', content: usr }
    ], action === 'rewrite' ? 1800 : 700, model || null, useOllama || false, useGroq || false, apiKey || null, groqKey || null);
    const cleaned = out.replace(/<think>[\s\S]*?<\/think>/g, '').replace(/CODE_JS::[\s\S]*?::END_CODE/g, '').trim();
    res.json({ feedback: cleaned, usage, action });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ───────────────────── MAIN CHAT ROUTE ───────────────────────────────────────
app.post('/api/chat', async (req, res) => {
  const { messages, docData, selectionInfo, summary, model, useOllama, useGroq, apiKey, groqKey, options, settings, webSearchResults } = req.body;

  const preferences   = options?.preferences   || '';
  const deepThink     = options?.deepThink     || false;
  const dynamicDepth  = options?.dynamicDepth  || false;
  const autoModel     = options?.autoModel     || false;
  const planFirst     = options?.plan          || false;
  const sectionByStep = options?.sectionByStep || false;
  const webSearchMode = (settings?.allowWebSearch || 'off').toLowerCase(); // 'off' | 'on' | 'disabled'

  let maxTokens = 4096;
  let effectiveModel = model || null;
  let selectedModel = null;

  const rawUserMessage = messages[messages.length - 1]?.content || '';
  const recentMessages = messages.slice(-6).map(m => ({ ...m }));

  // ── Auto Model ──────────────────────────────────────────────────────────
  if (autoModel && options?.availableModels?.length) {
    try {
      const { text: choice } = await callAI([
        { role: 'system', content: 'Select the best AI model for this task. JSON only: {"modelId":"id"}' },
        { role: 'user', content: `Task: "${rawUserMessage.slice(0, 300)}"\nPreferences: ${preferences || 'none'}\nCost limit: $${options.costLimit || 'unlimited'} per prompt\nModels:\n${options.availableModels.map(m=>`${m.id} $${m.in}/$${m.out}`).join('\n')}` }
      ], 60, null, false, false, apiKey || null, groqKey || null);
      const m = choice.replace(/<think>[\s\S]*?<\/think>/g,'').trim().match(/\{[\s\S]*?\}/);
      if (m) {
        const parsed = JSON.parse(m[0]);
        if (parsed.modelId && options.availableModels.find(x => x.id === parsed.modelId)) {
          effectiveModel = parsed.modelId;
          selectedModel  = parsed.modelId;
        }
      }
    } catch {}
  }

  // ── Dynamic Depth ───────────────────────────────────────────────────────
  if (dynamicDepth) {
    try {
      const { text: assess } = await callAI([
        { role: 'system', content: 'Rate complexity 1-3: 1=simple, 2=moderate, 3=complex. JSON: {"level":1|2|3}' },
        { role: 'user', content: rawUserMessage.slice(0, 300) }
      ], 30, null, false, false, apiKey || null, groqKey || null);
      const m = assess.match(/\{[\s\S]*?\}/);
      const level = m ? (JSON.parse(m[0]).level || 2) : 2;
      maxTokens = level >= 3 ? 8192 : level === 1 ? 2048 : 4096;
    } catch {}
  }

  // ── Deep Think ──────────────────────────────────────────────────────────
  if (deepThink) {
    const lastUserIdx = recentMessages.map(m=>m.role).lastIndexOf('user');
    if (lastUserIdx !== -1) {
      try {
        const orig = recentMessages[lastUserIdx].content;
        const { text: enhanced } = await callAI([
          { role: 'system', content: 'You are a prompt engineer for a Word AI assistant. Rewrite the user\'s request to be maximally clear, precise, and complete. Preserve intent exactly. Add explicit handling for edge cases (no selection, missing heading, target text not found). Reference specific paragraph indexes, heading texts, or selection contexts if visible. Output ONLY the rewritten prompt.' },
          { role: 'user', content: `Document context (heading outline + selection):\n${(docData||'').slice(0,2000)}\n\nOriginal request: ${orig}` }
        ], 700, null, false, false, apiKey || null, groqKey || null);
        const enhancedText = enhanced.replace(/<think>[\s\S]*?<\/think>/g,'').trim();
        recentMessages[lastUserIdx] = { ...recentMessages[lastUserIdx], content: enhancedText };
        maxTokens = Math.max(maxTokens, 8192);
      } catch {}
    }
  }

  // ── Plan First ──────────────────────────────────────────────────────────
  let planText = null;
  if (planFirst && !isQuestion(rawUserMessage)) {
    try {
      const prefsSection_ = preferences ? `\n\nUSER PREFERENCES:\n${preferences}` : '';
      const contextMessages_ = docData ? [
        { role: 'user', content: `Here is the current state of the document.\n\nSelection: ${selectionInfo || '(none)'}\n\n${docData}` },
        { role: 'assistant', content: 'I can see the document. What would you like me to do?' }
      ] : [];
      const { text: planReply } = await callAI([
        { role: 'system', content: SYSTEM_PROMPT + prefsSection_ + '\n\nIMPORTANT: The user has requested a plan. Do NOT output any CODE_JS block. Instead, write a short numbered plan (3-5 steps max) describing what you will do. Be concise.' },
        ...contextMessages_,
        { role: 'user', content: `Before executing, give me a brief plan for: ${rawUserMessage}` }
      ], 512, effectiveModel, useOllama || false, useGroq || false, apiKey || null, groqKey || null);
      planText = planReply.replace(/<think>[\s\S]*?<\/think>/g, '').trim();
    } catch {}
  }

  // ── Code reminder + section-by-section preview hint ────────────────────
  const lastIdx = recentMessages.map(m=>m.role).lastIndexOf('user');
  if (lastIdx !== -1) {
    let extra = '\n\n[REMINDER: If making changes, output a CODE_JS block with Office JS code. Do not skip it.]';
    if (sectionByStep) extra += '\n[If your change spans more than one section/heading, FIRST list the sections you will modify as a numbered preview, then emit the CODE_JS block.]';
    recentMessages[lastIdx] = {
      ...recentMessages[lastIdx],
      content: recentMessages[lastIdx].content + extra
    };
  }

  // ── Inject runtime hints (asset pack + web search mode) ─────────────────
  const status = packStatus();
  const assetTag = status.downloaded
    ? `<ASSET_PACK_READY categories="${status.categories.join(',')}"/>`
    : `<ASSET_PACK_MISSING/>`;
  const webTag = `<WEB_SEARCH_MODE value="${webSearchMode}"/>`;
  const runtimeHints = `\n\nRUNTIME STATE:\n${assetTag}\n${webTag}`;

  const prefsSection = preferences ? `\n\nUSER PREFERENCES (always follow these):\n${preferences}` : '';

  const contextMessages = docData ? [
    { role: 'user', content: `Here is the current state of the document (last synced before this message).\n\nSelection: ${selectionInfo || '(none)'}\n\n${docData}` },
    { role: 'assistant', content: 'I can see the document. What would you like me to do?' }
  ] : [];

  const summaryMessages = summary ? [
    { role: 'user', content: `Earlier in this session: ${summary}` },
    { role: 'assistant', content: 'Got it, I have the context of what we did earlier.' }
  ] : [];

  const webResultsMessages = (Array.isArray(webSearchResults) && webSearchResults.length) ? [{
    role: 'user',
    content: webSearchResults.map(wr =>
      `<WEB_SEARCH_RESULTS query="${(wr.query||'').replace(/"/g,'&quot;')}">\n` +
      (wr.results || []).map((r, i) => `${i+1}. ${r.title} — ${r.url}\n   ${r.snippet}`).join('\n') +
      `\n</WEB_SEARCH_RESULTS>`
    ).join('\n\n')
  }] : [];

  const allMessages = [
    { role: 'system', content: SYSTEM_PROMPT + prefsSection + runtimeHints },
    ...contextMessages,
    ...summaryMessages,
    ...webResultsMessages,
    ...recentMessages
  ];

  try {
    let { text: responseText, usage } = await callAI(allMessages, maxTokens, effectiveModel, useOllama || false, useGroq || false, apiKey || null, groqKey || null);
    responseText = responseText.replace(/<think>[\s\S]*?<\/think>/g, '').trim();

    // ── Web search interception ─────────────────────────────────────────
    let webSearchUsed = null;
    const wsReq = parseWebSearchRequest(responseText);
    if (wsReq && webSearchMode === 'on' && resolvedProvider()) {
      try {
        webSearchUsed = await runWebSearch(wsReq.query, 5);
        const followup = [
          ...allMessages,
          { role: 'assistant', content: responseText },
          { role: 'user', content:
            `<WEB_SEARCH_RESULTS query="${wsReq.query.replace(/"/g,'&quot;')}" provider="${webSearchUsed.provider}">\n` +
            webSearchUsed.results.map((r, i) => `${i+1}. ${r.title} — ${r.url}\n   ${r.snippet}`).join('\n') +
            `\n</WEB_SEARCH_RESULTS>\n\nNow proceed with the original request using these facts. Emit CODE_JS if changes are needed.`
          }
        ];
        const second = await callAI(followup, maxTokens, effectiveModel, useOllama || false, useGroq || false, apiKey || null, groqKey || null);
        responseText = second.text.replace(/<think>[\s\S]*?<\/think>/g, '').trim();
        usage = second.usage;
      } catch (err) {
        console.error('[web-search] auto-mode failed:', err.message);
      }
    } else if (wsReq && webSearchMode === 'off') {
      const { cleaned } = parseResponse(responseText);
      return res.json({
        response: cleaned,
        code: null,
        webSearchRequest: wsReq,
        usage,
        selectedModel,
        plan: planText
      });
    }

    if (modelForgotCode(responseText, rawUserMessage)) {
      console.log('[server] Model forgot CODE_JS — retrying...');
      const { text: retryText } = await callAI([
        ...allMessages,
        { role: 'assistant', content: responseText },
        { role: 'user', content: 'You forgot the CODE_JS block. Output ONLY the CODE_JS block now. Start with CODE_JS:: and end with ::END_CODE.' }
      ], 2048, effectiveModel, useOllama || false, useGroq || false, apiKey || null, groqKey || null);
      const retryClean = retryText.replace(/<think>[\s\S]*?<\/think>/g, '').trim();
      const { code: retryCode } = parseResponse(retryClean);
      if (retryCode) {
        const { cleaned } = parseResponse(responseText);
        return res.json({ response: cleaned, code: retryCode, usage, selectedModel, webSearchUsed });
      }
    }

    const { code, cleaned } = parseResponse(responseText);
    if (code) console.log('[server] Code to execute:\n' + code);
    res.json({ response: cleaned, code, usage, selectedModel, plan: planText, webSearchUsed });

  } catch (err) {
    console.error('AI error:', err);
    res.status(500).json({ error: `AI error: ${err.message}` });
  }
});

const server = https.createServer({ key: pems.private, cert: pems.cert }, app);
server.listen(3000, () => {
  console.log('Word AI Assistant running at https://localhost:3000');
  console.log(`Mode: ${USE_OPENROUTER ? 'OpenRouter' : USE_MLX ? 'MLX' : USE_GROQ ? 'Groq' : 'Ollama'}`);
  console.log(`Default model: ${DEFAULT_MODEL}`);
  const prov = resolvedProvider();
  console.log(`Web search: ${prov ? prov + ' (configured)' : 'not configured'}`);
  console.log(`  Google: ${(GOOGLE_KEY && GOOGLE_CX) ? 'ready' : 'no GOOGLE_KEY/GOOGLE_CX'} · Brave: ${BRAVE_KEY ? 'ready' : 'no BRAVE_KEY'} · preference: ${PROVIDER_PREF}`);
  console.log(`Asset pack: ${packStatus().downloaded ? `${packStatus().imageCount} images` : 'not downloaded'}`);
});
