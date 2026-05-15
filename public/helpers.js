/* Word AI Assistant — helpers.js
 *
 * Exposes static data the in-page execCode() block uses at runtime:
 *   - WORD_SAFE_FONTS        whitelist of fonts Word can render reliably
 *   - THEME_SEEDS            quick-start chip labels for the theme designer
 *   - TEMPLATES              14 document templates with placeholders + sections
 *   - OOXML                  small reusable OOXML snippets (TOC, columns, page-break,
 *                            watermark, sectPr, content-controls, equations)
 *   - latexToOmml()          tiny LaTeX → OMML converter for common patterns
 *   - wcagContrast()         contrast-ratio guard re-used both in browser + server
 *   - sanitizeFont()         pick the closest Word-safe substitute for any font name
 *
 * Helpers like addHeading, addParagraph, insertTable, applyTheme are defined inside
 * execCode() in index.html so they can capture the live Word.run context.
 *
 * Loaded BEFORE the inline script in index.html via <script src="helpers.js">.
 */

(function () {
  'use strict';

  /* ───────────────────────────── WORD-SAFE FONTS ───────────────────────────── */
  /* Fonts Word renders reliably on both desktop and Word on the web. The LLM is
   * constrained to this list; anything else is rerouted via sanitizeFont(). */
  const WORD_SAFE_FONTS = [
    'Calibri', 'Calibri Light', 'Cambria', 'Cambria Math',
    'Georgia', 'Times New Roman', 'Garamond', 'Book Antiqua',
    'Arial', 'Arial Black', 'Arial Narrow', 'Helvetica',
    'Verdana', 'Tahoma', 'Trebuchet MS', 'Lucida Sans Unicode',
    'Courier New', 'Consolas', 'Lucida Console',
    'Comic Sans MS', 'Impact', 'Palatino Linotype', 'Century Gothic',
    // Common Google / cloud fonts that render via Word's font substitution:
    'Inter', 'Roboto', 'Roboto Slab', 'Open Sans', 'Source Sans Pro',
    'Source Serif Pro', 'Source Code Pro', 'Lora', 'Merriweather',
    'Playfair Display', 'Montserrat', 'Poppins', 'Quicksand', 'Raleway',
    'Oswald', 'Bebas Neue', 'Lato', 'Nunito', 'Nunito Sans', 'PT Sans',
    'PT Serif', 'Fira Sans', 'Fira Code', 'Work Sans', 'IBM Plex Sans',
    'IBM Plex Serif', 'IBM Plex Mono', 'Crimson Text', 'EB Garamond',
    'Cormorant Garamond', 'Libre Baskerville', 'Bitter', 'Karla', 'Manrope',
    'Mulish', 'Rubik', 'Ubuntu', 'Futura', 'Avenir', 'Avenir Next',
    'Helvetica Neue', 'Optima', 'Baskerville', 'Didot', 'Hoefler Text',
  ];

  /* Map common LLM-hallucinated names → safe substitutes */
  const FONT_FALLBACKS = {
    'helvetica neue condensed': 'Helvetica Neue',
    'helvetica condensed': 'Helvetica',
    'avenir next condensed': 'Avenir Next',
    'gotham': 'Montserrat',
    'gotham bold': 'Montserrat',
    'proxima nova': 'Montserrat',
    'museo sans': 'Open Sans',
    'din': 'Oswald',
    'din alternate': 'Oswald',
    'helvetica neue thin': 'Helvetica Neue',
    'sf pro display': 'Inter',
    'sf pro text': 'Inter',
    'system': 'Calibri',
    'system-ui': 'Calibri',
    'sans-serif': 'Calibri',
    'serif': 'Cambria',
    'monospace': 'Consolas',
    'cursive': 'Comic Sans MS',
    'fantasy': 'Impact',
    'comic sans': 'Comic Sans MS',
    'comic sans replacement': 'Comic Sans MS',
    'arial unicode ms': 'Arial',
    'times': 'Times New Roman',
    'courier': 'Courier New',
    'roboto condensed': 'Roboto',
    'open sans condensed': 'Open Sans',
    'pt sans narrow': 'PT Sans',
    'futura bold': 'Futura',
  };

  function sanitizeFont(name) {
    if (!name) return { font: 'Calibri', changed: false };
    const trimmed = String(name).trim();
    if (!trimmed) return { font: 'Calibri', changed: false };
    // Exact match
    for (const f of WORD_SAFE_FONTS) {
      if (f.toLowerCase() === trimmed.toLowerCase()) return { font: f, changed: false };
    }
    // Known fallback
    const low = trimmed.toLowerCase();
    if (FONT_FALLBACKS[low]) return { font: FONT_FALLBACKS[low], changed: true, original: trimmed };
    // Family-name prefix match — "Roboto Mono" → "Roboto"
    for (const f of WORD_SAFE_FONTS) {
      if (low.startsWith(f.toLowerCase() + ' ')) return { font: f, changed: true, original: trimmed };
    }
    // Heuristics by suffix
    if (/mono|code|courier/.test(low)) return { font: 'Consolas', changed: true, original: trimmed };
    if (/serif|antiqua|roman/.test(low)) return { font: 'Cambria', changed: true, original: trimmed };
    if (/script|cursive|hand/.test(low)) return { font: 'Comic Sans MS', changed: true, original: trimmed };
    return { font: 'Calibri', changed: true, original: trimmed };
  }

  /* ───────────────────────────── WCAG contrast ───────────────────────────── */
  function hexToRgb(hex) {
    const m = String(hex || '').replace('#', '').match(/^([\da-f]{6})$/i);
    if (!m) return null;
    const n = parseInt(m[1], 16);
    return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255 };
  }
  function luminance({ r, g, b }) {
    const a = [r, g, b].map(v => {
      v /= 255;
      return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4);
    });
    return 0.2126 * a[0] + 0.7152 * a[1] + 0.0722 * a[2];
  }
  function wcagContrast(hex1, hex2) {
    const c1 = hexToRgb(hex1), c2 = hexToRgb(hex2);
    if (!c1 || !c2) return 0;
    const L1 = luminance(c1), L2 = luminance(c2);
    return (Math.max(L1, L2) + 0.05) / (Math.min(L1, L2) + 0.05);
  }

  /* ───────────────────────────── THEME SEEDS ───────────────────────────── */
  /* NOT preset palettes. Just chip labels that pre-fill the prompt box so
   * the LLM can generate a fresh theme on each click. */
  const THEME_SEEDS = [
    { label: 'Corporate clean',  prompt: 'a clean modern corporate theme with a navy primary and crisp sans-serif typography' },
    { label: 'Academic serif',   prompt: 'a refined academic theme with classic serif typography, generous margins, and conservative colors' },
    { label: 'Legal traditional',prompt: 'a traditional legal document theme — black text, Times New Roman body, slim accent, no decoration' },
    { label: 'Modern resume',    prompt: 'a modern minimalist resume theme with a single accent color and tight Inter typography' },
    { label: 'Creative warm',    prompt: 'a creative warm theme with sunset accents and friendly modern sans-serif headings' },
    { label: 'Minimal mono',     prompt: 'a strictly minimal monochrome theme — black, white, one neutral accent, no flourish' },
    { label: 'Bold editorial',   prompt: 'a bold editorial theme with high-contrast typography, oversized headings, and a single saturated accent' },
    { label: 'Technical dark',   prompt: 'a technical engineering document theme with a deep background, monospaced inline code, and a cool accent' },
  ];

  /* ───────────────────────────── DOCUMENT TEMPLATES ───────────────────────────── */
  /* Each section: { type, text, style?, options? }
   * type ∈ "heading1" | "heading2" | "heading3" | "title" | "subtitle"
   *      | "paragraph" | "quote" | "bullet" | "number" | "spacer" | "pageBreak" | "signature" */
  const TEMPLATES = {
    resume: {
      name: 'Modern Resume',
      description: 'One-page resume with skills, experience, education.',
      placeholders: ['Name','Title','Email','Phone','City','Summary','Skills','Experience','Education'],
      sections: [
        { type: 'title',      text: '{{Name}}' },
        { type: 'subtitle',   text: '{{Title}} · {{City}} · {{Email}} · {{Phone}}' },
        { type: 'heading2',   text: 'Summary' },
        { type: 'paragraph',  text: '{{Summary}}' },
        { type: 'heading2',   text: 'Skills' },
        { type: 'paragraph',  text: '{{Skills}}' },
        { type: 'heading2',   text: 'Experience' },
        { type: 'paragraph',  text: '{{Experience}}' },
        { type: 'heading2',   text: 'Education' },
        { type: 'paragraph',  text: '{{Education}}' },
      ],
    },
    'cover-letter': {
      name: 'Cover Letter',
      description: 'Formal cover letter with recipient block and signature.',
      placeholders: ['Name','Address','City','Date','Recipient','Company','RecipientAddress','RoleTitle','Body','SignOff'],
      sections: [
        { type: 'paragraph',  text: '{{Name}}' },
        { type: 'paragraph',  text: '{{Address}}' },
        { type: 'paragraph',  text: '{{City}}' },
        { type: 'paragraph',  text: '{{Date}}' },
        { type: 'spacer',     text: '' },
        { type: 'paragraph',  text: '{{Recipient}}' },
        { type: 'paragraph',  text: '{{Company}}' },
        { type: 'paragraph',  text: '{{RecipientAddress}}' },
        { type: 'spacer',     text: '' },
        { type: 'paragraph',  text: 'Dear {{Recipient}},' },
        { type: 'paragraph',  text: '{{Body}}' },
        { type: 'spacer',     text: '' },
        { type: 'paragraph',  text: '{{SignOff}}' },
        { type: 'signature',  text: '{{Name}}' },
      ],
    },
    'business-letter': {
      name: 'Business Letter',
      description: 'Formal business letter with letterhead block.',
      placeholders: ['SenderName','SenderCompany','SenderAddress','Date','Recipient','RecipientCompany','RecipientAddress','Subject','Body'],
      sections: [
        { type: 'paragraph',  text: '{{SenderName}}' },
        { type: 'paragraph',  text: '{{SenderCompany}}' },
        { type: 'paragraph',  text: '{{SenderAddress}}' },
        { type: 'spacer',     text: '' },
        { type: 'paragraph',  text: '{{Date}}' },
        { type: 'spacer',     text: '' },
        { type: 'paragraph',  text: '{{Recipient}}' },
        { type: 'paragraph',  text: '{{RecipientCompany}}' },
        { type: 'paragraph',  text: '{{RecipientAddress}}' },
        { type: 'spacer',     text: '' },
        { type: 'heading3',   text: 'Re: {{Subject}}' },
        { type: 'paragraph',  text: 'Dear {{Recipient}},' },
        { type: 'paragraph',  text: '{{Body}}' },
        { type: 'spacer',     text: '' },
        { type: 'paragraph',  text: 'Sincerely,' },
        { type: 'signature',  text: '{{SenderName}}' },
      ],
    },
    'formal-report': {
      name: 'Formal Report',
      description: 'Cover page, exec summary, sections, conclusion.',
      placeholders: ['Title','Author','Date','ExecutiveSummary','Background','Findings','Recommendations','Conclusion'],
      sections: [
        { type: 'title',      text: '{{Title}}' },
        { type: 'subtitle',   text: 'Prepared by {{Author}} · {{Date}}' },
        { type: 'pageBreak',  text: '' },
        { type: 'heading1',   text: 'Executive Summary' },
        { type: 'paragraph',  text: '{{ExecutiveSummary}}' },
        { type: 'heading1',   text: 'Background' },
        { type: 'paragraph',  text: '{{Background}}' },
        { type: 'heading1',   text: 'Findings' },
        { type: 'paragraph',  text: '{{Findings}}' },
        { type: 'heading1',   text: 'Recommendations' },
        { type: 'paragraph',  text: '{{Recommendations}}' },
        { type: 'heading1',   text: 'Conclusion' },
        { type: 'paragraph',  text: '{{Conclusion}}' },
      ],
    },
    proposal: {
      name: 'Project Proposal',
      description: 'Cover, problem, solution, scope, timeline, cost.',
      placeholders: ['Title','Client','Date','Problem','Solution','Scope','Timeline','Cost','Author'],
      sections: [
        { type: 'title',      text: '{{Title}}' },
        { type: 'subtitle',   text: 'Prepared for {{Client}} · {{Date}}' },
        { type: 'pageBreak',  text: '' },
        { type: 'heading1',   text: 'Problem Statement' },
        { type: 'paragraph',  text: '{{Problem}}' },
        { type: 'heading1',   text: 'Proposed Solution' },
        { type: 'paragraph',  text: '{{Solution}}' },
        { type: 'heading1',   text: 'Scope of Work' },
        { type: 'paragraph',  text: '{{Scope}}' },
        { type: 'heading1',   text: 'Timeline' },
        { type: 'paragraph',  text: '{{Timeline}}' },
        { type: 'heading1',   text: 'Investment' },
        { type: 'paragraph',  text: '{{Cost}}' },
        { type: 'spacer',     text: '' },
        { type: 'signature',  text: '{{Author}}' },
      ],
    },
    contract: {
      name: 'Service Contract',
      description: 'Binding agreement with parties, terms, signatures.',
      placeholders: ['Title','Date','PartyA','PartyB','Scope','Term','Compensation','Termination'],
      sections: [
        { type: 'title',      text: '{{Title}}' },
        { type: 'paragraph',  text: 'This agreement is entered into on {{Date}} between {{PartyA}} ("Service Provider") and {{PartyB}} ("Client").' },
        { type: 'heading2',   text: '1. Scope of Services' },
        { type: 'paragraph',  text: '{{Scope}}' },
        { type: 'heading2',   text: '2. Term' },
        { type: 'paragraph',  text: '{{Term}}' },
        { type: 'heading2',   text: '3. Compensation' },
        { type: 'paragraph',  text: '{{Compensation}}' },
        { type: 'heading2',   text: '4. Termination' },
        { type: 'paragraph',  text: '{{Termination}}' },
        { type: 'spacer',     text: '' },
        { type: 'paragraph',  text: 'IN WITNESS WHEREOF, the parties have executed this Agreement as of the date first above written.' },
        { type: 'spacer',     text: '' },
        { type: 'signature',  text: '{{PartyA}}' },
        { type: 'signature',  text: '{{PartyB}}' },
      ],
    },
    essay: {
      name: 'Essay',
      description: 'Title, intro, body, conclusion.',
      placeholders: ['Title','Author','Date','Intro','Body','Conclusion'],
      sections: [
        { type: 'title',      text: '{{Title}}' },
        { type: 'subtitle',   text: '{{Author}} · {{Date}}' },
        { type: 'heading2',   text: 'Introduction' },
        { type: 'paragraph',  text: '{{Intro}}' },
        { type: 'heading2',   text: 'Discussion' },
        { type: 'paragraph',  text: '{{Body}}' },
        { type: 'heading2',   text: 'Conclusion' },
        { type: 'paragraph',  text: '{{Conclusion}}' },
      ],
    },
    'academic-paper': {
      name: 'Academic Paper',
      description: 'Abstract, intro, methods, results, discussion, references.',
      placeholders: ['Title','Authors','Affiliation','Abstract','Introduction','Methods','Results','Discussion','References'],
      sections: [
        { type: 'title',      text: '{{Title}}' },
        { type: 'subtitle',   text: '{{Authors}} — {{Affiliation}}' },
        { type: 'heading2',   text: 'Abstract' },
        { type: 'paragraph',  text: '{{Abstract}}' },
        { type: 'heading1',   text: '1. Introduction' },
        { type: 'paragraph',  text: '{{Introduction}}' },
        { type: 'heading1',   text: '2. Methods' },
        { type: 'paragraph',  text: '{{Methods}}' },
        { type: 'heading1',   text: '3. Results' },
        { type: 'paragraph',  text: '{{Results}}' },
        { type: 'heading1',   text: '4. Discussion' },
        { type: 'paragraph',  text: '{{Discussion}}' },
        { type: 'heading1',   text: 'References' },
        { type: 'paragraph',  text: '{{References}}' },
      ],
    },
    'blog-post': {
      name: 'Blog Post',
      description: 'Title, lede, body sections, takeaway.',
      placeholders: ['Title','Author','Date','Lede','Section1Title','Section1','Section2Title','Section2','Takeaway'],
      sections: [
        { type: 'title',      text: '{{Title}}' },
        { type: 'subtitle',   text: 'By {{Author}} · {{Date}}' },
        { type: 'paragraph',  text: '{{Lede}}' },
        { type: 'heading2',   text: '{{Section1Title}}' },
        { type: 'paragraph',  text: '{{Section1}}' },
        { type: 'heading2',   text: '{{Section2Title}}' },
        { type: 'paragraph',  text: '{{Section2}}' },
        { type: 'heading2',   text: 'Takeaway' },
        { type: 'paragraph',  text: '{{Takeaway}}' },
      ],
    },
    'meeting-notes': {
      name: 'Meeting Notes',
      description: 'Header, attendees, agenda, decisions, action items.',
      placeholders: ['MeetingTitle','Date','Attendees','Agenda','Discussion','Decisions','ActionItems'],
      sections: [
        { type: 'title',      text: '{{MeetingTitle}}' },
        { type: 'subtitle',   text: '{{Date}}' },
        { type: 'heading3',   text: 'Attendees' },
        { type: 'paragraph',  text: '{{Attendees}}' },
        { type: 'heading3',   text: 'Agenda' },
        { type: 'paragraph',  text: '{{Agenda}}' },
        { type: 'heading3',   text: 'Discussion' },
        { type: 'paragraph',  text: '{{Discussion}}' },
        { type: 'heading3',   text: 'Decisions' },
        { type: 'paragraph',  text: '{{Decisions}}' },
        { type: 'heading3',   text: 'Action Items' },
        { type: 'paragraph',  text: '{{ActionItems}}' },
      ],
    },
    'project-plan': {
      name: 'Project Plan',
      description: 'Objective, scope, deliverables, timeline, risks, team.',
      placeholders: ['ProjectName','Owner','Date','Objective','Scope','Deliverables','Timeline','Risks','Team'],
      sections: [
        { type: 'title',      text: '{{ProjectName}}' },
        { type: 'subtitle',   text: 'Owner: {{Owner}} · {{Date}}' },
        { type: 'heading2',   text: 'Objective' },
        { type: 'paragraph',  text: '{{Objective}}' },
        { type: 'heading2',   text: 'Scope' },
        { type: 'paragraph',  text: '{{Scope}}' },
        { type: 'heading2',   text: 'Deliverables' },
        { type: 'paragraph',  text: '{{Deliverables}}' },
        { type: 'heading2',   text: 'Timeline' },
        { type: 'paragraph',  text: '{{Timeline}}' },
        { type: 'heading2',   text: 'Risks' },
        { type: 'paragraph',  text: '{{Risks}}' },
        { type: 'heading2',   text: 'Team' },
        { type: 'paragraph',  text: '{{Team}}' },
      ],
    },
    invoice: {
      name: 'Invoice',
      description: 'Sender, recipient, line items, total, payment terms.',
      placeholders: ['InvoiceNumber','Date','DueDate','FromName','FromAddress','ToName','ToAddress','LineItems','Total','PaymentTerms'],
      sections: [
        { type: 'title',      text: 'Invoice #{{InvoiceNumber}}' },
        { type: 'subtitle',   text: '{{Date}} · Due {{DueDate}}' },
        { type: 'heading3',   text: 'From' },
        { type: 'paragraph',  text: '{{FromName}}\n{{FromAddress}}' },
        { type: 'heading3',   text: 'Bill To' },
        { type: 'paragraph',  text: '{{ToName}}\n{{ToAddress}}' },
        { type: 'heading3',   text: 'Line Items' },
        { type: 'paragraph',  text: '{{LineItems}}' },
        { type: 'heading3',   text: 'Total Due' },
        { type: 'paragraph',  text: '{{Total}}' },
        { type: 'heading3',   text: 'Payment Terms' },
        { type: 'paragraph',  text: '{{PaymentTerms}}' },
      ],
    },
    memo: {
      name: 'Memo',
      description: 'Office memo with header block and body.',
      placeholders: ['To','From','Date','Subject','Body'],
      sections: [
        { type: 'title',      text: 'MEMORANDUM' },
        { type: 'paragraph',  text: 'TO: {{To}}' },
        { type: 'paragraph',  text: 'FROM: {{From}}' },
        { type: 'paragraph',  text: 'DATE: {{Date}}' },
        { type: 'paragraph',  text: 'SUBJECT: {{Subject}}' },
        { type: 'spacer',     text: '' },
        { type: 'paragraph',  text: '{{Body}}' },
      ],
    },
    whitepaper: {
      name: 'Whitepaper',
      description: 'Executive summary, problem, solution, evidence, conclusion.',
      placeholders: ['Title','Author','Date','ExecutiveSummary','Problem','Solution','Evidence','Conclusion'],
      sections: [
        { type: 'title',      text: '{{Title}}' },
        { type: 'subtitle',   text: '{{Author}} · {{Date}}' },
        { type: 'pageBreak',  text: '' },
        { type: 'heading1',   text: 'Executive Summary' },
        { type: 'paragraph',  text: '{{ExecutiveSummary}}' },
        { type: 'heading1',   text: 'The Problem' },
        { type: 'paragraph',  text: '{{Problem}}' },
        { type: 'heading1',   text: 'Our Solution' },
        { type: 'paragraph',  text: '{{Solution}}' },
        { type: 'heading1',   text: 'Evidence & Case Studies' },
        { type: 'paragraph',  text: '{{Evidence}}' },
        { type: 'heading1',   text: 'Conclusion' },
        { type: 'paragraph',  text: '{{Conclusion}}' },
      ],
    },
  };

  /* ───────────────────────────── OOXML SNIPPETS ───────────────────────────── */
  /* Used for things Word.js doesn't expose natively. All snippets must be wrapped
   * with the surrounding <pkg:package> envelope when calling insertOoxml on a
   * Body or Range; the helper functions in execCode() handle that wrapping. */
  const OOXML = {
    // ── Field-code based TOC. Word will offer to update fields on open. ──
    toc: function (opts) {
      opts = opts || {};
      const minLevel = opts.minLevel || 1;
      const maxLevel = opts.maxLevel || 3;
      const title = (opts.title == null) ? 'Table of Contents' : opts.title;
      return `
<w:p>
  <w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr>
  <w:r><w:t>${esc(title)}</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>
  <w:r><w:instrText xml:space="preserve"> TOC \\o "${minLevel}-${maxLevel}" \\h \\z \\u </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType="separate"/></w:r>
  <w:r><w:t>Right-click and choose "Update Field" to populate this table.</w:t></w:r>
  <w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:p>`;
    },

    // ── Section break that switches to N columns for the following content. ──
    columns: function (count) {
      const n = Math.max(1, Math.min(6, parseInt(count, 10) || 2));
      return `<w:p><w:pPr><w:sectPr><w:cols w:num="${n}" w:space="708"/><w:type w:val="continuous"/></w:sectPr></w:pPr></w:p>`;
    },

    pageBreak: `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`,

    sectionBreak: function (kind) {
      const k = ['nextPage','continuous','evenPage','oddPage'].includes(kind) ? kind : 'nextPage';
      return `<w:p><w:pPr><w:sectPr><w:type w:val="${k}"/></w:sectPr></w:pPr></w:p>`;
    },

    landscape: `<w:p><w:pPr><w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:type w:val="continuous"/></w:sectPr></w:pPr></w:p>`,

    portrait:  `<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/><w:type w:val="continuous"/></w:sectPr></w:pPr></w:p>`,

    // ── Diagonal watermark on every page using a WordArt fallback paragraph ──
    watermark: function (text, color) {
      const c = (color || '#C0C0C0').replace('#', '');
      const safe = esc(String(text || 'CONFIDENTIAL').toUpperCase().slice(0, 60));
      return `<w:p>
  <w:pPr>
    <w:jc w:val="center"/>
    <w:rPr><w:rFonts w:ascii="Arial Black" w:hAnsi="Arial Black"/><w:color w:val="${c}"/><w:sz w:val="96"/></w:rPr>
  </w:pPr>
  <w:r>
    <w:rPr><w:rFonts w:ascii="Arial Black" w:hAnsi="Arial Black"/><w:color w:val="${c}"/><w:sz w:val="96"/></w:rPr>
    <w:t>${safe}</w:t>
  </w:r>
</w:p>`;
    },

    // ── Page-number field in current paragraph ──
    pageNumberField: `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`,

    totalPagesField: `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`,

    // ── SDT (Structured Document Tag) content controls ──
    contentControl: function (kind, opts) {
      opts = opts || {};
      const placeholder = esc(opts.placeholder || 'Click to enter…');
      const tag = esc(opts.tag || kind);
      const title = esc(opts.title || tag);

      switch ((kind || 'richText').toLowerCase()) {
        case 'plaintext':
        case 'plain-text':
        case 'plain': return `
<w:sdt>
  <w:sdtPr>
    <w:alias w:val="${title}"/>
    <w:tag w:val="${tag}"/>
    <w:showingPlcHdr/>
    <w:text/>
  </w:sdtPr>
  <w:sdtContent>
    <w:r><w:t>${placeholder}</w:t></w:r>
  </w:sdtContent>
</w:sdt>`;
        case 'checkbox':
        case 'check': return `
<w:sdt>
  <w:sdtPr>
    <w:alias w:val="${title}"/>
    <w:tag w:val="${tag}"/>
    <w14:checkbox xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
      <w14:checked w14:val="${opts.checked ? '1' : '0'}"/>
      <w14:checkedState w14:val="2612" w14:font="MS Gothic"/>
      <w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/>
    </w14:checkbox>
  </w:sdtPr>
  <w:sdtContent>
    <w:r><w:t>${opts.checked ? '☒' : '☐'}</w:t></w:r>
  </w:sdtContent>
</w:sdt>`;
        case 'dropdown':
        case 'combobox':
        case 'combo': {
          const items = Array.isArray(opts.items) && opts.items.length
            ? opts.items.map(it => `      <w:listItem w:displayText="${esc(it)}" w:value="${esc(it)}"/>`).join('\n')
            : `      <w:listItem w:displayText="Choose an item" w:value=""/>`;
          const isCombo = kind.toLowerCase().startsWith('combo');
          return `
<w:sdt>
  <w:sdtPr>
    <w:alias w:val="${title}"/>
    <w:tag w:val="${tag}"/>
    <w:${isCombo ? 'comboBox' : 'dropDownList'}>
${items}
    </w:${isCombo ? 'comboBox' : 'dropDownList'}>
  </w:sdtPr>
  <w:sdtContent>
    <w:r><w:t>${placeholder}</w:t></w:r>
  </w:sdtContent>
</w:sdt>`;
        }
        case 'date':
        case 'datepicker':
        case 'date-picker': return `
<w:sdt>
  <w:sdtPr>
    <w:alias w:val="${title}"/>
    <w:tag w:val="${tag}"/>
    <w:date>
      <w:dateFormat w:val="${esc(opts.format || 'MMMM d, yyyy')}"/>
      <w:lid w:val="en-US"/>
      <w:storeMappedDataAs w:val="dateTime"/>
      <w:calendar w:val="gregorian"/>
    </w:date>
  </w:sdtPr>
  <w:sdtContent>
    <w:r><w:t>${placeholder}</w:t></w:r>
  </w:sdtContent>
</w:sdt>`;
        case 'richtext':
        case 'rich-text':
        default: return `
<w:sdt>
  <w:sdtPr>
    <w:alias w:val="${title}"/>
    <w:tag w:val="${tag}"/>
    <w:showingPlcHdr/>
  </w:sdtPr>
  <w:sdtContent>
    <w:p><w:r><w:t>${placeholder}</w:t></w:r></w:p>
  </w:sdtContent>
</w:sdt>`;
      }
    },
  };

  /* ───────────────────────────── pkg envelope helper ───────────────────────────── */
  /* Word.js's insertOoxml() accepts a flat XML fragment OR a full WordprocessingML
   * package. We wrap our snippets when they need to include extra namespaces. */
  function packageOoxml(bodyXml) {
    return `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
          ${bodyXml}
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
  }

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  /* ───────────────────────────── LaTeX → OMML ───────────────────────────── */
  /* Tiny converter that handles the most common patterns: fractions, square
   * roots, super/subscripts, Greek letters, basic operators. Falls back to a
   * plain-text run wrapped in an OMML <m:r> for anything it can't parse. */
  function latexToOmml(latex) {
    if (!latex) return '';
    let s = String(latex).trim();
    // Strip common $...$ wrappers
    s = s.replace(/^\$+|\$+$/g, '').trim();

    const GREEK = {
      alpha:'α', beta:'β', gamma:'γ', delta:'δ', epsilon:'ε', zeta:'ζ', eta:'η',
      theta:'θ', iota:'ι', kappa:'κ', lambda:'λ', mu:'μ', nu:'ν', xi:'ξ',
      pi:'π', rho:'ρ', sigma:'σ', tau:'τ', upsilon:'υ', phi:'φ', chi:'χ',
      psi:'ψ', omega:'ω',
      Alpha:'Α', Beta:'Β', Gamma:'Γ', Delta:'Δ', Epsilon:'Ε', Zeta:'Ζ', Eta:'Η',
      Theta:'Θ', Iota:'Ι', Kappa:'Κ', Lambda:'Λ', Mu:'Μ', Nu:'Ν', Xi:'Ξ',
      Pi:'Π', Rho:'Ρ', Sigma:'Σ', Tau:'Τ', Upsilon:'Υ', Phi:'Φ', Chi:'Χ',
      Psi:'Ψ', Omega:'Ω',
      infty:'∞', cdot:'·', times:'×', div:'÷', pm:'±', leq:'≤', geq:'≥',
      neq:'≠', approx:'≈', sum:'∑', prod:'∏', int:'∫', partial:'∂',
      rightarrow:'→', leftarrow:'←', Rightarrow:'⇒', Leftarrow:'⇐',
    };

    // Replace Greek letters & operators
    s = s.replace(/\\([A-Za-z]+)/g, (m, name) => GREEK[name] != null ? GREEK[name] : m);

    // Helper to emit a run
    const run = txt => `<m:r><m:t>${esc(txt)}</m:t></m:r>`;

    // Convert \frac{a}{b}
    s = s.replace(/\\frac\{([^{}]*)\}\{([^{}]*)\}/g,
      (m, a, b) => `«FRAC|${a}|${b}»`);
    // Convert \sqrt{a}
    s = s.replace(/\\sqrt\{([^{}]*)\}/g, (m, a) => `«SQRT|${a}»`);
    // Convert x^{...} and x_{...}
    s = s.replace(/([A-Za-z0-9\)\]])\^\{([^{}]*)\}/g, (m, a, b) => `«SUP|${a}|${b}»`);
    s = s.replace(/([A-Za-z0-9\)\]])\^([A-Za-z0-9])/g, (m, a, b) => `«SUP|${a}|${b}»`);
    s = s.replace(/([A-Za-z0-9\)\]])_\{([^{}]*)\}/g, (m, a, b) => `«SUB|${a}|${b}»`);
    s = s.replace(/([A-Za-z0-9\)\]])_([A-Za-z0-9])/g, (m, a, b) => `«SUB|${a}|${b}»`);

    // Build OMML from token stream
    const tokens = s.split(/(«[^»]+»)/g).filter(Boolean);
    const parts = [];
    for (const tok of tokens) {
      const m = /^«(FRAC|SQRT|SUP|SUB)\|([^|»]*)(?:\|([^|»]*))?»$/.exec(tok);
      if (!m) { parts.push(run(tok)); continue; }
      const [, kind, a, b] = m;
      if (kind === 'FRAC') {
        parts.push(`<m:f><m:num>${run(a)}</m:num><m:den>${run(b)}</m:den></m:f>`);
      } else if (kind === 'SQRT') {
        parts.push(`<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${run(a)}</m:e></m:rad>`);
      } else if (kind === 'SUP') {
        parts.push(`<m:sSup><m:e>${run(a)}</m:e><m:sup>${run(b)}</m:sup></m:sSup>`);
      } else if (kind === 'SUB') {
        parts.push(`<m:sSub><m:e>${run(a)}</m:e><m:sub>${run(b)}</m:sub></m:sub></m:sSub>`.replace('</m:sub></m:sSub>','</m:sub></m:sSub>'));
      }
    }
    return `<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:oMath>${parts.join('')}</m:oMath>
</m:oMathPara>`;
  }

  /* ───────────────────────────── Flesch readability ───────────────────────────── */
  function syllableCount(word) {
    word = String(word || '').toLowerCase().replace(/[^a-z]/g, '');
    if (!word) return 0;
    if (word.length <= 3) return 1;
    word = word.replace(/(?:[^laeiouy]es|ed|[^laeiouy]e)$/, '').replace(/^y/, '');
    const m = word.match(/[aeiouy]{1,2}/g);
    return m ? m.length : 1;
  }
  function readability(text) {
    const t = String(text || '').trim();
    if (!t) return { wordCount: 0, sentenceCount: 0, syllableCount: 0, flesch: null, grade: null };
    const sentences = (t.match(/[^.!?]+[.!?]+/g) || [t]).filter(s => s.trim());
    const words = t.match(/\b[\w']+\b/g) || [];
    const wc = words.length, sc = Math.max(1, sentences.length);
    const sylTot = words.reduce((s, w) => s + syllableCount(w), 0);
    if (!wc) return { wordCount: 0, sentenceCount: 0, syllableCount: 0, flesch: null, grade: null };
    const flesch = 206.835 - 1.015 * (wc / sc) - 84.6 * (sylTot / wc);
    const grade  = 0.39 * (wc / sc) + 11.8 * (sylTot / wc) - 15.59;
    return {
      wordCount: wc,
      sentenceCount: sc,
      syllableCount: sylTot,
      flesch: Math.round(flesch * 10) / 10,
      grade:  Math.round(grade  * 10) / 10,
    };
  }

  /* ───────────────────────────── EXPORTS ───────────────────────────── */
  window.WordHelpers = {
    WORD_SAFE_FONTS,
    THEME_SEEDS,
    TEMPLATES,
    OOXML,
    packageOoxml,
    sanitizeFont,
    wcagContrast,
    latexToOmml,
    readability,
    esc,
  };
})();
