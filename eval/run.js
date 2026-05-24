#!/usr/bin/env node
/**
 * Word AI Eval Runner
 *
 * - Scores AI responses 0–100 using an LLM judge
 * - Tracks per-category levels in progress.json
 * - When a category averages >= 95, auto-generates harder cases for it
 * - When overall avg >= 95, also generates a brand-new Word domain category
 * - When overall avg >= 90, generates a level-2 category (easier threshold)
 * - When a category is stuck on the same level for too long, generates easier cases
 * - Hard budget limit: $0.50 per session
 */
require('dotenv').config({ path: require('path').join(__dirname, '../.env') });

const fs   = require('fs');
const path = require('path');
const fetch = require('node-fetch');

const OPENROUTER_KEY   = process.env.OPENROUTER_KEY;
const MODEL            = 'meta-llama/llama-3.3-70b-instruct';
const JUDGE_MODEL      = 'meta-llama/llama-3.3-70b-instruct';
const COST_PER_1M_IN   = 0.07;
const COST_PER_1M_OUT  = 0.30;
const BUDGET_USD       = 0.04;
const MASTERY_THRESHOLD      = 95;   // per-category avg that promotes to a harder level
const SUITE_HEALTH_THRESHOLD = 80;   // overall avg required before ANY case generation runs
const NEW_CAT_THRESHOLD      = 97;   // overall avg AND every-category-mastered required for new categories
const STUCK_RUNS_THRESHOLD   = 3;
const HARDER_CASES_PER_GEN   = 2;    // new harder cases per mastered category (was 4 — too eager)
const NEW_CAT_CASES_PER_GEN  = 1;    // new cases per brand-new category (was 2 — too eager)
const EASIER_CASES_PER_GEN   = 2;    // recovery cases for stuck categories (was 3 — too eager)

const PATHS = {
  cases:     path.join(__dirname, 'cases.json'),
  generated: path.join(__dirname, 'generated-cases.json'),
  progress:  path.join(__dirname, 'progress.json'),
  results:   path.join(__dirname, 'results'),
};

if (!fs.existsSync(PATHS.results)) fs.mkdirSync(PATHS.results, { recursive: true });

// ── Cost tracking ─────────────────────────────────────────────────────────────
let totalCostUSD = 0;
function trackCost(inTok, outTok) {
  const c = (inTok / 1e6) * COST_PER_1M_IN + (outTok / 1e6) * COST_PER_1M_OUT;
  totalCostUSD += c;
  if (totalCostUSD > BUDGET_USD) throw new Error(`BUDGET_EXCEEDED: $${totalCostUSD.toFixed(4)} > $${BUDGET_USD}`);
}

// ── OpenRouter call ───────────────────────────────────────────────────────────
async function callLLM(messages, model = MODEL, maxTokens = 1024) {
  const res = await fetch('https://openrouter.ai/api/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type':  'application/json',
      'Authorization': `Bearer ${OPENROUTER_KEY}`,
      'HTTP-Referer':  'https://localhost:3000',
      'X-Title':       'Word AI Eval',
    },
    body: JSON.stringify({ model, messages, max_tokens: maxTokens, temperature: 0.1 }),
  });
  const data = await res.json();
  if (data.error) throw new Error(JSON.stringify(data.error));
  const u = data.usage || {};
  trackCost(u.prompt_tokens || 600, u.completion_tokens || 250);
  return data.choices[0].message.content;
}

// ── Load files ────────────────────────────────────────────────────────────────
function loadCases() {
  const base = JSON.parse(fs.readFileSync(PATHS.cases, 'utf8'));
  const gen  = fs.existsSync(PATHS.generated)
    ? JSON.parse(fs.readFileSync(PATHS.generated, 'utf8'))
    : [];
  return [...base, ...gen];
}
function loadProgress() {
  if (!fs.existsSync(PATHS.progress)) return {};
  return JSON.parse(fs.readFileSync(PATHS.progress, 'utf8'));
}
function saveProgress(p) { fs.writeFileSync(PATHS.progress, JSON.stringify(p, null, 2)); }
function loadLastResults() {
  const files = fs.readdirSync(PATHS.results).filter(f => f.endsWith('.json')).sort();
  if (!files.length) return null;
  return JSON.parse(fs.readFileSync(path.join(PATHS.results, files[files.length - 1]), 'utf8'));
}

// ── System prompt extractor ───────────────────────────────────────────────────
function loadSystemPrompt() {
  const src       = fs.readFileSync(path.join(__dirname, '../server.js'), 'utf8');
  const startMark = 'const SYSTEM_PROMPT = `';
  const startIdx  = src.indexOf(startMark);
  if (startIdx === -1) throw new Error('Cannot find SYSTEM_PROMPT in server.js');
  const from = startIdx + startMark.length;
  let i = from;
  while (i < src.length) {
    i = src.indexOf('`', i);
    if (i === -1) throw new Error('Cannot find end of SYSTEM_PROMPT');
    const after = src.slice(i + 1).trimStart();
    if (after.startsWith('+') || after.startsWith(';')) break;
    i++;
  }
  return src.slice(from, i);
}

// ── Extract code / web-search blocks ──────────────────────────────────────────
function extractCode(text) {
  const m = text.match(/CODE_JS::([\s\S]*?)::END_CODE/);
  return m ? m[1].trim() : null;
}
function extractWebSearch(text) {
  const m = text.match(/NEEDS_WEB_SEARCH::([\s\S]*?)::END/);
  if (!m) return null;
  try { return JSON.parse(m[1].trim()); } catch { return { raw: m[1].trim() }; }
}

// ── Pattern checks ────────────────────────────────────────────────────────────
function checkPatterns(text, code, required, forbidden) {
  const haystack = (code || '') + '\n' + (text || '');
  const missing  = (required || []).filter(p => !haystack.includes(p));
  const badFound = (forbidden || []).filter(p => haystack.includes(p));
  return { passed: !missing.length && !badFound.length, missing, badFound };
}

// ── LLM judge ────────────────────────────────────────────────────────────────
async function judge(tc, aiResponse, code) {
  const isQuestion = tc.mustHaveCode === false;
  const resp = await callLLM([
    { role: 'system', content: 'You are a strict expert evaluator of Office JavaScript API code for Microsoft Word. You score precisely — never give round multiples of 10 unless exactly right.' },
    {
      role: 'user',
      content: `Score this AI assistant response on FOUR dimensions, each out of 25 points.

USER REQUEST: "${tc.prompt}"
DOC CONTEXT: ${tc.docData}
${tc.note ? `SPECIAL NOTE: ${tc.note}` : ''}

AI RESPONSE:
${aiResponse}

${code ? `GENERATED CODE:\n${code}` : 'NO CODE WAS GENERATED.'}

${isQuestion
  ? `This is a QUESTION OR REFUSAL — no code expected. Score:
A) ACCURACY (0-25): Is the answer factually correct and complete? For refusals, does it correctly identify what Word Office.js cannot do (PDF export, print API, chart creation, slideshow control)?
B) CLARITY (0-25): Is it clear and easy to understand?
C) RELEVANCE (0-25): Does it directly address the question without padding?
D) BONUS (0-25): Extra credit for offering workarounds, suggesting Settings options, or noting Word-specific caveats (e.g. TOC refresh, wildcards vs regex).`
  : `Score each dimension carefully. Deduct proportionally for each flaw:
A) API_CORRECTNESS (0-25): Does the code use real Word Office.js APIs? HEAVILY penalise hallucinated methods: doc.exportToPdf(), doc.print(), doc.charts.add(), doc.smartArt.*, doc.startSlideShow(). Real APIs: Word.run, doc.body, body.paragraphs, paragraph.insertText, paragraph.styleBuiltIn, body.insertTable, table.getCell, range.insertFootnote, range.insertComment, body.search, range.font.*, range.insertOoxml, range.insertContentControl, doc.changeTrackingMode. Penalize JS regex inside body.search (use matchWildcards Word patterns) unless the user explicitly asked for JS regex post-processing.
B) COMPLETENESS (0-25): Does the code fully address the request? Partial solutions lose points.
C) WOULD_IT_WORK (0-25): Would this code actually execute in Word without runtime errors? Check load/sync order, variable scope, correct method signatures, range vs body distinctions.
D) APPROACH (0-25): Does it use best practices and the available helpers (addHeading, addParagraph, addList, applyStyle, replaceText, insertTable, insertImage, insertFootnote, insertComment, toggleTrackChanges, insertTableOfContents, applyTheme, designTheme, applyTemplate, insertContentControl, mailMergeReplace, insertEquation) instead of re-implementing them manually? For images, MUST use insertImage — never hallucinate a URL. For themes, MUST use designTheme()+applyTheme() — there are NO preset theme names.`}

Deduct points specifically for each issue you identify. Do NOT round to multiples of 5 — use precise values like 18, 22, 7.

Reply with exactly this format:
A: <number>
B: <number>
C: <number>
D: <number>
REASON: <one sentence summarising the main flaw or strength>`,
    },
  ], JUDGE_MODEL, 150);

  const aM = resp.match(/^A:\s*(\d+)/im);
  const bM = resp.match(/^B:\s*(\d+)/im);
  const cM = resp.match(/^C:\s*(\d+)/im);
  const dM = resp.match(/^D:\s*(\d+)/im);
  const reasonM = resp.match(/REASON:\s*(.+)/i);

  const a = aM ? Math.min(25, parseInt(aM[1])) : 0;
  const b = bM ? Math.min(25, parseInt(bM[1])) : 0;
  const c = cM ? Math.min(25, parseInt(cM[1])) : 0;
  const d = dM ? Math.min(25, parseInt(dM[1])) : 0;

  return {
    score:  a + b + c + d,
    reason: reasonM ? reasonM[1].trim() : resp.slice(0, 120),
    breakdown: { a, b, c, d },
  };
}

// ── Auto-patch: generate system prompt improvements ────────────────────────────
async function generateFix(failingResults, currentImprovements) {
  const cases = failingResults.map(r =>
    `  [${r.id}] score=${r.score}/100\n  prompt: "${r.prompt}"\n  reason: ${r.reason}\n  code:\n${(r.generatedCode || 'NONE').split('\n').map(l => '    ' + l).join('\n')}`
  ).join('\n\n');

  const resp = await callLLM([
    { role: 'system', content: 'You are an expert at improving system prompts for a Word AI assistant that generates Office JavaScript code.' },
    {
      role: 'user',
      content: `The following test cases are FAILING (score < 70). Analyse the errors and generate new rules or examples to add to the system prompt so the AI handles these cases correctly.

FAILING CASES:
${cases}

EXISTING IMPROVEMENTS ALREADY IN PROMPT:
${currentImprovements || '(none)'}

Constraints:
- Only add rules or examples that would fix the specific failures above
- Use the same format as the existing prompt (plain text rules or CODE_JS examples)
- Use the available helpers (addHeading, addParagraph, addList, applyStyle, replaceText, insertTable, insertImage, insertFootnote, insertComment, toggleTrackChanges, insertTableOfContents, applyTheme, designTheme, applyTemplate, insertContentControl, mailMergeReplace, insertEquation)
- Never propose APIs that do not exist (doc.exportToPdf, doc.charts.add, doc.smartArt.*, doc.startSlideShow, doc.print, slide.* — this is Word not PowerPoint)
- Be concise — 3–8 lines max
- Do NOT duplicate existing rules

Return ONLY the new text to add. No explanation, no preamble.`,
    },
  ], MODEL, 512);

  return resp.trim();
}

function applyImprovements(newText) {
  const impFile = path.join(__dirname, 'improvements.txt');
  const oldText = fs.existsSync(impFile) ? fs.readFileSync(impFile, 'utf8').trim() : '';
  fs.writeFileSync(impFile, newText + '\n');
  console.log('\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('  SYSTEM PROMPT PATCHED (eval/improvements.txt)');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  if (oldText) for (const line of oldText.split('\n')) console.log(`  - ${line}`);
  else console.log('  BEFORE: (empty)');
  console.log('AFTER:');
  for (const line of newText.split('\n')) console.log(`  + ${line}`);
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n');
}

async function generateNewCategory(existingCategories) {
  const knownList = existingCategories.join(', ');
  const resp = await callLLM([
    { role: 'system', content: 'You generate JSON test cases for a Word AI assistant evaluation suite. Return only valid JSON.' },
    {
      role: 'user',
      content: `The eval suite already covers: ${knownList}.

Pick ONE new Microsoft Word feature category not in that list (e.g. line-spacing, indentation, page-numbers, hyperlink-insert, bookmark, cross-reference, drop-cap, columns-balanced, etc.).

Generate exactly ${NEW_CAT_CASES_PER_GEN} level-1 test case(s) for it — simple, realistic prompts a real user would ask.
Each case must be solvable via Word Office JavaScript API + the available helpers (addHeading, addParagraph, addList, applyStyle, replaceText, insertTable, insertImage, insertFootnote, insertComment, toggleTrackChanges, insertTableOfContents, applyTheme, designTheme, applyTemplate, insertContentControl, mailMergeReplace, insertEquation). No PDF export, no slideshow, no chart creation, no print API.

Return ONLY a valid JSON object with this shape:
{
  "category": "<new category name>",
  "cases": [
    {
      "id": "gen-L1-<category-slug>-001",
      "level": 1,
      "category": "<new category name>",
      "prompt": "...",
      "docData": "Section 1: HEADING1=\\"...\\" | TEXT=\\"...\\"\\nSection 2: ...",
      "mustHaveCode": true,
      "requiredPatterns": ["someRealHelperOrApi"],
      "forbiddenPatterns": []
    }
  ]
}`,
    },
  ], MODEL, 1024);

  try {
    const jsonMatch = resp.match(/\{[\s\S]*\}/);
    if (!jsonMatch) throw new Error('No JSON found');
    const parsed = JSON.parse(jsonMatch[0]);
    if (!parsed.category || !Array.isArray(parsed.cases)) throw new Error('Bad shape');
    return parsed;
  } catch (err) {
    console.warn(`  Failed to parse new category: ${err.message}`);
    return null;
  }
}

// ── Run one test ──────────────────────────────────────────────────────────────
async function runCase(tc, systemPrompt) {
  const messages = [
    { role: 'system', content: systemPrompt + '\n\nRUNTIME STATE:\n<ASSET_PACK_READY categories="people,technology,finance,nature,charts-diagrams,abstract-backgrounds,transport,food,education,healthcare,cities,office-workplace,travel,sports,document-icons,watermarks,letterhead-graphics,page-borders"/>\n<WEB_SEARCH_MODE value="off"/>' },
    { role: 'user',      content: `Document state:\n\n${tc.docData}` },
    { role: 'assistant', content: 'I can see the document. What would you like me to do?' },
    { role: 'user',      content: tc.prompt + '\n\n[REMINDER: If making changes, output a CODE_JS block. Do not skip it.]' },
  ];

  let aiResponse, code, patterns, verdict;
  try {
    aiResponse = await callLLM(messages, MODEL, 1024);
    code       = extractCode(aiResponse);
    patterns   = checkPatterns(aiResponse, code, tc.requiredPatterns, tc.forbiddenPatterns);
    verdict    = await judge(tc, aiResponse, code);
  } catch (err) {
    if (err.message.startsWith('BUDGET_EXCEEDED')) throw err;
    aiResponse = `ERROR: ${err.message}`;
    code       = null;
    patterns   = { passed: false, missing: tc.requiredPatterns || [], badFound: [] };
    verdict    = { score: 0, reason: `Runtime error: ${err.message}` };
  }

  return {
    id:              tc.id,
    level:           tc.level,
    category:        tc.category,
    prompt:          tc.prompt,
    score:           verdict.score,
    reason:          verdict.reason,
    hasCode:         code !== null,
    codeExpected:    tc.mustHaveCode !== false,
    patternPass:     patterns.passed,
    missingPatterns: patterns.missing,
    badFound:        patterns.badFound,
    generatedCode:   code,
  };
}

async function generateHarderCases(category, currentLevel, existingCases, allCategories) {
  const nextLevel     = currentLevel + 1;
  const masteredCases = existingCases
    .filter(c => c.category === category)
    .map(c => `  - [${c.id}] "${c.prompt}"`)
    .join('\n');

  const isAllMastered = allCategories.every(cat => cat.mastered);
  const newCatHint    = isAllMastered
    ? `\nSince ALL categories are mastered, also include 1–2 cases for brand-new Word domains not yet tested.`
    : '';

  const prompt = `You generate test cases for a Word AI assistant evaluation suite.

The AI just scored 95+/100 on ALL these level ${currentLevel} "${category}" cases:
${masteredCases}

Generate exactly ${HARDER_CASES_PER_GEN} NEW test cases at level ${nextLevel} for the "${category}" category that are SIGNIFICANTLY HARDER. They must:
1. Be in the same category but test more complex, realistic, or edge-case scenarios
2. Not repeat any existing test idea
3. Be solvable via Word Office JavaScript API + helpers — no PDF export, slideshow, chart creation, print API
4. Include realistic document state
${newCatHint}

Return ONLY a valid JSON array, no explanation. Each object must have:
{
  "id": "gen-L${nextLevel}-${category.replace(/\s+/g, '-')}-NNN",
  "level": ${nextLevel},
  "category": "${category}",
  "prompt": "...",
  "docData": "Section 1: HEADING1=\\"...\\" | TEXT=\\"...\\"",
  "mustHaveCode": true,
  "requiredPatterns": ["pattern1"],
  "forbiddenPatterns": []
}`;

  console.log(`\n  Generating level ${nextLevel} cases for "${category}"...`);
  const raw = await callLLM([
    { role: 'system', content: 'You generate JSON test cases. Return only valid JSON arrays.' },
    { role: 'user',   content: prompt },
  ], MODEL, 2048);

  try {
    const jsonMatch = raw.match(/\[[\s\S]*\]/);
    if (!jsonMatch) throw new Error('No JSON array found in response');
    const newCases = JSON.parse(jsonMatch[0]);
    console.log(`  Generated ${newCases.length} new cases for "${category}" at level ${nextLevel}`);
    return newCases;
  } catch (err) {
    console.warn(`  Failed to parse generated cases for "${category}": ${err.message}`);
    return [];
  }
}

async function generateEasierCases(category, currentLevel, existingCases) {
  const stuckCases = existingCases
    .filter(c => c.category === category && c.level === currentLevel)
    .map(c => `  - [${c.id}] "${c.prompt}"`)
    .join('\n');

  const prompt = `You generate test cases for a Word AI assistant evaluation suite.

The AI has been STUCK on level ${currentLevel} of the "${category}" category across multiple runs.
These are the current cases it keeps failing:
${stuckCases}

Generate exactly ${EASIER_CASES_PER_GEN} NEW test cases at level ${currentLevel} for "${category}" that are EASIER and more approachable. They must:
1. Test the same core category skill but with simpler scenarios
2. Avoid the tricky edge-cases the AI is clearly struggling with
3. Be solvable via Word Office JavaScript API + helpers
4. Include simple, realistic document state

Return ONLY a valid JSON array, no explanation. Each object must have:
{
  "id": "gen-L${currentLevel}-${category.replace(/\s+/g, '-')}-easy-NNN",
  "level": ${currentLevel},
  "category": "${category}",
  "prompt": "...",
  "docData": "Section 1: HEADING1=\\"...\\" | TEXT=\\"...\\"",
  "mustHaveCode": true,
  "requiredPatterns": ["pattern1"],
  "forbiddenPatterns": []
}`;

  console.log(`\n  Generating easier cases for stuck category "${category}" (level ${currentLevel})...`);
  const raw = await callLLM([
    { role: 'system', content: 'You generate JSON test cases. Return only valid JSON arrays.' },
    { role: 'user',   content: prompt },
  ], MODEL, 2048);

  try {
    const jsonMatch = raw.match(/\[[\s\S]*\]/);
    if (!jsonMatch) throw new Error('No JSON array found in response');
    const newCases = JSON.parse(jsonMatch[0]);
    console.log(`  Generated ${newCases.length} easier cases for "${category}"`);
    return newCases;
  } catch (err) {
    console.warn(`  Failed to parse easier cases for "${category}": ${err.message}`);
    return [];
  }
}

function writeMarkdownSummary(results, prev, timestamp) {
  const prevById = {};
  if (prev) for (const r of prev.results) prevById[r.id] = r;

  const byCategory = {};
  for (const r of results) {
    if (!byCategory[r.category]) byCategory[r.category] = [];
    byCategory[r.category].push(r);
  }
  const overall = results.reduce((s, r) => s + r.score, 0) / results.length;
  const prevOverall = prev ? prev.results.reduce((s, r) => s + r.score, 0) / prev.results.length : null;

  const lines = [];
  lines.push(`# Word AI — Eval Results`);
  lines.push(`**Last run:** ${timestamp.replace('T', ' ').replace(/-/g, (m, o) => o < 10 ? '-' : ':')}  `);
  lines.push(`**Overall: ${overall.toFixed(1)}/100**${prevOverall !== null ? `  (prev: ${prevOverall.toFixed(1)})` : ''}  `);
  lines.push(`**Model:** ${MODEL}\n`);
  lines.push(`## Scores by category\n`);
  lines.push(`| Category | Score | Trend | Cases |`);
  lines.push(`|---|---|---|---|`);
  for (const [cat, items] of Object.entries(byCategory)) {
    const avg = items.reduce((s, r) => s + r.score, 0) / items.length;
    const prevAvg = prev ? items.map(r => prevById[r.id]?.score ?? r.score).reduce((a, b) => a + b, 0) / items.length : null;
    const delta = prevAvg !== null ? avg - prevAvg : null;
    const trend = delta === null ? '—' : delta > 2 ? `↑ +${delta.toFixed(1)}` : delta < -2 ? `↓ ${delta.toFixed(1)}` : `→`;
    const mastMark = avg >= MASTERY_THRESHOLD ? ' ✓' : '';
    lines.push(`| ${cat}${mastMark} | ${avg.toFixed(1)}/100 | ${trend} | ${items.length} |`);
  }
  lines.push(`\n## All test cases\n`);
  lines.push(`| ID | Score | Pass | Reason |`);
  lines.push(`|---|---|---|---|`);
  for (const r of results) {
    const prevR = prevById[r.id];
    const delta = prevR ? r.score - prevR.score : null;
    const deltaStr = delta !== null && Math.abs(delta) >= 1 ? ` (${delta > 0 ? '+' : ''}${delta})` : '';
    const warn = r.score < 60 ? ' ⚠️' : r.score >= 95 ? ' ✅' : '';
    lines.push(`| ${r.id} | ${r.score}${deltaStr}${warn} | ${r.patternPass ? '✓' : '✗'} | ${r.reason} |`);
  }
  const failing = results.filter(r => r.score < 60);
  if (failing.length) {
    lines.push(`\n## ⚠️ Needs attention\n`);
    for (const r of failing) {
      lines.push(`**[${r.id}]** score=${r.score} — ${r.reason}`);
      if (r.missingPatterns.length) lines.push(`- missing patterns: \`${r.missingPatterns.join('`, `')}\``);
      if (r.badFound?.length) lines.push(`- forbidden patterns found: \`${r.badFound.join('`, `')}\``);
      lines.push('');
    }
  }
  fs.writeFileSync(path.join(__dirname, 'RESULTS.md'), lines.join('\n'));
}

function printResults(results, prev) {
  const prevById = {};
  if (prev) for (const r of prev.results) prevById[r.id] = r;
  const byCategory = {};
  for (const r of results) {
    if (!byCategory[r.category]) byCategory[r.category] = [];
    byCategory[r.category].push(r);
  }
  console.log('\n════════════════════════════════════════════════════');
  console.log('  WORD AI EVAL RESULTS');
  console.log('════════════════════════════════════════════════════\n');
  for (const [cat, items] of Object.entries(byCategory)) {
    const avg     = items.reduce((s, r) => s + r.score, 0) / items.length;
    const prevAvg = prev ? items.map(r => prevById[r.id]?.score ?? r.score).reduce((a, b) => a + b, 0) / items.length : null;
    const delta   = prevAvg !== null ? avg - prevAvg : null;
    const trend   = delta === null ? '' : delta > 2 ? ' ↑' : delta < -2 ? ' ↓' : ' →';
    const mastMark = avg >= MASTERY_THRESHOLD ? ' ✓ MASTERED' : '';
    console.log(`${cat.toUpperCase()} — ${avg.toFixed(1)}/100${trend}${mastMark}`);
    for (const r of items) {
      const d    = prevById[r.id] ? r.score - prevById[r.id].score : null;
      const diff = d !== null && Math.abs(d) >= 1 ? ` (${d > 0 ? '+' : ''}${d})` : '';
      const warn = r.score < 60 ? ' ⚠' : '';
      const bar  = '█'.repeat(Math.round(r.score / 10)) + '░'.repeat(10 - Math.round(r.score / 10));
      console.log(`  ${r.id.padEnd(40)} ${bar} ${String(r.score).padStart(3)}${diff}${warn}`);
      if (r.score < 70) console.log(`    → ${r.reason}`);
      if (r.missingPatterns.length) console.log(`    missing: ${r.missingPatterns.join(', ')}`);
    }
    console.log();
  }
  const overall     = results.reduce((s, r) => s + r.score, 0) / results.length;
  const prevOverall = prev ? prev.results.reduce((s, r) => s + r.score, 0) / prev.results.length : null;
  console.log(`OVERALL: ${overall.toFixed(1)}/100${prevOverall !== null ? `  (prev: ${prevOverall.toFixed(1)})` : ''}`);
  console.log(`SESSION COST: $${totalCostUSD.toFixed(4)}`);
  console.log('════════════════════════════════════════════════════\n');
}

// ── Main ──────────────────────────────────────────────────────────────────────
async function main() {
  if (!OPENROUTER_KEY) { console.error('OPENROUTER_KEY not set'); process.exit(1); }
  const allCases     = loadCases();
  const systemPrompt = loadSystemPrompt();
  const prev         = loadLastResults();
  const progress     = loadProgress();
  console.log(`\nRunning ${allCases.length} Word eval cases  |  model: ${MODEL}  |  budget: $${BUDGET_USD}\n`);

  const results = [];
  for (const tc of allCases) {
    process.stdout.write(`  ${tc.id.padEnd(40)} `);
    try {
      const r = await runCase(tc, systemPrompt);
      results.push(r);
      const bar = '█'.repeat(Math.round(r.score / 10)) + '░'.repeat(10 - Math.round(r.score / 10));
      console.log(`${bar} ${r.score}/100`);
    } catch (err) {
      if (err.message.startsWith('BUDGET_EXCEEDED')) {
        console.log('\n⚠ Budget limit hit — saving partial results and stopping.');
        break;
      }
      console.log(`ERROR: ${err.message}`);
    }
  }

  const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const outPath   = path.join(PATHS.results, `${timestamp}.json`);
  fs.writeFileSync(outPath, JSON.stringify({ timestamp, model: MODEL, totalCostUSD, results }, null, 2));
  printResults(results, prev);
  writeMarkdownSummary(results, prev, timestamp);

  // Per-category mastery + generation
  const byCategory = {};
  for (const r of results) {
    if (!byCategory[r.category]) byCategory[r.category] = [];
    byCategory[r.category].push(r);
  }
  let newCasesGenerated = [];
  const categoryStatus  = Object.entries(byCategory).map(([cat, items]) => {
    const currentLevel = progress[cat]?.level ?? 1;
    const currentItems = items.filter(r => r.level === currentLevel);
    const avg          = currentItems.length ? currentItems.reduce((s, r) => s + r.score, 0) / currentItems.length : 0;
    return { cat, avg, currentLevel, mastered: avg >= MASTERY_THRESHOLD };
  });

  // Suite-health gate — don't grow the eval suite while the model still flunks
  // the basics. Generating "harder" cases for one mastered category while
  // overall avg is 60% just bloats the suite with noise. Only let the
  // generator run when the suite as a whole is performing well.
  const overallAvg = results.length ? results.reduce((s, r) => s + r.score, 0) / results.length : 0;
  const suiteHealthy = overallAvg >= SUITE_HEALTH_THRESHOLD;
  const everyCategoryMastered = categoryStatus.length > 0 && categoryStatus.every(c => c.mastered);

  if (!suiteHealthy) {
    console.log(`\n🚧 Overall avg ${overallAvg.toFixed(1)} < ${SUITE_HEALTH_THRESHOLD} — pausing all case generation until the suite stabilises.`);
  }

  for (const { cat, avg, currentLevel, mastered } of categoryStatus) {
    if (!progress[cat]) progress[cat] = { level: 1, masteredAt: null, runsAtLevel: 0 };
    progress[cat].lastScore  = Math.round(avg * 10) / 10;
    progress[cat].lastRun    = timestamp;
    if (mastered && suiteHealthy) {
      console.log(`✓ "${cat}" mastered at level ${currentLevel} (avg ${avg.toFixed(1)}) — generating level ${currentLevel + 1} cases`);
      try {
        const newCases = await generateHarderCases(cat, currentLevel, allCases, categoryStatus);
        newCasesGenerated = [...newCasesGenerated, ...newCases];
        progress[cat].level       = currentLevel + 1;
        progress[cat].masteredAt  = timestamp;
        progress[cat].runsAtLevel = 0;
      } catch (err) {
        if (err.message.startsWith('BUDGET_EXCEEDED')) { console.log('Budget hit — skipping.'); break; }
        console.warn(`Generation failed for "${cat}": ${err.message}`);
      }
    } else if (mastered) {
      // Mastered but suite is sick — still mark progress but don't generate.
      console.log(`✓ "${cat}" mastered at level ${currentLevel} (avg ${avg.toFixed(1)}) — holding generation (suite avg ${overallAvg.toFixed(1)} below ${SUITE_HEALTH_THRESHOLD}).`);
      progress[cat].masteredAt  = timestamp;
    } else {
      progress[cat].runsAtLevel = (progress[cat].runsAtLevel ?? 0) + 1;
      if (progress[cat].runsAtLevel >= STUCK_RUNS_THRESHOLD && suiteHealthy) {
        console.log(`⚠ "${cat}" stuck at level ${currentLevel} for ${progress[cat].runsAtLevel} runs — generating easier cases`);
        try {
          const easierCases = await generateEasierCases(cat, currentLevel, allCases);
          newCasesGenerated = [...newCasesGenerated, ...easierCases];
          progress[cat].runsAtLevel = 0;
        } catch (err) {
          if (err.message.startsWith('BUDGET_EXCEEDED')) { console.log('Budget hit — skipping.'); break; }
          console.warn(`Easier case generation failed: ${err.message}`);
        }
      }
    }
  }

  // Brand-new category — the strictest gate. ONLY fire when:
  //   • every existing category is at mastery (avg >= MASTERY_THRESHOLD), AND
  //   • overall avg >= NEW_CAT_THRESHOLD (97 by default).
  // Adds 1 case (configurable via NEW_CAT_CASES_PER_GEN). The old L2 path
  // that generated new categories at just 90 overall avg has been removed —
  // it was the main driver of suite bloat.
  const existingCategories = Object.keys(byCategory);
  if (everyCategoryMastered && overallAvg >= NEW_CAT_THRESHOLD) {
    console.log(`\n📂 Every category mastered AND overall avg ${overallAvg.toFixed(1)} >= ${NEW_CAT_THRESHOLD} — generating a single new category (${NEW_CAT_CASES_PER_GEN} case).`);
    try {
      const newCat = await generateNewCategory(existingCategories);
      if (newCat) {
        newCasesGenerated = [...newCasesGenerated, ...newCat.cases];
        if (!progress[newCat.category]) progress[newCat.category] = { level: 1, masteredAt: null };
        console.log(`  Added new category: "${newCat.category}" (${newCat.cases.length} case${newCat.cases.length === 1 ? '' : 's'})`);
      }
    } catch (err) {
      if (err.message.startsWith('BUDGET_EXCEEDED')) console.log('Budget hit — skipping.');
      else console.warn(`New category generation failed: ${err.message}`);
    }
  } else if (suiteHealthy) {
    // Log why we're holding back, so the trend is visible in CI output.
    const unmastered = categoryStatus.filter(c => !c.mastered);
    if (unmastered.length) {
      console.log(`\n📂 Holding new-category generation — ${unmastered.length}/${categoryStatus.length} categories not yet mastered (e.g. ${unmastered.slice(0,3).map(c => `${c.cat}=${c.avg.toFixed(0)}`).join(', ')}).`);
    } else if (overallAvg < NEW_CAT_THRESHOLD) {
      console.log(`\n📂 Holding new-category generation — overall avg ${overallAvg.toFixed(1)} below new-category threshold ${NEW_CAT_THRESHOLD}.`);
    }
  }

  if (newCasesGenerated.length > 0) {
    const existing  = fs.existsSync(PATHS.generated) ? JSON.parse(fs.readFileSync(PATHS.generated, 'utf8')) : [];
    fs.writeFileSync(PATHS.generated, JSON.stringify([...existing, ...newCasesGenerated], null, 2));
    console.log(`\nSaved ${newCasesGenerated.length} new generated cases to eval/generated-cases.json`);
  }
  saveProgress(progress);

  const failingResults = results.filter(r => r.score < 70);
  if (failingResults.length > 0) {
    console.log(`\n⚡ ${failingResults.length} test(s) failed (<70) — generating system prompt patch...`);
    try {
      const impFile = path.join(__dirname, 'improvements.txt');
      const currentImprovements = fs.existsSync(impFile) ? fs.readFileSync(impFile, 'utf8').trim() : '';
      const fix = await generateFix(failingResults, currentImprovements);
      applyImprovements(fix);
    } catch (err) {
      if (err.message.startsWith('BUDGET_EXCEEDED')) console.log('Budget hit — skipping.');
      else console.warn(`Auto-patch failed: ${err.message}`);
    }
  } else {
    console.log('\n✓ All tests passed (>=70) — no system prompt patch needed.');
  }
  console.log(`\nResults → eval/results/${timestamp}.json`);
  console.log(`Total cost: $${totalCostUSD.toFixed(4)}\n`);
}

main().catch(err => { console.error(err.message); process.exit(1); });
