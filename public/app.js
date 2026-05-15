/* Word AI Assistant — app.js
 *
 * All in-page logic for the task pane:
 *   - chat history, model picker, settings, cost limits
 *   - document sync (readDoc) — gathers heading outline + selection
 *   - execCode() — wraps AI-generated JS in Word.run with the full helper surface
 *   - Document Generator, Theme Designer, Template Gallery, Writing Coach,
 *     Find & Replace, Mail Merge panels
 *
 * Companion file: helpers.js exposes static data (templates, theme seeds,
 * OOXML snippets, font / contrast guards) on window.WordHelpers.
 */

// ── State ────────────────────────────────────────────────────────────────────
let chatId = null, history = [], summary = null, docData = '', selectionInfo = '';
let loading = false, ctrl = null, attached = null;
let totalTok = 0, autoSync = true, showCode = true, liveStats = true;
let opts = { deep: false, dynamic: false, auto: false, plan: false, sectionByStep: false };
let selectedModel = localStorage.getItem('selModel') || 'meta-llama/llama-3.3-70b-instruct';
let localModels = [], customModels = [], catalogModels = [], catalogLoaded = false, changeReady = false, syncTimer = null;
let plusOpen = false, timerInterval = null;
let recentModels = JSON.parse(localStorage.getItem('recentModels') || '[]');
let modelUsage   = JSON.parse(localStorage.getItem('modelUsage')   || '{}');
let activeModelFilter = null, modelSortMode = null, filterPopupOpen = false;
let lastRecommendation = null, recReasonOpen = false;
let userPreferences = localStorage.getItem('userPreferences') || '';
let costLimitType = localStorage.getItem('costLimitType') || 'prompt';
let costLimitIdx  = parseInt(localStorage.getItem('costLimitIdx') || '0');
let sessionCost = 0;
let pendingSendResolve = null;
let webSearchMode = localStorage.getItem('webSearchMode') || 'off';

// Theme designer state
let currentTheme = null;       // last applied theme (palette/typography object)
let pendingTheme = null;       // theme generated but not yet applied
let pendingThemeDesc = '';

// Generator state
let genTone = 'formal', genIncludeCover = false, genIncludeToc = false, genIncludeBib = false;

// Coach state
let coachScope = 'selection', coachFocus = 'clarity', coachAction = 'review';

// Templates state
let activeTemplateId = null;

// Mail-merge state
let mmDetectedFields = [];

// Find/replace state
let pendingFnrResults = [];
let pendingWebSearchResults = []; // accumulates per-turn for re-prompt

const PROMPT_LIMITS  = [0, 0.001, 0.002, 0.005, 0.01, 0.02, 0.05, 0.10, 0.20, 0.50, 1.00];
const SESSION_LIMITS = [0, 0.01,  0.02,  0.05,  0.10, 0.25, 0.50, 1.00, 2.00, 5.00, 10.00];

const OR_MODELS = [
  {id:'meta-llama/llama-3.3-70b-instruct',     name:'Llama 3.3 70B',    provider:'OpenRouter', in:0.07, out:0.30, tags:['rec']},
  {id:'meta-llama/llama-3.1-8b-instruct:free', name:'Llama 3.1 8B',     provider:'OpenRouter', in:0,    out:0,    tags:['free']},
  {id:'google/gemini-2.0-flash-001',           name:'Gemini 2.0 Flash', provider:'OpenRouter', in:0,    out:0,    tags:['free','fast']},
  {id:'google/gemini-2.5-pro-preview',         name:'Gemini 2.5 Pro',   provider:'OpenRouter', in:1.25, out:10,   tags:[]},
  {id:'anthropic/claude-sonnet-4-5',           name:'Claude Sonnet 4.5',provider:'OpenRouter', in:3,    out:15,   tags:[]},
  {id:'anthropic/claude-3-haiku',              name:'Claude 3 Haiku',   provider:'OpenRouter', in:0.25, out:1.25, tags:['fast']},
  {id:'openai/gpt-4o',                         name:'GPT-4o',           provider:'OpenRouter', in:2.50, out:10,   tags:[]},
  {id:'openai/gpt-4o-mini',                    name:'GPT-4o Mini',      provider:'OpenRouter', in:0.15, out:0.60, tags:['fast']},
  {id:'deepseek/deepseek-r1',                  name:'DeepSeek R1',      provider:'OpenRouter', in:0.55, out:2.19, tags:['reason']},
  {id:'qwen/qwen-2.5-72b-instruct',            name:'Qwen 2.5 72B',     provider:'OpenRouter', in:0.35, out:0.40, tags:[]},
];
const GROQ_MODELS = [
  {id:'llama-3.3-70b-versatile', name:'Llama 3.3 70B', provider:'Groq', in:0.59, out:0.79, tags:['groq','fast']},
  {id:'llama-3.1-8b-instant',    name:'Llama 3.1 8B',  provider:'Groq', in:0.05, out:0.08, tags:['groq','fast']},
  {id:'mixtral-8x7b-32768',      name:'Mixtral 8x7B',  provider:'Groq', in:0.24, out:0.24, tags:['groq']},
];

// ── Boot ─────────────────────────────────────────────────────────────────────
Office.onReady(async info => {
  if (info.host !== Office.HostType.Word) return;
  loadSettings();
  customModels = JSON.parse(localStorage.getItem('customModels') || '[]');
  await Promise.all([loadChats(), detectLocal(), refreshPackStatus(), refreshWebSearchStatus(), loadRecentThemes()]);
  renderModelBtn();
  renderThemeChips();
  await readDoc();
  registerListeners();
  if (!chatId) newChat(false);
  document.addEventListener('click', e => {
    if (plusOpen && !e.target.closest('.plus-wrap')) closePlus();
    if (filterPopupOpen && !e.target.closest('.model-ph')) closeFilterPopup();
  });
});

// ── Settings ─────────────────────────────────────────────────────────────────
function loadSettings() {
  const th = localStorage.getItem('uiTheme') || 'light';
  setUiTheme(th, false);
  autoSync   = localStorage.getItem('autoSync')   !== 'false';
  showCode   = localStorage.getItem('showCode')   !== 'false';
  liveStats  = localStorage.getItem('liveStats')  !== 'false';
  document.getElementById('autoSyncTog').className  = 'toggle' + (autoSync?' on':'');
  document.getElementById('showCodeTog').className  = 'toggle' + (showCode?' on':'');
  document.getElementById('liveStatsTog').className = 'toggle' + (liveStats?' on':'');
  const k = localStorage.getItem('apiKey');   if (k)  document.getElementById('apiKeyInp').value  = k;
  const gk= localStorage.getItem('groqKey');  if (gk) document.getElementById('groqKeyInp').value = gk;
  document.getElementById('prefsInp').value = userPreferences;
  document.getElementById('seg-prompt').className  = 'seg-btn' + (costLimitType==='prompt'?' active':'');
  document.getElementById('seg-session').className = 'seg-btn' + (costLimitType==='session'?' active':'');
  document.getElementById('costSlider').value = costLimitIdx;
  updateCostSlider(costLimitIdx, false);
  ['off','on','disabled'].forEach(m =>
    document.getElementById('ws-'+m).className = 'seg-btn' + (webSearchMode===m?' active':'')
  );
  updateWebSearchNote();
}
function setUiTheme(t, save=true) {
  document.documentElement.setAttribute('data-theme', t);
  document.getElementById('tLight').className = 'theme-opt' + (t==='light'?' active':'');
  document.getElementById('tDark').className  = 'theme-opt' + (t==='dark'?' active':'');
  if (save) localStorage.setItem('uiTheme', t);
}
function saveKey(provider) {
  const id = provider==='or'?'apiKeyInp':'groqKeyInp';
  const key = provider==='or'?'apiKey':'groqKey';
  const k = document.getElementById(id).value.trim();
  if (k) localStorage.setItem(key, k); else localStorage.removeItem(key);
}
function togSet(key, id) {
  if (key==='autoSync')      { autoSync=!autoSync;     localStorage.setItem('autoSync', autoSync);    document.getElementById(id).className='toggle'+(autoSync?' on':''); }
  else if (key==='showCode') { showCode=!showCode;     localStorage.setItem('showCode', showCode);    document.getElementById(id).className='toggle'+(showCode?' on':''); }
  else if (key==='liveStats'){ liveStats=!liveStats;   localStorage.setItem('liveStats', liveStats);  document.getElementById(id).className='toggle'+(liveStats?' on':''); document.getElementById('docStats').style.display = liveStats ? '' : 'none'; }
}
function savePrefs() {
  userPreferences = document.getElementById('prefsInp').value;
  localStorage.setItem('userPreferences', userPreferences);
}
function setCostLimitType(t) {
  costLimitType = t;
  localStorage.setItem('costLimitType', t);
  document.getElementById('seg-prompt').className  = 'seg-btn' + (t==='prompt'?' active':'');
  document.getElementById('seg-session').className = 'seg-btn' + (t==='session'?' active':'');
  updateCostSlider(costLimitIdx, false);
}
function updateCostSlider(val, save=true) {
  costLimitIdx = parseInt(val);
  if (save) localStorage.setItem('costLimitIdx', costLimitIdx);
  const lim = getCostLimit();
  document.getElementById('costSliderVal').textContent = lim === 0 ? 'Off' : '$' + lim.toFixed(lim < 0.01 ? 3 : 2);
  document.getElementById('costNote').textContent = lim === 0
    ? 'No limit — all prompts allowed.'
    : `Max $${lim.toFixed(lim < 0.01 ? 3 : 2)} per ${costLimitType}. You'll be warned before exceeding.`;
}
function getCostLimit() {
  const arr = costLimitType === 'prompt' ? PROMPT_LIMITS : SESSION_LIMITS;
  return arr[Math.min(costLimitIdx, arr.length - 1)] || 0;
}
function resetSessionCost() { sessionCost = 0; updateSessionCostDisplay(); }
function updateSessionCostDisplay() {
  const el = document.getElementById('sessionCostDisplay');
  if (el) el.textContent = '$' + sessionCost.toFixed(3);
  const note = document.getElementById('sessionCostNote');
  if (note) {
    const lim = getCostLimit();
    if (lim > 0 && costLimitType === 'session') {
      note.textContent = `Session limit: $${lim.toFixed(2)} — $${Math.max(0, lim - sessionCost).toFixed(3)} remaining.`;
    } else { note.textContent = ''; }
  }
}
function switchSettingsTab(tab) {
  ['general','prefs','data'].forEach(t => {
    document.getElementById('sbody-'+t).style.display = tab === t ? '' : 'none';
    document.getElementById('stab-'+t).className = 'stab' + (tab===t?' active':'');
  });
}
function openSettings()  { document.getElementById('settingsPanel').classList.add('open'); }
function closeSettings() { document.getElementById('settingsPanel').classList.remove('open'); }

// ── Web search settings ─────────────────────────────────────────────────────
function setWebSearchMode(m) {
  webSearchMode = m;
  localStorage.setItem('webSearchMode', m);
  ['off','on','disabled'].forEach(x =>
    document.getElementById('ws-'+x).className = 'seg-btn' + (m===x?' active':'')
  );
  updateWebSearchNote();
}
function updateWebSearchNote() {
  const note = document.getElementById('wsNote');
  if (!note) return;
  note.textContent = ({
    off: 'The AI will ask permission per search before reaching the web.',
    on: 'Searches happen automatically when the AI needs live data.',
    disabled: 'No web access; the AI will tell you to enable it if needed.',
  })[webSearchMode] || '';
}
async function refreshWebSearchStatus() {
  try {
    const r = await fetch('/api/web-search/status');
    const d = await r.json();
    const el = document.getElementById('wsProvNote');
    if (!el) return;
    if (!d.configured) {
      el.innerHTML = 'Provider: <strong>not configured</strong> — set <code>GOOGLE_KEY</code>+<code>GOOGLE_CX</code> (preferred) or <code>BRAVE_KEY</code> in <code>.env</code>.';
    } else {
      const others = [];
      if (d.available?.google && d.provider !== 'google') others.push('Google available');
      if (d.available?.brave  && d.provider !== 'brave')  others.push('Brave available');
      const extra = others.length ? ` · ${others.join(' · ')}` : '';
      el.innerHTML = `Provider: <strong>${d.provider === 'google' ? 'Google Custom Search' : 'Brave Search'}</strong> (configured)${extra}<br>Set <code>WEB_SEARCH_PROVIDER=google|brave</code> in <code>.env</code> to switch.`;
    }
  } catch {}
}

// ── Asset pack settings ─────────────────────────────────────────────────────
async function refreshPackStatus() {
  try {
    const r = await fetch('/api/assets/status');
    const d = await r.json();
    const el = document.getElementById('packStatus');
    if (!el) return;
    if (d.downloaded) {
      el.className = 'pack-status ok';
      const mb = (d.sizeBytes/1024/1024).toFixed(1);
      el.textContent = `Downloaded · ${d.imageCount} images · ${mb} MB · ${d.categories.length} categories`;
    } else if (d.imageCount > 0 && d.totalImages) {
      el.className = 'pack-status dl';
      el.textContent = `Partial · ${d.imageCount}/${d.totalImages} images downloaded`;
    } else {
      el.className = 'pack-status';
      el.textContent = `Not downloaded · ${d.totalImages || 0} images available`;
    }
  } catch {}
}
async function downloadPack() {
  const btn = document.getElementById('packDlBtn');
  btn.disabled = true; btn.textContent = 'Starting…';
  try {
    const r = await fetch('/api/assets/download-pack', { method: 'POST' });
    if (!r.ok) { btn.disabled = false; btn.textContent = 'Download asset pack'; alert('Download failed: ' + (await r.json()).error); return; }
    const ev = new EventSource('/api/assets/download-pack/progress');
    ev.onmessage = e => {
      const s = JSON.parse(e.data);
      btn.textContent = s.active ? `Downloading ${s.done}/${s.total}…` : 'Done';
      if (!s.active && s.done >= s.total) { ev.close(); btn.disabled = false; btn.textContent = 'Re-download pack'; refreshPackStatus(); }
    };
  } catch (e) {
    btn.disabled = false; btn.textContent = 'Download asset pack';
    alert('Download failed: ' + e.message);
  }
}
async function verifyPack() {
  await fetch('/api/assets/download-pack', { method: 'POST' }).catch(()=>{});
  refreshPackStatus();
}

// ── Theme designer ───────────────────────────────────────────────────────────
function openThemePicker() {
  document.getElementById('themeModal').classList.add('open');
  document.getElementById('themePrompt').focus();
}
function closeThemePicker() {
  document.getElementById('themeModal').classList.remove('open');
}
function renderThemeChips() {
  const wrap = document.getElementById('themeChips');
  wrap.innerHTML = '';
  (window.WordHelpers?.THEME_SEEDS || []).forEach(seed => {
    const c = document.createElement('button');
    c.className = 'theme-chip';
    c.textContent = seed.label;
    c.onclick = () => {
      document.getElementById('themePrompt').value = seed.prompt;
      generateTheme();
    };
    wrap.appendChild(c);
  });
}
async function loadRecentThemes() {
  try {
    const r = await fetch('/api/themes/recent');
    const d = await r.json();
    renderRecentThemes(d.recent || []);
  } catch {}
}
function renderRecentThemes(list) {
  const wrap = document.getElementById('themeRecent');
  const lbl  = document.getElementById('recentLbl');
  if (!list.length) { lbl.style.display = 'none'; wrap.innerHTML = ''; return; }
  lbl.style.display = '';
  wrap.innerHTML = '';
  list.slice(0, 6).forEach(r => {
    const row = document.createElement('div');
    row.className = 'theme-recent-row';
    const swatches = [r.theme.primary, r.theme.accent, r.theme.bg, r.theme.text]
      .map(c => `<div class="theme-sw" style="background:${c}"></div>`).join('');
    row.innerHTML = `<div class="theme-recent-desc">${esc(r.description)}</div><div class="theme-recent-mini">${swatches}</div>`;
    row.onclick = () => {
      pendingTheme = r.theme;
      pendingThemeDesc = r.description;
      renderThemePreview(r.theme, null, r.description);
    };
    wrap.appendChild(row);
  });
}
async function generateTheme() {
  const desc = document.getElementById('themePrompt').value.trim();
  if (!desc) { alert('Describe the look you want first.'); return; }
  const btn = document.getElementById('themeGenBtn');
  btn.disabled = true; btn.textContent = 'Generating…';
  try {
    const r = await fetch('/api/design-theme', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ description: desc, apiKey: localStorage.getItem('apiKey') || undefined }),
    });
    const d = await r.json();
    if (d.error) { alert('Theme generation failed: ' + d.error); return; }
    pendingTheme = d.theme;
    pendingThemeDesc = desc;
    renderThemePreview(d.theme, d.warnings, desc);
    loadRecentThemes();
  } catch (e) {
    alert('Theme generation failed: ' + e.message);
  } finally {
    btn.disabled = false; btn.textContent = 'Generate';
  }
}
async function surpriseTheme() {
  const btn = document.getElementById('themeGenBtn');
  btn.disabled = true; btn.textContent = 'Surprising…';
  try {
    const r = await fetch('/api/design-theme/surprise', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ apiKey: localStorage.getItem('apiKey') || undefined }),
    });
    const d = await r.json();
    if (d.error) { alert('Theme generation failed: ' + d.error); return; }
    pendingTheme = d.theme;
    pendingThemeDesc = d.description;
    document.getElementById('themePrompt').value = d.description;
    renderThemePreview(d.theme, d.warnings, d.description);
    loadRecentThemes();
  } catch (e) {
    alert('Surprise failed: ' + e.message);
  } finally {
    btn.disabled = false; btn.textContent = 'Generate';
  }
}
function renderThemePreview(t, warnings, desc) {
  const wrap = document.getElementById('themePreview');
  document.getElementById('previewLbl').style.display = '';
  wrap.style.display = '';
  wrap.style.background = t.bg;
  wrap.style.color = t.text;
  wrap.innerHTML = `
    <div class="theme-preview-name" style="color:${t.text};opacity:.6">${esc(t.name)}</div>
    <div class="theme-preview-h" style="color:${t.primary};font-family:'${t.headingFont}',sans-serif;font-size:${t.headingSize}px;font-weight:700">Heading example</div>
    <div class="theme-preview-sub" style="color:${t.secondary};font-family:'${t.headingFont}',sans-serif;font-size:${Math.max(12, Math.round(t.headingSize*0.65))}px;font-weight:600">Subheading example</div>
    <div class="theme-preview-p" style="color:${t.text};font-family:'${t.bodyFont}',serif;font-size:11pt;line-height:1.55">Body text in <strong>${esc(t.bodyFont)}</strong> on a clean background. A short paragraph proves contrast and rhythm — short sentences, varied length, no jargon. <a class="theme-preview-a" style="color:${t.accent}">accent link</a> for context.</div>
    <div class="theme-swatches-row">
      <span class="theme-sw-lbl">PALETTE</span>
      <div class="theme-sw" style="background:${t.primary}" title="primary ${t.primary}"></div>
      <div class="theme-sw" style="background:${t.secondary}" title="secondary ${t.secondary}"></div>
      <div class="theme-sw" style="background:${t.accent}" title="accent ${t.accent}"></div>
      <div class="theme-sw" style="background:${t.bg};border:1px solid #aaa" title="bg ${t.bg}"></div>
      <div class="theme-sw" style="background:${t.text}" title="text ${t.text}"></div>
    </div>
    ${t.rationale ? `<div class="theme-rationale">${esc(t.rationale)}</div>` : ''}
    ${(warnings && warnings.length) ? warnings.map(w => `<div class="theme-warning">⚠ ${esc(w)}</div>`).join('') : ''}
  `;
  document.getElementById('themeActions').style.display = '';
}
async function applyGeneratedTheme() {
  if (!pendingTheme) return;
  closeThemePicker();
  // Apply via execCode using applyTheme helper
  const code = `await applyTheme(${JSON.stringify(pendingTheme)});`;
  addMsg('user', `Apply theme: ${pendingThemeDesc}`);
  history.push({ role: 'user', content: `Apply theme: ${pendingThemeDesc}`, time: new Date().toISOString() });
  addTyping();
  const r = await execCode(code);
  removeTyping();
  if (r.ok) {
    currentTheme = pendingTheme;
    localStorage.setItem('activeTheme', JSON.stringify(pendingTheme));
    addMsg('assistant', `✓ Applied theme: ${pendingTheme.name}. ${r.result || ''}`);
    history.push({ role: 'assistant', content: `Applied theme ${pendingTheme.name}.` });
  } else {
    addMsg('assistant', '⚠ Apply failed: ' + r.err, true);
  }
  await saveChat();
}
async function tweakActiveTheme() {
  const desc = document.getElementById('themePrompt').value.trim();
  if (!desc) { alert('Describe the tweak first (e.g. "same but darker, swap accent for green").'); return; }
  const btn = document.getElementById('themeGenBtn');
  btn.disabled = true; btn.textContent = 'Tweaking…';
  try {
    const r = await fetch('/api/design-theme', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ description: desc, previous: pendingTheme || currentTheme, apiKey: localStorage.getItem('apiKey') || undefined }),
    });
    const d = await r.json();
    if (d.error) { alert('Tweak failed: ' + d.error); return; }
    pendingTheme = d.theme;
    pendingThemeDesc = desc;
    renderThemePreview(d.theme, d.warnings, desc);
    loadRecentThemes();
  } finally {
    btn.disabled = false; btn.textContent = 'Generate';
  }
}

// ── Document generator ───────────────────────────────────────────────────────
function openDocGenerator() { document.getElementById('docGenModal').classList.add('open'); }
function closeDocGenerator() {
  document.getElementById('docGenModal').classList.remove('open');
  document.getElementById('genProgress').style.display = 'none';
  document.getElementById('genGoBtn').disabled = false;
}
function setGenTone(t) {
  genTone = t;
  document.querySelectorAll('#genToneCtrl .seg-btn').forEach(b => b.classList.remove('active'));
  event.target.classList.add('active');
}
function togGen(k) {
  if (k==='cover') { genIncludeCover = !genIncludeCover; document.getElementById('genCoverTog').className = 'toggle' + (genIncludeCover?' on':''); }
  if (k==='toc')   { genIncludeToc   = !genIncludeToc;   document.getElementById('genTocTog').className   = 'toggle' + (genIncludeToc?' on':''); }
  if (k==='bib')   { genIncludeBib   = !genIncludeBib;   document.getElementById('genBibTog').className   = 'toggle' + (genIncludeBib?' on':''); }
}
async function generateDocument() {
  const topic = document.getElementById('genTopic').value.trim();
  if (!topic) { alert('Enter a topic first.'); return; }
  const audience = document.getElementById('genAudience').value.trim() || 'general';
  const length   = document.getElementById('genLength').value;
  const goBtn = document.getElementById('genGoBtn');
  goBtn.disabled = true;
  document.getElementById('genProgress').style.display = '';
  document.getElementById('genProgressText').textContent = 'Asking AI for outline…';
  document.getElementById('genBarFill').style.width = '5%';

  const isGroq = selectedModel.startsWith('groq:');
  const isOllama = selectedModel.startsWith('ollama:');
  const modelId = isGroq ? selectedModel.slice(5) : isOllama ? selectedModel.slice(7) : selectedModel;
  const apiKey = localStorage.getItem('apiKey') || undefined;
  const groqKey = localStorage.getItem('groqKey') || undefined;

  try {
    const r = await fetch('/api/generate-doc', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        topic, audience, tone: genTone, length,
        includeCover: genIncludeCover, includeToc: genIncludeToc, includeBibliography: genIncludeBib,
        model: modelId, useOllama: isOllama, useGroq: isGroq, apiKey, groqKey,
      })
    });
    const data = await r.json();
    if (data.error) { alert('Generation failed: ' + data.error); closeDocGenerator(); return; }
    const outline = data.outline || [];
    if (!outline.length) { alert('Got an empty outline.'); closeDocGenerator(); return; }

    addMsg('user', `Generate a ${genTone} ${length} document about: ${topic}`);
    history.push({ role: 'user', content: `Generate a ${genTone} ${length} document about: ${topic}`, time: new Date().toISOString() });
    addMsg('assistant', `📋 Outline ready — ${outline.length} sections. Writing now…`);

    // Stream section by section so the user can watch it appear.
    for (let i = 0; i < outline.length; i++) {
      const s = outline[i];
      document.getElementById('genProgressText').textContent = `Section ${i+1} of ${outline.length}: ${s.type}…`;
      document.getElementById('genBarFill').style.width = (5 + 90 * (i / outline.length)) + '%';
      const code = renderOutlineSection(s);
      if (code) {
        const r2 = await execCode(code);
        if (!r2.ok) addMsg('assistant', `⚠ Section ${i+1} (${s.type}) failed: ${r2.err}`, true);
      }
    }
    document.getElementById('genBarFill').style.width = '100%';
    document.getElementById('genProgressText').textContent = 'Done.';
    addMsg('assistant', `✓ Wrote ${outline.length} sections.`);
    history.push({ role: 'assistant', content: `Wrote ${outline.length} sections.` });
    await saveChat();
    setTimeout(closeDocGenerator, 800);
  } catch (e) {
    alert('Generation failed: ' + e.message);
    goBtn.disabled = false;
  }
}
function renderOutlineSection(s) {
  const t = JSON.stringify(s.text || '');
  switch (s.type) {
    case 'title':       return `await addTitle(${t});`;
    case 'subtitle':    return `await addSubtitle(${t});`;
    case 'heading1':    return `await addHeading(${t}, 1);`;
    case 'heading2':    return `await addHeading(${t}, 2);`;
    case 'heading3':    return `await addHeading(${t}, 3);`;
    case 'paragraph':   return `await addParagraph(${t});`;
    case 'bullet':      return `await addList([${t}], 'bullet');`;
    case 'number':      return `await addList([${t}], 'number');`;
    case 'quote':       return `await addQuote(${t});`;
    case 'pageBreak':   return `await insertPageBreak();`;
    case 'toc':         return `await insertTableOfContents({ title: ${t || '"Table of Contents"'} });`;
    case 'bibliography':return `await insertBibliography(${t || '"APA"'});`;
    default:            return `await addParagraph(${t});`;
  }
}

// ── Templates gallery ────────────────────────────────────────────────────────
function openTemplates() {
  const list = document.getElementById('tmplList');
  list.style.display = 'grid';
  document.getElementById('tmplForm').classList.remove('open');
  document.getElementById('tmplHeadText').textContent = '📋 Document Templates';
  list.innerHTML = '';
  const TEMPLATES = window.WordHelpers.TEMPLATES;
  Object.entries(TEMPLATES).forEach(([id, t]) => {
    const card = document.createElement('div');
    card.className = 'tmpl-card';
    card.innerHTML = `<div class="tmpl-name">${esc(t.name)}</div><div class="tmpl-desc">${esc(t.description)}</div>`;
    card.onclick = () => openTemplateForm(id);
    list.appendChild(card);
  });
  document.getElementById('tmplModal').classList.add('open');
}
function openTemplateForm(id) {
  activeTemplateId = id;
  const T = window.WordHelpers.TEMPLATES[id];
  document.getElementById('tmplHeadText').textContent = '📋 ' + T.name;
  document.getElementById('tmplList').style.display = 'none';
  const form = document.getElementById('tmplForm');
  form.classList.add('open');
  form.innerHTML = '';
  const back = document.createElement('button');
  back.className = 'tmpl-back';
  back.textContent = '← Back to templates';
  back.onclick = openTemplates;
  form.appendChild(back);
  const desc = document.createElement('div');
  desc.style = 'font-size:12px;color:var(--text2)';
  desc.textContent = T.description;
  form.appendChild(desc);
  T.placeholders.forEach(p => {
    const wrap = document.createElement('div');
    wrap.className = 'tmpl-field';
    wrap.innerHTML = `<div class="tmpl-field-lbl">${esc(p)}</div>`;
    const inp = document.createElement(p.toLowerCase().includes('body') || p.toLowerCase().includes('summary') || p.toLowerCase().includes('discussion') || p.toLowerCase().includes('experience') ? 'textarea' : 'input');
    inp.className = 'tmpl-field-inp';
    inp.dataset.placeholder = p;
    inp.placeholder = `Enter ${p}…`;
    if (inp.tagName === 'TEXTAREA') { inp.rows = 3; inp.style.minHeight = '70px'; inp.style.fontFamily = 'inherit'; }
    wrap.appendChild(inp);
    form.appendChild(wrap);
  });
  const goBtn = document.createElement('button');
  goBtn.className = 'theme-gen-btn';
  goBtn.style.marginTop = '6px';
  goBtn.textContent = '✓ Apply template';
  goBtn.onclick = () => applySelectedTemplate(id);
  form.appendChild(goBtn);
}
async function applySelectedTemplate(id) {
  const T = window.WordHelpers.TEMPLATES[id];
  const form = document.getElementById('tmplForm');
  const fields = {};
  form.querySelectorAll('.tmpl-field-inp').forEach(el => {
    fields[el.dataset.placeholder] = el.value || '';
  });
  closeTemplates();
  const code = `await applyTemplate(${JSON.stringify(id)}, ${JSON.stringify(fields)});`;
  addMsg('user', `Apply template: ${T.name}`);
  history.push({ role:'user', content: `Apply template: ${T.name}`, time: new Date().toISOString() });
  addTyping();
  const r = await execCode(code);
  removeTyping();
  addMsg('assistant', r.ok ? `✓ Applied ${T.name}.` : '⚠ Apply failed: ' + r.err, !r.ok);
  await saveChat();
}
function closeTemplates() { document.getElementById('tmplModal').classList.remove('open'); }

// ── Writing coach ────────────────────────────────────────────────────────────
function openCoach() { document.getElementById('coachModal').classList.add('open'); }
function closeCoach() { document.getElementById('coachModal').classList.remove('open'); }
function setCoachScope(v)  { coachScope = v;  syncSeg('#coachScopeCtrl', v); }
function setCoachFocus(v)  { coachFocus = v;  syncSeg('#coachFocusCtrl', v); }
function setCoachAction(v) { coachAction = v; syncSeg('#coachActionCtrl', v); document.getElementById('coachTransGrp').style.display = v === 'translate' ? '' : 'none'; }
function syncSeg(sel, v) {
  document.querySelectorAll(sel + ' .seg-btn').forEach(b => {
    b.classList.toggle('active', b.textContent.toLowerCase().replace(/\s+/g, '-') === v || b.textContent.toLowerCase() === v);
  });
}
async function runCoach() {
  const btn = document.getElementById('coachRunBtn');
  btn.disabled = true; btn.textContent = 'Running…';
  const out = document.getElementById('coachOutput');
  out.innerHTML = '';
  try {
    // 1) Read the chosen scope from the document
    let text = '';
    try {
      await Word.run(async ctx => {
        if (coachScope === 'selection') {
          const sel = ctx.document.getSelection();
          sel.load('text');
          await ctx.sync();
          text = sel.text || '';
        } else if (coachScope === 'paragraph') {
          const sel = ctx.document.getSelection();
          const para = sel.paragraphs.getFirst();
          para.load('text');
          await ctx.sync();
          text = para.text || '';
        } else {
          ctx.document.body.load('text');
          await ctx.sync();
          text = ctx.document.body.text || '';
        }
      });
    } catch (e) { text = ''; }
    if (!text.trim() && coachScope !== 'whole') {
      out.innerHTML = '<div class="msg-body" style="color:var(--err)">Nothing selected. Click in a paragraph or select text first.</div>';
      return;
    }

    const isGroq = selectedModel.startsWith('groq:');
    const isOllama = selectedModel.startsWith('ollama:');
    const modelId = isGroq ? selectedModel.slice(5) : isOllama ? selectedModel.slice(7) : selectedModel;
    const targetLang = coachAction === 'translate'
      ? (document.getElementById('coachTargetLang').value.trim() || 'Spanish')
      : null;
    const body = {
      text, scope: coachScope, focus: coachFocus, action: coachAction === 'translate' ? 'rewrite' : coachAction,
      targetLang, model: modelId, useOllama: isOllama, useGroq: isGroq,
      apiKey: localStorage.getItem('apiKey') || undefined,
      groqKey: localStorage.getItem('groqKey') || undefined,
    };
    if (coachAction === 'translate') body.focus = 'translate';

    const r = await fetch('/api/coach', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
    const d = await r.json();
    if (d.error) { out.innerHTML = '<div style="color:var(--err)">' + esc(d.error) + '</div>'; return; }

    const fb = d.feedback || '';
    out.innerHTML = `<div class="msg coach"><div class="msg-head"><span class="msg-role coach">🎤 Coach</span></div><div class="msg-body">${esc(fb)}</div></div>`;

    if (d.action === 'rewrite' || coachAction === 'rewrite' || coachAction === 'translate') {
      const applyBtn = document.createElement('button');
      applyBtn.className = 'coach-apply';
      applyBtn.textContent = '✓ Apply rewrite to ' + coachScope;
      applyBtn.onclick = async () => {
        applyBtn.disabled = true; applyBtn.textContent = 'Applying…';
        let code;
        if (coachScope === 'selection') {
          code = `const sel = doc.getSelection(); sel.insertText(${JSON.stringify(fb)}, 'Replace'); await context.sync();`;
        } else if (coachScope === 'paragraph') {
          code = `const sel = doc.getSelection(); const p = sel.paragraphs.getFirst(); p.insertText(${JSON.stringify(fb)}, 'Replace'); await context.sync();`;
        } else {
          code = `doc.body.clear(); doc.body.insertText(${JSON.stringify(fb)}, 'End'); await context.sync();`;
        }
        const rr = await execCode(code);
        applyBtn.textContent = rr.ok ? '✓ Applied' : '⚠ Failed: ' + rr.err;
      };
      out.appendChild(applyBtn);
    } else if (coachAction === 'review') {
      // Try to pull "SUGGESTED REWRITE" out
      const m = fb.match(/SUGGESTED REWRITE:?\s*([\s\S]+?)(?:\nREADING LEVEL:|$)/i);
      const suggested = m ? m[1].trim() : '';
      if (suggested) {
        const applyBtn = document.createElement('button');
        applyBtn.className = 'coach-apply';
        applyBtn.textContent = '✓ Apply suggested rewrite';
        applyBtn.onclick = async () => {
          applyBtn.disabled = true; applyBtn.textContent = 'Applying…';
          const code = `const sel = doc.getSelection(); sel.insertText(${JSON.stringify(suggested)}, 'Replace'); await context.sync();`;
          const rr = await execCode(code);
          applyBtn.textContent = rr.ok ? '✓ Applied' : '⚠ Failed: ' + rr.err;
        };
        out.appendChild(applyBtn);
      }
    }

    if (d.usage) addActualCost(d.usage, modelId);
  } catch (e) {
    out.innerHTML = '<div style="color:var(--err)">' + esc(e.message) + '</div>';
  } finally {
    btn.disabled = false; btn.textContent = 'Run coach';
  }
}

// ── Find & Replace ───────────────────────────────────────────────────────────
function openFnr()  { document.getElementById('fnrModal').classList.add('open'); document.getElementById('fnrFind').focus(); }
function closeFnr() { document.getElementById('fnrModal').classList.remove('open'); }
function fnrOpts() {
  return {
    matchCase: document.getElementById('fnrCase').checked,
    matchWholeWord: document.getElementById('fnrWhole').checked,
    matchWildcards: document.getElementById('fnrWild').checked,
  };
}
async function fnrFindNext() {
  const q = document.getElementById('fnrFind').value;
  if (!q) return;
  const o = fnrOpts();
  await execCode(`
    const results = doc.body.search(${JSON.stringify(q)}, ${JSON.stringify(o)});
    results.load("text");
    await context.sync();
    if (results.items.length) results.items[0].select();
    return 'Found ' + results.items.length + (results.items.length === 1 ? ' match.' : ' matches.');
  `).then(r => {
    document.getElementById('fnrResults').style.display = '';
    document.getElementById('fnrResults').textContent = r.ok ? r.result : '⚠ ' + r.err;
  });
}
async function fnrFindAll() {
  const q = document.getElementById('fnrFind').value;
  if (!q) return;
  const o = fnrOpts();
  const r = await execCode(`
    const results = doc.body.search(${JSON.stringify(q)}, ${JSON.stringify(o)});
    results.load("text");
    await context.sync();
    return results.items.length + ' match(es). First 8 contexts: ' + results.items.slice(0, 8).map(x => '"' + x.text.slice(0, 60) + '"').join(' | ');
  `);
  document.getElementById('fnrResults').style.display = '';
  document.getElementById('fnrResults').textContent = r.ok ? r.result : '⚠ ' + r.err;
}
async function fnrReplaceOne() {
  const q = document.getElementById('fnrFind').value;
  const repl = document.getElementById('fnrReplace').value;
  if (!q) return;
  const o = fnrOpts();
  const r = await execCode(`
    const results = doc.body.search(${JSON.stringify(q)}, ${JSON.stringify(o)});
    results.load("text");
    await context.sync();
    if (!results.items.length) return 'No matches.';
    results.items[0].insertText(${JSON.stringify(repl)}, 'Replace');
    await context.sync();
    return 'Replaced 1 occurrence.';
  `);
  document.getElementById('fnrResults').style.display = '';
  document.getElementById('fnrResults').textContent = r.ok ? r.result : '⚠ ' + r.err;
}
async function fnrReplaceAll() {
  const q = document.getElementById('fnrFind').value;
  const repl = document.getElementById('fnrReplace').value;
  if (!q) return;
  const o = fnrOpts();
  const r = await execCode(`await replaceText(${JSON.stringify(q)}, ${JSON.stringify(repl)}, ${JSON.stringify(o)});`);
  document.getElementById('fnrResults').style.display = '';
  document.getElementById('fnrResults').textContent = r.ok ? (r.result || '✓ Replaced all.') : '⚠ ' + r.err;
}

// ── Mail merge ───────────────────────────────────────────────────────────────
async function openMailMerge() {
  document.getElementById('mmModal').classList.add('open');
  const det = document.getElementById('mmDetected');
  det.textContent = 'Scanning…';
  try {
    await Word.run(async ctx => {
      ctx.document.body.load('text');
      await ctx.sync();
      const text = ctx.document.body.text || '';
      const matches = new Set();
      const re1 = /\{\{([A-Za-z0-9_]+)\}\}/g; let m;
      while ((m = re1.exec(text))) matches.add(m[1]);
      const re2 = /«([A-Za-z0-9_]+)»/g;
      while ((m = re2.exec(text))) matches.add(m[1]);
      mmDetectedFields = [...matches];
      det.innerHTML = mmDetectedFields.length
        ? mmDetectedFields.map(f => `<code style="background:var(--bg);padding:1px 5px;border-radius:3px;margin-right:4px">{{${esc(f)}}}</code>`).join(' ')
        : '<em>No {{Field}} or «Field» placeholders detected in document. Add some, then re-open this dialog.</em>';
    });
  } catch (e) { det.textContent = 'Scan failed: ' + e.message; }
}
function closeMailMerge() { document.getElementById('mmModal').classList.remove('open'); }
function parseMmInput() {
  const raw = document.getElementById('mmCsv').value.trim();
  if (!raw) return [];
  // Try JSON first
  if (raw.startsWith('[') || raw.startsWith('{')) {
    try { const j = JSON.parse(raw); return Array.isArray(j) ? j : [j]; } catch {}
  }
  // CSV
  const lines = raw.split(/\r?\n/).filter(Boolean);
  if (lines.length < 2) return [];
  const headers = lines[0].split(/[,\t]/).map(s => s.trim());
  return lines.slice(1).map(line => {
    const cells = line.split(/[,\t]/).map(s => s.trim());
    const o = {};
    headers.forEach((h, i) => { o[h] = cells[i] || ''; });
    return o;
  });
}
async function mmPreview() {
  const records = parseMmInput();
  const out = document.getElementById('mmResults');
  out.style.display = '';
  if (!records.length) { out.textContent = 'No records to preview.'; return; }
  out.textContent = `Will merge ${records.length} record(s): ` + records.slice(0, 4).map(r => JSON.stringify(r)).join(' | ') + (records.length > 4 ? ` and ${records.length - 4} more…` : '');
}
async function mmRun() {
  const records = parseMmInput();
  if (!records.length) { alert('No records.'); return; }
  closeMailMerge();
  const code = `await mailMergeReplace(${JSON.stringify(records)});`;
  addMsg('user', `Mail merge with ${records.length} record(s)`);
  history.push({ role: 'user', content: `Mail merge with ${records.length} record(s)`, time: new Date().toISOString() });
  addTyping();
  const r = await execCode(code);
  removeTyping();
  addMsg('assistant', r.ok ? (r.result || '✓ Merged.') : '⚠ Merge failed: ' + r.err, !r.ok);
  await saveChat();
}

// ── Quick toolbar actions ────────────────────────────────────────────────────
async function quickToggleTrack() {
  let on = null;
  await execCode(`
    return doc.changeTrackingMode || 'off';
  `).then(r => { if (r.ok) on = r.result !== 'off'; });
  const target = !on;
  const r = await execCode(`await toggleTrackChanges(${target});`);
  addMsg('assistant', r.ok ? (target ? '✎ Track changes ON.' : '✎ Track changes OFF.') : '⚠ ' + r.err, !r.ok);
}
async function quickTOC() {
  const r = await execCode(`await insertTableOfContents();`);
  addMsg('assistant', r.ok ? '📑 Inserted Table of Contents. Right-click the field and choose "Update Field" if it doesn\'t auto-populate.' : '⚠ ' + r.err, !r.ok);
}

// ── Cost estimation ──────────────────────────────────────────────────────────
function estimateCost(extraText) {
  const model = allModelsFlat().find(m => m.id === selectedModel);
  if (!model || (!model.in && !model.out)) return 0;
  const text = [...history.slice(-8), { content: extraText || '' }]
    .map(m => m.content || '').join(' ') + (docData || '');
  const inTok  = Math.ceil(text.length / 3.5);
  const outTok = 800;
  return (inTok * model.in + outTok * model.out) / 1_000_000;
}
function addActualCost(usage, modelId) {
  const model = allModelsFlat().find(m => m.id === (modelId || selectedModel));
  if (!model || (!model.in && !model.out)) return;
  const inTok  = usage?.prompt_tokens     || 0;
  const outTok = usage?.completion_tokens || 0;
  sessionCost += (inTok * model.in + outTok * model.out) / 1_000_000;
  updateSessionCostDisplay();
}
function checkCostBefore(text) {
  const lim = getCostLimit();
  if (lim === 0) return Promise.resolve(true);
  if (costLimitType === 'session' && sessionCost >= lim) {
    return showCostDialog(`Your session cost is already $${sessionCost.toFixed(3)}, which exceeds your $${lim.toFixed(2)} session limit.`);
  }
  const est = estimateCost(text);
  if (est > 0) {
    if (costLimitType === 'prompt' && est > lim) {
      return showCostDialog(`This prompt is estimated to cost $${est.toFixed(4)}, which exceeds your $${lim.toFixed(lim < 0.01 ? 3 : 2)} per-prompt limit.`);
    }
    if (costLimitType === 'session' && sessionCost + est > lim) {
      return showCostDialog(`This prompt (~$${est.toFixed(4)}) would bring your session total to $${(sessionCost + est).toFixed(3)}, exceeding the $${lim.toFixed(2)} limit.`);
    }
  }
  return Promise.resolve(true);
}
function showCostDialog(msg) {
  document.getElementById('costDialogBody').textContent = msg;
  document.getElementById('costDialog').classList.add('open');
  return new Promise(resolve => { pendingSendResolve = resolve; });
}
function confirmSend() { document.getElementById('costDialog').classList.remove('open'); if (pendingSendResolve) { pendingSendResolve(true); pendingSendResolve = null; } }
function cancelSend()  { document.getElementById('costDialog').classList.remove('open'); if (pendingSendResolve) { pendingSendResolve(false); pendingSendResolve = null; } }

// ── Plus popup ──────────────────────────────────────────────────────────────
function togglePlus() { plusOpen ? closePlus() : openPlus(); }
function openPlus()  { plusOpen=true; document.getElementById('plusPopup').classList.add('open'); }
function closePlus() { plusOpen=false; document.getElementById('plusPopup').classList.remove('open'); }
function toggleOpt(k) {
  opts[k] = !opts[k];
  const idMap = { deep:'deepCheck', dynamic:'dynamicCheck', auto:'autoCheck', plan:'planCheck', sectionByStep:'sectionByStepCheck' };
  if (idMap[k]) document.getElementById(idMap[k]).className = 'plus-check' + (opts[k]?' on':'');
  const anyOn = Object.values(opts).some(Boolean);
  document.getElementById('plusBtn').className = 'act-btn' + (anyOn?' active':'');
}

// ── Model picker ────────────────────────────────────────────────────────────
async function detectLocal() {
  try { const r = await fetch('/api/models/local'); localModels = await r.json(); } catch { localModels = []; }
}
function allModelsFlat() {
  return [
    ...OR_MODELS,
    ...GROQ_MODELS.map(m=>({...m, id:'groq:'+m.id})),
    ...localModels.map(m=>({id:'ollama:'+m.id, name:m.name, provider:'Ollama', in:0, out:0, tags:['local']})),
    ...customModels,
  ];
}
function renderModelBtn() {
  const f = allModelsFlat().find(m => m.id === selectedModel);
  document.getElementById('modelBtn').textContent = (f?.name || selectedModel) + ' ▾';
}
function addToRecent(id) {
  recentModels = [id, ...recentModels.filter(r => r !== id)].slice(0, 5);
  localStorage.setItem('recentModels', JSON.stringify(recentModels));
  modelUsage[id] = (modelUsage[id] || 0) + 1;
  localStorage.setItem('modelUsage', JSON.stringify(modelUsage));
}
function collectSpecs() {
  try {
    const canvas = document.createElement('canvas');
    const gl = canvas.getContext('webgl') || canvas.getContext('experimental-webgl');
    const ext = gl?.getExtension('WEBGL_debug_renderer_info');
    return {
      cores:     navigator.hardwareConcurrency || null,
      memory:    navigator.deviceMemory || null,
      gpu:       ext ? gl.getParameter(ext.UNMASKED_RENDERER_WEBGL) : null,
      gpuVendor: ext ? gl.getParameter(ext.UNMASKED_VENDOR_WEBGL)   : null,
      platform:  navigator.platform || null,
    };
  } catch { return {}; }
}
function getRecommendedModelId() {
  if (lastRecommendation?.modelId && allModelsFlat().find(m => m.id === lastRecommendation.modelId))
    return lastRecommendation.modelId;
  const entries = Object.entries(modelUsage);
  if (entries.length) {
    const [topId, topCount] = entries.sort((a,b)=>b[1]-a[1])[0];
    if (topCount >= 3 && allModelsFlat().find(m => m.id === topId)) return topId;
  }
  const cores = navigator.hardwareConcurrency || 4;
  const mem   = navigator.deviceMemory || 4;
  if (mem <= 4 || cores <= 4)  return 'google/gemini-2.0-flash-001';
  if (mem >= 16 && cores >= 8) return 'google/gemini-2.5-pro-preview';
  return 'meta-llama/llama-3.3-70b-instruct';
}
async function fetchRecommendation() {
  const apiKey = localStorage.getItem('apiKey') || undefined;
  try {
    const r = await fetch('/api/recommend-model', {
      method: 'POST', headers: {'Content-Type':'application/json'},
      body: JSON.stringify({
        specs: collectSpecs(),
        modelUsage,
        preferences: userPreferences,
        availableModels: allModelsFlat().map(m=>({id:m.id,name:m.name,in:m.in,out:m.out,tags:m.tags})),
        apiKey,
      })
    });
    if (r.ok) {
      const rec = await r.json();
      if (rec.modelId && allModelsFlat().find(m => m.id === rec.modelId)) {
        lastRecommendation = rec;
        renderModelList(document.getElementById('modelSearch')?.value || '');
      }
    }
  } catch {}
}
function toggleFilterPopup(e) {
  e.stopPropagation();
  filterPopupOpen = !filterPopupOpen;
  document.getElementById('filterPopup').classList.toggle('open', filterPopupOpen);
  document.getElementById('filterBtn').classList.toggle('active', filterPopupOpen);
}
function closeFilterPopup() {
  filterPopupOpen = false;
  document.getElementById('filterPopup')?.classList.remove('open');
  document.getElementById('filterBtn')?.classList.remove('active');
}
function toggleFilter(tag) {
  activeModelFilter = activeModelFilter === tag ? null : tag;
  ['free','fast','reason','groq'].forEach(t =>
    document.getElementById('fc-'+t).className = 'filter-check' + (activeModelFilter===t?' on':'')
  );
  updateFilterBtnState();
  renderModelList(document.getElementById('modelSearch').value);
}
function setSortMode(mode) {
  modelSortMode = modelSortMode === mode ? null : mode;
  ['asc','desc','free'].forEach(m =>
    document.getElementById('fc-sort-'+m).className = 'filter-check' + (modelSortMode===m?' on':'')
  );
  updateFilterBtnState();
  renderModelList(document.getElementById('modelSearch').value);
}
function updateFilterBtnState() {
  const active = activeModelFilter !== null || modelSortMode !== null;
  document.getElementById('filterBtn').className = 'filter-btn' + (active || filterPopupOpen ? ' active' : '');
}
function sortModels(models) {
  if (modelSortMode === 'asc')  return [...models].sort((a,b) => (a.in+a.out)-(b.in+b.out));
  if (modelSortMode === 'desc') return [...models].sort((a,b) => (b.in+b.out)-(a.in+a.out));
  if (modelSortMode === 'free') return models.filter(m => !m.in && !m.out);
  return models;
}
function filterTag(models) {
  if (!activeModelFilter) return models;
  return models.filter(m => (m.tags||[]).includes(activeModelFilter));
}
function allModelGroups() {
  return [
    { label:'OpenRouter', models:OR_MODELS },
    { label:'Groq',       models:GROQ_MODELS.map(m=>({...m,id:'groq:'+m.id})) },
    ...(customModels.length ? [{ label:'Custom', models:customModels, custom:true }] : []),
    ...(localModels.length  ? [{ label:'Local — Ollama', models:localModels.map(m=>({id:'ollama:'+m.id,name:m.name,provider:'Ollama',in:0,out:0,tags:['local']})) }] : []),
  ];
}
function mkModelOpt(m, isCustom) {
  const el = document.createElement('div');
  el.className = 'model-opt' + (m.id===selectedModel?' sel':'');
  const priceStr = (!m.in&&!m.out) ? (isCustom?'Custom · OpenRouter':'Free') : ('$'+m.in+' / $'+m.out+' per 1M tok');
  const badges = (m.tags||[]).map(t=>'<span class="badge '+t+'">'+t+'</span>').join('');
  el.innerHTML = '<div class="m-info"><div class="m-name">'+esc(m.name)+badges+'</div><div class="m-price">'+(isCustom?esc(m.id):esc(m.provider))+' · '+priceStr+'</div></div>';
  if (isCustom) {
    const del = document.createElement('button');
    del.className='model-opt-del'; del.title='Remove'; del.textContent='✕';
    del.onclick = e => { e.stopPropagation(); removeCustomModel(m.id); };
    el.appendChild(del);
  }
  el.onclick = e => {
    if (e.target.classList.contains('model-opt-del')) return;
    selectedModel=m.id; localStorage.setItem('selModel',m.id);
    addToRecent(m.id); renderModelBtn(); closeModelPicker();
  };
  return el;
}
function renderModelList(query) {
  const list = document.getElementById('modelList');
  list.innerHTML = '';
  const q = (query||'').toLowerCase().trim();

  if (!q) {
    const recId  = getRecommendedModelId();
    const recObj = allModelsFlat().find(m => m.id === recId);
    if (recObj && !activeModelFilter && !modelSortMode) {
      const row = document.createElement('div'); row.className='model-grp-row';
      const lbl = document.createElement('div'); lbl.className='model-grp'; lbl.textContent='Recommended';
      const infoBtn = document.createElement('button'); infoBtn.className='rec-info-btn'; infoBtn.textContent='ⓘ';
      infoBtn.title = lastRecommendation?.reason || 'Loading recommendation…';
      infoBtn.onclick = e => { e.stopPropagation(); toggleRecReason(); };
      row.appendChild(lbl); row.appendChild(infoBtn); list.appendChild(row);
      const reasonBox = document.createElement('div'); reasonBox.className='rec-reason-box'; reasonBox.id='recReasonBox';
      reasonBox.textContent = lastRecommendation?.reason || 'Analysing your system…';
      list.appendChild(reasonBox);
      list.appendChild(mkModelOpt(recObj, customModels.some(c=>c.id===recObj.id)));
    }

    const recentObjs = recentModels
      .filter(id => id !== recId)
      .map(id => allModelsFlat().find(m=>m.id===id)).filter(Boolean);
    if (recentObjs.length && !activeModelFilter && !modelSortMode) {
      const lbl=document.createElement('div'); lbl.className='model-grp'; lbl.style='padding:8px 12px 4px';lbl.textContent='Recently Used';
      list.appendChild(lbl);
      recentObjs.forEach(m=>list.appendChild(mkModelOpt(m, customModels.some(c=>c.id===m.id))));
    }

    for (const grp of allModelGroups()) {
      let models = filterTag(sortModels(grp.models));
      if (!models.length) continue;
      const lbl=document.createElement('div'); lbl.className='model-grp'; lbl.style='padding:8px 12px 4px';lbl.textContent=grp.label;
      list.appendChild(lbl);
      models.forEach(m=>list.appendChild(mkModelOpt(m, grp.custom||false)));
    }
    if (!localModels.length) {
      const d=document.createElement('div');d.className='no-local';d.textContent='No local Ollama models detected.';list.appendChild(d);
    }
    return;
  }

  const presetIds = new Set([...OR_MODELS.map(m=>m.id),...GROQ_MODELS.map(m=>'groq:'+m.id)]);
  const customIds = new Set(customModels.map(m=>m.id));
  let source = catalogLoaded ? catalogModels : allModelsFlat();
  let hits = source.filter(m=>(m.name||'').toLowerCase().includes(q)||(m.id||'').toLowerCase().includes(q));
  hits = filterTag(sortModels(hits));
  if (!hits.length) {
    const d=document.createElement('div');d.className='no-local';
    d.textContent='No models match "'+esc(query)+'".'+(catalogLoaded?'':' (catalog still loading)');
    list.appendChild(d); return;
  }
  hits.forEach(m=>{
    const isPreset=presetIds.has(m.id), isCustom=customIds.has(m.id);
    const el=document.createElement('div');
    el.className='model-opt'+(m.id===selectedModel?' sel':'');
    const priceStr=(!m.in&&!m.out)?'Free':('$'+m.in+' / $'+m.out+' per 1M tok');
    const ctxStr=m.context?' · '+Math.round(m.context/1000)+'k ctx':'';
    const badge=isCustom?'<span class="badge rec">pinned</span>':isPreset?'<span class="badge fast">preset</span>':'';
    el.innerHTML='<div class="m-info"><div class="m-name">'+esc(m.name)+badge+'</div><div class="m-price">'+esc(m.id)+' · '+priceStr+ctxStr+'</div></div>';
    if (isCustom){const del=document.createElement('button');del.className='model-opt-del';del.title='Remove';del.textContent='✕';del.onclick=e=>{e.stopPropagation();removeCustomModel(m.id);};el.appendChild(del);}
    el.onclick=e=>{
      if(e.target.classList.contains('model-opt-del'))return;
      if(!isPreset&&!isCustom){customModels.push({id:m.id,name:m.name,provider:'OpenRouter',in:m.in||0,out:m.out||0,tags:['custom']});localStorage.setItem('customModels',JSON.stringify(customModels));}
      selectedModel=m.id;localStorage.setItem('selModel',m.id);addToRecent(m.id);renderModelBtn();closeModelPicker();
    };
    list.appendChild(el);
  });
  if(!catalogLoaded){const d=document.createElement('div');d.className='no-local';d.textContent='Loading full catalog in background…';list.appendChild(d);}
}
function toggleRecReason() { recReasonOpen = !recReasonOpen; document.getElementById('recReasonBox')?.classList.toggle('open', recReasonOpen); }
async function loadCatalog() {
  if(catalogLoaded)return;
  try{const r=await fetch('/api/models/catalog');if(r.ok){catalogModels=await r.json();catalogLoaded=true;const q=document.getElementById('modelSearch')?.value;if(q)renderModelList(q);}}catch{}
}
function addCustomModel() {
  const inp=document.getElementById('addModelInp');const raw=inp.value.trim();if(!raw)return;
  const id=raw,name=id.includes('/')?id.split('/').pop():id;
  if(customModels.find(m=>m.id===id)){inp.value='';return;}
  const cat=catalogModels.find(m=>m.id===id);
  customModels.push({id,name:cat?.name||name,provider:id.startsWith('groq:')?'Groq':'OpenRouter',in:cat?.in||0,out:cat?.out||0,tags:['custom']});
  localStorage.setItem('customModels',JSON.stringify(customModels));
  inp.value='';renderModelList(document.getElementById('modelSearch').value);renderModelBtn();
}
function removeCustomModel(id) {
  customModels=customModels.filter(m=>m.id!==id);
  localStorage.setItem('customModels',JSON.stringify(customModels));
  if(selectedModel===id){selectedModel=OR_MODELS[0].id;localStorage.setItem('selModel',selectedModel);renderModelBtn();}
  renderModelList(document.getElementById('modelSearch').value);
}
function openModelPicker() {
  activeModelFilter=null; modelSortMode=null; recReasonOpen=false;
  closeFilterPopup();
  ['free','fast','reason','groq','sort-asc','sort-desc','sort-free'].forEach(t=>{const el=document.getElementById('fc-'+t);if(el)el.className='filter-check';});
  updateFilterBtnState();
  document.getElementById('modelModal').classList.add('open');
  const inp=document.getElementById('modelSearch'); inp.value='';
  renderModelList('');
  setTimeout(()=>inp.focus(),50);
  loadCatalog();
  fetchRecommendation();
}
function closeModelPicker() { document.getElementById('modelModal').classList.remove('open'); document.getElementById('modelSearch').value=''; }

// ── Sidebar / Chat history ──────────────────────────────────────────────────
function openSidebar()  { document.getElementById('sidebar').classList.add('open'); }
function closeSidebar() { document.getElementById('sidebar').classList.remove('open'); }
async function loadChats() {
  try{const r=await fetch('/api/chats');const chats=await r.json();renderChatList(chats);if(chats.length&&!chatId)await loadChat(chats[0].id);}catch{}
}
function renderChatList(chats) {
  const el=document.getElementById('chatList');
  if(!chats.length){el.innerHTML='<div class="no-chats">No chats yet</div>';return;}
  el.innerHTML=chats.map(c=>{
    const d=new Date(c.updatedAt).toLocaleDateString(undefined,{month:'short',day:'numeric'});
    return '<div class="chat-item'+(c.id===chatId?' active':'')+'" onclick="chatClick(event,\''+c.id+'\')">'+
      '<div class="chat-title">'+esc(c.title)+'</div>'+
      '<div class="chat-meta">'+d+' · '+(c.messageCount||0)+' msgs</div>'+
      '<button class="chat-del" onclick="delChat(event,\''+c.id+'\')" title="Delete">✕</button>'+
    '</div>';
  }).join('');
}
async function chatClick(e,id){if(!e.target.classList.contains('chat-del'))await loadChat(id);}
async function loadChat(id) {
  try{const r=await fetch('/api/chats/'+id);const c=await r.json();chatId=id;history=c.messages||[];summary=c.summary||null;totalTok=c.totalTokens||0;renderMsgs();updateTok();closeSidebar();refreshChatList();}catch{}
}
async function newChat(doSave=true) {
  if(doSave&&chatId)await saveChat();
  try{const r=await fetch('/api/chats',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({title:'New Chat',messages:[]})});const c=await r.json();chatId=c.id;history=[];summary=null;totalTok=0;renderEmpty();updateTok();closeSidebar();await loadChats();}catch{}
}
async function saveChat(title) {
  if(!chatId)return;
  const t=title||history.find(m=>m.role==='user')?.content?.split('\n')[0]?.slice(0,50)||'New Chat';
  try{await fetch('/api/chats/'+chatId,{method:'PATCH',headers:{'Content-Type':'application/json'},body:JSON.stringify({title:t,messages:history,summary,totalTokens:totalTok})});await loadChats();}catch{}
}
async function delChat(e,id) {
  e.stopPropagation();await fetch('/api/chats/'+id,{method:'DELETE'});
  if(chatId===id){chatId=null;history=[];renderEmpty();}await loadChats();
}
async function clearAllHistory() {
  if(!confirm('Delete all chat history? This cannot be undone.'))return;
  const r=await fetch('/api/chats');const chats=await r.json();
  await Promise.all(chats.map(c=>fetch('/api/chats/'+c.id,{method:'DELETE'})));
  chatId=null;history=[];renderEmpty();closeSettings();await loadChats();
}
async function refreshChatList(){try{const r=await fetch('/api/chats');renderChatList(await r.json());}catch{}}
async function generateTitle(userMsg,aiReply) {
  try{const k=localStorage.getItem('apiKey');const r=await fetch('/api/title',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({userMsg,aiReply,apiKey:k||undefined})});const d=await r.json();return d.title||null;}catch{return null;}
}

// ── Render ──────────────────────────────────────────────────────────────────
function renderEmpty() {
  document.getElementById('messages').innerHTML=
    '<div class="empty-state" id="emptyState">'+
    '<div class="empty-icon">✦</div>'+
    '<div class="empty-title">Word AI Assistant</div>'+
    '<div class="empty-sub">Ask me to write a draft, format a paragraph, design a theme, build a table, run a mail merge — anything Word can do.</div>'+
    '<div class="empty-actions">'+
      '<button class="empty-act-btn" onclick="openDocGenerator()">+ New Document</button>'+
      '<button class="empty-act-btn alt" onclick="openCoach()">✍ Coach me</button>'+
      '<button class="empty-act-btn alt" onclick="openTemplates()">📋 Templates</button>'+
      '<button class="empty-act-btn alt" onclick="openThemePicker()">🎨 Theme</button>'+
    '</div>'+
    '<div class="chips">'+
    '<button class="chip" onclick="useSugg(this)">Make the title bold and 28pt centered</button>'+
    '<button class="chip" onclick="useSugg(this)">Insert a 4×3 table with headers Quarter, Revenue, Notes</button>'+
    '<button class="chip" onclick="useSugg(this)">Design a calm trustworthy theme for a healthcare startup and apply it</button>'+
    '<button class="chip" onclick="useSugg(this)">Replace all &apos;Acme&apos; with &apos;Globex&apos;</button>'+
    '<button class="chip" onclick="useSugg(this)">Add a CONFIDENTIAL watermark to every page</button>'+
    '<button class="chip" onclick="useSugg(this)">How many words does this document have?</button>'+
    '</div></div>';
}
function renderMsgs() {
  const ui=history.filter(m=>m.role==='user'||m.role==='assistant');
  if(!ui.length){renderEmpty();return;}
  const box=document.getElementById('messages'); box.innerHTML='';
  ui.forEach(m=>box.appendChild(mkMsg(m.role,m.content,m.isErr,m.ms,m.tok,m.code,m.time,m.via)));
  box.scrollTop=box.scrollHeight;
}
function addMsg(role,text,isErr=false,ms=null,tok=null,code=null,via=null) {
  document.getElementById('emptyState')?.remove();
  const el=mkMsg(role,text,isErr,ms,tok,code,null,via);
  const box=document.getElementById('messages'); box.appendChild(el); box.scrollTop=box.scrollHeight;
  return el;
}
function mkMsg(role,text,isErr,ms,tok,code,savedTime,via,isCoach) {
  const el=document.createElement('div');
  el.className='msg '+role+(isErr?' error':'')+(isCoach?' coach':'');
  const timeStr=savedTime
    ?new Date(savedTime).toLocaleTimeString(undefined,{hour:'2-digit',minute:'2-digit'})
    :new Date().toLocaleTimeString(undefined,{hour:'2-digit',minute:'2-digit'});
  let meta=timeStr;
  if(role==='assistant'){
    const parts=[];
    if(ms!=null)parts.push((ms/1000).toFixed(1)+'s');
    if(tok!=null)parts.push(tok.toLocaleString()+' tok');
    parts.push(timeStr);
    meta=parts.join(' · ');
  }
  const viaTag = via ? '<span class="msg-via">via '+esc(via)+'</span>' : '';
  const roleLabel = role==='user' ? 'You' : (isCoach ? '🎤 Coach' : 'Word AI');
  const roleClass = role==='user' ? 'user' : (isCoach ? 'coach' : 'assistant');
  el.innerHTML='<div class="msg-head"><span class="msg-role '+roleClass+'">'+roleLabel+'</span>'+viaTag+'<span class="msg-meta">'+meta+'</span></div>'+
    '<div class="msg-body">'+esc(text)+'</div>';
  if(code&&showCode)el.appendChild(mkCode(code));
  const actions=document.createElement('div'); actions.className='msg-actions';
  const copyBtn=document.createElement('button'); copyBtn.className='copy-msg-btn'; copyBtn.innerHTML='⎘ Copy';
  copyBtn.onclick=()=>{navigator.clipboard.writeText(text).then(()=>{copyBtn.innerHTML='✓ Copied';copyBtn.classList.add('copied');setTimeout(()=>{copyBtn.innerHTML='⎘ Copy';copyBtn.classList.remove('copied');},1500);});};
  actions.appendChild(copyBtn);
  if(role==='assistant'&&!isErr&&!isCoach){
    const rb=document.createElement('button');rb.className='report-btn';rb.title='Report';rb.innerHTML='⚑ Report';
    rb.onclick=()=>reportMsg(rb,text,code);actions.appendChild(rb);
  }
  el.appendChild(actions);
  return el;
}
function mkCode(code) {
  const w=document.createElement('div');w.className='code-block';
  const h=document.createElement('div');h.className='code-head';
  h.innerHTML='<span>Generated Code</span><span class="code-toggle">hide ▲</span>';
  const b=document.createElement('div');b.className='code-body';
  const p=document.createElement('pre');p.textContent=code;b.appendChild(p);
  const footer=document.createElement('div');footer.className='code-footer';
  const cb=document.createElement('button');cb.className='copy-code-btn';cb.textContent='⎘ Copy code';
  cb.onclick=()=>{navigator.clipboard.writeText(code).then(()=>{cb.textContent='✓ Copied';cb.classList.add('copied');setTimeout(()=>{cb.textContent='⎘ Copy code';cb.classList.remove('copied');},1500);});};
  footer.appendChild(cb);b.appendChild(footer);
  h.onclick=()=>{const hide=b.classList.toggle('hidden');h.querySelector('.code-toggle').textContent=hide?'show ▼':'hide ▲';};
  w.appendChild(h);w.appendChild(b);return w;
}
function addTyping() {
  document.getElementById('emptyState')?.remove();
  const el=document.createElement('div');el.className='msg assistant';el.id='typing';
  el.innerHTML='<div class="msg-head"><span class="msg-role assistant">Word AI</span><span class="live-timer" id="liveTimer">0.0s</span></div>'+
    '<div class="typing"><div class="tdot"></div><div class="tdot"></div><div class="tdot"></div></div>';
  const box=document.getElementById('messages');box.appendChild(el);box.scrollTop=box.scrollHeight;
  let sec=0;
  timerInterval=setInterval(()=>{sec+=0.1;const t=document.getElementById('liveTimer');if(t)t.textContent=sec.toFixed(1)+'s';},100);
}
function removeTyping(){clearInterval(timerInterval);timerInterval=null;document.getElementById('typing')?.remove();}
function updateTok(){document.getElementById('tokInfo').textContent=totalTok>0?totalTok.toLocaleString()+' tok this chat':'';}

// ── Web search UI ───────────────────────────────────────────────────────────
function addWebSearchPill(parentEl, ws) {
  if (!ws) return;
  const pill = document.createElement('div');
  pill.className = 'ws-pill';
  pill.innerHTML = `🔍 searched the web · "${esc((ws.query||'').slice(0,60))}" · ${esc(ws.provider||'google')} · ${(ws.results||[]).length} results`;
  const panel = document.createElement('div');
  panel.className = 'ws-results-panel';
  panel.innerHTML = (ws.results||[]).map(r =>
    `<div class="ws-result-row"><a href="${esc(r.url)}" target="_blank">${esc(r.title)}</a><div>${esc(r.snippet||'')}</div></div>`
  ).join('') || '<i>No results.</i>';
  pill.onclick = () => panel.classList.toggle('open');
  parentEl.insertBefore(pill, parentEl.firstChild);
  parentEl.appendChild(panel);
}

function addWebSearchPermissionCard(parentEl, wsReq, retryCallback) {
  const card = document.createElement('div');
  card.className = 'ws-perm';
  card.innerHTML = `<div class="ws-perm-title">🔍 Live data needed</div>
    <div>${esc(wsReq.reason || 'The AI needs to look something up online.')}</div>
    <div class="ws-perm-q">${esc(wsReq.query || '')}</div>
    <div class="ws-perm-btns">
      <button class="ws-perm-allow">Allow this search</button>
      <button class="ws-perm-skip">Skip</button>
    </div>`;
  card.querySelector('.ws-perm-allow').onclick = async () => {
    card.querySelector('.ws-perm-allow').disabled = true;
    card.querySelector('.ws-perm-allow').textContent = 'Searching…';
    try {
      const r = await fetch('/api/web-search', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ query: wsReq.query, count: 5 }) });
      if (!r.ok) { card.querySelector('.ws-perm-allow').textContent = 'Search failed'; return; }
      const ws = await r.json();
      pendingWebSearchResults.push(ws);
      card.remove();
      retryCallback();
    } catch (e) { card.querySelector('.ws-perm-allow').textContent = 'Failed: ' + e.message; }
  };
  card.querySelector('.ws-perm-skip').onclick = () => card.remove();
  parentEl.appendChild(card);
}

// ── Sync ────────────────────────────────────────────────────────────────────
function setSyncStatus(state,text){const d=document.getElementById('syncDot');d.className='sync-dot'+(state==='busy'?' busy':state==='err'?' err':'');d.title=text;}

async function readDoc() {
  setSyncStatus('busy','Reading document…');
  try {
    await Word.run(async ctx => {
      const body = ctx.document.body;
      body.load('text');
      ctx.document.sections.load('items/getRange');
      const paras = body.paragraphs;
      paras.load('items/text, items/styleBuiltIn, items/style');
      const sel = ctx.document.getSelection();
      sel.load('text');
      await ctx.sync();

      const text = body.text || '';
      const words = (text.match(/\b[\w']+\b/g) || []).length;
      const sentences = (text.match(/[^.!?]+[.!?]+/g) || []).length;
      const flesch = words ? window.WordHelpers.readability(text) : null;

      // Build a heading + paragraph outline (max 40 items for context)
      const items = paras.items || [];
      const MAX = 40;
      const lines = [];
      for (let i = 0; i < Math.min(items.length, MAX); i++) {
        const p = items[i];
        const t = (p.text || '').slice(0, 200);
        if (!t.trim()) continue;
        const sBuilt = (p.styleBuiltIn || '').toString();
        const sCust  = (p.style || '').toString();
        const styleStr = sBuilt || sCust || 'Normal';
        lines.push(`P${i}: STYLE=${styleStr} | TEXT=${JSON.stringify(t)}`);
      }
      if (items.length > MAX) lines.push(`[...${items.length - MAX} more paragraphs not shown]`);

      docData = `Document has ${items.length} paragraphs, ${words} words.\nFlesch reading-ease: ${flesch ? flesch.flesch : 'n/a'} (grade ${flesch ? flesch.grade : 'n/a'}).\n\nOutline:\n${lines.join('\n')}`;
      selectionInfo = (sel.text || '').slice(0, 200) ? `selected "${(sel.text || '').slice(0,80)}…"` : '(none — cursor only)';

      // Update doc stats badge
      if (liveStats) {
        document.getElementById('docStats').textContent = `${words.toLocaleString()}w · grade ${flesch ? flesch.grade : '—'}`;
      }
      setSyncStatus('ok', `Synced — ${items.length} paragraphs, ${words} words`);
    });
  } catch (e) {
    setSyncStatus('err','Sync failed');
    console.error('readDoc error', e);
  }
}
async function registerListeners() {
  if (changeReady) return;
  try {
    await Word.run(async ctx => {
      try {
        ctx.document.onSelectionChanged.add(() => {
          if (autoSync) { clearTimeout(syncTimer); syncTimer = setTimeout(() => readDoc(), 1500); }
        });
      } catch {}
      await ctx.sync();
      changeReady = true;
    });
  } catch {}
}

/* ───────────────────────────── execCode() ────────────────────────────────────
 * Wraps an AI-generated body into Word.run, injecting every helper named in the
 * system prompt. Each helper is async and uses the captured `ctx` for the
 * Office.js calls. */
async function execCode(code) {
  let ret = null;
  try {
    await Word.run(async ctx => {
      const doc = ctx.document;
      const OOXML     = window.WordHelpers.OOXML;
      const TEMPLATES = window.WordHelpers.TEMPLATES;
      const packageOoxml = window.WordHelpers.packageOoxml;
      const sanitizeFont = window.WordHelpers.sanitizeFont;
      const latexToOmml  = window.WordHelpers.latexToOmml;
      const HelperUtils  = window.WordHelpers;

      // ── Internals ────────────────────────────────────────────────────────
      function safeNum(v, dflt) { const n = Number(v); return Number.isFinite(n) ? n : dflt; }
      function alignToWord(a) {
        const map = { left:'Left', right:'Right', center:'Centered', centered:'Centered', justify:'Justified', justified:'Justified' };
        return map[(a||'').toLowerCase()] || 'Left';
      }

      // ── Core insertion helpers ───────────────────────────────────────────
      async function addParagraph(text, opts) {
        opts = opts || {};
        const p = doc.body.insertParagraph(String(text == null ? '' : text), 'End');
        await applyParagraphOpts(p, opts);
        await ctx.sync();
        return p;
      }
      async function applyParagraphOpts(p, opts) {
        opts = opts || {};
        try { if (opts.alignment) p.alignment = alignToWord(opts.alignment); } catch {}
        try { if (opts.spaceBefore != null) p.spaceBefore = safeNum(opts.spaceBefore, 0); } catch {}
        try { if (opts.spaceAfter  != null) p.spaceAfter  = safeNum(opts.spaceAfter, 0); } catch {}
        try { if (opts.lineSpacing != null) p.lineSpacing = safeNum(opts.lineSpacing, 1) * 12; } catch {}
        try { if (opts.indent != null) p.leftIndent = safeNum(opts.indent, 0); } catch {}
        try {
          if (opts.bold)    p.font.bold = true;
          if (opts.italic)  p.font.italic = true;
          if (opts.size)    p.font.size = safeNum(opts.size, 11);
          if (opts.color)   p.font.color = String(opts.color);
          if (opts.fontName)p.font.name  = sanitizeFont(opts.fontName).font;
          if (opts.underline) p.font.underline = 'Single';
        } catch {}
      }

      async function addHeading(text, level, opts) {
        opts = opts || {};
        const lv = Math.max(1, Math.min(9, safeNum(level, 1)));
        const p = doc.body.insertParagraph(String(text == null ? '' : text), 'End');
        try { p.styleBuiltIn = 'Heading' + lv; } catch {}
        await applyParagraphOpts(p, opts);
        await ctx.sync();
        return p;
      }
      async function addTitle(text, opts) {
        const p = doc.body.insertParagraph(String(text == null ? '' : text), 'End');
        try { p.styleBuiltIn = 'Title'; } catch {}
        await applyParagraphOpts(p, opts || {});
        await ctx.sync();
        return p;
      }
      async function addSubtitle(text, opts) {
        const p = doc.body.insertParagraph(String(text == null ? '' : text), 'End');
        try { p.styleBuiltIn = 'Subtitle'; } catch {}
        await applyParagraphOpts(p, opts || {});
        await ctx.sync();
        return p;
      }
      async function addQuote(text, opts) {
        const p = doc.body.insertParagraph(String(text == null ? '' : text), 'End');
        try { p.styleBuiltIn = 'IntenseQuote'; } catch { try { p.styleBuiltIn = 'Quote'; } catch {} }
        await applyParagraphOpts(p, opts || {});
        await ctx.sync();
        return p;
      }
      async function addList(items, type, opts) {
        // Forgiving: accept a single string, a CSV-ish string, or an array.
        if (typeof items === 'string') {
          items = items.includes('\n') ? items.split(/\n+/) : items.split(/,\s*/);
        }
        if (!Array.isArray(items)) items = [String(items)];
        items = items.map(x => String(x).trim()).filter(Boolean);
        // Forgiving: type can be "bullet"/"number"/"numbered"/"ol"/"ul" or even omitted.
        const t = String(type || 'bullet').toLowerCase();
        const isNumber = /^(num|number|numbered|ol|ordered|1|decimal)/.test(t);
        opts = opts || {};
        const built = isNumber ? 'ListNumber' : 'ListBullet';
        const lv = Math.max(0, Math.min(8, safeNum(opts.level, 0)));
        for (const it of items) {
          const p = doc.body.insertParagraph(String(it), 'End');
          try { p.styleBuiltIn = built + (lv > 0 ? (lv + 1) : ''); } catch { try { p.styleBuiltIn = built; } catch {} }
          await applyParagraphOpts(p, opts);
        }
        await ctx.sync();
      }

      async function applyStyle(target, styleName) {
        // Forgiving: if called with a single arg, treat it as the style name on the selection.
        if (styleName === undefined && typeof target === 'string') {
          styleName = target;
          target = 'selection';
        }
        // Forgiving: tolerate "Heading 1" / "heading-1" → "Heading1".
        if (typeof styleName === 'string') {
          styleName = styleName.trim().replace(/[\s_-]+/g, '');
          styleName = styleName.charAt(0).toUpperCase() + styleName.slice(1);
        }
        if (target === 'selection') {
          const sel = doc.getSelection();
          try { sel.styleBuiltIn = styleName; } catch { sel.style = styleName; }
        } else if (target === 'all-headings') {
          doc.body.paragraphs.load('items/styleBuiltIn');
          await ctx.sync();
          for (const p of doc.body.paragraphs.items) {
            try {
              const s = String(p.styleBuiltIn || '');
              if (s.startsWith('Heading')) { try { p.styleBuiltIn = styleName; } catch { p.style = styleName; } }
            } catch {}
          }
        } else if (typeof target === 'number') {
          doc.body.paragraphs.load('items');
          await ctx.sync();
          const p = doc.body.paragraphs.items[target];
          if (!p) throw new Error('Paragraph index ' + target + ' out of range.');
          try { p.styleBuiltIn = styleName; } catch { p.style = styleName; }
        }
        await ctx.sync();
      }

      // ── Find & replace ───────────────────────────────────────────────────
      async function findText(query, opts) {
        opts = opts || {};
        const results = doc.body.search(String(query), {
          matchCase: !!opts.matchCase,
          matchWholeWord: !!opts.matchWholeWord,
          matchWildcards: !!opts.matchWildcards,
        });
        results.load('text');
        await ctx.sync();
        return results.items.map(r => ({ match: r.text }));
      }
      async function replaceText(find, replace, opts) {
        opts = opts || {};
        if (opts.regex) {
          // Client-side regex: replace within the full body text
          doc.body.load('text');
          await ctx.sync();
          const re = new RegExp(opts.regex.source || opts.regex, opts.regex.flags || (opts.matchCase ? 'g' : 'gi'));
          const newText = doc.body.text.replace(re, String(replace));
          doc.body.clear();
          doc.body.insertText(newText, 'End');
          await ctx.sync();
          return 'Regex replace done.';
        }
        const results = doc.body.search(String(find), {
          matchCase: !!opts.matchCase,
          matchWholeWord: !!opts.matchWholeWord,
          matchWildcards: !!opts.matchWildcards,
        });
        results.load('text');
        await ctx.sync();
        const n = results.items.length;
        results.items.forEach(r => r.insertText(String(replace), 'Replace'));
        await ctx.sync();
        return 'Replaced ' + n + ' occurrence(s).';
      }

      // ── Breaks / layout ──────────────────────────────────────────────────
      async function insertPageBreak() {
        doc.body.insertOoxml(packageOoxml(OOXML.pageBreak), 'End');
        await ctx.sync();
      }
      async function insertSectionBreak(kind) {
        doc.body.insertOoxml(packageOoxml(OOXML.sectionBreak(kind)), 'End');
        await ctx.sync();
      }
      async function insertColumns(count) {
        doc.body.insertOoxml(packageOoxml(OOXML.columns(count)), 'End');
        await ctx.sync();
      }
      async function setMargins(margins) {
        margins = margins || {};
        try {
          const sec = doc.sections.getFirst();
          // Word.js doesn't expose margins yet for all surfaces — fall back to OOXML for the section.
          const t = safeNum(margins.top,    72) * 20;
          const b = safeNum(margins.bottom, 72) * 20;
          const l = safeNum(margins.left,   72) * 20;
          const r = safeNum(margins.right,  72) * 20;
          const xml = `<w:p><w:pPr><w:sectPr><w:pgMar w:top="${t}" w:right="${r}" w:bottom="${b}" w:left="${l}" w:header="720" w:footer="720" w:gutter="0"/><w:type w:val="continuous"/></w:sectPr></w:pPr></w:p>`;
          doc.body.insertOoxml(packageOoxml(xml), 'End');
        } catch (e) { throw new Error('setMargins failed: ' + e.message); }
        await ctx.sync();
      }
      async function setPageOrientation(orient) {
        const xml = (String(orient || '').toLowerCase() === 'landscape') ? OOXML.landscape : OOXML.portrait;
        doc.body.insertOoxml(packageOoxml(xml), 'End');
        await ctx.sync();
      }

      // ── Headers / footers / page numbers ─────────────────────────────────
      async function setHeader(text, opts) {
        opts = opts || {};
        const section = doc.sections.getFirst();
        const which = ({ first:'FirstPage', even:'EvenPages', primary:'Primary' })[String(opts.kind || 'primary').toLowerCase()] || 'Primary';
        const header = section.getHeader(which);
        header.clear();
        const p = header.insertParagraph(String(text || ''), 'End');
        if (opts.alignment) try { p.alignment = alignToWord(opts.alignment); } catch {}
        if (opts.font) try { p.font.name = sanitizeFont(opts.font).font; } catch {}
        if (opts.size) try { p.font.size = safeNum(opts.size, 10); } catch {}
        if (opts.includePageNumber) {
          const sep = opts.includeTotalPages ? ' Page ' : ' Page ';
          header.insertText(sep, 'End');
          header.insertOoxml(packageOoxml(OOXML.pageNumberField), 'End');
          if (opts.includeTotalPages) {
            header.insertText(' of ', 'End');
            header.insertOoxml(packageOoxml(OOXML.totalPagesField), 'End');
          }
        }
        await ctx.sync();
      }
      async function setFooter(text, opts) {
        opts = opts || {};
        const section = doc.sections.getFirst();
        const which = ({ first:'FirstPage', even:'EvenPages', primary:'Primary' })[String(opts.kind || 'primary').toLowerCase()] || 'Primary';
        const footer = section.getFooter(which);
        footer.clear();
        const p = footer.insertParagraph(String(text || ''), 'End');
        if (opts.alignment) try { p.alignment = alignToWord(opts.alignment); } catch {}
        if (opts.font) try { p.font.name = sanitizeFont(opts.font).font; } catch {}
        if (opts.size) try { p.font.size = safeNum(opts.size, 10); } catch {}
        if (opts.includePageNumber) {
          footer.insertText(' Page ', 'End');
          footer.insertOoxml(packageOoxml(OOXML.pageNumberField), 'End');
          if (opts.includeTotalPages) {
            footer.insertText(' of ', 'End');
            footer.insertOoxml(packageOoxml(OOXML.totalPagesField), 'End');
          }
        }
        await ctx.sync();
      }
      async function addPageNumbers(opts) {
        opts = opts || {};
        const where = String(opts.location || 'footer').toLowerCase() === 'header' ? 'getHeader' : 'getFooter';
        const align = alignToWord(opts.alignment || 'center');
        const section = doc.sections.getFirst();
        const part = section[where]('Primary');
        part.clear();
        const p = part.insertParagraph('', 'End');
        try { p.alignment = align; } catch {}
        part.insertOoxml(packageOoxml(OOXML.pageNumberField), 'End');
        await ctx.sync();
      }

      // ── Table of Contents ────────────────────────────────────────────────
      async function insertTableOfContents(opts) {
        opts = opts || {};
        doc.body.insertOoxml(packageOoxml(OOXML.toc(opts)), 'Start');
        await ctx.sync();
        return 'TOC inserted. If it shows the placeholder text, right-click the field and choose "Update Field".';
      }

      // ── Tables ───────────────────────────────────────────────────────────
      async function insertTable(rows, cols, values, opts) {
        opts = opts || {};
        const r = safeNum(rows, 2), c = safeNum(cols, 2);
        const data = Array.isArray(values) ? values.map(row => Array.isArray(row) ? row.map(String) : [String(row)]) : null;
        const tbl = doc.body.insertTable(r, c, 'End', data);
        try { if (opts.style) tbl.styleBuiltIn = opts.style; } catch { try { tbl.style = opts.style; } catch {} }
        try { if (opts.headerRow) tbl.headerRowCount = 1; } catch {}
        await ctx.sync();
        return tbl;
      }
      async function setTableCell(tableIndex, row, col, text) {
        doc.body.tables.load('items');
        await ctx.sync();
        const t = doc.body.tables.items[tableIndex];
        if (!t) throw new Error('Table index ' + tableIndex + ' not found.');
        const cell = t.getCell(row, col);
        cell.value = String(text);
        await ctx.sync();
      }
      async function styleTable(tableIndex, styleName) {
        doc.body.tables.load('items');
        await ctx.sync();
        const t = doc.body.tables.items[tableIndex];
        if (!t) throw new Error('Table index ' + tableIndex + ' not found.');
        try { t.styleBuiltIn = styleName; } catch { t.style = styleName; }
        await ctx.sync();
      }
      async function mergeCellsBestEffort(/* tableIndex, fromRow, fromCol, toRow, toCol */) {
        // Word.js doesn't expose cell merging on all builds. Document this as a no-op
        // and let the user know.
        throw new Error('Cell merging is not supported by Word Office.js. Add an empty header row instead, or merge manually after applying the table.');
      }

      // ── Images ───────────────────────────────────────────────────────────
      async function insertImage(category, keywordOrTags, opts) {
        const tags = Array.isArray(keywordOrTags) ? keywordOrTags.join(',') : String(keywordOrTags || '');
        const r = await fetch('/api/assets/pick?category=' + encodeURIComponent(category) + '&keywords=' + encodeURIComponent(tags));
        if (!r.ok) {
          const err = await r.json().catch(() => ({ error: 'unknown' }));
          throw new Error('Image pick failed: ' + (err.error || r.status));
        }
        const img = await r.json();
        opts = opts || {};
        const range = doc.getSelection();
        const pic = range.insertInlinePictureFromBase64(img.base64, 'After');
        try {
          if (opts.width)  pic.width  = safeNum(opts.width, 320);
          if (opts.height) pic.height = safeNum(opts.height, 0) || undefined;
        } catch {}
        // Alignment: wrap by aligning the paragraph
        try {
          const para = range.paragraphs.getFirst();
          if (opts.alignment) para.alignment = alignToWord(opts.alignment);
        } catch {}
        await ctx.sync();
        return pic;
      }

      // ── Watermark ────────────────────────────────────────────────────────
      async function insertWatermark(text, opts) {
        opts = opts || {};
        // Put a faint centered watermark line into every section's primary header.
        // This is the most reliable way to get an on-every-page watermark via Word.js.
        const sections = doc.sections;
        sections.load('items');
        await ctx.sync();
        for (const s of sections.items) {
          try {
            const h = s.getHeader('Primary');
            h.insertOoxml(packageOoxml(OOXML.watermark(text || 'CONFIDENTIAL', opts.color || '#C0C0C0')), 'End');
          } catch {}
        }
        await ctx.sync();
        return 'Inserted watermark on every page header.';
      }

      // ── Footnotes / endnotes / comments / track changes ──────────────────
      async function insertFootnote(text) {
        const sel = doc.getSelection();
        sel.insertFootnote(String(text || ''));
        await ctx.sync();
      }
      async function insertEndnote(text) {
        const sel = doc.getSelection();
        try { sel.insertEndnote(String(text || '')); }
        catch { sel.insertFootnote(String(text || '')); } // some builds lack insertEndnote
        await ctx.sync();
      }
      async function insertComment(targetQuery, text) {
        const results = doc.body.search(String(targetQuery), { matchCase: false });
        results.load('text');
        await ctx.sync();
        if (!results.items.length) throw new Error('Could not find "' + targetQuery + '" to comment on.');
        try { results.items[0].insertComment(String(text)); }
        catch (e) { throw new Error('Comments need Word Online or desktop 1.4+: ' + e.message); }
        await ctx.sync();
      }
      async function toggleTrackChanges(on) {
        // Forgiving: accept truthy/falsy, 'on'/'off', 'true'/'false', undefined → on.
        const truthy = on === undefined ? true
                     : typeof on === 'string'
                       ? !/^(off|false|no|0)$/i.test(on)
                       : !!on;
        try { doc.changeTrackingMode = truthy ? 'TrackAll' : 'Off'; }
        catch (e) { throw new Error('Track changes not available: ' + e.message); }
        await ctx.sync();
      }
      async function acceptAllRevisions() {
        try { doc.body.getRange().acceptAllRevisions ? doc.body.getRange().acceptAllRevisions() : null; } catch {}
        try { if (doc.body.acceptAllRevisions) doc.body.acceptAllRevisions(); } catch {}
        // Fall back to revisions collection if available
        try {
          if (doc.revisions) { doc.revisions.load('items'); await ctx.sync(); for (const r of doc.revisions.items) r.accept && r.accept(); }
        } catch {}
        await ctx.sync();
      }
      async function rejectAllRevisions() {
        try { if (doc.body.rejectAllRevisions) doc.body.rejectAllRevisions(); } catch {}
        try {
          if (doc.revisions) { doc.revisions.load('items'); await ctx.sync(); for (const r of doc.revisions.items) r.reject && r.reject(); }
        } catch {}
        await ctx.sync();
      }

      // ── Citations & bibliography ────────────────────────────────────────
      const _biblio = (window.__wordAiBibliography ||= []);
      function formatInline(source, style) {
        const author = String(source.author || 'Anon').split(/\s+/).slice(-1)[0]; // last name
        const year   = String(source.year   || 'n.d.');
        const page   = source.page ? String(source.page) : '';
        switch ((style || 'APA').toUpperCase()) {
          case 'MLA':     return page ? `(${author} ${page})` : `(${author})`;
          case 'CHICAGO': return page ? `(${author} ${year}, ${page})` : `(${author} ${year})`;
          case 'HARVARD': return page ? `(${author} ${year}: ${page})` : `(${author} ${year})`;
          case 'APA':
          default:        return `(${author}, ${year})`;
        }
      }
      function formatEntry(source, style) {
        const a = String(source.author || 'Anonymous');
        const y = String(source.year   || 'n.d.');
        const t = String(source.title  || 'Untitled');
        const s = String(source.source || '');
        switch ((style || 'APA').toUpperCase()) {
          case 'MLA':     return `${a}. "${t}." ${s}, ${y}.`;
          case 'CHICAGO': return `${a}. ${y}. "${t}." ${s}.`;
          case 'HARVARD': return `${a} (${y}) ${t}. ${s}.`;
          case 'APA':
          default:        return `${a}. (${y}). ${t}. ${s}.`;
        }
      }
      async function insertCitation(source) {
        const style = source.style || 'APA';
        const inline = formatInline(source, style);
        const sel = doc.getSelection();
        sel.insertText(inline, 'After');
        _biblio.push({ ...source, style });
        await ctx.sync();
        return inline;
      }
      async function insertBibliography(style) {
        style = style || 'APA';
        const entries = _biblio.length ? _biblio : [];
        await addHeading('References', 1);
        if (!entries.length) {
          await addParagraph('(No citations recorded yet — insertCitation() first.)');
          return;
        }
        for (const e of entries) {
          const p = doc.body.insertParagraph(formatEntry(e, style), 'End');
          try { p.styleBuiltIn = 'Bibliography'; } catch {}
        }
        await ctx.sync();
      }

      // ── Equation ────────────────────────────────────────────────────────
      async function insertEquation(latex) {
        const omml = window.WordHelpers.latexToOmml(latex);
        const xml  = `<w:p>${omml}</w:p>`;
        doc.body.insertOoxml(packageOoxml(xml), 'End');
        await ctx.sync();
      }

      // ── Content controls / form fields ──────────────────────────────────
      async function insertContentControl(kind, opts) {
        opts = opts || {};
        // Forgiving: normalize common kind aliases.
        const k = String(kind || 'richText').toLowerCase().replace(/[\s_-]+/g, '');
        const ALIAS = {
          rich:'richText', richtext:'richText',
          plain:'plainText', plaintext:'plainText', text:'plainText',
          check:'checkbox', checkbox:'checkbox',
          drop:'dropdown', dropdown:'dropdown', dropdownlist:'dropdown',
          combo:'comboBox', combobox:'comboBox',
          date:'datePicker', datepicker:'datePicker',
        };
        kind = ALIAS[k] || kind;
        const sel = doc.getSelection();
        if ((kind || 'richText').toLowerCase() === 'richtext' || (kind || '').toLowerCase() === 'rich-text') {
          // Native Word.js content control
          try {
            const cc = sel.insertContentControl();
            if (opts.title) cc.title = opts.title;
            if (opts.tag) cc.tag = opts.tag;
            if (opts.placeholder) cc.placeholderText = opts.placeholder;
            await ctx.sync();
            return cc;
          } catch (e) { /* fall through to OOXML */ }
        }
        const xml = `<w:p>${OOXML.contentControl(kind, opts)}</w:p>`;
        doc.body.insertOoxml(packageOoxml(xml), 'End');
        await ctx.sync();
      }
      async function insertFormField(kind, opts) {
        return insertContentControl(kind, opts);
      }

      // ── Mail merge ──────────────────────────────────────────────────────
      async function mailMergeReplace(records) {
        records = Array.isArray(records) ? records : [];
        if (!records.length) throw new Error('mailMergeReplace: no records provided.');
        // Strategy: capture the body OOXML, then for each record, append a copy
        // with placeholders replaced. For the first record, replace in place.
        doc.body.load('text');
        await ctx.sync();
        const originalText = doc.body.text;
        // Find unique placeholders
        const placeholders = new Set();
        const re1 = /\{\{([A-Za-z0-9_]+)\}\}/g;
        const re2 = /«([A-Za-z0-9_]+)»/g;
        let m;
        while ((m = re1.exec(originalText))) placeholders.add(m[1]);
        while ((m = re2.exec(originalText))) placeholders.add(m[1]);

        for (let i = 0; i < records.length; i++) {
          const rec = records[i];
          // For i > 0, insert a page break and re-add the original template text
          if (i > 0) {
            doc.body.insertOoxml(packageOoxml(OOXML.pageBreak), 'End');
            doc.body.insertText(originalText, 'End');
            await ctx.sync();
          }
          // Replace placeholders for THIS record on the LAST chunk
          for (const ph of placeholders) {
            const val = rec[ph] != null ? String(rec[ph]) : '';
            // First record: replace globally in doc; subsequent: only newly appended chunk
            const searchPatterns = [`{{${ph}}}`, `«${ph}»`];
            for (const pat of searchPatterns) {
              const res = doc.body.search(pat, { matchCase: true });
              res.load('text');
              await ctx.sync();
              // For i>0, only replace from the latest set of matches (LAST occurrences).
              const targets = i === 0 ? res.items : res.items.slice(-1);
              for (const r of targets) r.insertText(val, 'Replace');
            }
            await ctx.sync();
          }
        }
        return 'Merged ' + records.length + ' record(s).';
      }

      // ── Theme designer ──────────────────────────────────────────────────
      async function designTheme(description) {
        const r = await fetch('/api/design-theme', {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ description: String(description), apiKey: localStorage.getItem('apiKey') || undefined })
        });
        if (!r.ok) throw new Error('Theme generation failed (' + r.status + ').');
        const d = await r.json();
        if (d.error) throw new Error('Theme generation failed: ' + d.error);
        if (d.warnings && d.warnings.length) console.warn('[designTheme]', d.warnings);
        return d.theme;
      }
      async function tweakTheme(prev, instruction) {
        const r = await fetch('/api/design-theme', {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ description: String(instruction), previous: prev, apiKey: localStorage.getItem('apiKey') || undefined })
        });
        if (!r.ok) throw new Error('Theme tweak failed.');
        const d = await r.json();
        if (d.error) throw new Error('Theme tweak failed: ' + d.error);
        return d.theme;
      }
      async function recolorDocument(palette) {
        palette = palette || {};
        const headingFont = sanitizeFont(palette.headingFont || 'Calibri Light').font;
        const bodyFont    = sanitizeFont(palette.bodyFont    || 'Calibri').font;
        doc.body.paragraphs.load('items/styleBuiltIn');
        await ctx.sync();
        let touched = 0;
        for (const p of doc.body.paragraphs.items) {
          try {
            const style = String(p.styleBuiltIn || '');
            const isHeading = style.startsWith('Heading') || style === 'Title' || style === 'Subtitle';
            if (palette.text) p.font.color = isHeading ? (palette.primary || palette.text) : palette.text;
            p.font.name = isHeading ? headingFont : bodyFont;
            if (isHeading && palette.headingSize) p.font.size = Math.max(12, Math.min(48, palette.headingSize));
            touched++;
          } catch {}
        }
        await ctx.sync();
        return 'Recolored ' + touched + ' paragraph(s).';
      }
      async function applyTheme(themeOrName) {
        // If a string came in, it was probably a user typo — generate it.
        if (typeof themeOrName === 'string') {
          themeOrName = await designTheme(themeOrName);
        }
        if (!themeOrName || typeof themeOrName !== 'object') throw new Error('applyTheme expects a theme object.');
        const summary = await recolorDocument(themeOrName);
        // persist as current theme
        try { localStorage.setItem('activeTheme', JSON.stringify(themeOrName)); } catch {}
        currentTheme = themeOrName;
        return summary + ' Theme: "' + (themeOrName.name || 'custom') + '".';
      }

      // ── Templates ───────────────────────────────────────────────────────
      function expandPlaceholders(s, fields) {
        return String(s).replace(/\{\{([A-Za-z0-9_]+)\}\}/g, (_, k) => (fields[k] != null ? String(fields[k]) : ''));
      }
      async function applyTemplate(id, fields) {
        const T = TEMPLATES[id];
        if (!T) throw new Error('Unknown template "' + id + '". Available: ' + Object.keys(TEMPLATES).join(', '));
        fields = fields || {};
        for (const sec of T.sections) {
          const text = expandPlaceholders(sec.text, fields);
          switch (sec.type) {
            case 'title':      await addTitle(text); break;
            case 'subtitle':   await addSubtitle(text); break;
            case 'heading1':   await addHeading(text, 1); break;
            case 'heading2':   await addHeading(text, 2); break;
            case 'heading3':   await addHeading(text, 3); break;
            case 'paragraph':  await addParagraph(text); break;
            case 'quote':      await addQuote(text); break;
            case 'spacer':     await addParagraph(''); break;
            case 'pageBreak':  await insertPageBreak(); break;
            case 'signature':  {
              const p = doc.body.insertParagraph('_________________________', 'End');
              await ctx.sync();
              const p2 = doc.body.insertParagraph(text, 'End');
              try { p2.font.italic = true; } catch {}
              await ctx.sync();
              break;
            }
            default:           await addParagraph(text);
          }
        }
        return 'Applied template ' + T.name + '.';
      }

      // ── Introspection helpers ───────────────────────────────────────────
      async function countWords() {
        doc.body.load('text');
        await ctx.sync();
        return (doc.body.text.match(/\b[\w']+\b/g) || []).length;
      }
      async function countCharacters() {
        doc.body.load('text');
        await ctx.sync();
        return doc.body.text.length;
      }
      async function countParagraphs() {
        doc.body.paragraphs.load('items');
        await ctx.sync();
        return doc.body.paragraphs.items.length;
      }
      async function getReadability() {
        doc.body.load('text');
        await ctx.sync();
        return window.WordHelpers.readability(doc.body.text);
      }
      async function getSelectionText() {
        const sel = doc.getSelection();
        sel.load('text');
        await ctx.sync();
        return sel.text || '';
      }
      async function getSelectionHtml() {
        const sel = doc.getSelection();
        const html = sel.getHtml();
        await ctx.sync();
        return html.value;
      }
      async function listSections() {
        doc.sections.load('items');
        await ctx.sync();
        return doc.sections.items.map((_, i) => ({ index: i }));
      }
      async function listHeadings() {
        doc.body.paragraphs.load('items/text, items/styleBuiltIn');
        await ctx.sync();
        return doc.body.paragraphs.items
          .map((p, i) => ({ index: i, text: p.text, style: String(p.styleBuiltIn || '') }))
          .filter(p => p.style.startsWith('Heading') || p.style === 'Title');
      }
      async function getDocumentSummary() {
        const w = await countWords();
        const p = await countParagraphs();
        const r = await getReadability();
        const h = await listHeadings();
        return { wordCount: w, paragraphCount: p, readability: r, headings: h };
      }

      // ── Common-mistake aliases — if the AI invents an "intuitive" helper name,
      // route it to the real one rather than throwing ReferenceError. ──
      async function setBold(on) {
        const sel = doc.getSelection(); sel.font.bold = on === false ? false : true; await ctx.sync();
      }
      async function setItalic(on) {
        const sel = doc.getSelection(); sel.font.italic = on === false ? false : true; await ctx.sync();
      }
      async function setUnderline(on) {
        const sel = doc.getSelection(); sel.font.underline = on === false ? 'None' : 'Single'; await ctx.sync();
      }
      async function setFontColor(color) {
        const sel = doc.getSelection(); sel.font.color = String(color); await ctx.sync();
      }
      async function setFontSize(size) {
        const sel = doc.getSelection(); sel.font.size = safeNum(size, 11); await ctx.sync();
      }
      async function setFontName(name) {
        const sel = doc.getSelection(); sel.font.name = sanitizeFont(name).font; await ctx.sync();
      }
      async function setAlignment(a) {
        const sel = doc.getSelection(); const p = sel.paragraphs.getFirst();
        try { p.alignment = alignToWord(a); } catch {} await ctx.sync();
      }
      const insertHeading = addHeading;
      const insertParagraph = addParagraph;
      const insertText = async (t) => { doc.getSelection().insertText(String(t || ''), 'After'); await ctx.sync(); };
      const insertBulletList = (items, opts) => addList(items, 'bullet', opts);
      const insertNumberedList = (items, opts) => addList(items, 'number', opts);
      const insertList = addList;
      const find = findText;
      const replaceAll = replaceText;
      const setTrackChanges = toggleTrackChanges;
      const acceptRevisions = acceptAllRevisions;
      const rejectRevisions = rejectAllRevisions;
      const insertTOC = insertTableOfContents;
      const insertToc = insertTableOfContents;
      const setOrientation = setPageOrientation;

      // ── Execute AI code ─────────────────────────────────────────────────
      const fn = new Function(
        'context','doc','Word',
        'addHeading','addParagraph','addTitle','addSubtitle','addQuote','addList','applyStyle',
        'replaceText','findText','insertPageBreak','insertSectionBreak','insertColumns','setHeader',
        'setFooter','addPageNumbers','insertTableOfContents','setMargins','setPageOrientation',
        'insertTable','setTableCell','styleTable','mergeCellsBestEffort','insertImage','insertWatermark',
        'insertFootnote','insertEndnote','insertComment','toggleTrackChanges','acceptAllRevisions','rejectAllRevisions',
        'insertCitation','insertBibliography','insertEquation','insertContentControl','insertFormField',
        'mailMergeReplace','applyTheme','recolorDocument','tweakTheme','designTheme',
        'countWords','countCharacters','countParagraphs','getReadability',
        'getSelectionText','getSelectionHtml','listSections','listHeadings','getDocumentSummary',
        'applyTemplate','activeTheme','TEMPLATES','OOXML',
        // forgiving aliases
        'setBold','setItalic','setUnderline','setFontColor','setFontSize','setFontName','setAlignment',
        'insertHeading','insertParagraph','insertText','insertBulletList','insertNumberedList','insertList',
        'find','replaceAll','setTrackChanges','acceptRevisions','rejectRevisions','insertTOC','insertToc','setOrientation',
        'return (async()=>{ ' + code + ' })();'
      );
      const rv = await fn(
        ctx, doc, Word,
        addHeading, addParagraph, addTitle, addSubtitle, addQuote, addList, applyStyle,
        replaceText, findText, insertPageBreak, insertSectionBreak, insertColumns, setHeader,
        setFooter, addPageNumbers, insertTableOfContents, setMargins, setPageOrientation,
        insertTable, setTableCell, styleTable, mergeCellsBestEffort, insertImage, insertWatermark,
        insertFootnote, insertEndnote, insertComment, toggleTrackChanges, acceptAllRevisions, rejectAllRevisions,
        insertCitation, insertBibliography, insertEquation, insertContentControl, insertFormField,
        mailMergeReplace, applyTheme, recolorDocument, tweakTheme, designTheme,
        countWords, countCharacters, countParagraphs, getReadability,
        getSelectionText, getSelectionHtml, listSections, listHeadings, getDocumentSummary,
        applyTemplate, currentTheme, TEMPLATES, OOXML,
        setBold, setItalic, setUnderline, setFontColor, setFontSize, setFontName, setAlignment,
        insertHeading, insertParagraph, insertText, insertBulletList, insertNumberedList, insertList,
        find, replaceAll, setTrackChanges, acceptRevisions, rejectRevisions, insertTOC, insertToc, setOrientation
      );
      if (rv !== undefined && rv !== null) ret = String(rv);
    });
    await readDoc();
    return { ok: true, err: null, result: ret };
  } catch (e) {
    return { ok: false, err: e.message, result: null };
  }
}

// ── Input helpers ───────────────────────────────────────────────────────────
function handleKey(e){if(e.key==='Enter'&&!e.shiftKey){e.preventDefault();sendMessage();}}
function autoResize(el){el.style.height='auto';el.style.height=Math.min(el.scrollHeight,120)+'px';}
function useSugg(btn){document.getElementById('msgInput').value=btn.textContent;sendMessage();}
function attachFile(e){const f=e.target.files[0];if(!f)return;const r=new FileReader();r.onload=ev=>{attached={name:f.name,content:ev.target.result};document.getElementById('attachBadge').innerHTML='📎 <span style="color:var(--text2)">'+esc(f.name)+'</span> <button class="attach-x" onclick="clearAttach()">✕</button>';};r.readAsText(f);e.target.value='';}
function clearAttach(){attached=null;document.getElementById('attachBadge').innerHTML='';}

// ── Summarise ───────────────────────────────────────────────────────────────
async function maybeSummarize() {
  if(history.length<10)return;
  try{const k=localStorage.getItem('apiKey');const r=await fetch('/api/summarize',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({messages:history.slice(0,-4),previousSummary:summary||undefined,apiKey:k||undefined})});const d=await r.json();if(d.summary){summary=d.summary;history=history.slice(-4);}}catch{}
}

// ── Stop ────────────────────────────────────────────────────────────────────
function stopGen(){ctrl?.abort();loading=false;removeTyping();const b=document.getElementById('sendBtn');b.textContent='Send';b.className='send';b.onclick=sendMessage;}

// ── Send ────────────────────────────────────────────────────────────────────
async function sendMessage() {
  const inp=document.getElementById('msgInput');
  let text=inp.value.trim();
  if(!text&&!attached)return;
  if(loading)return;
  closePlus();

  if(attached){text+=(text?'\n\n':'')+'[Attached: '+attached.name+']\n'+attached.content.slice(0,8000);clearAttach();}

  const allowed = await checkCostBefore(text);
  if (!allowed) return;

  inp.value='';inp.style.height='auto';loading=true;
  const sb=document.getElementById('sendBtn');sb.textContent='■ Stop';sb.className='send stop';sb.onclick=stopGen;

  const userMsg={role:'user',content:text,time:new Date().toISOString()};
  history.push(userMsg);addMsg('user',text);
  pendingWebSearchResults = [];
  addTyping();
  await maybeSummarize();
  await runChatTurn();

  loading=false;
  const sb2=document.getElementById('sendBtn');sb2.textContent='Send';sb2.className='send';sb2.onclick=sendMessage;
}

async function runChatTurn() {
  const t0=Date.now();ctrl=new AbortController();
  const isGroq   =selectedModel.startsWith('groq:');
  const isOllama =selectedModel.startsWith('ollama:');
  const modelId  =isGroq?selectedModel.slice(5):isOllama?selectedModel.slice(7):selectedModel;
  const apiKey   =localStorage.getItem('apiKey')||undefined;
  const groqKey  =localStorage.getItem('groqKey')||undefined;

  try{
    const res=await fetch('/api/chat',{
      method:'POST',headers:{'Content-Type':'application/json'},signal:ctrl.signal,
      body:JSON.stringify({
        messages:history.slice(-8),docData,selectionInfo,summary,
        model:modelId,useOllama:isOllama,useGroq:isGroq,apiKey,groqKey,
        settings: { allowWebSearch: webSearchMode },
        webSearchResults: pendingWebSearchResults.length ? pendingWebSearchResults : undefined,
        options:{
          deepThink:  opts.deep,
          dynamicDepth: opts.dynamic,
          autoModel:  opts.auto,
          plan:       opts.plan,
          sectionByStep: opts.sectionByStep,
          preferences: userPreferences,
          costLimit:  getCostLimit(),
          availableModels: opts.auto
            ? allModelsFlat().filter(m=>!m.id.startsWith('ollama:')).map(m=>({id:m.id,name:m.name,in:m.in,out:m.out,tags:m.tags}))
            : undefined,
        }
      })
    });
    const data=await res.json();
    const ms=Date.now()-t0;
    removeTyping();

    if(data.error){
      addMsg('assistant','⚠ '+data.error,true,ms);
      return;
    }

    const{response:reply,code,usage,selectedModel:autoSelected,plan,webSearchRequest,webSearchUsed}=data;

    if(plan){history.push({role:'assistant',content:'📋 Plan:\n'+plan});addMsg('assistant','📋 Plan:\n'+plan);}
    const tok=usage?.completion_tokens??null;
    if(tok){totalTok+=tok;updateTok();}
    if(usage)addActualCost(usage, autoSelected||modelId);

    const viaName = autoSelected ? (allModelsFlat().find(m=>m.id===autoSelected)?.name||autoSelected) : null;
    const isFirst=history.filter(m=>m.role==='assistant').length===0;

    let msgEl = null;
    if(reply){
      history.push({role:'assistant',content:reply,ms,tok,time:new Date().toISOString(),via:viaName});
      msgEl = addMsg('assistant',reply,false,ms,tok,null,viaName);
      if (webSearchUsed) addWebSearchPill(msgEl, webSearchUsed);
    }

    // Permission card if AI requested a search and mode=off
    if (webSearchRequest && !code) {
      const parent = msgEl || addMsg('assistant', 'I need to look something up online.', false);
      addWebSearchPermissionCard(parent, webSearchRequest, async () => {
        addTyping();
        await runChatTurn();
      });
      return;
    }

    if(code){
      const r=await execCode(code);
      if(showCode){const last=document.getElementById('messages').lastElementChild;if(last&&!last.querySelector('.code-block'))last.appendChild(mkCode(code));}
      if(r.ok){
        if(r.result){history.push({role:'assistant',content:r.result});addMsg('assistant',r.result,false,null,null);}
        else if(!reply){addMsg('assistant','✓ Done.');}
      }else{
        addMsg('assistant','⚠ Error: '+r.err,true);
        // If it's a ReferenceError on a helper name, give the model the full helper list to ground its retry.
        const refMatch = /(\w+) is not defined/.exec(r.err || '');
        const hint = refMatch
          ? `The code failed: "${r.err}". You called a helper that doesn't exist. The ONLY helpers you may use are: addHeading, addParagraph, addTitle, addSubtitle, addQuote, addList, applyStyle, replaceText, findText, insertPageBreak, insertSectionBreak, insertColumns, setHeader, setFooter, addPageNumbers, insertTableOfContents, setMargins, setPageOrientation, insertTable, setTableCell, styleTable, insertImage, insertWatermark, insertFootnote, insertEndnote, insertComment, toggleTrackChanges, acceptAllRevisions, rejectAllRevisions, insertCitation, insertBibliography, insertEquation, insertContentControl, insertFormField, mailMergeReplace, applyTheme, designTheme, tweakTheme, recolorDocument, applyTemplate, countWords, countCharacters, countParagraphs, getReadability, getSelectionText, listSections, listHeadings, getDocumentSummary. Pick the closest existing helper and emit a corrected CODE_JS block.`
          : `The code failed: "${r.err}". Fix it with a corrected CODE_JS block.`;
        history.push({role:'user',content:hint});
        const rr=await fetch('/api/chat',{method:'POST',headers:{'Content-Type':'application/json'},signal:ctrl.signal,body:JSON.stringify({messages:history.slice(-8),docData,selectionInfo,summary,model:modelId,useOllama:isOllama,useGroq:isGroq,apiKey,groqKey,settings:{allowWebSearch:webSearchMode}})});
        const rd=await rr.json();
        if(rd.code){const rres=await execCode(rd.code);if(rd.response)history.push({role:'assistant',content:rd.response});addMsg('assistant',rres.ok?'✓ Fixed.':'⚠ Retry failed: '+rres.err,!rres.ok);}
      }
    }

    if(getCostLimit()>0&&costLimitType==='session'&&sessionCost>getCostLimit()){
      const over=(sessionCost-getCostLimit()).toFixed(3);
      addMsg('assistant',`⚠ Session cost ($${sessionCost.toFixed(3)}) has exceeded your $${getCostLimit().toFixed(2)} limit by $${over}. You'll need to confirm each subsequent prompt.`,true);
    }

    if(isFirst&&reply){const text=history.find(m=>m.role==='user')?.content||'';const title=await generateTitle(text,reply);if(title)await saveChat(title);else await saveChat();}
    else await saveChat();
  }catch(e){
    removeTyping();
    if(e.name!=='AbortError')addMsg('assistant','⚠ Could not reach the server. Is it running?',true);
  }
}

async function reportMsg(btn,response,code) {
  if(btn.classList.contains('sent'))return;
  const prompt=history.slice().reverse().find(m=>m.role==='user')?.content||'';
  try{await fetch('/api/feedback',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({prompt,response,code:code||null,model:selectedModel,timestamp:new Date().toISOString()})});btn.innerHTML='✓ Reported';btn.classList.add('sent');}catch{btn.innerHTML='⚑ Report';}
}

function esc(s){return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}

// Load active theme from storage at startup
try {
  const saved = localStorage.getItem('activeTheme');
  if (saved) currentTheme = JSON.parse(saved);
} catch {}
