/**
 * Frontend unit tests for ETL_ODE static/index.html
 *
 * Run:  node tests/test_frontend.js
 *
 * Uses jsdom to execute the real app script in a simulated DOM environment
 * with a controllable fetch mock. Each test creates an isolated environment.
 *
 * What IS tested here:
 *   - collectSerieMap: selector sync, deletion, preservation of other-report mappings
 *   - doAnalyze: selector values flushed into config_json payload before request
 *   - doAnalyze: A started → B replaces A → A's finally doesn't touch B's controller/button
 *   - doAnalyze: single request aborted by input change, button restores when request completes
 *   - doGenerate: blocked when analysisCompletedRevision !== inputRevision
 *   - doGenerate: blocked when there is no valid current analysis / stale inputs
 *   - doGenerate: blocked when pending selector changes exist
 *   - doGenerate: G's finally doesn't touch a newer generation's state
 *   - doGenerate: input changes during a single request restore controller/button state
 *   - invalidateAnalysis: clears downloads (object URLs) and disables generate
 *
 * What is NOT tested here (requires real browser or Playwright):
 *   - Visual rendering (badges, colours, download link clicks)
 *   - Actual file picker / drag-and-drop interactions
 *   - Real network latency races beyond what the mock can reproduce
 */
'use strict';

const { JSDOM }  = require('jsdom');
const assert     = require('assert/strict');
const fs         = require('fs');
const path       = require('path');

// ---------------------------------------------------------------------------
// Extract script from HTML
// ---------------------------------------------------------------------------
const HTML_PATH = path.join(__dirname, '..', 'static', 'index.html');
const htmlContent = fs.readFileSync(HTML_PATH, 'utf8');
const scriptMatch = htmlContent.match(/<script>([\s\S]*?)<\/script>/);
if (!scriptMatch) throw new Error('No <script> block found in index.html');
const APP_SCRIPT = scriptMatch[1];

// ---------------------------------------------------------------------------
// Minimal DOM markup with all IDs the script references
// ---------------------------------------------------------------------------
const MINIMAL_HTML = `<!DOCTYPE html><html><body>
  <button id="btn-analyze" disabled>Analizar</button>
  <button id="btn-generate" disabled>Generar</button>
  <button id="btn-remap" disabled>Guardar mapeo</button>
  <span id="remap-hint"></span>
  <button id="btn-add-bank"></button>
  <button id="btn-confirm-add-bank"></button>
  <button id="btn-cancel-add-bank"></button>
  <button id="btn-export-config"></button>
  <button id="btn-logout"></button>
  <input type="file"   id="file-reporte">
  <input type="file"   id="file-modelo">
  <input type="file"   id="inp-import-config" accept=".json">
  <div   id="file-name"></div>
  <input type="number" id="inp-year"  value="2026">
  <select id="inp-month"><option value="8" selected>Agosto</option></select>
  <input type="date" id="inp-fecha-desde" value="2026-08-01">
  <input type="date" id="inp-fecha-hasta" value="2026-08-31">
  <input type="number" id="inp-correlativo-inicial" value="1">
  <button id="btn-open-guide">Ayuda</button>
  <button id="btn-close-guide">Cerrar</button>
  <div id="guide-modal"></div>
  <div id="help-popover" hidden></div>
  <input type="text" id="cfg-cuenta-clientes"  value="12120001">
  <input type="text" id="cfg-sub-diario"       value="07">
  <input type="text" id="cfg-tipo-anexo"       value="02">
  <input type="text" id="cfg-moneda"           value="MN">
  <input type="text" id="cfg-tipo-conversion"  value="VTA">
  <input type="text" id="cfg-doc-anulado"      value="0">
  <input type="text" id="cfg-glosa"            value="COBRANZA DEL MES">
  <input type="text" id="new-bank-nombre">
  <input type="text" id="new-bank-cuenta">
  <input type="text" id="new-bank-medio">
  <div id="banks-grid"></div>
  <div id="add-bank-form"    style="display:none"></div>
  <div id="cfg-toggle"></div>
  <div id="cfg-body"         class="hidden"></div>
  <div id="card-config"></div>
  <div id="section-results"  style="display:none"></div>
  <div id="section-downloads"style="display:none"></div>
  <div id="downloads-body"></div>
  <div id="stale-notice"     style="display:none"></div>
  <div id="result-badge"></div>
  <div id="errors-block"     style="display:none"></div>
  <ul  id="errors-list"></ul>
  <div id="warnings-block"   style="display:none"></div>
  <ul  id="warnings-list"></ul>
  <div id="info-block"       style="display:none"></div>
  <ul  id="info-list"></ul>
  <div id="type-summary-block" style="display:none"></div>
  <p   id="type-summary-text"></p>
  <div id="stats-block"      style="display:none"></div>
  <div id="preview-block"    style="display:none"></div>
  <tbody id="preview-body"></tbody>
  <table><tbody id="series-table-body"></tbody></table>
  <div id="series-section"   style="display:none"></div>
  <div id="unconfigured-banks-block" style="display:none"></div>
  <ul  id="unconfigured-banks-list"></ul>
  <span id="analyze-status"></span>
  <span id="generate-status"></span>
</body></html>`;

// ---------------------------------------------------------------------------
// Environment factory
// ---------------------------------------------------------------------------
/**
 * Creates an isolated jsdom environment with the app script loaded.
 *
 * @param {object} opts
 * @param {boolean} opts.noFile   – if true, reporteFile is null (btn-analyze stays disabled)
 * @returns {{ win, fetchCtrl, G }}
 *   win       – the jsdom Window
 *   fetchCtrl – controllable fetch helper (see below)
 *   G         – alias for win (global access)
 */
function createEnv({ noFile = false } = {}) {
  const dom = new JSDOM(MINIMAL_HTML, {
    runScripts: 'dangerously',
    url: 'http://localhost',
  });
  const win = dom.window;

  // Polyfill btoa / atob for older jsdom
  if (!win.btoa) win.btoa = s => Buffer.from(s, 'binary').toString('base64');
  if (!win.atob) win.atob = s => Buffer.from(s, 'base64').toString('binary');

  // Polyfill URL.createObjectURL / revokeObjectURL
  win.URL.createObjectURL = () => 'blob:fake-url';
  win.URL.revokeObjectURL = () => {};

  // Controllable fetch: each call enqueues a deferred promise.
  const queue = [];
  win.fetch = (url, opts) => {
    if (url === '/api/config/default') {
      return Promise.resolve({ ok: true, json: async () => ({}) });
    }
    let resolveFn, rejectFn;
    const p = new Promise((res, rej) => { resolveFn = res; rejectFn = rej; });
    const entry = { url, opts, resolve: resolveFn, reject: rejectFn, settled: false };

    // Wire up AbortController signal
    const signal = opts && opts.signal;
    if (signal) {
      if (signal.aborted) {
        // Already aborted before fetch was called
        const e = new Error('AbortError'); e.name = 'AbortError';
        entry.settled = true;
        rejectFn(e);
      } else {
        signal.addEventListener('abort', () => {
          if (!entry.settled) {
            entry.settled = true;
            const e = new Error('AbortError'); e.name = 'AbortError';
            rejectFn(e);
          }
        });
      }
    }
    queue.push(entry);
    return p;
  };

  // Intercept alert/confirm so tests don't block
  win.alert   = () => {};
  win.confirm = () => false;

  // Suppress scrollIntoView (not implemented in jsdom)
  // Applied after script runs — see patchDOM below.

  // Run app script
  const s = win.document.createElement('script');
  s.textContent = APP_SCRIPT;
  win.document.body.appendChild(s);

  const lexicalStateVars = new Set([
    'reporteFile',
    'modeloFile',
    'analysisValid',
    'analysisData',
    'serieMap',
    'defaultConfig',
    'inputRevision',
    'analysisCompletedRevision',
    'activeAnalysisController',
    'activeGenerationController',
    'activeDownloadUrls',
    'userBanks',
  ]);

  const G = new Proxy(win, {
    get(target, prop, receiver) {
      if (typeof prop === 'string' && lexicalStateVars.has(prop)) {
        return win.eval(prop);
      }
      const val = Reflect.get(target, prop, receiver);
      if (typeof val === 'function') return val.bind(target);
      return val;
    },
    set(target, prop, value, receiver) {
      if (typeof prop === 'string' && lexicalStateVars.has(prop)) {
        win.__tmp_state_value__ = value;
        win.eval(`${prop} = __tmp_state_value__`);
        delete win.__tmp_state_value__;
        return true;
      }
      return Reflect.set(target, prop, value, receiver);
    },
  });

  // Patch DOM methods not available in jsdom
  win.document.querySelectorAll = (sel => {
    const orig = win.document.querySelectorAll.bind(win.document);
    return orig;
  })(win.document.querySelectorAll);

  // fetchCtrl helpers
  const fetchCtrl = {
    /** How many fetch calls have been made so far. */
    get count() { return queue.length; },

    /**
     * Resolve the n-th fetch (0-based) with an ok JSON response.
     * @param {number} n
     * @param {object} data
     */
    resolveOk(n, data) {
      const e = queue[n];
      if (!e) throw new Error(`No fetch[${n}] yet (total: ${queue.length})`);
      if (e.settled) return;
      e.settled = true;
      e.resolve({ ok: true, json: async () => data });
    },

    /**
     * Resolve the n-th fetch with an error response (ok=false).
     */
    resolveErr(n, detail) {
      const e = queue[n];
      if (!e) throw new Error(`No fetch[${n}] yet`);
      if (e.settled) return;
      e.settled = true;
      e.resolve({ ok: false, json: async () => ({ detail }) });
    },

    /** Parsed body of the n-th fetch as a plain object. */
    bodyJson(n) {
      const e = queue[n];
      if (!e) throw new Error(`No fetch[${n}] yet`);
      // FormData: get config_json field
      const fd = e.opts && e.opts.body;
      if (fd && typeof fd.get === 'function') {
        const raw = fd.get('config_json');
        return raw ? JSON.parse(raw) : null;
      }
      return null;
    },

    queue,
  };

  // Simulate a loaded file (so btn-analyze can be enabled)
  if (!noFile) {
    // jsdom doesn't support real File objects via picker; set programmatically
    G.reporteFile = { name: 'test.xlsx' };
    win.document.getElementById('btn-analyze').disabled = false;
  }

  return { win, fetchCtrl, G };
}

// ---------------------------------------------------------------------------
// Helpers for test scenarios
// ---------------------------------------------------------------------------

/** Standard successful analyze response */
function okAnalyze(seriesAll = []) {
  return {
    errors: [],
    warnings: [],
    info: ['Archivo legible'],
    preview: [],
    type_summary: { facturas: 1, boletas: 0, sin_resolver: 0,
                    importe_facturas: 100, importe_boletas: 0, importe_sin_resolver: 0 },
    stats: { cantidad_original: 1, cantidad_excluida: 0, cantidad_procesada: 1, total_procesado: 100 },
    series_all: seriesAll,
    unresolved_series: [],
    unconfigured_banks: [],
  };
}

/** Inject a .serie-type-select element into the series table */
function addSerieSelector(win, serie, currentValue) {
  const tbody = win.document.getElementById('series-table-body');
  const sel = win.document.createElement('select');
  sel.className = 'serie-type-select';
  sel.dataset.serie = serie;
  sel.setAttribute('onchange', 'onSerieSelectChange()');
  [['', '-- Sin asignar --'], ['01', '01 -- Factura'], ['03', '03 -- Boleta']].forEach(([v, t]) => {
    const opt = win.document.createElement('option');
    opt.value = v; opt.textContent = t;
    if (v === currentValue) opt.selected = true;
    sel.appendChild(opt);
  });
  const td = win.document.createElement('td');
  td.appendChild(sel);
  const tr = win.document.createElement('tr');
  tr.appendChild(td);
  tbody.appendChild(tr);
  return sel;
}

/** Wait one microtask tick (lets pending promises resolve) */
const tick = () => new Promise(r => setImmediate(r));
/** Wait multiple ticks */
const ticks = n => Array.from({length: n}).reduce(p => p.then(tick), Promise.resolve());

// ---------------------------------------------------------------------------
// Test runner
// ---------------------------------------------------------------------------
let passed = 0, failed = 0;
const failures = [];

async function test(name, fn) {
  try {
    await fn();
    passed++;
    console.log(`  ✓  ${name}`);
  } catch (err) {
    failed++;
    failures.push({ name, err });
    console.error(`  ✗  ${name}`);
    console.error(`     ${err.message}`);
  }
}

// Wrap test body in async IIFE so top-level await works in CommonJS
(async () => {

// ===========================================================================
// 1. collectSerieMap — pure selector sync
// ===========================================================================
console.log('\n--- collectSerieMap ---');

await test('syncs 01→03 selector change into serieMap', async () => {
  const { win, G } = createEnv();
  // Pre-existing map: FE01='01'
  G.serieMap = { FE01: '01' };
  // Simulate analysisData so collectSerieMap knows FE01 is in current report
  G.analysisData = { series_all: [{ serie: 'FE01' }] };
  // Render a selector for FE01 and set it to '03'
  const sel = addSerieSelector(win, 'FE01', '03');
  sel.value = '03';
  G.collectSerieMap();
  assert.equal(G.serieMap.FE01, '03', 'FE01 should be 03 after selector change');
});

await test('empty selector removes mapping for current-report serie', async () => {
  const { win, G } = createEnv();
  G.serieMap = { FE01: '01' };
  G.analysisData = { series_all: [{ serie: 'FE01' }] };
  const sel = addSerieSelector(win, 'FE01', '');
  sel.value = '';
  G.collectSerieMap();
  assert.equal(G.serieMap.FE01, undefined, 'FE01 mapping should be removed');
});

await test('series not in current report preserve their mapping', async () => {
  const { win, G } = createEnv();
  // BE01 is in serieMap but NOT in current report's series_all
  G.serieMap = { FE01: '01', BE01: '03' };
  G.analysisData = { series_all: [{ serie: 'FE01' }] };
  // Only FE01 selector exists in DOM
  const sel = addSerieSelector(win, 'FE01', '01');
  sel.value = '01';
  G.collectSerieMap();
  assert.equal(G.serieMap.BE01, '03', 'BE01 mapping must survive (not in current report)');
  assert.equal(G.serieMap.FE01, '01', 'FE01 should remain 01');
});

await test('column-detected type not auto-promoted into serieMap', async () => {
  const { win, G } = createEnv();
  // serieMap has no entry for FE01; selector shows '-- Sin asignar --'
  G.serieMap = {};
  G.analysisData = { series_all: [{ serie: 'FE01', fuente: 'columna', tipos: ['01'] }] };
  const sel = addSerieSelector(win, 'FE01', '');
  sel.value = '';  // Sin asignar — even though columna resolves it
  G.collectSerieMap();
  assert.equal(G.serieMap.FE01, undefined, 'Column-resolved series must not be promoted into serieMap');
});

await test('no analysisData: keeps existing map entries, ignores selectors for unknown series', async () => {
  const { win, G } = createEnv();
  G.serieMap = { FE01: '01' };
  G.analysisData = null;
  // Selector with arbitrary serie
  const sel = addSerieSelector(win, 'XX99', '03');
  sel.value = '03';
  G.collectSerieMap();
  // With no analysisData, currentSeries is empty; XX99 is NOT in current report
  // → selector is NOT synced (no removal from empty set)
  // Actually with empty currentSeries, all selectors are synced as "current-report series"?
  // Check collectSerieMap logic: currentSeries = new Set([]) so XX99 is not in it
  // → XX99 is not in currentSeries → it goes to "preserve" path
  // Wait: the code does: if (!currentSeries.has(k)) newMap[k] = v; for serieMap entries
  // And for selectors: if (sel.value) newMap[serie] = sel.value; regardless of currentSeries
  // Actually let me re-read collectSerieMap...
  // The selector sync section runs for ALL selectors, not just current-series.
  // So XX99 selector with value '03' WOULD be added.
  // This seems like it could be an issue — let me check what the spec says.
  // Spec: "preserve mappings for series not in the current report"
  // The selectors in the DOM should only exist for current-report series (they're rendered by renderSeriesTable)
  // So if there's a selector for XX99 and it's not in currentSeries, that's a stale selector.
  // The current code does sync all selectors regardless. We accept this for now.
  assert.equal(G.serieMap.FE01, '01', 'FE01 preserved (not in null analysisData)');
  assert.equal(G.serieMap.XX99, undefined, 'Stale selector for unknown series must be ignored');
});

// ===========================================================================
// 2. doAnalyze flushes selector state into the request payload
// ===========================================================================
console.log('\n--- doAnalyze payload ---');

await test('doAnalyze sends updated serieMap (01→03 change) in config_json', async () => {
  const { win, fetchCtrl, G } = createEnv();

  // Pre-existing map: FE01='01'
  G.serieMap = { FE01: '01' };
  G.analysisData = { series_all: [{ serie: 'FE01' }] };
  // User changed selector to '03'
  const sel = addSerieSelector(win, 'FE01', '03');
  sel.value = '03';

  // Trigger doAnalyze (no await yet — let it start)
  const analyzePromise = G.doAnalyze();
  await tick();  // let fetch be called

  assert.equal(fetchCtrl.count, 1, 'fetch should have been called once');
  const cfg = fetchCtrl.bodyJson(0);
  assert.ok(cfg, 'config_json should be present in request');
  assert.equal(cfg.serie_map && cfg.serie_map.FE01, '03',
    `Expected FE01=03 in payload, got: ${JSON.stringify(cfg.serie_map)}`);

  // Resolve to clean up
  fetchCtrl.resolveOk(0, okAnalyze());
  await analyzePromise;
});

await test('doAnalyze sends removal of FE01 mapping when set to Sin asignar', async () => {
  const { win, fetchCtrl, G } = createEnv();

  G.serieMap = { FE01: '01' };
  G.analysisData = { series_all: [{ serie: 'FE01' }] };
  const sel = addSerieSelector(win, 'FE01', '');
  sel.value = '';  // Sin asignar — explicit removal

  const analyzePromise = G.doAnalyze();
  await tick();

  const cfg = fetchCtrl.bodyJson(0);
  assert.ok(cfg, 'config_json should be present');
  assert.equal(cfg.serie_map && cfg.serie_map.FE01, undefined,
    `FE01 must be absent from serie_map, got: ${JSON.stringify(cfg.serie_map)}`);

  fetchCtrl.resolveOk(0, okAnalyze());
  await analyzePromise;
});

await test('doAnalyze ignores stale selectors after file change', async () => {
  const { win, fetchCtrl, G } = createEnv();

  G.serieMap = {};
  G.analysisData = { series_all: [{ serie: 'FE01' }] };
  const sel = addSerieSelector(win, 'FE01', '03');
  sel.value = '03';

  G.setupListeners();
  const input = win.document.getElementById('file-reporte');
  Object.defineProperty(input, 'files', {
    value: [{ name: 'reporte_nuevo.xlsx' }],
    configurable: true,
  });
  input.dispatchEvent(new win.Event('change'));

  const analyzePromise = G.doAnalyze();
  await tick();
  const cfg = fetchCtrl.bodyJson(0);
  assert.equal(cfg.serie_map && cfg.serie_map.FE01, undefined,
    `FE01 from stale selectors must not be sent, got: ${JSON.stringify(cfg.serie_map)}`);

  fetchCtrl.resolveOk(0, okAnalyze());
  await analyzePromise;
});

await test('file change clears stale series selectors from previous report', async () => {
  const { win, G } = createEnv();
  G.analysisData = { series_all: [{ serie: 'FE01' }] };
  addSerieSelector(win, 'FE01', '03');
  win.document.getElementById('series-section').style.display = 'block';
  win.document.getElementById('btn-remap').disabled = false;
  win.document.getElementById('remap-hint').textContent = 'Cambios pendientes';

  G.setupListeners();
  const input = win.document.getElementById('file-reporte');
  Object.defineProperty(input, 'files', {
    value: [{ name: 'reporte_nuevo.xlsx' }],
    configurable: true,
  });
  input.dispatchEvent(new win.Event('change'));

  assert.equal(win.document.querySelectorAll('.serie-type-select').length, 0,
    'series editor should be cleared on file change');
  assert.equal(win.document.getElementById('series-section').style.display, 'none',
    'series section should be hidden on file change');
  assert.equal(win.document.getElementById('btn-remap').disabled, true,
    'remap button should be disabled after clearing series editor');
  assert.equal(win.document.getElementById('remap-hint').textContent, '',
    'remap hint should be cleared after file change');
});

// ===========================================================================
// 3. Race condition: A started, B replaces A, A's finally must not touch B
// ===========================================================================
console.log('\n--- Race condition: A → B ---');

await test("A's finally does not reset B's activeAnalysisController to null", async () => {
  const { win, fetchCtrl, G } = createEnv();
  G.serieMap = {};
  G.analysisData = null;

  // Start A
  const promiseA = G.doAnalyze();
  await tick();
  assert.equal(fetchCtrl.count, 1, 'fetch A called');

  // While A is in flight, start B (this aborts A)
  const promiseB = G.doAnalyze();
  await tick();
  assert.equal(fetchCtrl.count, 2, 'fetch B called');

  // Now A's request has been aborted (AbortError). A's catch returns early.
  // A's finally should NOT clear activeAnalysisController since B is now active.
  await promiseA;  // A settles (AbortError → finally runs)
  await tick();

  assert.notEqual(G.activeAnalysisController, null,
    "A's finally must not set activeAnalysisController=null while B is still in flight");

  // Resolve B normally
  fetchCtrl.resolveOk(1, okAnalyze());
  await promiseB;
  await ticks(3);

  assert.equal(G.activeAnalysisController, null,
    "After B completes, activeAnalysisController should be null");
});

await test("A's finally does not restore btn-analyze label/state while B runs", async () => {
  const { win, fetchCtrl, G } = createEnv();
  G.serieMap = {};
  G.analysisData = null;

  const btn = win.document.getElementById('btn-analyze');

  const promiseA = G.doAnalyze();
  await tick();

  // B replaces A
  const promiseB = G.doAnalyze();
  await tick();

  // A settles with AbortError
  await promiseA;
  await tick();

  // Button should still show spinner/disabled — B is still running
  assert.ok(btn.disabled, 'btn-analyze should still be disabled while B runs');
  assert.ok(btn.innerHTML.includes('spinner') || btn.textContent.includes('naliz'),
    `btn should show in-progress text while B runs, got: ${btn.innerHTML}`);

  // Resolve B
  fetchCtrl.resolveOk(1, okAnalyze());
  await promiseB;
  await ticks(3);

  // Now B's finally ran — button should be restored
  assert.equal(btn.disabled, false, 'After B completes, btn-analyze should be enabled');
  assert.equal(btn.textContent, 'Analizar archivos');
});

// ===========================================================================
// 4. Single request: input changes during flight, button still restores
// ===========================================================================
console.log('\n--- Input change during single request ---');

await test('changing input during analysis allows retry after request settles', async () => {
  const { win, fetchCtrl, G } = createEnv();
  G.serieMap = {};
  G.analysisData = null;

  const btn    = win.document.getElementById('btn-analyze');
  const status = win.document.getElementById('analyze-status');

  const promiseA = G.doAnalyze();
  await tick();

  // User changes an input while request is in flight (no new analysis started)
  G.invalidateAnalysis();
  // revision bumped; myRevision for A is now stale

  // Resolve A's fetch — but revision has changed
  fetchCtrl.resolveOk(0, okAnalyze());
  await promiseA;
  await ticks(3);

  // A's finally: activeAnalysisController === myControllerA still (no B started)
  // So it SHOULD restore the button
  assert.equal(btn.disabled, false, 'After single aborted analysis, button should be re-enabled for retry');
  assert.equal(btn.textContent, 'Analizar archivos');
  // But analysisValid must remain false (result was discarded)
  assert.equal(G.analysisValid, false, 'analysisValid must remain false — result was stale');
  assert.equal(G.analysisCompletedRevision, -1, 'analysisCompletedRevision must not be set for stale result');
});

// ===========================================================================
// 5. doGenerate guards
// ===========================================================================
console.log('\n--- doGenerate guards ---');

await test('doGenerate blocked when analysisCompletedRevision !== inputRevision', async () => {
  const { win, fetchCtrl, G } = createEnv();
  const status = win.document.getElementById('generate-status');

  // Simulate valid analysis for revision 5
  G.analysisValid = true;
  G.analysisCompletedRevision = 5;
  G.inputRevision = 7;  // revision changed after analysis
  G.serieMap = {};

  await G.doGenerate();

  assert.equal(fetchCtrl.count, 0, 'No fetch should have been called');
  assert.ok(
    status.textContent.toLowerCase().includes('analiz') ||
    status.textContent.toLowerCase().includes('cambiaron'),
    `Expected stale-analysis message, got: "${status.textContent}"`
  );
});

await test('doGenerate blocked when analysis is no longer valid', async () => {
  const { win, fetchCtrl, G } = createEnv();
  const status = win.document.getElementById('generate-status');

  G.analysisValid = false;
  G.analysisCompletedRevision = G.inputRevision;

  await G.doGenerate();

  assert.equal(fetchCtrl.count, 0, 'No fetch should have been called');
  assert.ok(
    status.textContent.toLowerCase().includes('cambiaron'),
    `Expected stale-input message, got: "${status.textContent}"`
  );
});

await test('doGenerate blocked when pending selector changes exist', async () => {
  const { win, fetchCtrl, G } = createEnv();
  const status = win.document.getElementById('generate-status');

  G.analysisValid = true;
  G.analysisCompletedRevision = G.inputRevision;  // same revision — analysis is current
  G.serieMap = { FE01: '01' };
  G.analysisData = { series_all: [{ serie: 'FE01' }] };

  // Selector shows 03 but serieMap says 01 — pending change
  const sel = addSerieSelector(win, 'FE01', '03');
  sel.value = '03';

  await G.doGenerate();

  assert.equal(fetchCtrl.count, 0, 'No fetch should be called with pending selector changes');
  assert.ok(
    status.textContent.toLowerCase().includes('cambiaron'),
    `Expected stale-input message, got: "${status.textContent}"`
  );
});

await test('doGenerate proceeds when revision matches and no pending changes', async () => {
  const { win, fetchCtrl, G } = createEnv();

  G.analysisValid = true;
  G.analysisCompletedRevision = G.inputRevision;
  G.serieMap = { FE01: '01' };
  G.analysisData = { series_all: [{ serie: 'FE01' }] };
  // Selector matches serieMap
  const sel = addSerieSelector(win, 'FE01', '01');
  sel.value = '01';

  const p = G.doGenerate();
  await tick();

  assert.equal(fetchCtrl.count, 1, 'fetch should have been called once');
  fetchCtrl.resolveOk(0, { files: [], summary: { resultado: 'CUADRADO', diferencia: 0, cantidad_procesada: 1, total_procesado: 100 } });
  await p;
});

// ===========================================================================
// 6. Generate race condition: G_A → G_B, G_A's finally doesn't touch G_B
// ===========================================================================
console.log('\n--- Generate race condition ---');

await test("G_A's finally does not clear G_B's controller", async () => {
  const { win, fetchCtrl, G } = createEnv();

  G.analysisValid = true;
  G.analysisCompletedRevision = G.inputRevision;
  G.serieMap = {};

  // Start G_A
  const promiseGA = G.doGenerate();
  await tick();
  assert.equal(fetchCtrl.count, 1, 'G_A fetch called');

  // Simulate another generate replacing G_A
  // To bypass guards, bump revision and reset analysisCompletedRevision
  G.analysisValid = true;
  G.analysisCompletedRevision = G.inputRevision;
  const promiseGB = G.doGenerate();
  await tick();
  assert.equal(fetchCtrl.count, 2, 'G_B fetch called');

  // Resolve G_A (it's stale — revision changed when G_B started)
  fetchCtrl.resolveOk(0, { files: [], summary: {} });
  await promiseGA;
  await tick();

  assert.notEqual(G.activeGenerationController, null,
    "G_A's finally must not null G_B's controller");

  fetchCtrl.resolveOk(1, { files: [], summary: {} });
  await promiseGB;
  await ticks(2);
  assert.equal(G.activeGenerationController, null, 'After G_B, controller should be null');
});

await test('generation result is discarded when inputs change mid-request', async () => {
  const { win, fetchCtrl, G } = createEnv();
  const section = win.document.getElementById('section-downloads');
  const body = win.document.getElementById('downloads-body');

  G.analysisValid = true;
  G.analysisCompletedRevision = G.inputRevision;
  G.serieMap = {};

  const p = G.doGenerate();
  await tick();
  assert.equal(fetchCtrl.count, 1, 'generate request should be sent');

  G.invalidateAnalysis();
  fetchCtrl.resolveOk(0, {
    files: [{ label: 'Archivo principal', name: 'carga.xls', mime: 'application/vnd.ms-excel', data: '' }],
    summary: { resultado: 'CUADRADO', diferencia: 0, cantidad_procesada: 1, total_procesado: 100 },
  });
  await p;
  await ticks(2);

  assert.notEqual(section.style.display, 'block', 'stale generation must not show downloads');
  assert.equal(body.textContent.trim(), '', 'stale generation must not render download cards');
});

await test('single generation with input change restores controller/button state', async () => {
  const { win, fetchCtrl, G } = createEnv();
  const btn = win.document.getElementById('btn-generate');

  G.analysisValid = true;
  G.analysisCompletedRevision = G.inputRevision;
  G.serieMap = {};

  const p = G.doGenerate();
  await tick();
  assert.equal(fetchCtrl.count, 1, 'generate request should be sent');

  G.invalidateAnalysis();
  fetchCtrl.resolveOk(0, { files: [], summary: {} });
  await p;
  await ticks(2);

  assert.equal(G.activeGenerationController, null,
    'controller should be cleared after single stale generation settles');
  assert.equal(btn.textContent, 'Generar carga',
    'button label should be restored after stale generation settles');
  assert.equal(btn.disabled, true,
    'button should remain disabled because analysis is stale after input change');
});

// ===========================================================================
// 7. invalidateAnalysis clears downloads
// ===========================================================================
console.log('\n--- invalidateAnalysis clears downloads ---');

await test('invalidateAnalysis hides section-downloads', async () => {
  const { win, G } = createEnv();
  const sec = win.document.getElementById('section-downloads');
  sec.style.display = 'block';  // pretend downloads are shown
  G.analysisValid = true;
  G.invalidateAnalysis();
  assert.equal(sec.style.display, 'none', 'section-downloads should be hidden');
});

await test('invalidateAnalysis disables btn-generate', async () => {
  const { win, G } = createEnv();
  const btn = win.document.getElementById('btn-generate');
  btn.disabled = false;
  G.analysisValid = true;
  G.invalidateAnalysis();
  assert.equal(btn.disabled, true, 'btn-generate should be disabled');
});

await test('invalidateAnalysis increments inputRevision even when already invalid', async () => {
  const { G } = createEnv();
  G.analysisValid = false;
  const rev0 = G.inputRevision;
  G.invalidateAnalysis();
  assert.ok(G.inputRevision > rev0, 'inputRevision must increment even when analysisValid=false');
});

// ===========================================================================
// 8. Verify confirmed Sharon values reach the API payload
// ===========================================================================
console.log('\n--- Sharon config in payload ---');

await test('readConfig includes schema_version:2', async () => {
  const { G } = createEnv();
  const cfg = G.readConfig();
  assert.equal(cfg.schema_version, 2, 'readConfig must emit schema_version:2');
});

await test('readConfig default sub_diario is 07', async () => {
  const { G } = createEnv();
  const cfg = G.readConfig();
  assert.equal(cfg.sub_diario, '07', `Expected sub_diario=07, got ${cfg.sub_diario}`);
});

await test('readConfig includes BCP in default bancos', async () => {
  const { G } = createEnv();
  const cfg = G.readConfig();
  assert.ok(cfg.bancos.BCP, 'BCP should be in default bancos');
  assert.equal(cfg.bancos.BCP.cuenta_contable, '10400011');
});

await test('normaliseBankName maps DE CREDITO DEL PERU to BCP', async () => {
  const { G } = createEnv();
  assert.equal(G.normaliseBankName('DE CREDITO DEL PERU'), 'BCP');
});

// ===========================================================================
// Summary
// ===========================================================================
console.log(`\n${'='.repeat(60)}`);
console.log(`Results: ${passed} passed, ${failed} failed`);
if (failures.length) {
  console.log('\nFailed tests:');
  failures.forEach(f => console.error(`  ✗ ${f.name}\n    ${f.err.stack || f.err.message}`));
  process.exitCode = 1;
} else {
  console.log('All frontend tests passed.');
}

})().catch(e => { console.error(e.stack || e); process.exitCode = 1; });
