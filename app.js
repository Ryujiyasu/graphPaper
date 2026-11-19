// SPDX-License-Identifier: MIT OR Apache-2.0
// 方眼紙専用ソフト: draw boxes on a square grid, fill them in, and save an
// ordinary .xlsx that Excel opens as the same 方眼紙. Opening, saving and PDF
// go through Oxi (Rust + WebAssembly) inside the browser; nothing is sent
// anywhere.
import {
  KINDS, makeBox, isInput, isField, free, ref, nameInputs, toWorkbook, fromWorkbook, toMarkdown,
  titleOf, checkOptions, edges, pages, pageBands, movePage, addPage, removePage, resizeGrid,
  newSheet, newGrid, sheetName, MAX_COLS, MAX_ROWS,
} from './model.js';
import { sampleForm } from './sample.js';
import { BLOCKS, GROUPS, MINE, bands, bottomRow, insertBlock, moveBlock, removeBlock, captureBlock, registerMine, unregisterMine, readShared } from './blocks.js';

const VERSION = '0.10.0';
// The desktop app carries the engine beside its pages; the site takes it from npm.
const DESKTOP = !!window.__TAURI__;
const LOCAL = DESKTOP || new URLSearchParams(location.search).get('engine') === 'local';
const ENGINE = LOCAL ? './engine/' : `https://unpkg.com/@oxidocs/wasm@${VERSION}/`;
const FONTS = LOCAL ? './engine/fonts-plugin.js' : `https://cdn.jsdelivr.net/npm/@oxidocs/wasm-fonts@${VERSION}/index.js`;

const $ = (id) => document.getElementById(id);
const paper = $('paper');
const STORE = 'graphPaper:v1';

let sheets = [];
let active = 0;
let selected = null;
let tool = 'label';
let mode = 'design';
let protect = true;
let history = [], future = [];
let editing = null;
let xs = [0], ys = [0];

const sheet = () => sheets[active];
const grid = () => sheet().grid;
const byId = (id) => sheet().boxes.find((b) => b.id === id);
const setSheet = (next) => { sheets[active] = next; };

// ---- engine -------------------------------------------------------------
let enginePromise = null, withFonts = (api, f) => f();
function engine() {
  enginePromise ||= (async () => {
    status('変換エンジン（Oxi）を読み込んでいます…');
    const t0 = performance.now();
    const oxi = await import(ENGINE + 'oxidocs_wasm.js');
    await oxi.default();
    try {
      const [helper, plugin] = await Promise.all([import(ENGINE + 'fonts.js'), import(FONTS)]);
      helper.useFonts(oxi, plugin.default);
      withFonts = helper.withFonts;
    } catch (e) { console.warn('font plugin not loaded', e); }
    status(`変換エンジンの準備ができました（${Math.round(performance.now() - t0)} ms）`);
    return oxi;
  })();
  enginePromise.catch(() => { enginePromise = null; });
  return enginePromise;
}
let template = null;
async function templateBytes() {
  template ||= new Uint8Array(await (await fetch('./template.xlsx')).arrayBuffer());
  return template;
}

// ---- state --------------------------------------------------------------
const snapshot = () => JSON.stringify({ sheets, active });
function remember() { history.push(snapshot()); if (history.length > 200) history.shift(); future = []; }
function restore(json) { ({ sheets, active } = JSON.parse(json)); if (!byId(selected)) selected = null; render(); }
function undo() { if (!history.length) return; future.push(snapshot()); restore(history.pop()); }
function redo() { if (!future.length) return; history.push(snapshot()); restore(future.pop()); }
function persist() { try { localStorage.setItem(STORE, JSON.stringify({ sheets, active, protect })); } catch { /* storage off */ } }

// ---- drawing ------------------------------------------------------------
function render() {
  const g = grid();
  xs = edges(g.colPx); ys = edges(g.rowPx);
  nameInputs(sheet().boxes);
  paper.textContent = '';
  Object.assign(paper.style, { width: `${xs[g.cols]}px`, height: `${ys[g.rows]}px` });
  paper.appendChild(gridLines());
  const sel = selected && byId(selected);
  const band = sel?.block && mode === 'design' && bands(sheet().boxes, sheet().blocks).find((k) => k.id === sel.block);
  if (band) {
    const el = document.createElement('div'); el.className = 'band';
    Object.assign(el.style, { top: `${ys[band.r]}px`, height: `${ys[band.end] - ys[band.r]}px` });
    paper.appendChild(el);
    // The block's own bar: its name and what can be done to it, on the page.
    const bar = document.createElement('div'); bar.className = 'bandbar';
    bar.style.top = `${ys[band.r]}px`;
    const name = document.createElement('span'); name.textContent = `ブロック「${band.title || BLOCKS[band.name]?.title || ''}」`;
    bar.appendChild(name);
    for (const [act, text, title] of [['up', '↑', '上へ'], ['down', '↓', '下へ'], ['save', '★ マイブロックに', 'いまの形をマイブロックとして保存'], ['remove', '✕ 削除', 'このブロックを消す（Ctrl+Z で戻せます）']]) {
      const b = document.createElement('button'); b.textContent = text; b.title = title;
      if (act === 'remove') b.className = 'danger';
      b.onclick = (e) => {
        e.stopPropagation();
        if (act === 'save') makeMine({ r: band.r, c: 0, h: band.end - band.r, w: grid().cols }, `${band.title || BLOCKS[band.name]?.title || 'ブロック'}（マイ）`);
        else blockAction(band.id, act);
      };
      bar.appendChild(b);
    }
    paper.appendChild(bar);
  }
  pageMarks();
  for (const b of sheet().boxes) paper.appendChild(boxElement(b));
  tabs();
  inspector();
  footer();
  document.body.classList.toggle('has-selection', !!selected && mode === 'design');
  persist();
}

function gridLines() {
  const g = grid();
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.setAttribute('class', 'lines');
  svg.setAttribute('width', xs[g.cols]); svg.setAttribute('height', ys[g.rows]);
  const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
  let d = '';
  for (const x of xs) d += `M${x + 0.5} 0V${ys[g.rows]}`;
  for (const y of ys) d += `M0 ${y + 0.5}H${xs[g.cols]}`;
  path.setAttribute('d', d);
  svg.appendChild(path);
  return svg;
}

function pageMarks() {
  const list = pageBands(grid());
  list.forEach((p, k) => {
    if (k > 0) {
      const line = document.createElement('div'); line.className = 'pagebreak';
      line.style.top = `${ys[p.start]}px`;
      paper.appendChild(line);
    }
    if (mode !== 'design') return;
    const tag = document.createElement('div'); tag.className = 'pagetag';
    tag.style.top = `${ys[p.start]}px`;
    const label = document.createElement('span'); label.textContent = `${k + 1}ページ`;
    tag.appendChild(label);
    if (list.length > 1) {
      for (const [act, text, title] of [['up', '↑', 'このページを前へ'], ['down', '↓', 'このページを後ろへ'], ['remove', '×', 'このページを消す']]) {
        const b = document.createElement('button'); b.textContent = text; b.title = title;
        b.onclick = (e) => { e.stopPropagation(); pageAction(k, act); };
        tag.appendChild(b);
      }
    }
    paper.appendChild(tag);
  });
}

function boxElement(b) {
  const el = document.createElement('div');
  el.className = `box ${b.kind} b-${b.border}` + (b.id === selected && mode === 'design' ? ' selected' : '') +
    (isInput(b) ? ' input' : '') + (b.value ? ' filled' : '') + (b.kind === 'text' && b.h > 1 ? ' top' : '');
  el.dataset.id = b.id;
  const line = b.border === 'none' ? 0 : 1;
  Object.assign(el.style, {
    left: `${xs[b.c]}px`, top: `${ys[b.r]}px`, width: `${xs[b.c + b.w] - xs[b.c] + line}px`,
    height: `${ys[b.r + b.h] - ys[b.r] + line}px`, fontSize: `${b.size * 4 / 3}px`, fontWeight: b.bold ? 700 : 400,
    background: b.fill ? `#${b.fill}` : '',
  });
  const t = document.createElement('div');
  t.className = 't';
  t.style.textAlign = b.align === 'distributed' ? 'justify' : b.align;
  if (b.align === 'distributed') t.style.textAlignLast = 'justify';
  if (b.kind === 'check' && mode === 'fill') {
    const picked = new Set(String(b.value || '').split('\n'));
    const opts = document.createElement('div'); opts.className = 'opts';
    for (const o of checkOptions(b)) {
      const l = document.createElement('label');
      l.textContent = `${picked.has(o) ? '■' : '□'} ${o}`;
      l.onclick = (e) => { e.stopPropagation(); remember(); picked.has(o) ? picked.delete(o) : picked.add(o); b.value = [...picked].filter(Boolean).join('\n'); render(); };
      opts.appendChild(l);
    }
    t.appendChild(opts);
  } else if (isInput(b)) {
    t.textContent = display(b);
  } else t.textContent = b.text;
  el.appendChild(t);
  if (mode === 'design' && isField(b)) { const tag = document.createElement('span'); tag.className = 'tag'; tag.textContent = b.name; el.appendChild(tag); }
  if (mode === 'design') { const h = document.createElement('div'); h.className = 'handle'; el.appendChild(h); }
  return el;
}

function display(b) {
  const v = String(b.value ?? '');
  if (!v) return '';
  if (b.type === 'number' && !isNaN(Number(v.replace(/,/g, '')))) return Number(v.replace(/,/g, '')).toLocaleString('ja-JP');
  if (b.type === 'date' && /^\d{4}-\d{2}-\d{2}$/.test(v)) {
    return new Intl.DateTimeFormat('ja-JP-u-ca-japanese', { era: 'long', year: 'numeric', month: 'long', day: 'numeric' }).format(new Date(v + 'T00:00:00'));
  }
  return v;
}

// ---- pointer --------------------------------------------------------------
let drag = null;
/** The last edge at or before `v`: the cell a point falls in. */
function indexAt(list, v) {
  let lo = 0, hi = list.length - 2;
  while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (list[mid] <= v) lo = mid; else hi = mid - 1; }
  return lo;
}
function cellAt(e) {
  const r = paper.getBoundingClientRect();
  const scale = r.width / xs[xs.length - 1];
  return { r: indexAt(ys, (e.clientY - r.top) / scale), c: indexAt(xs, (e.clientX - r.left) / scale) };
}
function ghost(area, ok) {
  let g = paper.querySelector('.ghost');
  if (!g) { g = document.createElement('div'); g.className = 'ghost'; paper.appendChild(g); }
  const r = Math.max(0, area.r), c = Math.max(0, area.c);
  const r2 = Math.min(ys.length - 1, area.r + area.h), c2 = Math.min(xs.length - 1, area.c + area.w);
  Object.assign(g.style, { left: `${xs[c]}px`, top: `${ys[r]}px`, width: `${xs[c2] - xs[c]}px`, height: `${ys[r2] - ys[r]}px` });
  g.classList.toggle('bad', !ok);
}
const clearGhost = () => paper.querySelector('.ghost')?.remove();
const span = (a, b) => ({ r: Math.min(a.r, b.r), c: Math.min(a.c, b.c), h: Math.abs(a.r - b.r) + 1, w: Math.abs(a.c - b.c) + 1 });
const fitsHere = (b) => free(sheet().boxes, b, grid().cols, grid().rows);

paper.addEventListener('pointerdown', (e) => {
  if (mode !== 'design' || e.button !== 0 || e.target.closest('textarea,input,select,button')) return;
  commitEdit();
  const at = cellAt(e);
  const el = tool === 'capture' ? null : e.target.closest('.box');
  if (el) {
    const b = byId(el.dataset.id);
    selected = b.id;
    drag = { kind: e.target.classList.contains('handle') ? 'resize' : 'move', id: b.id, start: at, from: { ...b }, moved: false };
  } else {
    selected = null;
    drag = { kind: 'create', start: at, moved: false };
  }
  render();
  paper.setPointerCapture(e.pointerId);
});
paper.addEventListener('pointermove', (e) => {
  if (!drag) return;
  const at = cellAt(e);
  if (at.r !== drag.start.r || at.c !== drag.start.c) drag.moved = true;
  if (!drag.moved) return;
  if (drag.kind === 'create') {
    drag.area = span(drag.start, at);
    ghost(drag.area, tool === 'capture' || fitsHere({ id: '', ...drag.area }));
    paper.querySelector('.ghost')?.classList.toggle('capture', tool === 'capture');
    return;
  }
  const b = byId(drag.id), f = drag.from;
  const next = drag.kind === 'move'
    ? { ...b, r: f.r + at.r - drag.start.r, c: f.c + at.c - drag.start.c }
    : { ...b, h: Math.max(1, f.h + at.r - drag.start.r), w: Math.max(1, f.w + at.c - drag.start.c) };
  drag.area = next;
  ghost(next, fitsHere(next));
});
paper.addEventListener('pointerup', () => {
  if (!drag) return;
  const d = drag; drag = null; clearGhost();
  if (!d.moved || !d.area) return;
  if (d.kind === 'create' && tool === 'capture') { makeMine(d.area); return; }
  if (!fitsHere({ id: d.id || '', ...d.area })) { toast('ほかの枠と重なるところには置けません（セル結合は重ねられないので）'); return; }
  remember();
  if (d.kind === 'create') {
    const b = makeBox(tool, d.area.r, d.area.c, d.area.h, d.area.w);
    sheet().boxes.push(b); selected = b.id;
    render();
    if (!isInput(b) && b.kind !== 'stamp') startEdit(b, true);
  } else {
    Object.assign(byId(d.id), { r: d.area.r, c: d.area.c, h: d.area.h, w: d.area.w });
    render();
  }
});
paper.addEventListener('dblclick', (e) => {
  const el = e.target.closest('.box');
  if (el && mode === 'design') startEdit(byId(el.dataset.id), true);
});
paper.addEventListener('click', (e) => {
  if (mode !== 'fill') return;
  const el = e.target.closest('.box');
  const b = el && byId(el.dataset.id);
  if (b && isInput(b)) startEdit(b, false);
});

// ---- text editing -----------------------------------------------------------
function startEdit(b, selectAll, initial) {
  commitEdit();
  const el = paper.querySelector(`[data-id="${b.id}"]`);
  if (!el) return;
  const field = mode === 'fill' && isInput(b) && b.type !== 'text' ? document.createElement('input') : document.createElement('textarea');
  if (field.tagName === 'INPUT') field.type = b.type === 'date' ? 'date' : 'text';
  if (field.tagName === 'INPUT' && b.type === 'number') field.inputMode = 'numeric';
  const key = mode === 'fill' || isInput(b) ? 'value' : 'text';
  field.value = initial ?? String(b[key] ?? '');
  field.style.textAlign = b.align === 'distributed' ? 'left' : b.align;
  el.appendChild(field);
  editing = { id: b.id, key, field, before: snapshot() };
  field.focus();
  if (selectAll && initial == null) field.select();
  field.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') { editing = null; render(); }
    else if (e.key === 'Enter' && !e.shiftKey && !e.isComposing && !(b.h > 2 && mode === 'fill')) { e.preventDefault(); commitEdit(); if (mode === 'fill') nextInput(b); }
    else if (e.key === 'Tab' && mode === 'fill') { e.preventDefault(); commitEdit(); nextInput(b, e.shiftKey); }
    e.stopPropagation();
  });
  field.addEventListener('blur', () => setTimeout(() => { if (editing?.field === field) commitEdit(); }, 0));
}
function commitEdit() {
  if (!editing) return;
  const { id, key, field, before } = editing;
  editing = null;
  const b = byId(id);
  if (b && b[key] !== field.value) { history.push(before); future = []; b[key] = field.value; }
  render();
}
function nextInput(b, back = false) {
  const inputs = sheet().boxes.filter(isInput).sort((x, y) => x.r - y.r || x.c - y.c);
  const i = inputs.findIndex((x) => x.id === b.id);
  const n = inputs[(i + (back ? -1 : 1) + inputs.length) % inputs.length];
  if (n) startEdit(n, true);
}

// ---- keyboard -------------------------------------------------------------
document.addEventListener('keydown', (e) => {
  if (editing || e.target.closest?.('input,textarea,select,dialog')) return;
  const mod = e.ctrlKey || e.metaKey;
  if (mod && e.key.toLowerCase() === 'z') { e.preventDefault(); e.shiftKey ? redo() : undo(); return; }
  if (mod && e.key.toLowerCase() === 'y') { e.preventDefault(); redo(); return; }
  if (mod && e.key.toLowerCase() === 's') { e.preventDefault(); saveXlsx(); return; }
  const b = selected && byId(selected);
  if (!b || mode !== 'design') return;
  if ((e.key === 'Delete' || e.key === 'Backspace') && e.shiftKey && b.block) { e.preventDefault(); blockAction(b.block, 'remove'); return; }
  if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); remember(); sheet().boxes = sheet().boxes.filter((x) => x.id !== b.id); selected = null; render(); return; }
  const arrows = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
  if (arrows[e.key]) {
    e.preventDefault();
    const [dr, dc] = arrows[e.key];
    const next = e.shiftKey ? { ...b, h: b.h + dr, w: b.w + dc } : { ...b, r: b.r + dr, c: b.c + dc };
    if (fitsHere(next)) { remember(); Object.assign(b, next); render(); }
    return;
  }
  if (mod && e.key.toLowerCase() === 'd') {
    e.preventDefault();
    const copy = { ...structuredClone(b), id: makeBox(b.kind, 0, 0, 1, 1).id, r: b.r + b.h, name: '', value: '', block: undefined };
    if (fitsHere(copy)) { remember(); sheet().boxes.push(copy); selected = copy.id; render(); } else toast('下に空きがありません');
    return;
  }
  if (e.key === 'Enter' || e.key === 'F2') { e.preventDefault(); startEdit(b, true); return; }
  if (e.key.length === 1 && !mod && !isInput(b)) { e.preventDefault(); startEdit(b, false, e.key); }
});

// ---- blocks, pages -----------------------------------------------------------
let blockTab = GROUPS[0];
function blockPalette() {
  const tabsEl = $('block-tabs'); tabsEl.textContent = '';
  for (const g of GROUPS) {
    const t = document.createElement('button'); t.textContent = g; t.className = g === blockTab ? 'on' : '';
    t.onclick = () => { blockTab = g; blockPalette(); };
    tabsEl.appendChild(t);
  }
  const host = $('blocks'); host.textContent = '';
  if (blockTab === MINE) minePalette(host);
  for (const [name, def] of Object.entries(BLOCKS)) {
    if (def.group !== blockTab) continue;
    if (def.mine) {
      const row = document.createElement('div'); row.className = 'mine';
      const btn = document.createElement('button'); btn.textContent = def.title;
      btn.onclick = () => { hidePreview(); addBlock(name); };
      btn.onmouseenter = btn.onfocus = () => showPreview(name, btn);
      btn.onmouseleave = btn.onblur = hidePreview;
      const del = document.createElement('button'); del.textContent = '✕'; del.className = 'danger'; del.title = 'このマイブロックを消す';
      del.onclick = async () => { if (await ask(`マイブロック「${def.title}」を消しますか？`, '消す')) { delete mine[name]; unregisterMine(name); saveMine(); blockPalette(); } };
      row.append(btn, del);
      host.appendChild(row);
      continue;
    }
    const btn = document.createElement('button');
    btn.textContent = def.title;
    btn.onclick = () => { hidePreview(); addBlock(name); };
    btn.onmouseenter = btn.onfocus = () => showPreview(name, btn);
    btn.onmouseleave = btn.onblur = hidePreview;
    host.appendChild(btn);
  }
}
// ---- BlockMaker ------------------------------------------------------------------
const MINE_STORE = 'graphPaper:blocks';
let mine = {};
function saveMine() { try { localStorage.setItem(MINE_STORE, JSON.stringify(mine)); } catch { /* storage off */ } }
function loadMine() {
  try { mine = JSON.parse(localStorage.getItem(MINE_STORE) || '{}') || {}; } catch { mine = {}; }
  for (const [key, def] of Object.entries(mine)) registerMine(key, def);
}
function addMine(def) {
  const key = `mine:${Date.now().toString(36)}${Math.random().toString(36).slice(2, 5)}`;
  mine[key] = def; registerMine(key, def);
  return key;
}
/** Save the boxes inside an area as a block of one's own. */
async function makeMine(area, suggested) {
  const res = captureBlock(sheet().boxes, area);
  if (res.error) { toast(res.error); return; }
  const title = await askText('マイブロックの名前', suggested || '');
  if (title == null) return;
  res.def.title = title.trim().slice(0, 30) || 'マイブロック';
  addMine(res.def); saveMine();
  tool = 'label'; palette();
  blockTab = MINE; blockPalette();
  toast(`マイブロック「${res.def.title}」を作りました（${res.def.rows} 行）` + (res.partial ? `。囲みからはみ出した枠 ${res.partial} 個は入れていません` : ''));
}
function minePalette(host) {
  const tools = document.createElement('div'); tools.className = 'minetools';
  const make = document.createElement('button'); make.textContent = tool === 'capture' ? '囲んでください…（やめる）' : '＋ 範囲を囲んで作る';
  make.className = tool === 'capture' ? 'on' : '';
  make.title = '方眼紙の上をドラッグして囲んだ枠が、マイブロックになります';
  make.onclick = () => { if (mode !== 'design') setMode('design'); tool = tool === 'capture' ? 'label' : 'capture'; palette(); blockPalette(); if (tool === 'capture') toast('ブロックにしたい枠を、ドラッグで囲んでください'); };
  const out = document.createElement('button'); out.textContent = '書き出す';
  out.title = 'マイブロックをファイルにして、ほかの人に渡せます';
  out.onclick = () => {
    const list = Object.values(mine);
    if (!list.length) { toast('マイブロックがまだありません'); return; }
    download(JSON.stringify({ graphPaper: 'blocks', version: 1, blocks: list }, null, 1), 'マイブロック.graphpaper-blocks.json', 'application/json');
  };
  const inp = document.createElement('label'); inp.className = 'button'; inp.textContent = '読み込む';
  const file = document.createElement('input'); file.type = 'file'; file.accept = '.json'; file.hidden = true;
  file.onchange = async () => {
    const f = file.files[0]; if (!f) return;
    try {
      const list = readShared(await f.text());
      if (!list.length) { toast('読み込めるブロックがありませんでした'); return; }
      for (const def of list) addMine(def);
      saveMine(); blockPalette();
      toast(`マイブロックを ${list.length} 個読み込みました`);
    } catch (e) { toast('読み込めませんでした: ' + (e.message || e)); }
  };
  inp.appendChild(file);
  tools.append(make, out, inp);
  host.appendChild(tools);
  if (!Object.keys(mine).length) {
    const p = document.createElement('p'); p.className = 'hint';
    p.textContent = 'よく使う欄の組み合わせを、自分のブロックにできます。枠を並べてから「範囲を囲んで作る」で囲むか、ブロックの上のバーの「★」から。';
    host.appendChild(p);
  }
}

/** Where a block goes: below the selected box's block (or the selected box), else at the bottom. */
function insertRow() {
  const s = sheet();
  const sel = selected && byId(selected);
  const band = sel?.block && bands(s.boxes, s.blocks).find((k) => k.id === sel.block);
  return band ? band.end : sel ? sel.r + sel.h : bottomRow(s.boxes, s.blocks);
}

/** A small drawing of the block beside the palette, and a line on the page where it goes. */
function showPreview(name, btn) {
  const def = BLOCKS[name], g = grid();
  const boxes = def.make(0, g.cols);
  const width = 300, cell = width / g.cols, height = Math.max(1, def.rows) * cell;
  const pop = $('preview');
  pop.textContent = '';
  const head = document.createElement('div'); head.className = 'phead';
  const at = insertRow();
  const room = g.rows - bottomRow(sheet().boxes, sheet().blocks);
  head.innerHTML = `<strong>${def.title}</strong><span>${def.rows} 行</span>`;
  const art = document.createElement('div'); art.className = 'pbody';
  Object.assign(art.style, { width: `${width}px`, height: `${height}px`, backgroundSize: `${cell}px ${cell}px` });
  for (const b of boxes) {
    const el = document.createElement('div');
    el.className = `pbox b-${b.border}`;
    Object.assign(el.style, {
      left: `${b.c * cell}px`, top: `${b.r * cell}px`, width: `${b.w * cell}px`, height: `${b.h * cell}px`,
      background: b.fill ? `#${b.fill}` : isInput(b) ? 'rgba(31,111,235,.08)' : '', fontSize: `${Math.max(6, b.size * 4 / 3 * cell / 20)}px`,
      fontWeight: b.bold ? 700 : 400, justifyContent: b.align === 'center' ? 'center' : b.align === 'right' ? 'flex-end' : 'flex-start',
    });
    el.textContent = isInput(b) ? '' : b.kind === 'stamp' ? (b.text || '') : b.text;
    art.appendChild(el);
  }
  const foot = document.createElement('div'); foot.className = 'pfoot';
  foot.textContent = `${at + 1} 行目に入ります。下にあるものは ${def.rows} 行下がります。` +
    (def.rows > room ? '入りきらない分は、ページを足します。' : '');
  pop.append(head, art, foot);
  const r = btn.getBoundingClientRect();
  Object.assign(pop.style, { left: `${r.right + 10}px`, top: `${Math.min(r.top, innerHeight - height - 110)}px` });
  pop.hidden = false;
  if (mode === 'design') {
    const line = document.createElement('div'); line.className = 'insertline';
    line.style.top = `${ys[at]}px`;
    line.dataset.label = `ここに「${def.title}」`;
    paper.appendChild(line);
  }
}
function hidePreview() {
  $('preview').hidden = true;
  paper.querySelector('.insertline')?.remove();
}

/** Pages to add so a block of `rows` fits below everything: 0 when it already does. */
function pagesNeeded(s, rows) {
  let grown = s, added = 0;
  while (grown.grid.rows - bottomRow(grown.boxes, grown.blocks) < rows) {
    const next = addPage(grown);
    if (next.error) return { error: next.error };
    grown = next; added++;
  }
  return { sheet: grown, added };
}

function addBlock(name) {
  if (mode !== 'design') setMode('design');
  const at = insertRow();
  const rows = BLOCKS[name].rows;
  // A full page grows by a page rather than turning the block away, and a
  // block that would run over a page break starts on the next page instead.
  let room = pagesNeeded(sheet(), rows);
  if (room.error) { toast(room.error); return; }
  const breakAt = pages(room.sheet.grid).starts.find((p) => p > at && p < at + rows);
  const pad = breakAt != null && rows <= breakAt - (pages(room.sheet.grid).starts.filter((p) => p < breakAt).pop() ?? 0) ? breakAt - at : 0;
  if (pad) { room = pagesNeeded(room.sheet, rows + pad); if (room.error) { toast(room.error); return; } room.added = pagesNeeded(sheet(), rows + pad).added; }
  const s = room.sheet;
  const res = insertBlock(s.boxes, s.blocks, name, at, s.grid, pad);
  if (res.error) { toast(res.error); return; }
  remember();
  setSheet({ ...s, boxes: res.boxes, blocks: res.blocks });
  selected = res.boxes.find((b) => b.block === res.id)?.id || null;
  render();
  if (room.added) toast(`入りきらないので、ページを ${room.added} 枚足しました（${pages(grid()).starts.length} ページ）`);
  paper.querySelector('.band')?.scrollIntoView({ block: 'nearest', inline: 'nearest', behavior: 'smooth' });
}
function apply(res) {
  if (res.error) { toast(res.error); return false; }
  remember();
  setSheet(res.sheet || { ...sheet(), ...res });
  if (!byId(selected)) selected = null;
  render();
  return true;
}
function blockAction(id, action) {
  const s = sheet();
  const res = action === 'remove' ? removeBlock(s.boxes, s.blocks, id) : moveBlock(s.boxes, s.blocks, id, action === 'up' ? -1 : 1);
  if (apply(res.error ? res : { boxes: res.boxes, blocks: res.blocks }) && action === 'remove') toast('ブロックを消しました（Ctrl+Z で戻せます）');
}
async function pageAction(k, action) {
  if (action === 'remove') {
    const p = pageBands(grid())[k];
    const lost = sheet().boxes.filter((b) => b.r >= p.start && b.r < p.end).length;
    if (lost && !(await ask(`${k + 1}ページ目と、その上の枠 ${lost} 個を消しますか？`, '消す'))) return;
  }
  const res = action === 'remove' ? removePage(sheet(), k) : movePage(sheet(), k, action === 'up' ? -1 : 1);
  apply(res.error ? res : { sheet: res });
}

// ---- tool palette & inspector --------------------------------------------------
function palette() {
  const host = $('kinds'); host.textContent = '';
  for (const [k, v] of Object.entries(KINDS)) {
    const btn = document.createElement('button');
    btn.textContent = v.title; btn.className = k === tool ? 'on' : '';
    btn.onclick = () => { tool = k; palette(); };
    host.appendChild(btn);
  }
}

function inspector() {
  const host = $('inspector'); host.textContent = '';
  const row = (title, node) => { const d = document.createElement('div'); d.className = 'row'; const h = document.createElement('h2'); h.textContent = title; d.append(h, node); host.appendChild(d); };
  const seg = (options, current, set) => {
    const d = document.createElement('div'); d.className = 'seg';
    for (const [value, label] of options) {
      const btn = document.createElement('button'); btn.textContent = label; if (value === current) btn.className = 'on';
      btn.onclick = () => set(value);
      d.appendChild(btn);
    }
    return d;
  };
  const change = (fn) => (v) => { remember(); fn(v); render(); };
  const b = selected && byId(selected);
  if (mode === 'fill') {
    const n = sheet().boxes.filter(isField).length;
    host.innerHTML = `<div class="empty">記入するモードです。入力欄（${n} 個）をクリックして書き込みます。Tab で次の欄へ。<br><br>終わったら「Excel で保存」か「PDF」。</div>`;
    return;
  }
  if (!b) { sheetPanel(row, seg); return; }
  const s = sheet();
  const band = b.block && bands(s.boxes, s.blocks).find((k) => k.id === b.block);
  if (band) {
    const d = document.createElement('div'); d.className = 'seg';
    for (const [act, label] of [['up', '↑ 上へ'], ['down', '↓ 下へ'], ['remove', '✕ ブロックを消す']]) {
      const btn = document.createElement('button'); btn.textContent = label; btn.onclick = () => blockAction(band.id, act);
      if (act === 'remove') btn.className = 'danger';
      d.appendChild(btn);
    }
    row(`ブロック「${BLOCKS[band.name]?.title || ''}」`, d);
  }
  const where = document.createElement('div'); where.className = 'where';
  where.textContent = `${ref(b)}（${b.w} × ${b.h} マス）`;
  row('場所', where);
  row('種類', seg(Object.entries(KINDS).map(([k, v]) => [k, v.title]), b.kind, change((k) => {
    const d = KINDS[k];
    Object.assign(b, { kind: k, border: d.border, fill: d.fill, align: d.align, size: d.size || 11, bold: !!d.bold });
    if (!isInput(b) && !b.text) b.text = d.text;
    if (isInput(b)) b.text = '';
  })));
  if (isField(b)) {
    const name = document.createElement('input'); name.type = 'text'; name.value = b.name;
    name.onchange = () => { remember(); b.name = name.value.trim(); render(); };
    row('名前（Excel の名前の定義）', name);
  }
  if (isInput(b)) row('型', seg([['text', '文字'], ['number', '数値'], ['date', '日付']], b.type, change((t) => { b.type = t; })));
  else {
    const text = document.createElement('textarea'); text.value = b.text;
    text.onchange = () => { remember(); b.text = text.value; render(); };
    row(b.kind === 'check' ? '選択肢（□ で区切る）' : '文字', text);
  }
  row('罫線', seg([['box', '囲む'], ['thick', '太枠'], ['under', '下線'], ['none', 'なし']], b.border, change((v) => { b.border = v; })));
  row('塗り', seg([[null, 'なし'], ['F2F2F2', 'グレー'], ['DDEBF7', '水色'], ['FFF2CC', '黄色']], b.fill, change((v) => { b.fill = v; })));
  row('文字の大きさ', seg([[9, '9'], [10.5, '10.5'], [11, '11'], [14, '14'], [16, '16'], [20, '20']], b.size, change((v) => { b.size = v; })));
  row('揃え', seg([['left', '左'], ['center', '中央'], ['right', '右'], ['distributed', '均等割付']], b.align, change((v) => { b.align = v; })));
  row('太字', seg([[false, 'なし'], [true, '太字']], b.bold, change((v) => { b.bold = v; })));
  const del = document.createElement('button'); del.textContent = 'この枠を消す';
  del.onclick = () => { remember(); sheet().boxes = sheet().boxes.filter((x) => x.id !== b.id); selected = null; render(); };
  host.appendChild(del);
}

/** With nothing selected: the sheet, its grid and its pages. */
function sheetPanel(row, seg) {
  const s = sheet(), g = s.grid;
  const name = document.createElement('input'); name.type = 'text'; name.value = s.name;
  name.onchange = () => renameSheet(active, name.value);
  row('シート名', name);
  row('向き（A4）', seg([[false, '縦'], [true, '横']], g.landscape, (v) => {
    remember();
    const pristine = !s.boxes.length && g.cols === (g.landscape ? 50 : 34) && g.rows === (g.landscape ? 34 : 50);
    setSheet({ ...s, grid: pristine ? newGrid(v) : { ...g, landscape: v } });
    render();
  }));
  const size = document.createElement('div'); size.className = 'size';
  const num = (value, max, title) => { const i = document.createElement('input'); i.type = 'number'; i.min = 1; i.max = max; i.value = value; i.title = title; return i; };
  const cols = num(g.cols, MAX_COLS, '列の数'), rows = num(g.rows, MAX_ROWS, '行の数');
  const go = document.createElement('button'); go.textContent = '広げる・狭める';
  go.onclick = () => resize(Number(cols.value), Number(rows.value));
  size.append(cols, document.createTextNode(' 列 × '), rows, document.createTextNode(' 行 '), go);
  row('方眼紙の大きさ', size);
  const grow = document.createElement('div'); grow.className = 'seg';
  for (const [label, fn] of [['下に1ページ足す', () => { const r = addPage(sheet()); apply(r.error ? r : { sheet: r }); }], ['右に5列広げる', () => resize(g.cols + 5, g.rows)]]) {
    const b = document.createElement('button'); b.textContent = label; b.onclick = fn; grow.appendChild(b);
  }
  row('広げる', grow);
  const tabs = document.createElement('div'); tabs.className = 'seg';
  for (const [label, fn] of [['複製', () => duplicateSheet()], ['← 左へ', () => moveSheet(-1)], ['右へ →', () => moveSheet(1)], ['シートを消す', () => deleteSheet()]]) {
    const b = document.createElement('button'); b.textContent = label; b.onclick = fn; tabs.appendChild(b);
  }
  row('シート', tabs);
  const opt = document.createElement('label'); opt.className = 'inline';
  opt.innerHTML = `<input type="checkbox" ${protect ? 'checked' : ''}> Excel で開いたとき、入力欄以外を保護する`;
  opt.querySelector('input').onchange = (e) => { protect = e.target.checked; persist(); };
  host().appendChild(opt);
  const hint = document.createElement('p'); hint.className = 'empty';
  hint.textContent = '枠を選ぶと、ここで種類・罫線・文字を変えられます。入力欄には Excel の「名前の定義」が付きます。名前は左（なければ上）の見出しから自動で決まります。';
  host().appendChild(hint);
  function host() { return $('inspector'); }
}

function resize(cols, rows) {
  const g = grid();
  if (!(cols >= 1 && rows >= 1 && cols <= MAX_COLS && rows <= MAX_ROWS)) { toast(`1〜${MAX_COLS} 列、1〜${MAX_ROWS} 行の範囲で`); return; }
  const cut = sheet().boxes.filter((b) => b.c + b.w > cols || b.r + b.h > rows);
  if (cut.length) { toast(`狭めると枠が ${cut.length} 個はみ出します。先に枠を動かすか消してください`); return; }
  apply({ grid: resizeGrid(g, cols, rows) });
  const { scale } = pages(grid());
  if (scale < 1) toast(`横幅がページを超えたので、印刷は ${Math.round(scale * 100)}% に縮みます`);
}

// ---- sheets (tabs) ---------------------------------------------------------------
function tabs() {
  const host = $('tabs'); host.textContent = '';
  sheets.forEach((s, i) => {
    const t = document.createElement('button');
    t.className = 'tab' + (i === active ? ' on' : '');
    t.textContent = s.name;
    t.title = 'ダブルクリックで名前を変える';
    t.onclick = () => switchSheet(i);
    t.ondblclick = async () => { const n = await askText('シート名', s.name); if (n != null) renameSheet(i, n); };
    host.appendChild(t);
  });
  const add = document.createElement('button'); add.className = 'tab add'; add.textContent = '＋'; add.title = 'シートを足す';
  add.onclick = () => { remember(); sheets.push(newSheet(sheetName('Sheet' + (sheets.length + 1), sheets.map((s) => s.name)))); switchSheet(sheets.length - 1); };
  host.appendChild(add);
}
function switchSheet(i) { commitEdit(); active = i; selected = null; render(); }
function renameSheet(i, name) {
  const n = sheetName(name, sheets.filter((_, k) => k !== i).map((s) => s.name));
  if (n === sheets[i].name) { render(); return; }
  remember(); sheets[i] = { ...sheets[i], name: n }; render();
}
function duplicateSheet() {
  remember();
  const copy = structuredClone(sheet());
  copy.name = sheetName(copy.name, sheets.map((s) => s.name));
  for (const b of copy.boxes) { b.id = makeBox(b.kind, 0, 0, 1, 1).id; b.value = ''; }
  sheets.splice(active + 1, 0, copy);
  switchSheet(active + 1);
}
function moveSheet(dir) {
  const j = active + dir;
  if (j < 0 || j >= sheets.length) return;
  remember();
  [sheets[active], sheets[j]] = [sheets[j], sheets[active]];
  switchSheet(j);
}
async function deleteSheet() {
  if (sheets.length < 2) { toast('シートは1枚は残します'); return; }
  if (sheet().boxes.length && !(await ask(`シート「${sheet().name}」を消しますか？`, '消す'))) return;
  remember();
  sheets.splice(active, 1);
  switchSheet(Math.min(active, sheets.length - 1));
}

function footer() {
  const g = grid();
  const { scale, starts } = pages(g);
  const square = g.colPx.every((p) => p === g.colPx[0]) && g.rowPx.every((p) => p === g.rowPx[0]);
  $('gridinfo').textContent = `A4 ${g.landscape ? '横' : '縦'}・${g.cols} × ${g.rows} マス・${starts.length} ページ` +
    (square ? `・1 マス ${g.colPx[0]} × ${g.rowPx[0]} ピクセル` : '・マスの大きさは Excel のまま') +
    (scale < 1 ? `・印刷 ${Math.round(scale * 100)}%` : '');
}

// ---- modes, files -------------------------------------------------------------
function setMode(m) {
  commitEdit();
  mode = m;
  document.body.classList.toggle('fill', m === 'fill');
  $('mode-design').classList.toggle('on', m === 'design');
  $('mode-fill').classList.toggle('on', m === 'fill');
  render();
}
$('mode-design').onclick = () => setMode('design');
$('mode-fill').onclick = () => setMode('fill');

$('new').onclick = async () => {
  if (sheets.some((s) => s.boxes.length) && !(await ask('いまの方眼紙を消して、白紙から始めますか？（Ctrl+Z で戻せます）', '白紙にする'))) return;
  remember(); sheets = [newSheet('様式')]; active = 0; selected = null; render();
  status('白紙の方眼紙です。左の「ブロック」を押すと、上から順に積めます。');
};
$('sample').onclick = async () => {
  if (sheets.some((s) => s.boxes.length) && !(await ask('いまの方眼紙を、見本の申請書に置き換えますか？（Ctrl+Z で戻せます）', '置き換える'))) return;
  remember(); sheets = [sampleForm()]; active = 0; selected = null; render(); status('見本の申請書を開きました。');
};

/** The pixel width of the digit 0 in a workbook's normal font, as Excel counts it. */
function digitWidth(style) {
  const known = { '游ゴシック': 8, 'Yu Gothic': 8, '游明朝': 8, 'Calibri': 7, 'ＭＳ Ｐゴシック': 8, 'ＭＳ ゴシック': 8, 'ＭＳ Ｐ明朝': 8, 'ＭＳ 明朝': 8, 'Meiryo UI': 8, 'メイリオ': 9, 'Arial': 7 };
  const size = style?.font_size || 11;
  const face = style?.font_name || 'Calibri';
  if (known[face] && size === 11) return known[face];
  try {
    const pen = document.createElement('canvas').getContext('2d');
    pen.font = `${size * 96 / 72}px "${face}", sans-serif`;
    return Math.round(pen.measureText('0').width) || 8;
  } catch { return 8; }
}

$('file').onchange = async (e) => {
  const f = e.target.files[0]; e.target.value = '';
  if (!f) return;
  try {
    status(`${f.name} を読み込んでいます…`);
    const api = await engine();
    const wb = api.parse_spreadsheet(new Uint8Array(await f.arrayBuffer()));
    const digit = digitWidth(wb.default_style);
    const opened = [];
    let inputs = 0, boxes = 0, clipped = 0;
    wb.sheets.forEach((s, i) => {
      if (s.visibility && !/^visible$/i.test(String(s.visibility))) return;
      const res = fromWorkbook(wb, i, digit);
      res.sheet.name = sheetName(res.sheet.name, opened.map((x) => x.name));
      opened.push(res.sheet);
      inputs += res.sheet.boxes.filter(isField).length; boxes += res.sheet.boxes.length; clipped += res.clipped;
    });
    if (!opened.length) { status('開けるシートがありませんでした', true); return; }
    remember(); sheets = opened; active = 0; selected = null; render();
    status(`${f.name}: ${opened.length} シート、枠 ${boxes} 個、うち入力欄 ${inputs} 個を見つけました` + (clipped ? `（範囲の外の ${clipped} 個は取り込んでいません）` : ''));
  } catch (err) { console.error(err); status('開けませんでした: ' + (err.message || err), true); }
};

async function buildXlsx() {
  const [api, bytes] = await Promise.all([engine(), templateBytes()]);
  const wb = toWorkbook(api.parse_spreadsheet(bytes), structuredClone(sheets), { protect, filled: true });
  return { api, xlsx: api.edit_xlsx_from_workbook(bytes, wb) };
}
const bookTitle = () => titleOf(sheets[0].boxes);
const fileTitle = () => bookTitle().replace(/[\\/:*?"<>|]/g, '_');
/** Save through the system's save dialog in the desktop app, as a download on the site. */
async function download(data, name, type) {
  if (DESKTOP) {
    const T = window.__TAURI__;
    const ext = name.split('.').pop();
    const filters = [{ name: { xlsx: 'Excel ブック', pdf: 'PDF', md: 'Markdown' }[ext] || ext, extensions: [ext] }];
    const path = T.dialog ? await T.dialog.save({ defaultPath: name, filters })
      : await T.core.invoke('plugin:dialog|save', { options: { defaultPath: name, filters } });
    if (!path) return false;
    const bytes = typeof data === 'string' ? new TextEncoder().encode(data) : data;
    if (T.fs) await T.fs.writeFile(path, bytes);
    else await T.core.invoke('plugin:fs|write_file', bytes, { headers: { path: encodeURIComponent(path), options: '{}' } });
    toast(`保存しました: ${path}`);
    return true;
  }
  const url = URL.createObjectURL(new Blob([data], { type }));
  const a = Object.assign(document.createElement('a'), { href: url, download: name });
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}
async function saveXlsx() {
  commitEdit();
  try {
    status('Excel ファイルを作っています…');
    const { xlsx } = await buildXlsx();
    if (await download(xlsx, fileTitle() + '.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') === false) { status('保存をやめました'); return; }
    status(`保存しました（${sheets.length} シート、${Math.round(xlsx.length / 1024)} KB）。Excel で開くと、同じ方眼紙です。`);
  } catch (err) { console.error(err); status('保存できませんでした: ' + (err.message || err), true); }
}
$('save-xlsx').onclick = saveXlsx;
$('pdf').onclick = async () => {
  commitEdit();
  try {
    status('PDF を作っています…');
    const t0 = performance.now();
    const { api, xlsx } = await buildXlsx();
    const pdf = await withFonts(api, () => api.xlsx_to_pdf(xlsx, bookTitle()));
    const url = URL.createObjectURL(new Blob([pdf], { type: 'application/pdf' }));
    const frame = document.createElement('iframe'); frame.src = url;
    const save = document.createElement('button'); save.textContent = 'PDF を保存';
    save.onclick = () => download(pdf, fileTitle() + '.pdf', 'application/pdf');
    dialog('PDF（保存した .xlsx を Oxi で印刷したもの）', frame, [save], () => URL.revokeObjectURL(url));
    status(`PDF を作りました（${Math.round(performance.now() - t0)} ms、ブラウザの中だけで変換）`);
  } catch (err) { console.error(err); status('PDF にできませんでした: ' + (err.message || err), true); }
};
$('markdown').onclick = () => {
  commitEdit();
  const md = toMarkdown(sheets, bookTitle());
  const area = document.createElement('textarea'); area.value = md; area.readOnly = true;
  const copy = document.createElement('button'); copy.textContent = 'コピー';
  copy.onclick = async () => { try { await navigator.clipboard.writeText(md); toast('コピーしました'); } catch { area.select(); } };
  const save = document.createElement('button'); save.textContent = '.md を保存';
  save.onclick = () => download(md, fileTitle() + '.md', 'text/markdown');
  dialog('AI 用に書き出す（どの欄がどのセルか、何が書いてあるか）', area, [copy, save]);
};

function dialog(title, body, actions = [], onClose) {
  const d = $('dlg');
  $('dlg-title').textContent = title;
  $('dlg-body').replaceChildren(body);
  const close = document.createElement('button'); close.textContent = '閉じる'; close.onclick = () => d.close();
  $('dlg-actions').replaceChildren(...actions, close);
  d.onclose = () => onClose?.();
  d.showModal();
}
/**
 * A question in the page itself, answered later: Promise<boolean>.
 * Not window.confirm: inside the desktop app Tauri swaps that for a dialog
 * that does not wait for the answer, so a "消しますか？" would not hold.
 */
function ask(message, yes = 'OK') {
  return askDialog(message, null, yes).then((v) => v !== null);
}
/** A line of text asked for in the page: Promise<string | null>. */
function askText(label, value = '') {
  return askDialog(label, value, 'OK');
}
function askDialog(message, value, yes) {
  return new Promise((resolve) => {
    const d = $('ask');
    $('ask-text').textContent = message;
    const input = $('ask-input');
    input.hidden = value === null;
    input.value = value ?? '';
    $('ask-yes').textContent = yes;
    let answer = null;
    $('ask-yes').onclick = () => { answer = value === null ? '' : input.value; d.close(); };
    $('ask-no').onclick = () => d.close();
    input.onkeydown = (e) => { if (e.key === 'Enter' && !e.isComposing) { e.preventDefault(); $('ask-yes').click(); } };
    d.onclose = () => resolve(answer);
    d.showModal();
    (value === null ? $('ask-yes') : input).focus();
    if (value !== null) input.select();
  });
}
let toastTimer = 0;
function toast(text) {
  const t = $('toast'); t.textContent = text; t.classList.add('show');
  clearTimeout(toastTimer); toastTimer = setTimeout(() => t.classList.remove('show'), 3600);
}
function status(text, error = false) { const s = $('status'); s.textContent = text; s.classList.toggle('error', error); }

let zoom = 1;
function setZoom(z) { zoom = Math.max(0.4, Math.min(2, Math.round(z * 10) / 10)); $('sheet').style.zoom = zoom; $('zoom').textContent = `${Math.round(zoom * 100)}%`; }
$('zoom-in').onclick = () => setZoom(zoom + 0.1);
$('zoom-out').onclick = () => setZoom(zoom - 0.1);

// ---- start ----------------------------------------------------------------------
try {
  const saved = JSON.parse(localStorage.getItem(STORE) || 'null');
  if (saved?.sheets?.length) { sheets = saved.sheets; active = Math.min(saved.active || 0, sheets.length - 1); protect = saved.protect ?? true; }
} catch { /* storage off */ }
if (!sheets.length) sheets = [sampleForm()];
loadMine();
palette();
blockPalette();
render();
engine().catch((e) => status('変換エンジンを読み込めませんでした: ' + (e.message || e), true));
