// SPDX-License-Identifier: MIT OR Apache-2.0
// The form as boxes on a square grid, and the way to and from an Excel sheet.
//
// A box is a rectangle of whole cells with one role: a heading, an input,
// a seal, a check list or free text. In Excel it becomes one merged area with
// its borders on the cells along its edges, and an input box also gets a
// defined name, so the saved workbook is an ordinary Excel 方眼紙 that still
// knows which cell is which.
//
// A sheet is { name, grid, boxes, blocks }. The grid holds each column's and
// row's size in screen pixels (96 dpi), the way Excel draws them, so a grid
// opened from any Excel 方眼紙 keeps its own cells.

export const CELL = 20;
/** The printable part of an A4 page inside 15 mm margins, in pixels. */
export const PAGE = { portrait: { w: 680, h: 1009 }, landscape: { w: 1009, h: 680 } };

export function newGrid(landscape = false, cols = landscape ? 50 : 34, rows = landscape ? 34 : 50) {
  return { cols, rows, colPx: Array(cols).fill(CELL), rowPx: Array(rows).fill(CELL), landscape };
}
export const newSheet = (name = '様式', landscape = false) => ({ name, grid: newGrid(landscape), boxes: [], blocks: [] });

/** Positions of the grid lines: edges(sizes)[i] is where cell i starts. */
export function edges(sizes) {
  const out = [0];
  for (const s of sizes) out.push(out[out.length - 1] + s);
  return out;
}

/** How Excel prints the grid: the width fitted to the page, the rows run on
 * to further pages. Returns { scale, starts } with the first row of each page. */
export function pages(grid) {
  const page = PAGE[grid.landscape ? 'landscape' : 'portrait'];
  const width = grid.colPx.reduce((a, b) => a + b, 0);
  const scale = Math.min(1, page.w / Math.max(1, width));
  const tall = page.h / scale;
  const starts = [0];
  let used = 0;
  grid.rowPx.forEach((h, r) => {
    if (used + h > tall + 0.01 && used > 0) { starts.push(r); used = 0; }
    used += h;
  });
  return { scale, starts };
}

/** Grow or shrink the grid; new cells take the size of the last one. */
export function resizeGrid(grid, cols, rows) {
  const grow = (list, n) => (n <= list.length ? list.slice(0, n) : [...list, ...Array(n - list.length).fill(list[list.length - 1] || CELL)]);
  return { ...grid, cols, rows, colPx: grow(grid.colPx, cols), rowPx: grow(grid.rowPx, rows) };
}

export const KINDS = {
  label: { title: '見出し', border: 'box', fill: 'F2F2F2', align: 'center', text: '項目' },
  input: { title: '入力欄', border: 'box', fill: null, align: 'left', text: '' },
  line: { title: '下線の入力欄', border: 'under', fill: null, align: 'left', text: '' },
  check: { title: 'チェック', border: 'none', fill: null, align: 'left', text: '□ はい　□ いいえ' },
  stamp: { title: '印', border: 'box', fill: null, align: 'center', text: '印', size: 9 },
  text: { title: '文字', border: 'none', fill: null, align: 'left', text: '文字' },
  heading: { title: '表題', border: 'none', fill: null, align: 'center', text: '表題', size: 16, bold: true },
};
export const isInput = (b) => b.kind === 'input' || b.kind === 'line';
/** Boxes that hold what the person filling in writes or picks: named in Excel. */
export const isField = (b) => isInput(b) || b.kind === 'check';

let serial = 0;
export function makeBox(kind, r, c, h, w, extra = {}) {
  const k = KINDS[kind];
  return {
    id: `b${Date.now().toString(36)}${(serial++).toString(36)}`, kind, r, c, h, w,
    text: k.text, border: k.border, fill: k.fill, align: k.align,
    size: k.size || 11, bold: !!k.bold, type: 'text', name: '', value: '', ...extra,
  };
}

export const overlaps = (a, b) => a.r < b.r + b.h && b.r < a.r + a.h && a.c < b.c + b.w && b.c < a.c + a.w;
export const fits = (b, cols, rows) => b.r >= 0 && b.c >= 0 && b.h >= 1 && b.w >= 1 && b.r + b.h <= rows && b.c + b.w <= cols;
export const free = (boxes, b, cols, rows) => fits(b, cols, rows) && !boxes.some((o) => o.id !== b.id && overlaps(o, b));

export function colName(i) {
  let s = '';
  for (let n = i + 1; n; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}
export const ref = (b) => `${colName(b.c)}${b.r + 1}` + (b.h > 1 || b.w > 1 ? `:${colName(b.c + b.w - 1)}${b.r + b.h}` : '');
export const quoteSheet = (name) => `'${String(name).replace(/'/g, "''")}'`;
const absRef = (sheet, b) => `${quoteSheet(sheet)}!` + ref(b).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2');

/** A sheet name Excel accepts, unique among `taken`. */
export function sheetName(name, taken = []) {
  let n = String(name).replace(/[\[\]:*?/\\]/g, '_').replace(/^'+|'+$/g, '').trim().slice(0, 31) || 'Sheet';
  const used = new Set(taken.map((t) => t.toLowerCase()));
  for (let k = 2, base = n; used.has(n.toLowerCase()); k++) n = `${base.slice(0, 28)}(${k})`;
  return n;
}

/** A string Excel accepts as a defined name, close to the one asked for. */
export function excelName(name) {
  let n = String(name).trim().replace(/[\s　()（）\[\]{}<>「」【】『』+\-*/=,、。;:：!！?？"'`&%$#@^|~・]/g, '_').replace(/_+/g, '_').replace(/^_|_$/g, '');
  if (!n) return '';
  if (/^\d/.test(n) || /^[A-Za-z]{1,3}\d+$/.test(n) || /^[RrCc]\d*$/.test(n) || /^(TRUE|FALSE)$/i.test(n)) n = '_' + n;
  return n.slice(0, 255);
}

/** Labels for inputs that have none: the heading to the left (and the one
 * left of it that spans more rows, as in 受注者 / 住所), else the one above. */
export function nameInputs(boxes) {
  const labels = boxes.filter((b) => !isInput(b) && b.kind !== 'check' && b.kind !== 'stamp' && b.text.trim());
  const rowsMeet = (a, b) => a.r < b.r + b.h && b.r < a.r + a.h;
  const colsMeet = (a, b) => a.c < b.c + b.w && b.c < a.c + a.w;
  const clean = (t) => t.replace(/\s+/g, '').replace(/[:：]$/, '');
  const leftOf = (b) => labels.filter((l) => rowsMeet(l, b) && l.c + l.w <= b.c).sort((x, y) => (y.c + y.w) - (x.c + x.w))[0];
  const used = new Set(boxes.filter((b) => isField(b) && b.name).map((b) => b.name));
  let n = 0;
  for (const b of boxes) {
    if (!isField(b) || b.name) continue;
    let name = '';
    const left = leftOf(b);
    if (left) {
      name = clean(left.text);
      const parent = leftOf(left);
      if (parent && parent.h > left.h) name = clean(parent.text) + '_' + name;
    } else {
      const above = labels.filter((l) => colsMeet(l, b) && l.r + l.h <= b.r).sort((x, y) => (y.r + y.h) - (x.r + x.h))[0];
      if (above && b.r - (above.r + above.h) <= 1) name = clean(above.text);
    }
    name = excelName(name) || `欄${++n}`;
    let unique = name;
    for (let k = 2; used.has(unique); k++) unique = `${name}_${k}`;
    used.add(unique);
    b.name = unique;
  }
  return boxes;
}

const DATE_FORMAT = '[$-411]ggge"年"m"月"d"日";@';
const NUMBER_FORMAT = '#,##0';
const EPOCH = Date.UTC(1899, 11, 30);

function checkText(b, filled) {
  if (!filled || !b.value) return b.text;
  const picked = new Set(String(b.value).split('\n'));
  return b.text.replace(/[□■]\s*([^□■]*?)(?=\s*[□■]|$)/g, (all, opt) => (picked.has(opt.trim()) ? all.replace('□', '■') : all));
}
export const checkOptions = (b) => [...b.text.matchAll(/[□■]\s*([^□■]*?)(?=\s*[□■]|$)/g)].map((m) => m[1].trim()).filter(Boolean);

/**
 * The sheets as a workbook, built on the parsed template (parse_spreadsheet
 * JSON of template.xlsx: an empty portrait and an empty landscape sheet with
 * their page setup). Every sheet is a copy of one of the two, so it keeps
 * that page setup, resized to its grid and filled with its boxes.
 */
export function toWorkbook(wb, sheets, { protect = true, filled = true } = {}) {
  const [portrait, landscape] = wb.sheets;
  const base = wb.default_style;
  const all = sheets.flatMap((sh) => sh.boxes.filter(isField).map((b) => b.name));
  const count = new Map();
  for (const n of all) count.set(n, (count.get(n) || 0) + 1);
  wb.copied_sheet_origins = {};
  wb.defined_names = [];
  wb.defined_name_attributes = {};
  wb.sheets = sheets.map((sheet, i) => {
    const from = sheet.grid.landscape ? landscape : portrait;
    const sh = structuredClone(from);
    sh.name = sheet.name;
    sh.origin_id = 1000 + i;
    wb.copied_sheet_origins[1000 + i] = from.origin_id;
    const { cols, rows, colPx, rowPx } = sheet.grid;
    // 8-pixel digits (Yu Gothic 11): a stored width of px / 8 is px wide.
    sh.col_widths = colPx.map((px) => Math.floor((px / 8) * 256) / 256);
    sh.col_count = cols;
    sh.col_custom_widths = Object.fromEntries(colPx.map((_, c) => [c, true]));
    sh.hidden_cols = colPx.flatMap((px, c) => (px > 0 ? [] : [c]));
    sh.rows = rowPx.map((px, r) => ({ index: r + 1, cells: [], height: px * 0.75, custom_height: true, ...(px > 0 ? {} : { hidden: true }) }));
    sh.merge_cells = [];
    const at = (r, c) => {
      const row = sh.rows[r];
      let cell = row.cells.find((x) => x.col === c);
      if (!cell) { cell = { col: c, value: 'Empty', style: structuredClone(base) }; row.cells.push(cell); }
      return cell;
    };
    nameInputs(sheet.boxes);
    for (const b of sheet.boxes) {
      if (!fits(b, cols, rows)) continue;
      const thick = b.border === 'thick';
      const line = { style: thick ? 'medium' : 'thin', color: '000000' };
      const ruled = b.border === 'box' || thick;
      for (let r = b.r; r < b.r + b.h; r++) for (let c = b.c; c < b.c + b.w; c++) {
        const s = at(r, c).style;
        if (ruled && r === b.r) s.border_top = line;
        if ((ruled || b.border === 'under') && r === b.r + b.h - 1) s.border_bottom = line;
        if (ruled && c === b.c) s.border_left = line;
        if (ruled && c === b.c + b.w - 1) s.border_right = line;
        if (b.fill) s.bg_color = b.fill;
        s.font_size = b.size; s.bold = b.bold;
        s.horizontal_align = b.align === 'distributed' ? 'distributed' : b.align;
        s.vertical_align = b.h > 1 && b.kind === 'text' ? 'top' : 'center';
        if (b.h > 1 || /\n/.test(b.text)) s.wrap_text = true;
        if (b.kind === 'stamp') s.font_color = '808080';
        if (isField(b)) s.unlocked = true;
        if (isInput(b) && b.type === 'date') s.number_format = DATE_FORMAT;
        if (isInput(b) && b.type === 'number') s.number_format = NUMBER_FORMAT;
      }
      const first = at(b.r, b.c);
      let value = null;
      if (isInput(b)) {
        const v = filled ? String(b.value ?? '').trim() : '';
        if (v && b.type === 'number' && !isNaN(Number(v.replace(/,/g, '')))) value = { Number: Number(v.replace(/,/g, '')) };
        else if (v && b.type === 'date' && /^\d{4}-\d{2}-\d{2}$/.test(v)) {
          const [y, m, d] = v.split('-').map(Number);
          value = { Number: (Date.UTC(y, m - 1, d) - EPOCH) / 86400000 };
        } else if (v) value = { String: v };
      } else if (b.text) value = { String: b.kind === 'check' ? checkText(b, filled) : b.text };
      if (value) first.value = value;
      if (b.h > 1 || b.w > 1) sh.merge_cells.push({ start_row: b.r + 1, end_row: b.r + b.h, start_col: b.c, end_col: b.c + b.w - 1 });
      if (isField(b)) {
        // A name used on one sheet only is the book's; one that recurs (a
        // copy of the form on another sheet) belongs to its sheet.
        if (count.get(b.name) > 1) {
          const key = `${quoteSheet(sheet.name)}!${b.name}`;
          wb.defined_names.push([key, absRef(sheet.name, b)]);
          wb.defined_name_attributes[key] = { localSheetId: String(i), name: b.name };
        } else wb.defined_names.push([b.name, absRef(sheet.name, b)]);
      }
    }
    for (const row of sh.rows) row.cells.sort((x, y) => x.col - y.col);
    const area = `${quoteSheet(sheet.name)}!Print_Area`;
    wb.defined_names.push([area, `${quoteSheet(sheet.name)}!$A$1:$${colName(cols - 1)}$${rows}`]);
    wb.defined_name_attributes[area] = { localSheetId: String(i), name: '_xlnm.Print_Area' };
    sh.protected = !!protect;
    return sh;
  });
  return wb;
}

/** Excel's pixels for a stored column width, with the digit width in pixels. */
export const columnPixels = (width, digit) => Math.trunc(((256 * width + Math.trunc(128 / digit)) / 256) * digit);

/** A cell reference or range as 0-based { r, c, h, w }. */
export function parseRef(text) {
  const m = String(text).replace(/\$/g, '').replace(/^.*!/, '').match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
  if (!m) return null;
  const col = (s) => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const r0 = Number(m[2]) - 1, c0 = col(m[1]);
  const r1 = m[3] ? Number(m[4]) - 1 : r0, c1 = m[3] ? col(m[3]) : c0;
  return { r: Math.min(r0, r1), c: Math.min(c0, c1), h: Math.abs(r1 - r0) + 1, w: Math.abs(c1 - c0) + 1 };
}

const textOf = (v) => (v && typeof v === 'object' ? (v.String ?? (v.Number != null ? String(v.Number) : v.Bool != null ? String(v.Bool) : '')) : '');
const valueOf = (v, type) => {
  if (v && typeof v === 'object' && v.Number != null && type === 'date') {
    return new Date(EPOCH + Math.round(v.Number) * 86400000).toISOString().slice(0, 10);
  }
  return textOf(v);
};
const widthOf = (t, size) => [...t].reduce((n, ch) => n + (ch.charCodeAt(0) < 0x2000 ? 0.55 : 1), 0) * size * 4 / 3;

export const MAX_COLS = 256, MAX_ROWS = 1000;

/** The grid of an Excel sheet: its print area, or what it uses, with each
 * column and row at the size Excel draws it. */
export function gridOf(wb, sh, digit = 8) {
  let cols = 0, rows = 0;
  const area = (wb.defined_names || []).find(([n]) => /Print_Area$/.test(n) && n.replace(/^'|'?!Print_Area$/g, '') === sh.name);
  const printed = area && String(area[1]).split(',').map(parseRef).filter(Boolean);
  if (printed?.length) {
    for (const r of printed) { cols = Math.max(cols, r.c + r.w); rows = Math.max(rows, r.r + r.h); }
  } else {
    for (const row of sh.rows) for (const c of row.cells) {
      const s = c.style || {};
      if (c.value !== 'Empty' || s.border_top || s.border_bottom || s.border_left || s.border_right || s.bg_color) {
        cols = Math.max(cols, c.col + 1); rows = Math.max(rows, row.index);
      }
    }
    for (const m of sh.merge_cells || []) { cols = Math.max(cols, m.end_col + 1); rows = Math.max(rows, m.end_row); }
  }
  cols = Math.min(MAX_COLS, Math.max(cols, 1));
  rows = Math.min(MAX_ROWS, Math.max(rows, 1));
  const hiddenCols = new Set(sh.hidden_cols || []);
  const width = (c) => sh.col_widths?.[c] > 0 ? sh.col_widths[c] : sh.default_col_width > 0 ? sh.default_col_width : null;
  const colPx = Array.from({ length: cols }, (_, c) => (hiddenCols.has(c) ? 0 : width(c) != null ? columnPixels(width(c), digit) : Math.trunc(8.43 * digit + 5)));
  const byIndex = new Map(sh.rows.map((r) => [r.index, r]));
  const fallback = sh.default_row_height > 0 ? sh.default_row_height : 15;
  const rowPx = Array.from({ length: rows }, (_, r) => {
    const row = byIndex.get(r + 1);
    return row?.hidden ? 0 : Math.round((row?.height ?? fallback) * 4 / 3);
  });
  const w = colPx.reduce((a, b) => a + b, 0), h = rowPx.reduce((a, b) => a + b, 0);
  return { cols, rows, colPx, rowPx, landscape: w > h * 1.05 && w > PAGE.portrait.w };
}

/** Boxes from any Excel sheet drawn as 方眼紙: merged areas, ruled or filled
 * cells and lone text, with inputs found from defined names, unlocked cells
 * and empty ruled areas. Returns a sheet and how many boxes fell outside. */
export function fromWorkbook(wb, sheetIndex = 0, digit = 8) {
  const sh = wb.sheets[sheetIndex];
  const grid = gridOf(wb, sh, digit);
  const { cols, rows } = grid;
  const cells = new Map();
  for (const row of sh.rows) for (const c of row.cells) cells.set(`${row.index - 1},${c.col}`, c);
  const get = (r, c) => cells.get(`${r},${c}`);
  const names = new Map();
  for (const [key, target] of wb.defined_names || []) {
    if (/Print_|_xlnm|^_Filter/.test(key)) continue;
    // A sheet's own name comes as Sheet!Name.
    const local = key.match(/^'?(.*?)'?!(.+)$/);
    if (local && local[1] !== sh.name) continue;
    const name = local ? local[2] : key;
    const sheet = String(target).match(/^'?(.*?)'?!/);
    if (sheet && sheet[1].replace(/''/g, "'") !== sh.name) continue;
    const r = parseRef(target);
    if (r) names.set(`${r.r},${r.c},${r.h},${r.w}`, name);
  }
  // A book that names its fields (or protects all but its inputs) says which
  // cells are inputs; only a book that does neither is guessed at.
  const declared = names.size > 0 || !!sh.protected;
  const taken = new Set();
  const boxes = [];
  let clipped = 0;
  const add = (area) => {
    if (!fits(area, cols, rows)) { clipped++; return; }
    const first = get(area.r, area.c);
    const s = first?.style || {};
    const edge = (side, list) => list.some(([r, c]) => get(r, c)?.style?.['border_' + side]);
    const span = (f) => Array.from({ length: f.n }, (_, i) => f.at(i));
    const top = edge('top', span({ n: area.w, at: (i) => [area.r, area.c + i] }));
    const bottom = edge('bottom', span({ n: area.w, at: (i) => [area.r + area.h - 1, area.c + i] }));
    const left = edge('left', span({ n: area.h, at: (i) => [area.r + i, area.c] }));
    const right = edge('right', span({ n: area.h, at: (i) => [area.r + i, area.c + area.w - 1] }));
    const thick = ['medium', 'thick', 'double'].includes(s.border_top?.style);
    const border = top && bottom && left && right ? (thick ? 'thick' : 'box') : bottom && !top ? 'under' : top || left || right ? 'box' : 'none';
    const text = textOf(first?.value);
    const name = names.get(`${area.r},${area.c},${area.h},${area.w}`) || '';
    const fmt = s.number_format || '';
    const type = /g|yy|m"月"|d"日"/.test(fmt) && !/#/.test(fmt) ? 'date' : /[#0]/.test(fmt) && !/@$/.test(fmt) ? 'number' : 'text';
    const template = /^(令和|平成)?[\s　]*年[\s　]*月[\s　]*日$|^[\s　]*年[\s　]*月[\s　]*日$/.test(text);
    let kind;
    const blank = !text || template;
    const looksInput = declared ? (name || (s.unlocked && blank && !names.size)) : blank && (border !== 'none' || s.unlocked || s.bg_color);
    if (/[□■]/.test(text)) kind = 'check';
    else if (looksInput) kind = border === 'under' ? 'line' : 'input';
    else if (/^[\s　]*(印|㊞)?[\s　]*$/.test(text) && border !== 'none') kind = 'stamp';
    else if (border !== 'none' || s.bg_color) kind = 'label';
    else kind = (s.font_size || 11) >= 14 ? 'heading' : 'text';
    const box = makeBox(kind, area.r, area.c, area.h, area.w, {
      text: kind === 'input' || kind === 'line' ? '' : kind === 'check' ? text.replace(/■/g, '□') : text, border, fill: s.bg_color && !/^F{6}$/i.test(s.bg_color) ? s.bg_color : null,
      align: ['center', 'right', 'distributed', 'left'].includes(s.horizontal_align) ? s.horizontal_align : kind === 'label' ? 'center' : 'left',
      size: s.font_size || 11, bold: !!s.bold, name: isField({ kind }) ? name : '', type: isInput({ kind }) ? (template ? 'date' : type) : 'text',
      value: kind === 'check' ? [...text.matchAll(/■\s*([^□■]*?)(?=\s*[□■]|$)/g)].map((m) => m[1].trim()).join('\n')
        : isInput({ kind }) && !template ? valueOf(first?.value, type) : '',
    });
    boxes.push(box);
    for (let r = area.r; r < area.r + area.h; r++) for (let c = area.c; c < area.c + area.w; c++) taken.add(`${r},${c}`);
  };
  for (const m of sh.merge_cells || []) {
    add({ r: m.start_row - 1, c: m.start_col, h: m.end_row - m.start_row + 1, w: m.end_col - m.start_col + 1 });
  }
  for (const [key, cell] of cells) {
    if (taken.has(key)) continue;
    const [r, c] = key.split(',').map(Number);
    const s = cell.style || {};
    const text = textOf(cell.value);
    const named = [...names.keys()].some((k) => k.startsWith(`${r},${c},1,1`));
    const ruled = s.border_top && s.border_bottom && s.border_left && s.border_right;
    if (text) {
      // Lone text spills right over empty cells, as Excel draws it.
      const need = widthOf(text, s.font_size || 11);
      let k = 1, have = grid.colPx[c] || 0;
      while (!(ruled || s.bg_color) && have < need && c + k < cols && !taken.has(`${r},${c + k}`) && !textOf(get(r, c + k)?.value) && !get(r, c + k)?.style?.border_left) have += grid.colPx[c + k++];
      add({ r, c, h: 1, w: k });
    } else if (named || s.unlocked || (ruled && s.bg_color)) {
      add({ r, c, h: 1, w: 1 });
    }
  }
  boxes.sort((a, b) => a.r - b.r || a.c - b.c);
  nameInputs(boxes);
  return { sheet: { name: sh.name, grid, boxes, blocks: [] }, clipped };
}

/** The form as text for people and for AI: headings, inputs with their cells. */
export function toMarkdown(sheets, title) {
  return [`# ${title}`, '', ...sheets.map((sh) => sheetMarkdown(sh.boxes, sh.name))].join('\n');
}
function sheetMarkdown(boxes, name) {
  const sorted = [...boxes].sort((a, b) => a.r - b.r || a.c - b.c);
  const lines = [`## ${name}`, ''];
  lines.push('| 欄 | 種類 | セル | 記入値 |', '| --- | --- | --- | --- |');
  const kindName = (b) => (b.kind === 'check' ? `選択（${checkOptions(b).join(' / ')}）` : b.type === 'date' ? '日付' : b.type === 'number' ? '数値' : '文字');
  for (const b of sorted) {
    if (isField(b)) {
      lines.push(`| ${b.name} | ${kindName(b)} | ${ref(b)} | ${String(b.value || '').replace(/\n/g, ' / ')} |`);
    }
  }
  lines.push('', '### 様式の文字（上から順に）', '');
  for (const b of sorted) if (!isInput(b) && b.text.trim()) lines.push(`- ${ref(b)}: ${b.text.replace(/\n/g, ' ')}`);
  return lines.join('\n') + '\n';
}

export function titleOf(boxes) {
  const t = [...boxes].filter((b) => !isInput(b) && b.text.trim()).sort((a, b) => b.size - a.size || a.r - b.r)[0];
  return t ? t.text.replace(/\s+/g, '') : '方眼紙';
}

/** The rows of each printed page: [{ start, end }]. */
export function pageBands(grid) {
  const { starts } = pages(grid);
  return starts.map((start, i) => ({ start, end: starts[i + 1] ?? grid.rows }));
}

const straddles = (boxes, row) => boxes.some((b) => b.r < row && b.r + b.h > row);

/** Swap page k with the page before (dir -1) or after (+1), rows and all. */
export function movePage(sheet, k, dir) {
  const list = pageBands(sheet.grid);
  const j = k + dir;
  if (j < 0 || j >= list.length) return { error: dir < 0 ? '最初のページです' : '最後のページです' };
  const [a, b] = j < k ? [list[j], list[k]] : [list[k], list[j]];
  if ([a.start, b.start, b.end].some((row) => straddles(sheet.boxes, row))) return { error: 'ページの境目にまたがる枠があるので入れ替えられません' };
  const ha = a.end - a.start, hb = b.end - b.start;
  const shift = (r) => (r >= a.start && r < a.end ? hb : r >= b.start && r < b.end ? -ha : 0);
  const rowPx = sheet.grid.rowPx;
  return {
    ...sheet,
    grid: { ...sheet.grid, rowPx: [...rowPx.slice(0, a.start), ...rowPx.slice(b.start, b.end), ...rowPx.slice(a.start, a.end), ...rowPx.slice(b.end)] },
    boxes: sheet.boxes.map((x) => ({ ...x, r: x.r + shift(x.r) })),
    blocks: sheet.blocks.map((x) => ({ ...x, r: x.r + shift(x.r) })),
  };
}

/** Add a page's worth of rows at the bottom. */
export function addPage(sheet) {
  const { scale } = pages(sheet.grid);
  const tall = PAGE[sheet.grid.landscape ? 'landscape' : 'portrait'].h / scale;
  const px = sheet.grid.rowPx[sheet.grid.rowPx.length - 1] || CELL;
  const more = Math.max(1, Math.floor(tall / px));
  if (sheet.grid.rows + more > MAX_ROWS) return { error: `${MAX_ROWS} 行までです` };
  return { ...sheet, grid: resizeGrid(sheet.grid, sheet.grid.cols, sheet.grid.rows + more) };
}

/** Remove page k with what is on it, closing the gap. */
export function removePage(sheet, k) {
  const list = pageBands(sheet.grid);
  if (list.length < 2) return { error: '1ページしかありません' };
  const p = list[k];
  if ([p.start, p.end].some((row) => straddles(sheet.boxes, row))) return { error: 'ページの境目にまたがる枠があるので消せません' };
  const h = p.end - p.start;
  const keep = (x) => !(x.r >= p.start && x.r < p.end);
  const up = (x) => (x.r >= p.end ? { ...x, r: x.r - h } : x);
  const rowPx = [...sheet.grid.rowPx.slice(0, p.start), ...sheet.grid.rowPx.slice(p.end)];
  return {
    ...sheet,
    grid: { ...sheet.grid, rows: rowPx.length, rowPx },
    boxes: sheet.boxes.filter(keep).map(up),
    blocks: sheet.blocks.filter(keep).map(up),
  };
}
