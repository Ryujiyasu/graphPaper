// SPDX-License-Identifier: MIT OR Apache-2.0
// Ready-made blocks: the parts most forms are built from, each a band of
// whole rows across the page. Stacked, they make a form; inside, they are
// ordinary boxes that can be edited on the grid like any other.
import { makeBox, KINDS } from './model.js';

const L = 10; // the width of a heading column

/** The palette's tabs, by what the blocks are for. */
export const GROUPS = ['表題', '申請者', '入力欄', '表', '文章', '事務', 'マイ'];
export const MINE = 'マイ';

/** name → { title, group, rows, make(r, W) → boxes }, W the grid's width in cells */
export const BLOCKS = {
  title: { title: '表題', group: '表題', rows: 3, make: (r, W) => [
    makeBox('heading', r, 0, 2, W, { text: '○○申請書' }),
  ] },
  number: { title: '様式番号・日付', group: '表題', rows: 2, make: (r, W) => [
    makeBox('text', r, 0, 1, 8, { text: '様式第1号', size: 9 }),
    makeBox('text', r, W - 13, 1, 3, { text: '日付', align: 'right' }),
    makeBox('line', r, W - 10, 1, 10, { type: 'date', align: 'right' }),
  ] },
  address: { title: '宛名', group: '表題', rows: 2, make: (r, W) => [
    makeBox('text', r, 0, 1, 14, { text: '○○長　様' }),
  ] },
  applicant: { title: '申請者', group: '申請者', rows: 8, make: (r, W) => [
    makeBox('label', r, 0, 7, 4, { text: '申請者' }),
    makeBox('label', r, 4, 2, 6, { text: '住所' }),
    makeBox('input', r, 10, 2, W - 10),
    makeBox('label', r + 2, 4, 1, 6, { text: 'フリガナ', size: 9 }),
    makeBox('input', r + 2, 10, 1, W - 10, { size: 9 }),
    makeBox('label', r + 3, 4, 2, 6, { text: '氏名' }),
    makeBox('input', r + 3, 10, 2, W - 14),
    makeBox('stamp', r + 3, W - 4, 2, 4),
    makeBox('label', r + 5, 4, 2, 6, { text: '電話番号' }),
    makeBox('input', r + 5, 10, 2, W - 10),
  ] },
  lead: { title: '本文', group: '文章', rows: 2, make: (r, W) => [
    makeBox('text', r, 0, 1, W, { text: '下記のとおり申請します。' }),
  ] },
  ki: { title: '記', group: '文章', rows: 2, make: (r, W) => [
    makeBox('text', r, 0, 1, W, { text: '記', align: 'center' }),
  ] },
  field: { title: '項目と入力欄', group: '入力欄', rows: 2, stack: true, make: (r, W) => [
    makeBox('label', r, 0, 2, L, { text: '項目' }),
    makeBox('input', r, L, 2, W - L),
  ] },
  date: { title: '日付', group: '入力欄', rows: 2, stack: true, make: (r, W) => [
    makeBox('label', r, 0, 2, L, { text: '日付' }),
    makeBox('input', r, L, 2, 12, { type: 'date', align: 'center' }),
    makeBox('label', r, L + 12, 2, W - L - 12, { text: '', fill: null, border: 'box' }),
  ] },
  money: { title: '金額', group: '入力欄', rows: 2, stack: true, make: (r, W) => [
    makeBox('label', r, 0, 2, L, { text: '金額' }),
    makeBox('input', r, L, 2, W - L - 2, { type: 'number', align: 'right' }),
    makeBox('label', r, W - 2, 2, 2, { text: '円', fill: null }),
  ] },
  period: { title: '期間', group: '入力欄', rows: 2, stack: true, make: (r, W) => [
    makeBox('label', r, 0, 2, L, { text: '期間' }),
    makeBox('input', r, L, 2, 11, { type: 'date', align: 'center', name: '開始日' }),
    makeBox('label', r, L + 11, 2, 2, { text: '〜', fill: null }),
    makeBox('input', r, L + 13, 2, W - L - 13, { type: 'date', align: 'center', name: '終了日' }),
  ] },
  choice: { title: '選択肢', group: '入力欄', rows: 2, stack: true, make: (r, W) => [
    makeBox('label', r, 0, 2, L, { text: '区分' }),
    makeBox('check', r, L, 2, W - L, { text: '□ はい　□ いいえ', border: 'box' }),
  ] },
  long: { title: '複数行', group: '入力欄', rows: 6, stack: true, make: (r, W) => [
    makeBox('label', r, 0, 6, L, { text: '内容' }),
    makeBox('input', r, L, 6, W - L),
  ] },
  items: { title: '明細表', group: '表', rows: 9, make: (r, W) => {
    const q = Math.round(W * 7 / 34), n = Math.round(W * 4 / 34);
    const cols = [['品名', 0, W - n - 2 * q], ['数量', W - n - 2 * q, n], ['単価', W - 2 * q, q], ['金額', W - q, q]];
    const out = cols.map(([t, c, w]) => makeBox('label', r, c, 1, w, { text: t }));
    for (let i = 1; i <= 6; i++) {
      for (const [t, c, w] of cols) {
        out.push(makeBox('input', r + i, c, 1, w, { name: `明細${i}_${t}`, type: t === '品名' ? 'text' : 'number', align: t === '品名' ? 'left' : 'right' }));
      }
    }
    out.push(makeBox('label', r + 7, W - 2 * q, 1, q, { text: '合計' }));
    out.push(makeBox('input', r + 7, W - q, 1, q, { name: '合計', type: 'number', align: 'right' }));
    return out;
  } },
  approval: { title: '承認印欄', group: '事務', rows: 6, make: (r, W) => {
    const out = [];
    ['担当', '係長', '課長', '部長'].forEach((t, i) => {
      const c = W - 20 + i * 5;
      out.push(makeBox('label', r, c, 1, 5, { text: t, size: 9 }));
      out.push(makeBox('stamp', r + 1, c, 4, 5, { text: '' }));
    });
    return out;
  } },
  office: { title: '受付欄', group: '事務', rows: 6, make: (r, W) => [
    makeBox('text', r, 0, 1, W, { text: '※ 以下は記入しないでください', size: 9 }),
    makeBox('label', r + 1, 0, 1, 8, { text: '受付日', size: 9 }),
    makeBox('label', r + 1, 8, 1, 8, { text: '受付番号', size: 9 }),
    makeBox('input', r + 2, 0, 3, 8, { type: 'date', align: 'center' }),
    makeBox('input', r + 2, 8, 3, 8, { align: 'center' }),
  ] },
  note: { title: '注意書き', group: '文章', rows: 3, make: (r, W) => [
    makeBox('text', r, 0, 2, W, { text: '（注）\n1 ○○は、○○を添付してください。', size: 9 }),
  ] },
  space: { title: '余白', group: '文章', rows: 1, make: () => [] },
};

/** Rows a block's band covers: its boxes' rows, or the band it was made with. */
export function bands(boxes, blocks) {
  return blocks.map((k) => {
    const own = boxes.filter((b) => b.block === k.id);
    const top = own.length ? Math.min(...own.map((b) => b.r)) : k.r;
    const bottom = own.length ? Math.max(...own.map((b) => b.r + b.h)) : k.r + k.rows;
    return { ...k, r: Math.min(k.r, top), end: Math.max(k.r + k.rows, bottom) };
  }).sort((a, b) => a.r - b.r);
}

/** The first row below everything on the page. */
export const bottomRow = (boxes, blocks) => Math.max(0, ...boxes.map((b) => b.r + b.h), ...bands(boxes, blocks).map((k) => k.end));

/**
 * Insert a block at row `at` (after `pad` blank rows), moving everything at
 * or below that row down by the block's height and the padding.
 * Returns { boxes, blocks, id } or { error }.
 */
export function insertBlock(boxes, blocks, name, at, grid, pad = 0) {
  const def = BLOCKS[name];
  const h = def.rows + pad;
  const room = grid.rows - bottomRow(boxes, blocks);
  if (h > room) return { error: `入りきりません（あと ${room} 行、このブロックは ${h} 行）。方眼紙を下に広げてください。` };
  const need = def.cols || 24;
  if (grid.cols < need) return { error: `このブロックを置くには、方眼紙の幅が ${need} マス以上いります` };
  const cut = boxes.filter((b) => b.r < at && b.r + b.h > at);
  if (cut.length) return { error: 'ここには挟めません。枠の途中です。' };
  const id = `k${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
  const moved = boxes.map((b) => (b.r >= at ? { ...b, r: b.r + h } : b));
  const movedBlocks = blocks.map((k) => (k.r >= at ? { ...k, r: k.r + h } : k));
  const made = def.make(at + pad, grid.cols).map((b) => ({ ...b, block: id }));
  return { boxes: [...moved, ...made], blocks: [...movedBlocks, { id, name, title: def.title, r: at + pad, rows: def.rows }], id };
}

/** Swap a block with the one above (dir -1) or below (+1). */
export function moveBlock(boxes, blocks, id, dir) {
  const list = bands(boxes, blocks);
  const i = list.findIndex((k) => k.id === id);
  const j = i + dir;
  if (i < 0 || j < 0 || j >= list.length) return { error: dir < 0 ? 'いちばん上です' : 'いちばん下です' };
  const [a, b] = dir < 0 ? [list[j], list[i]] : [list[i], list[j]];
  // Only the two bands and what lies between them move; a loose box there is in the way.
  const inside = (x, k) => x.r >= k.r && x.r + x.h <= k.end;
  const loose = boxes.filter((x) => x.r < b.end && x.r + x.h > a.r && !inside(x, a) && !inside(x, b));
  if (loose.length) return { error: 'ブロックの外の枠が間にあるので動かせません' };
  const ha = a.end - a.r, hb = b.end - b.r, gap = b.r - a.end;
  const shift = (x) => (x.block === a.id || (inside(x, a) && !x.block) ? hb + gap : x.block === b.id || (inside(x, b) && !x.block) ? -(ha + gap) : 0);
  return {
    boxes: boxes.map((x) => ({ ...x, r: x.r + shift(x) })),
    blocks: blocks.map((k) => (k.id === a.id ? { ...k, r: k.r + hb + gap } : k.id === b.id ? { ...k, r: k.r - ha - gap } : k)),
  };
}

/** Remove a block and close the gap it leaves. */
export function removeBlock(boxes, blocks, id) {
  const k = bands(boxes, blocks).find((x) => x.id === id);
  if (!k) return { boxes, blocks };
  const h = k.end - k.r;
  const rest = boxes.filter((b) => b.block !== id && !(b.r >= k.r && b.r + b.h <= k.end));
  const blocking = rest.some((b) => b.r < k.end && b.r + b.h > k.r);
  return {
    boxes: blocking ? rest : rest.map((b) => (b.r >= k.end ? { ...b, r: b.r - h } : b)),
    blocks: blocks.filter((x) => x.id !== id).map((x) => (!blocking && x.r >= k.end ? { ...x, r: x.r - h } : x)),
  };
}

// ---- BlockMaker: blocks people make from their own boxes ----------------------

const KEEP = ['kind', 'r', 'c', 'h', 'w', 'text', 'border', 'fill', 'align', 'size', 'bold', 'type', 'name'];

/**
 * A block made from the boxes inside rows [top, bottom) and columns
 * [left, right): their places relative to the area, without what was written
 * in them. Returns { def } or { error }.
 */
export function captureBlock(boxes, area, title) {
  const inside = boxes.filter((b) => b.r >= area.r && b.r + b.h <= area.r + area.h && b.c >= area.c && b.c + b.w <= area.c + area.w);
  if (!inside.length) return { error: '囲んだ中に枠がありません' };
  const left = Math.min(...inside.map((b) => b.c));
  const right = Math.max(...inside.map((b) => b.c + b.w));
  const parts = inside.map((b) => Object.fromEntries(KEEP.map((k) => [k, k === 'r' ? b.r - area.r : k === 'c' ? b.c - left : b[k]])));
  const partial = boxes.filter((b) => !inside.includes(b) && b.r < area.r + area.h && b.r + b.h > area.r && b.c < area.c + area.w && b.c + b.w > area.c).length;
  return { def: { title: String(title).trim().slice(0, 30) || 'マイブロック', rows: area.h, cols: right, boxes: parts, made: new Date().toISOString().slice(0, 10) }, partial };
}

/** Make a saved block usable: an entry in BLOCKS under the マイ tab. */
export function registerMine(key, def) {
  BLOCKS[key] = {
    title: def.title, group: MINE, rows: def.rows, cols: def.cols, mine: true,
    make: (r) => def.boxes.map((p) => makeBox(p.kind, r + p.r, p.c, p.h, p.w, { ...p, r: r + p.r, value: '' })),
  };
}
export function unregisterMine(key) { delete BLOCKS[key]; }

/** Check a file of shared blocks and give back the usable ones. */
export function readShared(json) {
  const data = typeof json === 'string' ? JSON.parse(json) : json;
  const list = Array.isArray(data?.blocks) ? data.blocks : [];
  return list.filter((d) => d && typeof d.title === 'string' && Number.isInteger(d.rows) && d.rows > 0 && d.rows <= 200
    && Array.isArray(d.boxes) && d.boxes.every((b) => b && Number.isInteger(b.r) && Number.isInteger(b.c) && b.h > 0 && b.w > 0 && Object.hasOwn(KINDS, b.kind)))
    .map((d) => ({ ...d, cols: Math.max(1, ...d.boxes.map((b) => b.c + b.w)), boxes: d.boxes.map((b) => ({ ...b, text: String(b.text ?? '').slice(0, 2000) })) }));
}
