// SPDX-License-Identifier: MIT OR Apache-2.0
// node --test test/  (no engine needed: the model alone)
import test from 'node:test';
import assert from 'node:assert/strict';
import { newGrid, pages, pageBands, movePage, addPage, removePage, makeBox, nameInputs, excelName, columnPixels, parseRef, sheetName } from '../model.js';
import { insertBlock, bottomRow } from '../blocks.js';
import { sampleForm } from '../sample.js';

test('an A4 grid is one page at full size; a second page of rows follows on', () => {
  const g = newGrid();
  assert.deepEqual(pages(g), { scale: 1, starts: [0] });
  const s = addPage({ name: 's', grid: g, boxes: [], blocks: [] });
  assert.deepEqual(pageBands(s.grid), [{ start: 0, end: 50 }, { start: 50, end: 100 }]);
});

test('pages swap with what is on them, and a box across the break stops it', () => {
  let s = addPage(sampleForm());
  s.boxes.push(makeBox('text', 60, 0, 1, 4, { text: 'p2' }));
  const moved = movePage(s, 1, -1);
  assert.equal(moved.boxes.find((b) => b.text === 'p2').r, 10);
  assert.equal(moved.boxes.find((b) => b.text === '補助金交付申請書').r, 52);
  s.boxes.push(makeBox('text', 49, 10, 2, 4, { text: 'across' }));
  assert.ok(movePage(s, 1, -1).error);
  assert.equal(removePage(moved, 0).grid.rows, 50);
});

test('inputs are named from the heading to their left, with its parent', () => {
  const boxes = [
    makeBox('label', 0, 0, 3, 4, { text: '申請者' }), makeBox('label', 0, 4, 1, 6, { text: '氏 名' }),
    makeBox('input', 0, 10, 1, 10), makeBox('label', 4, 0, 1, 10, { text: '電話番号' }), makeBox('input', 5, 0, 1, 10),
  ];
  nameInputs(boxes);
  assert.equal(boxes[2].name, '申請者_氏名');
  assert.equal(boxes[4].name, '電話番号');
  assert.equal(excelName('A1'), '_A1');
  assert.equal(excelName('1回目'), '_1回目');
});

test('blocks stack and refuse to run off the grid', () => {
  const grid = newGrid();
  let boxes = [], blocks = [];
  for (const name of ['title', 'applicant', ...Array(8).fill('long')]) {
    const r = insertBlock(boxes, blocks, name, bottomRow(boxes, blocks), grid);
    if (r.error) { assert.match(r.error, /広げて/); return; }
    boxes = r.boxes; blocks = r.blocks;
  }
  assert.fail('expected the grid to fill');
});

test('Excel widths, references and sheet names', () => {
  assert.equal(columnPixels(2.5, 8), 20);
  assert.equal(columnPixels(9.140625, 7), 64);
  assert.deepEqual(parseRef("'様式'!$B$4:$I$5"), { r: 3, c: 1, h: 2, w: 8 });
  assert.equal(sheetName('a/b', []), 'a_b');
  assert.equal(sheetName('様式', ['様式']), '様式(2)');
});
