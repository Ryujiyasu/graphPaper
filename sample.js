// SPDX-License-Identifier: MIT OR Apache-2.0
// The sample form: an application stacked from blocks, then filled in.
import { insertBlock, bottomRow } from './blocks.js';
import { isInput, newGrid } from './model.js';

export function sampleForm() {
  const grid = newGrid();
  let boxes = [], blocks = [];
  const add = (name, edit) => {
    const res = insertBlock(boxes, blocks, name, bottomRow(boxes, blocks), grid);
    boxes = res.boxes; blocks = res.blocks;
    const own = boxes.filter((b) => b.block === res.id);
    const labels = own.filter((b) => !isInput(b));
    const inputs = own.filter(isInput);
    edit?.({ own, labels, inputs });
  };
  add('number', ({ labels, inputs }) => { labels[1].text = '申請日'; inputs[0].value = '2026-12-01'; });
  add('title', ({ labels }) => { labels[0].text = '補助金交付申請書'; });
  add('address', ({ labels }) => { labels[0].text = '○○市長　様'; });
  add('applicant', ({ inputs }) => {
    ['○○市北区○○町1-2-3', 'ミホン　タロウ', '見本　太郎', '000-0000-0000'].forEach((v, i) => { inputs[i].value = v; });
  });
  add('lead', ({ labels }) => { labels[0].text = '下記のとおり補助金の交付を申請します。'; });
  add('ki');
  add('choice', ({ own }) => { own[0].text = '申請区分'; own[1].text = '□ 新規　□ 継続　□ 変更'; own[1].value = '新規'; });
  add('field', ({ labels, inputs }) => { labels[0].text = '事業の名称'; inputs[0].value = '商店街にぎわい創出事業'; });
  add('money', ({ labels, inputs }) => { labels[0].text = '申請額'; inputs[0].value = '500000'; });
  add('period', ({ labels, inputs }) => {
    labels[0].text = '事業期間';
    inputs[0].name = '事業開始日'; inputs[0].value = '2027-04-01';
    inputs[1].name = '事業終了日'; inputs[1].value = '2028-03-31';
  });
  add('long', ({ labels, inputs }) => {
    labels[0].text = '事業の内容';
    inputs[0].value = '空き店舗を活用した週末マルシェを月2回開催し、\n来街者の増加と出店者の育成を図る。';
  });
  add('choice', ({ own }) => { own[0].text = '添付書類'; own[1].text = '□ 事業計画書　□ 収支予算書　□ 見積書'; own[1].value = '事業計画書\n収支予算書'; });
  add('space');
  add('office');
  return { name: '申請書', grid, boxes, blocks };
}
