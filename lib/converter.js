const XLSX = require('xlsx');
const { TEMPLATE_B64 } = require('./template');

/**
 * PDFテキストから見積行を抽出
 */
function parseEstimateRows(text) {
  const rows = [];
  const lines = text.split(/\n|\r/).map(l => l.trim()).filter(Boolean);

  // 単位パターン
  const UNIT_RE = /^(式|㎡|ヶ|台|本|枚|ｹ|ケ|m2|m²|組|回|ヵ所|箇所)$/;
  const NUM_RE = /^[\d,.]+$/;

  for (const line of lines) {
    const tokens = line.split(/\s+/);
    if (tokens.length < 3) continue;

    const no = parseInt(tokens[0]);
    if (isNaN(no) || no < 1 || no > 100) continue;

    // ヘッダー行をスキップ
    if (['並順', '場所', 'NO', 'No'].includes(tokens[1])) continue;

    const basho = tokens[1] || '';
    const kasho = tokens[2] || '';

    // 残りトークンから工事項目・数量・単位・単価を解析
    const rest = tokens.slice(3);
    let koji = '';
    let suryo = null;
    let tani = '';
    let tanka = null;

    // 単位の位置を探す
    let uIdx = -1;
    for (let i = 0; i < rest.length; i++) {
      if (UNIT_RE.test(rest[i])) { uIdx = i; break; }
    }

    if (uIdx >= 0) {
      tani = rest[uIdx];
      // 単位の直前が数字なら数量
      if (uIdx > 0 && NUM_RE.test(rest[uIdx - 1])) {
        suryo = parseFloat(rest[uIdx - 1].replace(/,/g, ''));
        koji = rest.slice(0, uIdx - 1).join('');
      } else {
        koji = rest.slice(0, uIdx).join('');
      }
      // 単位の後ろの数字群から単価を取得
      const afterNums = rest.slice(uIdx + 1).filter(t => NUM_RE.test(t));
      if (afterNums.length > 0) {
        tanka = parseFloat(afterNums[0].replace(/,/g, ''));
      }
    } else {
      koji = rest.join('');
    }

    if (!basho || basho.length < 1) continue;
    if (!koji || koji.length < 2) continue;

    rows.push({ no, basho, kasho, koji, suryo, tani, tanka });
  }

  return rows.sort((a, b) => a.no - b.no);
}

/**
 * 抽出データをテンプレートに書き込みバッファを返す
 */
function buildExcel(rows) {
  const buf = Buffer.from(TEMPLATE_B64, 'base64');
  const wb = XLSX.read(buf, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];

  const set = (ws, col, row, val) => {
    const addr = col + row;
    if (val === null || val === undefined || val === '') return;
    if (!ws[addr]) ws[addr] = {};
    ws[addr].v = val;
    ws[addr].t = typeof val === 'number' ? 'n' : 's';
  };

  for (const r of rows) {
    const row = r.no + 7; // 並順1 → 行8
    set(ws, 'B', row, r.basho);
    set(ws, 'C', row, r.kasho);
    set(ws, 'D', row, r.koji);
    if (r.suryo !== null) set(ws, 'M', row, r.suryo);
    if (r.tani) set(ws, 'N', row, r.tani);
    if (r.tanka !== null) set(ws, 'O', row, r.tanka);
  }

  return XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });
}

module.exports = { parseEstimateRows, buildExcel };
