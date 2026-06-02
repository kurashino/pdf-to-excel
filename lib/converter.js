const XLSX = require('xlsx');
const { TEMPLATE_B64 } = require('./template');

function parseEstimateRows(text) {
  const allRows = [];
  const pages = text.split(/\f/);

  for (const page of pages) {
    const lines = page.split(/\n|\r/).map(l => l.trim()).filter(Boolean);
    const pageRows = [];

    for (const line of lines) {
      const tokens = line.split(/\s+/);
      if (tokens.length < 4) continue;

      const noVal = parseInt(tokens[0]);
      if (isNaN(noVal) || noVal < 1 || noVal > 100) continue;
      if (['並順','場所','箇所','工事項目','NO','No'].includes(tokens[1])) continue;

      const unitPattern = /^(式|㎡|ヶ|台|本|枚|ｹ|ケ|m2|m²)$/;
      const numPattern = /^[\d,.]+$/;
      const rest = tokens.slice(1);

      // 単位の位置を探す
      let uIdx = -1;
      for (let i = 0; i < rest.length; i++) {
        if (unitPattern.test(rest[i])) { uIdx = i; break; }
      }

      let basho = '', kasho = '', koji = '';
      let suryo = null, tani = '', tanka = null;

      if (uIdx >= 0) {
        tani = rest[uIdx];

        // 単位の直前が数字なら数量
        if (uIdx > 0 && numPattern.test(rest[uIdx - 1])) {
          suryo = parseFloat(rest[uIdx - 1].replace(/,/g, ''));
          const before = rest.slice(0, uIdx - 1);
          basho = before[0] || '';
          kasho = before[1] || '';
          koji  = before.slice(2).join('');
        } else {
          const before = rest.slice(0, uIdx);
          basho = before[0] || '';
          kasho = before[1] || '';
          koji  = before.slice(2).join('');
        }

        // 単位の後の数字が単価
        const afterNums = rest.slice(uIdx + 1).filter(t => numPattern.test(t));
        if (afterNums.length > 0) tanka = parseFloat(afterNums[0].replace(/,/g, ''));

      } else {
        // 単位なし：場所・箇所・工事項目だけ
        basho = rest[0] || '';
        kasho = rest[1] || '';
        koji  = rest.slice(2).join('');
      }

      if (!basho || basho.length < 1) continue;

      pageRows.push({ no: noVal, basho, kasho, koji, suryo, tani, tanka });
    }

    if (pageRows.length >= 3) allRows.push(...pageRows);
  }

  const seen = new Set();
  return allRows.filter(r => {
    if (seen.has(r.no)) return false;
    seen.add(r.no);
    return true;
  }).sort((a, b) => a.no - b.no);
}

function buildExcel(rows) {
  const buf = Buffer.from(TEMPLATE_B64, 'base64');
  const wb = XLSX.read(buf, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];

  // ヘッダー行（7行目）を読んで列番号を自動判定
  const headers = {};
  const headerRow = 7;
  for (let c = 1; c <= 20; c++) {
    const addr = XLSX.utils.encode_cell({ r: headerRow - 1, c: c - 1 });
    const cell = ws[addr];
    if (cell && cell.v) {
      const v = cell.v.toString().trim();
      headers[v] = XLSX.utils.encode_col(c - 1);
    }
  }

  const colOf = (name) => headers[name] || null;

  const set = (col, row, val) => {
    if (!col || val === null || val === undefined || val === '') return;
    const addr = col + row;
    if (!ws[addr]) ws[addr] = {};
    ws[addr].v = val;
    ws[addr].t = typeof val === 'number' ? 'n' : 's';
  };

  for (const r of rows) {
    const row = r.no + 7;
    set(colOf('場所'),     row, r.basho);
    set(colOf('箇所'),     row, r.kasho);
    set(colOf('工事項目'), row, r.koji);
    if (r.suryo !== null) set(colOf('数量'), row, r.suryo);
    if (r.tani)           set(colOf('単位'), row, r.tani);
    if (r.tanka !== null) set(colOf('単価'), row, r.tanka);
  }

  return XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });
}

module.exports = { parseEstimateRows, buildExcel };