const XLSX = require('xlsx');
const { TEMPLATE_B64 } = require('./template');

function parseEstimateRows(text) {
  const rows = [];
  const pages = text.split(/\f/);

  for (const page of pages) {
    const lines = page.split(/\n|\r/).map(l => l.trim()).filter(Boolean);
    const pageRows = [];

    for (const line of lines) {
      const m = line.match(/^(\d+)\s+(.+?)\s+(.+?)\s+(.+?)\s+([\d.]+)?\s*(式|㎡|ヶ|台|本|枚|ｹ|ケ)?\s*([\d,]+)?/);
      if (!m) continue;
      const no = parseInt(m[1]);
      if (no < 1 || no > 100) continue;

      const tokens = line.split(/\s+/);
      if (tokens.length < 4) continue;

      let idx = 0;
      const noVal = parseInt(tokens[idx]); if (isNaN(noVal)) continue; idx++;
      const basho = tokens[idx]; idx++;
      const kasho = tokens[idx]; idx++;

      let koji = '', suryo = null, tani = '', tanka = null;
      const unitPattern = /^(式|㎡|ヶ|台|本|枚|ｹ|ケ|m2|m²)$/;
      const numPattern = /^[\d,.]+$/;
      const rest = tokens.slice(idx);

      let uIdx = -1;
      for (let i = 0; i < rest.length; i++) {
        if (unitPattern.test(rest[i])) { uIdx = i; break; }
      }

      if (uIdx >= 0) {
        tani = rest[uIdx];
        if (uIdx > 0 && numPattern.test(rest[uIdx - 1])) {
          suryo = parseFloat(rest[uIdx - 1].replace(',', ''));
          koji = rest.slice(0, uIdx - 1).join('');
        } else {
          koji = rest.slice(0, uIdx).join('');
        }
        const afterUnit = rest.slice(uIdx + 1).filter(t => numPattern.test(t));
        if (afterUnit.length > 0) tanka = parseFloat(afterUnit[0].replace(',', ''));
      } else {
        koji = rest.join('');
      }

      if (!koji || koji.length < 2) continue;
      if (['並順','場所','箇所','工事項目','数量','単位','単価','金額','備考'].includes(basho)) continue;

      pageRows.push({ no: noVal, basho, kasho, koji, suryo, tani, tanka });
    }

    // 見積行が3件以上あるページだけ採用（図面・表紙は無視）
    if (pageRows.length >= 3) {
      allRows.push(...pageRows);
    }
  }

  // 重複排除
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

  const set = (col, row, val) => {
    if (val === null || val === undefined || val === '') return;
    const addr = col + row;
    if (!ws[addr]) ws[addr] = {};
    ws[addr].v = val;
    ws[addr].t = typeof val === 'number' ? 'n' : 's';
  };

  for (const r of rows) {
    const row = r.no + 7;
    set('B', row, r.basho);
    set('C', row, r.kasho);
    set('D', row, r.koji);
    if (r.suryo !== null) set('F', row, r.suryo);
    if (r.tani) set('G', row, r.tani);
    if (r.tanka !== null) set('H', row, r.tanka);
  }

  return XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });
}

module.exports = { parseEstimateRows, buildExcel };