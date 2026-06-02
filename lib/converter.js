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

      const basho = tokens[1];
      const kasho = tokens[2];

      if (['並順','場所','箇所','工事項目','NO','No'].includes(basho)) continue;

      const unitPattern = /^(式|㎡|ヶ|台|本|枚|ｹ|ケ|m2|m²)$/;
      const numPattern = /^[\d,.]+$/;
      const rest = tokens.slice(3);

      let koji = '', suryo = null, tani = '', tanka = null;

      let uIdx = -1;
      for (let i = 0; i < rest.length; i++) {
        if (unitPattern.test(rest[i])) { uIdx = i; break; }
      }

      if (uIdx >= 0) {
        tani = rest[uIdx];
        if (uIdx > 0 && numPattern.test(rest[uIdx - 1])) {
          suryo = parseFloat(rest[uIdx - 1].replace(/,/g, ''));
          koji = rest.slice(0, uIdx - 1).join('');
        } else {
          koji = rest.slice(0, uIdx).join('');
        }
        const afterUnit = rest.slice(uIdx + 1).filter(t => numPattern.test(t));
        if (afterUnit.length > 0) tanka = parseFloat(afterUnit[0].replace(/,/g, ''));
      } else {
        koji = rest.join('');
      }

      if (!koji || koji.length < 2) continue;

      pageRows.push({ no: noVal, basho, kasho, koji, suryo, tani, tanka });
    }

    if (pageRows.length >= 3) {
      allRows.push(...pageRows);
    }
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