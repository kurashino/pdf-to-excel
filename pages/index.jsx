import { useState, useRef } from 'react';
import Head from 'next/head';
import * as XLSX from 'xlsx';

export default function Home() {
  const [state, setState] = useState('idle');
  const [rowCount, setRowCount] = useState(0);
  const [errorMsg, setErrorMsg] = useState('');
  const [fileName, setFileName] = useState('');
  const [dlUrl, setDlUrl] = useState('');
  const [dlName, setDlName] = useState('');
  const [dragging, setDragging] = useState(false);
  const inputRef = useRef(null);

  const TEMPLATE_B64 = process.env.NEXT_PUBLIC_TEMPLATE_B64;

  const reset = () => {
    setState('idle');
    setRowCount(0);
    setErrorMsg('');
    setFileName('');
    if (dlUrl) URL.revokeObjectURL(dlUrl);
    setDlUrl('');
    setDlName('');
    if (inputRef.current) inputRef.current.value = '';
  };

  const parseEstimateRows = (text) => {
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
        let uIdx = -1;
        for (let i = 0; i < rest.length; i++) {
          if (unitPattern.test(rest[i])) { uIdx = i; break; }
        }
        let basho = '', kasho = '', koji = '', suryo = null, tani = '', tanka = null;
        if (uIdx >= 0) {
          tani = rest[uIdx];
          if (uIdx > 0 && numPattern.test(rest[uIdx - 1])) {
            suryo = parseFloat(rest[uIdx - 1].replace(/,/g, ''));
            const before = rest.slice(0, uIdx - 1);
            basho = before[0] || ''; kasho = before[1] || ''; koji = before.slice(2).join('');
          } else {
            const before = rest.slice(0, uIdx);
            basho = before[0] || ''; kasho = before[1] || ''; koji = before.slice(2).join('');
          }
          const afterNums = rest.slice(uIdx + 1).filter(t => numPattern.test(t));
          if (afterNums.length > 0) tanka = parseFloat(afterNums[0].replace(/,/g, ''));
        } else {
          basho = rest[0] || ''; kasho = rest[1] || ''; koji = rest.slice(2).join('');
        }
        if (!basho) continue;
        pageRows.push({ no: noVal, basho, kasho, koji, suryo, tani, tanka });
      }
      if (pageRows.length >= 3) allRows.push(...pageRows);
    }
    const seen = new Set();
    return allRows.filter(r => { if (seen.has(r.no)) return false; seen.add(r.no); return true; })
      .sort((a, b) => a.no - b.no);
  };

  const buildExcel = (rows, templateB64) => {
    const buf = Uint8Array.from(atob(templateB64), c => c.charCodeAt(0));
    const wb = XLSX.read(buf, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const headers = {};
    for (let c = 0; c < 20; c++) {
      const addr = XLSX.utils.encode_cell({ r: 6, c });
      const cell = ws[addr];
      if (cell && cell.v) headers[cell.v.toString().trim()] = XLSX.utils.encode_col(c);
    }
    const set = (col, row, val) => {
      if (!col || val === null || val === undefined || val === '') return;
      const addr = col + row;
      if (!ws[addr]) ws[addr] = {};
      ws[addr].v = val;
      ws[addr].t = typeof val === 'number' ? 'n' : 's';
    };
    for (const r of rows) {
      const row = r.no + 7;
      set(headers['場所'], row, r.basho);
      set(headers['箇所'], row, r.kasho);
      set(headers['工事項目'], row, r.koji);
      if (r.suryo !== null) set(headers['数量'], row, r.suryo);
      if (r.tani) set(headers['単位'], row, r.tani);
      if (r.tanka !== null) set(headers['単価'], row, r.tanka);
    }
    return XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  };

  const process = async (file) => {
    if (!file || !file.name.toLowerCase().endsWith('.pdf')) {
      setErrorMsg('PDFファイルを選択してください'); setState('error'); return;
    }
    setFileName(file.name);
    setState('uploading');
    try {
      const pdfjsLib = await import('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.mjs');
      pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.mjs';
      const ab = await file.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({ data: ab }).promise;
      let allText = '';
      for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const tc = await page.getTextContent();
        allText += tc.items.map(x => x.str).join(' ') + '\f';
      }
      const rows = parseEstimateRows(allText);
      if (!rows.length) throw new Error('見積データが抽出できませんでした');
      setRowCount(rows.length);
      const res = await fetch('/api/template');
      const { b64 } = await res.json();
      const xlsxData = buildExcel(rows, b64);
      const blob = new Blob([xlsxData], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      setDlUrl(URL.createObjectURL(blob));
      setDlName(file.name.replace('.pdf', '') + '_変換.xlsx');
      setState('done');
    } catch (e) {
      setErrorMsg(e.message); setState('error');
    }
  };

  const onDrop = (e) => { e.preventDefault(); setDragging(false); const f = e.dataTransfer.files[0]; if (f) process(f); };

  return (
    <>
      <Head><title>見積PDF → エクセル変換</title></Head>
      <style>{`
        *{box-sizing:border-box;margin:0;padding:0}
        body{font-family:'Hiragino Sans','Meiryo',sans-serif;background:#f5f5f0;min-height:100vh;display:flex;align-items:center;justify-content:center;padding:2rem}
        .card{background:#fff;border-radius:16px;box-shadow:0 2px 24px rgba(0,0,0,0.08);padding:2.5rem 2rem;width:100%;max-width:520px}
        .drop{border:2px dashed #d0d0d0;border-radius:12px;padding:2.5rem 1.5rem;text-align:center;cursor:pointer;transition:all 0.15s;background:#fafaf8}
        .drop.over,.drop:hover{border-color:#3b7dd8;background:#eff5ff}
        input[type=file]{display:none}
        .status{margin-top:1.25rem;background:#f7f7f5;border-radius:10px;padding:1rem 1.25rem}
        .row{display:flex;align-items:center;gap:.6rem;font-size:.9rem;color:#555;padding:.3rem 0}
        .row.done{color:#2a7a2a}.row.active{color:#1a1a1a;font-weight:600}
        .dot{width:18px;height:18px;border-radius:50%;border:2px solid currentColor;display:flex;align-items:center;justify-content:center;flex-shrink:0;font-size:10px}
        .spin{width:18px;height:18px;border:2px solid #ccc;border-top-color:#3b7dd8;border-radius:50%;animation:spin .7s linear infinite;flex-shrink:0}
        @keyframes spin{to{transform:rotate(360deg)}}
        .dl-btn{display:flex;align-items:center;justify-content:center;gap:.5rem;margin-top:1.25rem;width:100%;padding:.85rem;background:#1a6fd4;color:#fff;border:none;border-radius:10px;font-size:1rem;font-weight:600;cursor:pointer;text-decoration:none}
        .dl-btn:hover{background:#145bb0}
        .reset{display:block;margin:.75rem auto 0;background:none;border:none;color:#aaa;font-size:.82rem;cursor:pointer}
        .err{margin-top:1rem;background:#fff5f5;border:1px solid #fcc;border-radius:10px;padding:.85rem 1rem;color:#c0392b;font-size:.88rem}
      `}</style>
      <div className="card">
        <p style={{fontSize:'1.25rem',fontWeight:700,marginBottom:'.25rem'}}>見積PDF → エクセル変換</p>
        <p style={{fontSize:'.85rem',color:'#888',marginBottom:'1.75rem'}}>PDFをアップロードするだけでＡＡ原紙フォーマットに自動変換</p>
        {state === 'idle' && (
          <div className={`drop${dragging?' over':''}`}
            onDragOver={e=>{e.preventDefault();setDragging(true)}}
            onDragLeave={()=>setDragging(false)}
            onDrop={onDrop}
            onClick={()=>inputRef.current?.click()}>
            <input ref={inputRef} type="file" accept=".pdf" onChange={e=>process(e.target.files[0])}/>
            <div style={{fontSize:'2.5rem',marginBottom:'.75rem'}}>📄</div>
            <p style={{fontSize:'1rem',fontWeight:600}}>PDFをここにドロップ</p>
            <p style={{fontSize:'.82rem',color:'#aaa',marginTop:'.25rem'}}>またはクリックして選択</p>
          </div>
        )}
        {(state==='uploading'||state==='done') && (
          <div className="status">
            <div className={`row done`}><div className="dot">✓</div><span>{fileName} を読み込みました</span></div>
            <div className={`row ${state==='uploading'?'active':'done'}`}>
              {state==='uploading'?<div className="spin"/>:<div className="dot">✓</div>}
              <span>{state==='uploading'?'解析・変換中...':`${rowCount}件抽出しました`}</span>
            </div>
          </div>
        )}
        {state==='done' && dlUrl && (
          <>
            <a className="dl-btn" href={dlUrl} download={dlName}>⬇ エクセルをダウンロード</a>
            <button className="reset" onClick={reset}>別のPDFを変換する</button>
          </>
        )}
        {state==='error' && (
          <>
            <div className="err">⚠ {errorMsg}</div>
            <button className="reset" onClick={reset} style={{color:'#555',marginTop:'1rem'}}>やり直す</button>
          </>
        )}
      </div>
    </>
  );
}