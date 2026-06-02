import { useState, useRef, useEffect } from 'react';
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

  useEffect(() => {
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
    script.onload = () => {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc =
        'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    };
    document.head.appendChild(script);
  }, []);

  const reset = () => {
    setState('idle'); setRowCount(0); setErrorMsg(''); setFileName('');
    if (dlUrl) URL.revokeObjectURL(dlUrl);
    setDlUrl(''); setDlName('');
    if (inputRef.current) inputRef.current.value = '';
  };

  const parseEstimateRows = (text) => {
    const rows = [];
    const seen = new Set();
    const flat = text.replace(/\s+/g, ' ').trim();

    // 数量ありパターン：番号 場所 箇所 工事項目 数量 単位
    const p1 = /\b(\d{1,2})\s+([\u3000-\u9FFFﾀ-ﾟA-Za-z()（）・\-]+)\s+([\u3000-\u9FFFﾀ-ﾟA-Za-z()（）・\-\/]+)\s+([\u3000-\u9FFFﾀ-ﾟA-Za-z()（）・\-\/,？\?]+?)\s+([\d.]+)\s*(式|㎡|ヶ|台|本|枚|ｹ|ケ)/g;
    let m;
    while ((m = p1.exec(flat)) !== null) {
      const no = parseInt(m[1]);
      if (no < 1 || no > 100 || seen.has(no)) continue;
      seen.add(no);
      let tanka = null;
      const after = flat.slice(m.index + m[0].length, m.index + m[0].length + 30);
      const tm = after.match(/^\s*([\d,]+)/);
      if (tm) tanka = parseFloat(tm[1].replace(/,/g, ''));
      rows.push({ no, basho: m[2], kasho: m[3], koji: m[4], suryo: parseFloat(m[5]), tani: m[6], tanka });
    }

    // 数量なしパターン：番号 場所 箇所 工事項目 単位
    const p2 = /\b(\d{1,2})\s+([\u3000-\u9FFFﾀ-ﾟA-Za-z()（）・\-]+)\s+([\u3000-\u9FFFﾀ-ﾟA-Za-z()（）・\-\/]+)\s+([\u3000-\u9FFFﾀ-ﾟA-Za-z()（）・\-\/,？\?]+?)\s+(式|㎡|ヶ|台|本|枚|ｹ|ケ)/g;
    while ((m = p2.exec(flat)) !== null) {
      const no = parseInt(m[1]);
      if (no < 1 || no > 100 || seen.has(no)) continue;
      seen.add(no);
      rows.push({ no, basho: m[2], kasho: m[3], koji: m[4], suryo: null, tani: m[5], tanka: null });
    }

    return rows.sort((a, b) => a.no - b.no);
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
      const pdfjsLib = window.pdfjsLib;
      if (!pdfjsLib) throw new Error('PDF読み込みライブラリが準備中です。少し待ってから再度お試しください。');
      const ab = await file.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({ data: ab }).promise;
      let allText = '';
      for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const tc = await page.getTextContent();
        allText += tc.items.map(x => x.str).join(' ') + ' ';
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
        .srow{display:flex;align-items:center;gap:.6rem;font-size:.9rem;color:#555;padding:.3rem 0}
        .srow.done{color:#2a7a2a}.srow.active{color:#1a1a1a;font-weight:600}
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
            <div className="srow done"><div className="dot">✓</div><span>{fileName} を読み込みました</span></div>
            <div className={`srow ${state==='uploading'?'active':'done'}`}>
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