import { useState, useRef, useEffect } from 'react';
import Head from 'next/head';

export default function Home() {
  const [state, setState] = useState('idle');
  const [rowCount, setRowCount] = useState(0);
  const [errorMsg, setErrorMsg] = useState('');
  const [fileName, setFileName] = useState('');
  const [dlUrl, setDlUrl] = useState('');
  const [dlName, setDlName] = useState('');
  const [preview, setPreview] = useState([]);
  const [dragging, setDragging] = useState(false);
  const [rawText, setRawText] = useState('');
  const inputRef = useRef(null);
  const xlsxRef = useRef(null);

  useEffect(() => {
    // pdf.js
    const s1 = document.createElement('script');
    s1.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
    s1.onload = () => {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc =
        'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    };
    document.head.appendChild(s1);
    // xlsx
    const s2 = document.createElement('script');
    s2.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    document.head.appendChild(s2);
  }, []);

  const reset = () => {
    setState('idle'); setRowCount(0); setErrorMsg(''); setFileName('');
    setPreview([]); setRawText('');
    if (dlUrl) URL.revokeObjectURL(dlUrl);
    setDlUrl(''); setDlName('');
    if (inputRef.current) inputRef.current.value = '';
  };

  const parseEstimateRows = (text) => {
    const rows = [];
    const seen = new Set();
    const tokens = text.replace(/\s+/g, ' ').trim().split(' ').filter(t => t.length > 0);
    const UNIT_SET = new Set(['式','㎡','ヶ','ヵ','台','本','枚','ｹ','ケ','組','回']);
    const NUM_RE = /^[\d.]+$/;

    let i = 0;
    while (i < tokens.length) {
      const no = parseInt(tokens[i]);
      if (isNaN(no) || no < 1 || no > 100 || seen.has(no)) { i++; continue; }
      // tokens[i]が純粋な数字かチェック
      if (!/^\d{1,2}$/.test(tokens[i])) { i++; continue; }

      // 単位を探す（最大15トークン先まで）
      let uIdx = -1;
      for (let j = i + 1; j < Math.min(i + 15, tokens.length); j++) {
        if (UNIT_SET.has(tokens[j])) { uIdx = j; break; }
      }
      if (uIdx === -1) { i++; continue; }

      const between = tokens.slice(i + 1, uIdx);
      if (between.length < 2) { i++; continue; }

      let basho, kasho, koji, suryo = null;
      if (NUM_RE.test(between[between.length - 1])) {
        suryo = parseFloat(between[between.length - 1]);
        basho = between[0] || '';
        kasho = between[1] || '';
        koji = between.slice(2, -1).join('') || '';
      } else {
        basho = between[0] || '';
        kasho = between[1] || '';
        koji = between.slice(2).join('') || '';
      }

      if (!basho || basho.length < 1) { i++; continue; }

      let tanka = null;
      for (let j = uIdx + 1; j < Math.min(uIdx + 5, tokens.length); j++) {
        if (/^[\d,]+$/.test(tokens[j])) {
          tanka = parseFloat(tokens[j].replace(/,/g, ''));
          break;
        }
      }

      seen.add(no);
      rows.push({ no, basho, kasho, koji, suryo, tani: tokens[uIdx], tanka });
      i = uIdx + 1;
    }

    return rows.sort((a, b) => a.no - b.no);
  };

  const process = async (file) => {
    if (!file || !file.name.toLowerCase().endsWith('.pdf')) {
      setErrorMsg('PDFファイルを選択してください'); setState('error'); return;
    }
    setFileName(file.name);
    setState('uploading');
    try {
      const pdfjsLib = window.pdfjsLib;
      if (!pdfjsLib) throw new Error('ライブラリ準備中です。3秒後に再試行してください。');
      const ab = await file.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({ data: ab }).promise;
      let allText = '';
      for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const tc = await page.getTextContent();
        allText += tc.items.map(x => x.str).join(' ') + ' ';
      }
      setRawText(allText.slice(0, 300));
      const rows = parseEstimateRows(allText);
      if (!rows.length) throw new Error('抽出0件。テキスト: ' + allText.slice(0, 200));
      setRowCount(rows.length);
      setPreview(rows.slice(0, 10));

      const XLSX = window.XLSX;
      if (!XLSX) throw new Error('XLSXライブラリ準備中です。再試行してください。');
      const res = await fetch('/api/template');
      const { b64 } = await res.json();
      const buf = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
      const wb = XLSX.read(buf, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const set = (col, row, val) => {
        if (!col || val === null || val === undefined || val === '') return;
        const addr = col + row;
        if (!ws[addr]) ws[addr] = {};
        ws[addr].v = val; ws[addr].t = typeof val === 'number' ? 'n' : 's';
      };
      for (const r of rows) {
        const row = r.no + 7;
        set('B', row, r.basho); set('C', row, r.kasho); set('D', row, r.koji);
        if (r.suryo !== null) set('M', row, r.suryo);
        if (r.tani) set('N', row, r.tani);
        if (r.tanka !== null) set('O', row, r.tanka);
      }
      const xlsxData = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
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
        .card{background:#fff;border-radius:16px;box-shadow:0 2px 24px rgba(0,0,0,0.08);padding:2.5rem 2rem;width:100%;max-width:640px}
        .drop{border:2px dashed #d0d0d0;border-radius:12px;padding:2.5rem 1.5rem;text-align:center;cursor:pointer;transition:all 0.15s;background:#fafaf8}
        .drop.over,.drop:hover{border-color:#3b7dd8;background:#eff5ff}
        input[type=file]{display:none}
        .status{margin-top:1.25rem;background:#f7f7f5;border-radius:10px;padding:1rem 1.25rem}
        .srow{display:flex;align-items:center;gap:.6rem;font-size:.9rem;color:#555;padding:.3rem 0}
        .srow.done{color:#2a7a2a}.srow.active{color:#1a1a1a;font-weight:600}
        .dot{width:18px;height:18px;border-radius:50%;border:2px solid currentColor;display:flex;align-items:center;justify-content:center;flex-shrink:0;font-size:10px}
        .spin{width:18px;height:18px;border:2px solid #ccc;border-top-color:#3b7dd8;border-radius:50%;animation:spin .7s linear infinite;flex-shrink:0}
        @keyframes spin{to{transform:rotate(360deg)}}
        .preview{width:100%;border-collapse:collapse;font-size:12px;margin-top:1rem}
        .preview th{background:#f0f0ee;padding:5px 8px;text-align:left;border:0.5px solid #ddd;font-weight:500;color:#555}
        .preview td{padding:4px 8px;border:0.5px solid #ddd;color:#333}
        .preview tr:nth-child(even) td{background:#fafaf8}
        .dl-btn{display:flex;align-items:center;justify-content:center;gap:.5rem;margin-top:1.25rem;width:100%;padding:.85rem;background:#1a6fd4;color:#fff;border:none;border-radius:10px;font-size:1rem;font-weight:600;cursor:pointer;text-decoration:none}
        .dl-btn:hover{background:#145bb0}
        .reset{display:block;margin:.75rem auto 0;background:none;border:none;color:#aaa;font-size:.82rem;cursor:pointer}
        .err{margin-top:1rem;background:#fff5f5;border:1px solid #fcc;border-radius:10px;padding:.85rem 1rem;color:#c0392b;font-size:.88rem;word-break:break-all}
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
        {state==='done' && preview.length > 0 && (
          <>
            <p style={{fontSize:'12px',color:'#888',marginTop:'1rem',marginBottom:'4px'}}>抽出データプレビュー（先頭10件）</p>
            <div style={{overflowX:'auto'}}>
              <table className="preview">
                <thead><tr><th>No</th><th>場所</th><th>箇所</th><th>工事項目</th><th>数量</th><th>単位</th><th>単価</th></tr></thead>
                <tbody>{preview.map(r=>(
                  <tr key={r.no}>
                    <td>{r.no}</td><td>{r.basho}</td><td>{r.kasho}</td><td>{r.koji}</td>
                    <td>{r.suryo??''}</td><td>{r.tani}</td><td>{r.tanka!=null?r.tanka.toLocaleString():''}</td>
                  </tr>
                ))}</tbody>
              </table>
            </div>
          </>
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