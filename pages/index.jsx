import { useState, useRef } from 'react';
import Head from 'next/head';

export default function Home() {
  const [state, setState] = useState('idle'); // idle | uploading | done | error
  const [rowCount, setRowCount] = useState(0);
  const [errorMsg, setErrorMsg] = useState('');
  const [fileName, setFileName] = useState('');
  const [dlUrl, setDlUrl] = useState('');
  const [dlName, setDlName] = useState('');
  const [dragging, setDragging] = useState(false);
  const inputRef = useRef(null);

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

  const upload = async (file) => {
    if (!file || !file.name.toLowerCase().endsWith('.pdf')) {
      setErrorMsg('PDFファイルを選択してください');
      setState('error');
      return;
    }
    setFileName(file.name);
    setState('uploading');

    const form = new FormData();
    form.append('pdf', file);

    try {
      const res = await fetch('/api/convert', { method: 'POST', body: form });
      if (!res.ok) {
        const json = await res.json().catch(() => ({ error: '変換に失敗しました' }));
        throw new Error(json.error);
      }
      const count = res.headers.get('X-Row-Count') || '?';
      setRowCount(count);

      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const cd = res.headers.get('Content-Disposition') || '';
      const match = cd.match(/filename\*=UTF-8''(.+)/);
      const name = match ? decodeURIComponent(match[1]) : file.name.replace('.pdf', '') + '_変換.xlsx';

      setDlUrl(url);
      setDlName(name);
      setState('done');
    } catch (e) {
      setErrorMsg(e.message);
      setState('error');
    }
  };

  const onDrop = (e) => {
    e.preventDefault();
    setDragging(false);
    const f = e.dataTransfer.files[0];
    if (f) upload(f);
  };

  return (
    <>
      <Head>
        <title>見積PDF → エクセル変換</title>
        <meta name="viewport" content="width=device-width, initial-scale=1" />
      </Head>

      <style>{`
        *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Hiragino Sans', 'Meiryo', sans-serif; background: #f5f5f0; min-height: 100vh; display: flex; align-items: center; justify-content: center; padding: 2rem; }
        .card { background: #fff; border-radius: 16px; box-shadow: 0 2px 24px rgba(0,0,0,0.08); padding: 2.5rem 2rem; width: 100%; max-width: 520px; }
        .title { font-size: 1.25rem; font-weight: 700; color: #1a1a1a; margin-bottom: 0.25rem; }
        .subtitle { font-size: 0.85rem; color: #888; margin-bottom: 1.75rem; }
        .drop { border: 2px dashed #d0d0d0; border-radius: 12px; padding: 2.5rem 1.5rem; text-align: center; cursor: pointer; transition: all 0.15s; background: #fafaf8; }
        .drop.over { border-color: #3b7dd8; background: #eff5ff; }
        .drop:hover { border-color: #aaa; }
        .drop-icon { font-size: 2.5rem; margin-bottom: 0.75rem; }
        .drop-label { font-size: 1rem; font-weight: 600; color: #333; margin-bottom: 0.25rem; }
        .drop-sub { font-size: 0.82rem; color: #aaa; }
        input[type=file] { display: none; }
        .status { margin-top: 1.25rem; background: #f7f7f5; border-radius: 10px; padding: 1rem 1.25rem; }
        .status-row { display: flex; align-items: center; gap: 0.6rem; font-size: 0.9rem; color: #555; padding: 0.3rem 0; }
        .status-row.done { color: #2a7a2a; }
        .status-row.active { color: #1a1a1a; font-weight: 600; }
        .status-row.error { color: #c0392b; }
        .dot { width: 18px; height: 18px; border-radius: 50%; border: 2px solid currentColor; display:flex; align-items:center; justify-content:center; flex-shrink:0; font-size: 10px; }
        .spinner { width: 18px; height: 18px; border: 2px solid #ccc; border-top-color: #3b7dd8; border-radius: 50%; animation: spin 0.7s linear infinite; flex-shrink:0; }
        @keyframes spin { to { transform: rotate(360deg); } }
        .dl-btn { display: flex; align-items: center; justify-content: center; gap: 0.5rem; margin-top: 1.25rem; width: 100%; padding: 0.85rem; background: #1a6fd4; color: #fff; border: none; border-radius: 10px; font-size: 1rem; font-weight: 600; cursor: pointer; text-decoration: none; transition: background 0.15s; }
        .dl-btn:hover { background: #145bb0; }
        .reset { display: block; margin: 0.75rem auto 0; background: none; border: none; color: #aaa; font-size: 0.82rem; cursor: pointer; }
        .reset:hover { color: #555; }
        .err-box { margin-top: 1rem; background: #fff5f5; border: 1px solid #fcc; border-radius: 10px; padding: 0.85rem 1rem; color: #c0392b; font-size: 0.88rem; }
      `}</style>

      <div className="card">
        <p className="title">見積PDF → エクセル変換</p>
        <p className="subtitle">PDFをアップロードするだけでＡＡ原紙フォーマットに自動変換</p>

        {state === 'idle' && (
          <div
            className={`drop${dragging ? ' over' : ''}`}
            onDragOver={e => { e.preventDefault(); setDragging(true); }}
            onDragLeave={() => setDragging(false)}
            onDrop={onDrop}
            onClick={() => inputRef.current?.click()}
          >
            <input ref={inputRef} type="file" accept=".pdf" onChange={e => upload(e.target.files[0])} />
            <div className="drop-icon">📄</div>
            <p className="drop-label">PDFをここにドロップ</p>
            <p className="drop-sub">またはクリックして選択</p>
          </div>
        )}

        {(state === 'uploading' || state === 'done') && (
          <div className="status">
            <div className={`status-row ${state === 'done' ? 'done' : 'done'}`}>
              <div className="dot">✓</div>
              <span>{fileName} を読み込みました</span>
            </div>
            <div className={`status-row ${state === 'uploading' ? 'active' : 'done'}`}>
              {state === 'uploading' ? <div className="spinner" /> : <div className="dot">✓</div>}
              <span>{state === 'uploading' ? 'テキスト解析・変換中...' : `${rowCount}件のデータを抽出しました`}</span>
            </div>
            {state === 'done' && (
              <div className="status-row done">
                <div className="dot">✓</div>
                <span>エクセルファイルを生成しました</span>
              </div>
            )}
          </div>
        )}

        {state === 'done' && dlUrl && (
          <>
            <a className="dl-btn" href={dlUrl} download={dlName}>
              ⬇ エクセルをダウンロード
            </a>
            <button className="reset" onClick={reset}>別のPDFを変換する</button>
          </>
        )}

        {state === 'error' && (
          <>
            <div className="err-box">⚠ {errorMsg}</div>
            <button className="reset" onClick={reset} style={{ color: '#555', marginTop: '1rem' }}>やり直す</button>
          </>
        )}
      </div>
    </>
  );
}
