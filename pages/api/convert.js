import { IncomingForm } from 'formidable';
import fs from 'fs';
import pdfParse from 'pdf-parse';
import { parseEstimateRows, buildExcel } from '../../lib/converter';

export const config = {
  api: { bodyParser: false },
};

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    // ファイル受け取り
    const form = new IncomingForm({ keepExtensions: true });
    const [, files] = await form.parse(req);
    const file = Array.isArray(files.pdf) ? files.pdf[0] : files.pdf;

    if (!file) return res.status(400).json({ error: 'PDFファイルが見つかりません' });

    // PDF → テキスト
    const pdfBuf = fs.readFileSync(file.filepath);
    const parsed = await pdfParse(pdfBuf);
    const text = parsed.text;

    // テキスト → 見積行
    const rows = parseEstimateRows(text);
    if (rows.length === 0) {
      return res.status(422).json({ error: '見積データが抽出できませんでした。PDFの形式をご確認ください。' });
    }

    // Excel生成
    const xlsxBuf = buildExcel(rows);

    // ファイル名を元のPDF名から生成
    const originalName = file.originalFilename || 'output';
    const baseName = originalName.replace(/\.pdf$/i, '');
    const outputName = `${baseName}_変換.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(outputName)}`);
    res.setHeader('X-Row-Count', rows.length);
    res.send(xlsxBuf);

    // 一時ファイル削除
    fs.unlinkSync(file.filepath);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: `変換エラー: ${err.message}` });
  }
}
