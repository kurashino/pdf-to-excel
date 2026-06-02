# 見積PDF → エクセル変換アプリ

PDFの見積依頼書をアップロードするだけでＡＡ原紙フォーマットのエクセルに自動変換します。

## 動作

1. PDF をドロップ or クリック選択
2. サーバー側でテキスト解析
3. 場所・箇所・工事項目・数量・単位・単価を抽出
4. ＡＡ原紙テンプレート（B/C/D/M/N/O列）に書き込み
5. エクセルをダウンロード

## ローカル起動

```bash
npm install
npm run dev
```
→ http://localhost:3000

## Vercel デプロイ手順

1. このリポジトリを GitHub に push
2. https://vercel.com で「New Project」→ リポジトリを選択
3. 設定はデフォルトのまま「Deploy」
4. 完了 🎉

## 技術スタック

- Next.js 14 (Pages Router)
- pdf-parse（PDF テキスト抽出）
- xlsx（エクセル生成）
- Vercel（ホスティング）
