# SharePoint URL Decoder

SharePoint URLを簡単にデコード・加工するChrome拡張機能です。

## 機能

- **URLデコード**: SharePoint URLに含まれるパーセントエンコードされた文字を日本語を含めてデコードします
- **URL加工**:
  - `:w:/s`、`:v:/r` などのSharePoint特有の区切りを削除
  - URL末尾のパラメータ（`?web=1` など）を削除
- **クリップボードコピー**: デコード結果を自動的にクリップボードにコピー
- **Officeアプリ連携**: 拡張子に応じてWord/Excel/PowerPointで読み取り専用で開くリンクを生成

## 対応ファイル

以下のOfficeファイル形式に対応しています:

- **Word**: `.doc`, `.docx`, `.docm`
- **Excel**: `.xls`, `.xlsx`, `.xlsm`
- **PowerPoint**: `.ppt`, `.pptx`, `.pptm`

## 使用方法

1. Chrome拡張機能画面を開く
2. デベロッパーモードをオンにする
3. 「パッケージ化されていない拡張機能を読み込む」をクリック
4. このフォルダを選択
5. SharePoint URLをテキストエリアに貼り付け
6. 「実行」ボタンをクリック
7. デコード済みURLがテキストエリアに表示され、クリップボードにコピーされます
8. 必要に応じて「アプリで読み取り専用で開く」リンクをクリック

## 使用例

**入力**:
```
https://example.sharepoint.com/:w:/s/sites/test/Shared%20Documents/%E3%83%95%E3%82%A9%E3%83%AB%E3%83%80%EF%BC%91/%E3%83%95%E3%82%A1%E3%82%A4%E3%83%AB.docx?web=1
```

**出力**:
```
https://example.sharepoint.com/sites/test/Shared Documents/フォルダ１/ファイル.docx
```

## ファイル構成

- `manifest.json` - 拡張機能の設定
- `popup.html` - UIのマークアップ
- `popup.css` - UIのスタイル
- `popup.js` - メイン処理とロジック
- `README.md` - このファイル

## 対応サイト

- `*.sharepoint.com` - Microsoft SharePointのみ対応

## ライセンス

自由に利用・改変できます。
