# 台語字練習簿編輯器

靜態網頁程式，可產生能列印的 A4 台語文書寫練習單，具所見即所得編輯與 PDF 輸出。

## 使用方式

1. **字體**：請將 `Iansui-Regular.woff2` 放在專案內的 `fonts/` 資料夾：

   ```
   台語字練習簿/
   └── fonts/
       └── Iansui-Regular.woff2
   ```

   預覽與輸出的 PDF 都會使用此字體顯示範例文字。

2. **開啟網頁**：用瀏覽器直接開啟 `index.html`（若從檔案路徑開啟，部分瀏覽器可能限制載入本地字體，建議用本機 HTTP 伺服器）。

   本機預覽範例：

   ```bash
   # Python 3
   python3 -m http.server 8000
   # 或 npx
   npx serve .
   ```

   然後在瀏覽器開啟 http://localhost:8000

3. **編輯內容**
   - 左側「每一行範例文字」：可新增、刪除行，並編輯每行行首的範例字。
   - 「底線樣式」：切換「一條底線」或「英文式三條線」。
   - 右側即時顯示 A4 預覽。

4. **輸出 PDF**：點「輸出 PDF」即可下載可列印的 A4 PDF 檔。

## 檔案說明

- `index.html` — 主頁（介面與預覽區）
- `styles.css` — 版型、字體、底線樣式
- `app.js` — 行編輯、預覽更新、PDF 產生
- `fonts/Iansui-Regular.woff2` — 需自行放入字體檔

## 技術

- 純前端，無後端
- PDF 透過 [html2pdf.js](https://github.com/eKoopmans/html2pdf.js)（html2canvas + jsPDF）於瀏覽器產生
