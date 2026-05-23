# QuickSign Pro Ultimate v2

完整 Mobile-first 電子文件快速簽署網站。

## v2 修復及新增
- 修復 PDF 匯入後不能新增簽名：點擊 PDF 可新增簽署框
- PDF 簽署框支援姓名、職位、日期、手寫簽名
- PDF 簽署框可拖曳重新定位、刪除
- 匯出時以 pdf-lib 嵌入原 PDF，保留原 PDF 所有頁面
- 建立頁新增「重設內容」雙重確認
- 輸出 PDF / PNG 移除所有 QuickSign 水印及法律提示
- 簽署頁新增「完成簽署」按鈕
- 完成後自動儲存至最近文件
- 最近文件可刪除
- 文件排版專業化
- 簽署欄因應簽署人數自動調整
- 預設範本可刪除，並可恢復預設範本

## 安裝
```bash
npm install
npm run dev
npm run build
```

## 部署
- Vercel：直接匯入 GitHub repo
- GitHub Pages：部署 `dist` 資料夾
