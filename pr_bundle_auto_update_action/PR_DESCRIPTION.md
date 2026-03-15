# Auto update workflow & script

此 PR 加入：

- GitHub Actions 工作流程（`.github/workflows/update.yml`）：
  - 當 `data/marksix_latest_200.xlsx` 有新推送時觸發
  - 或手動觸發、每日排程觸發（UTC 03:05）
  - 讀取 Excel → 重新計算最近20/50/200 → 覆寫 `index.html` 三個表格與「數據更新至」
  - 複製最新 Excel 至 repo 根目錄供下載按鈕使用
  - 另輸出 `marksix_latest_200_updated.xlsx`（raw/freq20/pairs50/pairs200）

- Python 腳本（`scripts/update.py`）：
  - 只計 N1–N6（不含特別號）
  - 以 `h2` 標題定位對應表格的 `<tbody>` 進行覆寫

> 頁面結構與標題（`index.html`）已符合需求：
> - 最近20期：各號碼出現次數（由高至低）
> - 最近50期：同期出現（號碼對）次數排名
> - 最近200期：同期出現（號碼對）次數排名
> - 「數據更新至：……」段落

合併後的使用方式：每次只要將最新 Excel 推到 `data/marksix_latest_200.xlsx`，Action 會自動更新頁面並生成可下載的最新 Excel。
