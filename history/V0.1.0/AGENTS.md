# Codex 專案指示

## 專案

- 主程式：`WordMerge_V0.1.0.pyw`
- 視窗標題：`Word組合工`
- 目前版本：`V0.1.0`
- 主要平台：Windows 10／11
- 建議 Python：3.13

## 修改規則

- 一般錯誤修正或小功能增加第三碼，例如 `V0.1.1`。
- 較大功能或架構變更才增加第二碼，例如 `V0.2.0`。
- 進版時同步修改程式內 `VERSION`、檔名、README、CHANGELOG 與兩份 Nuitka 建置批次檔。
- 保持 `.pyw` 執行時不顯示命令提示字元。
- GUI 操作不得在主執行緒執行長時間合併工作。
- 不得直接覆蓋任何來源 Word 文件。
- 保留中文路徑及含空白路徑的相容性。
- 修改後至少執行 `python -m py_compile WordMerge_版本.pyw`。
- 影響合併引擎時，需以含文字、表格、圖片、頁首頁尾及不同樣式的多份文件進行實際驗收。

## 檔案規則

- 不提交 `.venv/`、`Nuitka_Output/`、`*.build/`、`*.dist/`、`*.onefile-build/` 或 `__pycache__/`。
- 保留 `CODEX_TRANSFER.md`，每次交接時更新現況、已知限制及下一步。
