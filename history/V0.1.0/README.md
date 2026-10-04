# WordMerge V0.1.0（Word組合工）

把多個 Word `.docx` 文件依清單順序合併成一個新的 Word 文件。

## 功能

- 將多個 `.docx` 拖曳到視窗載入。
- 也可按「加入檔案」一次選取多份文件。
- 上移、下移、移除及清空合併清單。
- 可選擇讓每份文件從新的一頁開始。
- 透過 `docxcompose` 合併文字、表格、圖片、樣式等內容。
- 合併作業在背景執行，避免視窗失去回應。
- 以暫存檔完成輸出後才取代目標檔，降低輸出中斷造成損壞的風險。
- 來源檔與輸出檔同名時會阻止覆蓋。

## 執行環境

- Windows 10／11
- Python 3.11～3.13（建議 Python 3.13）
- 輸入格式：`.docx`
- 輸出格式：`.docx`

## 直接執行

雙擊：

```text
WordMerge_V0.1.0.pyw
```

第一次執行若缺少套件，程式會詢問是否自動安裝。也可先執行：

```bat
python -m pip install -r requirements.txt
```

或直接雙擊 `run_windows.bat`，批次檔會建立 `.venv`、安裝套件並啟動程式。

## 建置 EXE

先安裝 Microsoft Visual Studio Build Tools 的 C++ 編譯工具，再使用以下其中一個批次檔：

- `build_nuitka_standalone.bat`：先建立較容易除錯的資料夾版。
- `build_nuitka_onefile.bat`：建立單一 EXE 檔。

建議先確認 Standalone 版可正常拖曳及合併，再建立 Onefile 版。

## 已知限制

- V0.1.0 僅支援 `.docx`，不直接支援舊式 `.doc`。
- 不同文件若使用同名但定義不同的樣式，合併後可能以主文件的樣式為準。
- 主文件是清單中的第一份文件，因此頁面設定、預設樣式及部分頁首頁尾行為可能以第一份文件為基礎。
- 含有巨集的 `.docm` 不在本版支援範圍內。
