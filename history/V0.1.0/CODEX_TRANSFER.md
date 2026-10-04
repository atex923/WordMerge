# WordMerge V0.1.0 — Codex 接手說明

## 一、專案目標

製作 Windows Python 視窗工具，讓使用者把多份 Word 文件拖曳載入、調整順序，最後合併輸出成單一 `.docx`。

## 二、目前完成內容

- Tkinter／ttk 視窗，尺寸 `780 × 560`，最小尺寸 `640 × 450`。
- `tkinterdnd2` 負責 Windows 拖放。
- Treeview 顯示順序、檔名與來源資料夾。
- 支援多選移除、Ctrl+A、Delete、上移、下移與清空。
- `docxcompose.Composer` 負責文件合併。
- 可在文件之間插入分頁符號。
- 合併工作使用背景執行緒，完成後回到 Tk 主執行緒更新 UI。
- 先輸出至隱藏暫存檔，成功後使用 `os.replace()` 取代正式輸出。
- 第一次執行若套件缺少，先詢問使用者，再呼叫 pip 安裝。

## 三、核心流程

1. 使用者拖曳或選取 `.docx`。
2. `add_paths()` 驗證副檔名、檔案存在及重複項目。
3. 使用者透過清單決定合併順序。
4. `start_merge()` 驗證清單及輸出路徑。
5. `merge_worker()` 在背景呼叫 `merge_docx_files()`。
6. 第一份文件建立主文件，後續文件透過 Composer 依序附加。
7. 暫存檔完成後原子取代輸出檔。
8. UI 詢問是否開啟結果。

## 四、相依套件

- `python-docx`：Word Open XML 文件讀寫。
- `docxcompose`：合併文件並處理樣式及媒體關聯。
- `tkinterdnd2`：檔案拖放。
- `lxml`：由相關套件使用，Nuitka 建置時明確收錄。

## 五、重要限制

- 只支援 `.docx`，不支援 `.doc`、`.docm`。
- Word 文件合併的格式結果會受 Open XML 樣式、節與頁首頁尾繼承規則影響。
- 第一份文件是主文件；同名樣式衝突時可能以主文件定義為準。
- 尚未提供文件預覽、拖曳列項排序、最近輸出資料夾記憶或批次合併方案儲存。
- 尚未在實際 Windows Word 文件集合及 Nuitka EXE 環境完成端到端驗收。

## 六、版本規則

- 起始版本：`V0.1.0`。
- 一般修正、小功能：增加第三碼，最高可到 99。
- 大型功能或架構變更：增加第二碼並將第三碼歸零。
- 每次進版需同步程式內常數、程式檔名、說明文件及建置批次檔。

## 七、建議驗收項目

1. 拖曳單一、多個及含中文／空白路徑的 `.docx`。
2. 重複載入相同檔案時不重複加入。
3. 非 `.docx` 檔案會被略過並顯示清單。
4. 上移、下移及多選移動後順序正確。
5. 合併文字、表格、圖片、頁首頁尾、頁碼及不同頁面方向的文件。
6. 分頁選項開啟與關閉的輸出結果。
7. 輸出檔已在 Word 開啟時，程式能顯示可理解的錯誤。
8. 不允許輸出檔覆蓋任何來源檔。
9. Standalone 與 Onefile EXE 均可拖曳、合併及開啟輸出。
10. 未安裝 Python 的電腦可直接執行 Onefile EXE。

## 八、Nuitka 建置順序

1. 執行 `build_nuitka_standalone.bat`。
2. 測試 `Nuitka_Output\Standalone` 內的 EXE。
3. Standalone 驗收通過後執行 `build_nuitka_onefile.bat`。
4. 在未安裝 Python 的 Windows 測試 Onefile EXE。
5. 保存 Nuitka XML report，若有缺少模組可由報告追蹤。

## 九、後續開發建議

- 優先以實際複雜 Word 文件驗證頁首頁尾、分節、頁碼與樣式衝突。
- 若要求完全忠實保留 Word 排版，可評估 Windows Microsoft Word COM 合併引擎，並保留目前純 Python 引擎作為備援。
- 增加拖曳清單項目重新排序。
- 增加記住最後輸出資料夾。
- 需要支援 `.doc` 時，可先透過 Word COM 或 LibreOffice 轉成 `.docx` 再合併。
