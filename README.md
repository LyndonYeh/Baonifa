# Baonifa — 上版清單產生器

自動比對 Git 兩個 commit 之間的差異，一鍵產生上版所需的 **程式修改清單 (Excel)**、**程式修改說明 (Word)** 及 **差異檔壓縮包**，最終打包為一份完整的上版清單 ZIP。

## 功能

- 自動抓取 `develop` 分支的本地與遠端 commit
- 產生完整 diff patch 及逐檔 patch
- 匯出差異原始檔並壓縮
- 解析 patch 產生「程式修改清單.xlsx」（含模組代號、路徑、commit 說明）
- 將 Excel 轉換為「程式修改說明.docx」（含 OLE 附件嵌入）
- 產生刪除檔案清單
- 最終將所有產出打包為 `{版本}上版清單.zip`，並自動清理中間檔案

## 環境需求

- Windows 10/11
- Python 3.10+
- Git
- Microsoft Word（用於 OLE 附件嵌入）

## 安裝

```bash
python install.py
```

或手動安裝相依套件：

```bash
pip install -r requirements.txt
```

### 相依套件

| 套件 | 用途 |
|------|------|
| pandas | 資料處理 |
| openpyxl | Excel 讀寫 |
| python-docx | Word 文件產生 |
| pywin32 | Word OLE 附件嵌入 |

## 使用方式

### 直接執行

```bash
python go.py
```

### 使用打包後的 EXE

```bash
dist\Baonifa.exe
```

### 操作步驟

1. 設定 **Git Repo 路徑**（含 `develop` 分支的專案）
2. 設定 **輸出資料夾**
3. 確認自動帶入的 Source / Target Commit
4. 點擊「**產生上版清單**」

## 打包為 EXE

```bash
build.bat
```

產出的執行檔位於 `dist\Baonifa.exe`。

## 專案結構

```
├── go.py               # 主程式（GUI + 流程控制）
├── patch_to_excel.py   # diff patch → Excel
├── excel_to_word.py    # Excel → Word（含 OLE 附件）
├── install.py          # 一鍵安裝相依套件
├── build.bat           # PyInstaller 打包腳本
├── requirements.txt    # Python 相依套件
├── robot.png           # 應用程式圖示（PNG）
├── robot.ico           # 應用程式圖示（ICO）
└── Baonifa.spec        # PyInstaller spec 檔
```
