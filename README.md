# QC 自動化工具

自動替換 Word 文件中的內嵌圖片與日期，並支援匯出 PDF，適用於查驗照片報告的批次更新作業。

手動 v.s 自動 
![介面截圖](https://github.com/user-attachments/assets/83a709cb-beb0-4097-8ca3-7411755edbdb)

demo:在僅需替換 6 張圖片的輕度作業中，處理時間由 78s 縮減至 23s，節省約 70.5% 的操作耗時。

由於自動化工具主要耗時在於「環境初始化」與「Word 轉檔引擎載入」，而非圖片替換動作本身，意即在圖片數量越多的作業中，邊際效益將隨之遞增，預計可節省高達 90% 以上 的重複性工時。

---

## 系統需求

| 項目 | 需求 |
|---|---|
| 作業系統 | Windows 10 / 11（64 位元） |
| Microsoft Word | 需已安裝（2016 以上建議），供圖片替換與 PDF 匯出使用 |
| .NET Runtime | **不需要**（exe 已內含完整 runtime） |

---

## 使用方式

直接執行 `WordAutoTool.exe`，不需要安裝，也不需要附帶任何其他檔案。

開啟後會出現瀏覽器介面，依序操作以下步驟：

### 步驟 1 — 選擇 Word 檔案

支援 `.doc`（Word 97-2003）與 `.docx` 格式。
選擇後會自動掃描文件內容（浮動圖形、內嵌圖片、文字方塊、段落文字），顯示於「文件內容掃描」區塊。

### 步驟 2 — 選擇新圖片資料夾

選擇包含新圖片的資料夾。資料夾內的圖片會依**檔名排序**，依序取代文件中的**內嵌圖片**，並保留原圖片的位置與尺寸。

支援格式：`.png` `.jpg` `.jpeg` `.gif` `.bmp` `.tiff` `.webp`

> 圖片數量多於文件內嵌圖片時，多餘的圖片略過；數量少時，只替換前 N 張。

### 步驟 3 — 選擇日期

從日期選擇器選取日期，右側即時預覽轉換後的民國日期（如 `民國 115.3.5`）。

**個位數月/日前補零**（預設不勾選）：
- 不勾選：`115.3.5`
- 勾選：`115.03.05`

日期替換範圍：
- 文件內所有**文字方塊**（整個文字方塊內容取代為新日期）
- 段落文字中符合 `日期：XXX` 或 `日期:XXX` 格式的文字

### 步驟 4 — 輸出選項

**同時輸出 PDF**（預設勾選）：
- **勾選**：將 `.doc` 與 `.pdf` 打包成一個 `.zip` 下載
- **不勾選**：只下載 `.doc`

---

## 輸出檔案

| 情況 | 輸出 |
|---|---|
| 不含 PDF | `8_查驗照片MMDD.doc` |
| 含 PDF | `8_查驗照片MMDD.zip`（內含 `.doc` + `.pdf`） |

月日格式（`MMDD`）固定補零，如 3 月 5 日為 `0305`。

---

## 處理流程

```
上傳 .doc/.docx
      │
      ▼ (若為 .doc)
  .doc → .docx 轉換（Word COM）
      │
      ▼
  替換內嵌圖片（Word COM，保留尺寸位置）
      │
      ▼
  替換文字方塊與段落日期（OpenXML）
      │
      ▼
  .docx → .doc 轉換（Word COM）
      │
      ▼ (若勾選 PDF)
  .doc → .pdf 轉換（Word COM）
  打包 .zip
      │
      ▼
  下載結果
```

---

## 開發與建置

### 技術架構

- **前端**：HTML / CSS / JavaScript（嵌入 exe，透過 WebView2 呈現）
- **後端**：ASP.NET Core 9 + Kestrel（內嵌於 WinForms 應用程式中）
- **Word 操作**：Word COM Automation（圖片替換、格式轉換、PDF 匯出）
- **OpenXML**：DocumentFormat.OpenXml（文字替換）

### 建置需求

- [.NET 9 SDK](https://dotnet.microsoft.com/download/dotnet/9.0)
- Windows（使用 Word COM，需 Windows 環境）

### 建置指令

```bash
# 偵錯執行
dotnet run --project WordAutoTool

# 發布成單一 exe
dotnet publish WordAutoTool -c Release -r win-x64
# 輸出：WordAutoTool/bin/Release/net9.0-windows/win-x64/publish/WordAutoTool.exe
```
