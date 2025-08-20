# OIR Report System v6 - Prototype

## 系統概述 / System Overview

OIR (Outgoing Inspection Report) 報告系統是為Johnson Electric製造部門設計的數位化檢驗報告解決方案。本系統旨在將傳統的紙本檢驗流程數位化，提高效率並減少人工錯誤。

The OIR (Outgoing Inspection Report) system is a digital inspection report solution designed for Johnson Electric's manufacturing department. This system aims to digitize traditional paper-based inspection processes, improve efficiency, and reduce manual errors.

## 主要功能 / Key Features

### 🌐 多語言支援 / Multi-language Support
- 英文 (English)
- 繁體中文 (Traditional Chinese)
- 簡體中文 (Simplified Chinese)
- 介面支援多語言，報告固定為英文格式

### 📊 核心功能 / Core Functions
1. **新增檢驗報告 / New Inspection Reports**
   - 數位化數據輸入
   - 支援多項目檢驗 (Items 1-5)
   - 每項目10個數據點輸入
   - 即時數據統計計算

2. **歷史報告查詢 / Historical Report Search**
   - 依日期範圍搜尋
   - 依型號、批號、操作員篩選
   - 快速篩選選項（今天、本週、本月）

3. **報告生成與匯出 / Report Generation & Export**
   - Excel格式報告生成
   - PDF格式匯出（開發中）
   - 基於OIR_Report_Sample.xlsx模板

4. **資料庫管理 / Database Management**
   - Excel為基礎的原型資料庫
   - OIS (Outgoing Inspection Standard) 數據管理
   - 檢驗記錄數據存儲

## 系統架構 / System Architecture

```
v6/
├── app.py                 # Flask主應用程式
├── database.py            # Excel資料庫操作模組
├── languages.py           # 多語言支援模組
├── requirements.txt       # Python依賴套件
├── README.md             # 系統說明文件
├── OIR_database.xlsx     # 主資料庫檔案（自動生成）
├── templates/            # HTML模板資料夾
│   ├── base.html
│   ├── index.html
│   ├── new_report_step1.html
│   ├── data_input.html
│   ├── confirm_data.html
│   ├── preview_report.html
│   ├── history_report.html
│   ├── history_results.html
│   └── error.html
└── static/              # 靜態檔案資料夾（CSS/JS）
```

## 安裝與設置 / Installation & Setup

### 1. 環境需求 / Requirements
- Python 3.8 或更高版本
- Windows 10/11 (開發環境)
- 支援Excel檔案讀寫

### 2. 安裝步驟 / Installation Steps

```bash
# 1. 進入專案目錄
cd "C:\Users\aaron\OneDrive\桌面\intern\JE\OIR Report\v6"

# 2. 安裝Python依賴套件
pip install -r requirements.txt

# 3. 運行應用程式
python app.py
```

### 3. 系統啟動 / System Startup
- 開啟瀏覽器訪問：http://127.0.0.1:5000
- 系統將自動創建所需的資料庫檔案

## 使用流程 / Usage Workflow

### 新增檢驗報告 / Creating New Reports

1. **步驟1：基本資訊輸入**
   - 輸入型號 (Model No.)
   - 輸入OIS編號 (OIS No.)
   - 選擇檢驗員 (Inspector: Aaron/Alan/Brain)

2. **步驟2：數據輸入**
   - 依序輸入各項目的10個數據點
   - 支援快速填入功能
   - 即時統計計算

3. **步驟3：確認數據**
   - 檢視已輸入的數據摘要
   - 輸入額外資訊（訂單號、出貨數量、批號、位置）

4. **步驟4：預覽與匯出**
   - 預覽完整報告格式
   - 下載Excel或PDF格式

### 歷史報告查詢 / Historical Report Search

1. **設定搜尋條件**
   - 日期範圍
   - 型號 (部分匹配)
   - 批號 (部分匹配)
   - 操作員

2. **檢視搜尋結果**
   - 表格顯示匹配記錄
   - 詳細資料檢視
   - 報告重新生成

## 資料庫結構 / Database Structure

### OIS工作表 (OIS Sheet)
- OIS No. - OIS編號
- Model Code - 型號代碼
- Model Desc. - 型號描述
- Model Version - 型號版本
- Item - 項目編號 (1-5)
- SC Symbol - SC符號
- Description - 描述
- Minimum Limit - 最小限值
- Maximum Limit - 最大限值
- Median - 中位數
- Unit - 單位
- A.QAL(%) of Sample Size - 樣本大小
- Type of Data - 數據類型
- Measurement Equipment - 測量設備

### Database工作表 (Database Sheet)
- Date - 日期
- Model No. - 型號
- Model Description - 型號描述
- OIS No. - OIS編號
- Lot No. - 批號
- Item - 項目
- Datapoint_1 to Datapoint_10 - 數據點1-10
- Operator - 操作員

## 技術規格 / Technical Specifications

### 後端技術 / Backend Technologies
- **Flask 2.3.3** - Web框架
- **pandas 2.1.1** - 數據處理
- **openpyxl 3.1.2** - Excel檔案操作
- **Python 3.8+** - 程式語言

### 前端技術 / Frontend Technologies
- **Bootstrap 5.1.3** - UI框架
- **jQuery 3.6.0** - JavaScript庫
- **Font Awesome 6.0.0** - 圖標庫
- **HTML5/CSS3** - 標記語言

### 系統限制 / System Limitations
- **單用戶環境** - 原型階段僅支援單一用戶
- **Excel資料庫** - 使用Excel作為資料庫（原型限制）
- **本地伺服器** - 運行於本地環境
- **無離線支援** - 需要網路連接

## 開發注意事項 / Development Notes

### 原型階段特點 / Prototype Characteristics
- 🔧 **單用戶設計** - 暫不考慮並發存取
- 📊 **Excel資料庫** - 未來可擴展至SQL資料庫
- 🌐 **本地部署** - 適合原型測試
- 📝 **基本驗證** - 最小化數據驗證
- 🚫 **無備份機制** - 原型階段不包含

### 未來擴展計劃 / Future Enhancement Plans
- 多用戶並發支援
- SQL資料庫整合
- 進階數據驗證
- 自動備份機制
- 雲端部署支援

## 故障排除 / Troubleshooting

### 常見問題 / Common Issues

**Q: 系統啟動失敗**
A: 檢查Python版本和依賴套件是否正確安裝

**Q: Excel檔案讀寫錯誤**
A: 確保Excel檔案未被其他程式開啟，檢查檔案權限

**Q: OIS編號未找到**
A: 檢查OIR_database.xlsx中OIS工作表是否包含相應數據

**Q: 語言切換無效果**
A: 清除瀏覽器快取並重新整理頁面

### 錯誤代碼 / Error Codes
- **404** - 頁面未找到
- **500** - 內部伺服器錯誤
- **檔案錯誤** - Excel檔案操作失敗

## 聯絡資訊 / Contact Information

**開發者**: Aaron (Junior Engineer)  
**部門**: Johnson Electric Manufacturing Department  
**專案**: OIR Report System Prototype  
**版本**: v6  
**更新日期**: 2025年1月

---

## 版本歷史 / Version History

### v6 (Current) - 2025年1月
- ✅ 完整多語言支援系統
- ✅ Flask Web應用架構
- ✅ Excel資料庫整合
- ✅ 響應式UI設計
- ✅ 數據輸入與驗證
- ✅ 報告預覽與匯出
- ✅ 歷史查詢功能
- ✅ **基於OIR_Report_Sample.xlsx的精確報告格式**
- ✅ **自動匯入Standards數據**
- 🔄 PDF匯出功能（開發中）

### 開發狀態 / Development Status
- 🟢 **核心功能** - 已完成
- 🟡 **PDF匯出** - 開發中
- 🔴 **多用戶支援** - 未開始
- 🔴 **雲端部署** - 未開始

---

**注意**: 本系統為原型版本，僅供內部測試使用。生產環境部署需要進一步的安全性和效能優化。

