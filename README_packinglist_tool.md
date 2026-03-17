# Packing List Converter（Diptyque PDF → Excel）

此工具可將 Diptyque packing list PDF 轉成 `.xlsx`，並輸出固定欄位：

- Reference
- Item Name
- Batch Number
- ELD / DLP
- Quantity
- Net Weight
- Alcohol Vol

## 已實作規則

1. **不輸出 barcode**（例如 13 碼 EAN）
2. **不輸出表頭 / summary / shipment info**
3. **不輸出 pallet / colis header**（SSCC、包材、尺寸、毛重等）
4. **每一行保留，不合併同品項**
5. **Excel 檔名自動使用 Order Ref**（如 `SO261549-X79251.xlsx`）

## 環境需求

- Python 3.9+
- 本工具目前使用 Python 標準函式庫，不需額外套件。

> 若你要改回 `pdfplumber / pandas` 版本，再自行加回對應 requirements。

## 專案結構

- `packinglist_to_excel.py`：主程式
- `sample_pdfs/`：測試 PDF
- `output/`：輸出 Excel（執行後自動建立）

## 使用方式

### 1) 批次轉換整個資料夾

```bash
python packinglist_to_excel.py sample_pdfs -o output
```

### 2) 轉換單一 PDF

```bash
python packinglist_to_excel.py "sample_pdfs/pla_SO261549-X79251.pdf" -o output
```

## 輸出說明

- 每個 PDF 會輸出一個 Excel
- 檔名格式：`<OrderRef>.xlsx`
- 例如：`output/SO262498-X79255.xlsx`

## 常見問題

### Q1: 為什麼有些列的 Batch Number 或 ELD / DLP 是空白？
A: 原始 PDF 該列若沒有提供批號或效期，工具會保留空白，不會猜測或合併其他列資料。

### Q2: Quantity、Net Weight、Alcohol Vol 會被合併計算嗎？
A: 不會。工具以「逐行輸出」為原則，保留原始列。
