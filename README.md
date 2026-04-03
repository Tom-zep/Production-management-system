# 📦 Multi-Source Inventory Data Integrator (DX Tool)

### 🚀 Business Impact: Reduced Processing Time from 4 Hours to 10 Minutes (95.8% Reduction)

This project solves a common operational bottleneck: **manual consolidation of disparate inventory reports.** By leveraging Python, this tool automates the integration of multiple CSV exports into a single, standardized Excel form ready for physical audits.

---

## 🇯🇵 プロジェクト概要
現場で複数の在庫管理システムが乱立している際、実地棚卸のための「統一フォーム」を作成するのは非常に重労働でした。
本ツールは、各システムから出力されるバラバラなCSVをPythonで自動集約し、実査（手入力）に最適なExcelフォームを一瞬で生成します。

### ✅ 解決した課題 (Key Benefits)
* **工数削減**: 手作業でのコピペ・名寄せ作業（約4時間）を**10分以内**に短縮。
* **現場改善**: 「作業の押し付け合い」が発生していた不毛なルーチンを自動化し、チームの士気を向上。
* **柔軟な名寄せ**: システムごとに異なる列名（例：「品番」「SKU」「item_id」）をロジックで自動判別。

---

## 🛠️ Features (機能)
1.  **Multi-Encoding Support**: Handles Shift-JIS (Legacy Japanese systems) and UTF-8 automatically.
2.  **Smart Column Mapping**: Automatically identifies and extracts Item IDs, Locations, and Stock quantities even if headers differ.
3.  **Human-in-the-Loop Design**: Outputs a clean Excel file with empty columns for manual input during physical audits.

---

## 📂 Usage (使い方)
1.  Place all your exported CSV files in the `input_data/` folder.
2.  Run the Python script.
3.  An integrated Excel file (`Inventory_Audit_Form_YYYYMMDD.xlsx`) will be generated instantly.

---

## 📈 Future Roadmap
* [ ] **AI-Powered PDF OCR**: Automate data entry by reading handwritten audit sheets via AI/OCR.
* [ ] **Direct API Integration**: Connect directly to ERP systems to eliminate CSV exports.

---
