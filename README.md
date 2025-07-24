# 📊 AI-Powered Sales Analysis & Reporting Tool

This project is a **Python-based AI-driven sales reporting system** that automates the entire flow from data ingestion to Excel report generation and visualization.

With integrated **K-Means customer segmentation**, **anomaly detection**, and **interactive Excel output**, it empowers businesses to better understand customer behavior and revenue trends—all in just a few lines of code.

---

## 🚀 Features

- ✅ Merges multiple datasets: `customers`, `orders`, `products`
- 🤖 Applies AI techniques:
  - Customer segmentation using **KMeans**
  - Order value anomaly detection using **Z-score**
- 📈 Generates a **3-sheet Excel report**:
  - `Transaction Data`
  - `Summary` (with KPIs like total revenue, top product)
  - `Charts` (bar chart of revenue by customer segment)
- 📤 Automatically opens the report in Microsoft Excel

---

## 📂 Folder Structure
📦 ai_sales_report
├── 📁 data                   # Input CSV files
│   ├── customers.csv
│   ├── orders.csv
│   └── products.csv
│
├── 📁 reports                # Auto-generated Excel reports
│   └── sales_report.xlsx     # (generated after script run)
│
├── 📁 utils                  # Optional: helper functions (modular code)
│   ├── __init__.py
│   └── db_handler.py         # (if needed to load from DB or preprocess)
│
├── main.py                   # Main script to run the pipeline
├── requirements.txt          # Dependencies to install
└── README.md                 # Project overview and instructions


