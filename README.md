# Inventory and Stock Management System

A desktop inventory application built with **Python, Tkinter and Excel (XLSX)**.

## Features
- Add, update and delete products
- Product ID, name, category, quantity, price, supplier and reorder level
- Automatic stock value calculation
- Low-stock status detection
- Dashboard: number of products, total stock, low-stock products and inventory value
- Live search
- Excel file storage (`inventory.xlsx`)
- Open Excel directly from the application

## Requirements
- Python 3.9+
- Tkinter (usually included with standard Windows Python)
- openpyxl

## Installation
1. Open a terminal in this project folder.
2. Run:
   `pip install -r requirements.txt`
3. Run:
   `python main.py`

The Excel database is created automatically as `inventory.xlsx`.

## Suggested project structure
inventory_stock_management_system/
- main.py
- excel_manager.py
- requirements.txt
- README.md
- inventory.xlsx (created after first run)

## Future enhancements
- Stock IN / Stock OUT transactions
- Sales and purchase modules
- Supplier management
- Login/admin authentication
- Monthly stock reports
- Excel export/report generation
- Barcode support
- Charts and analytics
