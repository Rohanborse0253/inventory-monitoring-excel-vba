# Excel/VBA Inventory Management System

## Overview
Macro-enabled Excel workbook to manage products, suppliers, customers, purchase orders, and sales orders. Tracks **Available Stock in Inventory**, computes line/item totals, and centralizes PO/SO headers with detailed line items—ready for alerts and reporting.

## File
- **Inventory Management System.xlsm** (enable macros on open)

## Core Sheets (from this workbook)
- **Product List** — Product No, Supplier Name, Product Name, Weight in Grams, Description, **Available Stock in Inventory**, Purchase Price, Sales Price (MRP), …
- **Supplier List** — Supplier ID, Supplier Name, Address, City, State, Zip Code, Phone, Email, …
- **Customer List** — Customer ID, Type, Customer Name, Address, Village, Zip Code, Phone, Email, …
- **Purchase Order** — Order ID, Supplier Name, Order Date, **Total Order Amount**
- **Details of Purchase Order** — Order ID, Order Date, Supplier Name, Product Name, Weight in Grams, Quantity, Expiry Date, Purchase Price, **Total Amount**, Item Row, DB Row
- **Sales Order** — Order ID, Customer Name, Order Date, **Total Order Amount**
- **Details of Sales Order** — Order ID, Order Date, Customer Name, Supplier Name, Product Name, Weight in Grams, Quantity, **Sales Price**, **Total Amount**, Item Row, DB Row
- **Dashboard** — PO/SO form layout and summary canvas

## Typical Workflow
1. Update **Product/Supplier/Customer** masters.
2. Create **Purchase Orders** and **Sales Orders**; line-level details populate in the corresponding **Details** sheets.
3. Review **Available Stock in Inventory** in *Product List*; investigate negatives/low stock for replenishment.
4. Use **Dashboard** for formatted outputs (PO/SO forms) or quick review.

## Suggested Enhancements (optional)
- Low/negative stock **alerts** via conditional formatting or a VBA routine that writes to an `Alerts` sheet.
- Reorder logic (`ROP = AvgDailyDemand * LeadTime + SafetyStock`) if demand/lead-time fields are added to Product List.
- Export routines (CSV/XLSX/PDF) for daily inventory and PO/SO summaries.

## Requirements
- Microsoft Excel (Windows recommended for full macro support)

## Repo Structure
```text
.
├── Inventory Management System.xlsm
├── data/
│   └── sample/                 # optional sample CSVs for import
├── docs/
│   └── screenshots/            # add UI screenshots here
└── README.md
