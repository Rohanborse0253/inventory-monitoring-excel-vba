# Excel/VBA Inventory Management System

## Overview
Macro-enabled Excel workbook that monitors stock for 400+ SKUs, calculates reorder points, flags phantom inventory, and triggers **automatic stock alerts** via VBA with conditional formatting . Includes minimum stock thresholds, lead-time based safety stock, and one-click daily report export.

## File
- **Inventory Management System.xlsm** (macro-enabled; open and click **Enable Content**)

## Features
-  **Live KPIs:** on-hand, reserved, in-transit, days of supply
-  **Auto alerts:** low/min stock, negative availability, phantom inventory variance
-  **Reorder logic:** `ROP = AvgDailyDemand * LeadTime + SafetyStock`
-  **Daily report:** one-click export (CSV/XLSX/PDF)

## How to use
1. **Open** the `.xlsm` → click **Enable Content** (Trust Center will prompt).
2. Go to **Settings** and set: LeadTimeDays default, z-factor (e.g., 1.65), alert email on/off, export folder.
3. Import your latest inventory CSV into **Data** (or click **Refresh** if button provided).
4. Click **Recalculate** → review **Inventory** and **Alerts**.
5. (Optional) Click **Send Emails** to notify stakeholders.
6. Click **Export Report** to save the daily summary.

> **Excel Macro settings:** File → Options → Trust Center → Trust Center Settings → Macro Settings → *Disable with notification* (recommended). Then use **Enable Content** per file.
