# Construction Estimate Template Excel | Contractor Bid & Job Costing Spreadsheet, Markup Margin Calculator, Proposal Generator Google Sheets

**Author:** `novality store`  
**Formula Protection Password:** `premium`  
**File Name:** `construction_estimate_template.xlsx`  
**Compatibility:** Microsoft Excel (.xlsx) & Google Sheets (no macros required)

---

## 📥 Direct Download Links (GitHub)

If you are unable to download the file directly through the file viewer, you can download it immediately from the public GitHub repository for this session:

- **[Download Complete Excel (.xlsx) Template (Direct Raw Download)](https://raw.githubusercontent.com/arbabkhan007/Excel/arena/019fb67f-excel/construction_estimate_template.xlsx)**  
  *(Clicking this link downloads `construction_estimate_template.xlsx` directly to your computer)*
- **[View on GitHub (with Download Button)](https://github.com/arbabkhan007/Excel/blob/arena/019fb67f-excel/construction_estimate_template.xlsx)**  
  *(On this page, click the download/raw icon on the right side of the file header)*
- **[Browse Individual Google Sheets CSV Import Files](https://github.com/arbabkhan007/Excel/tree/arena/019fb67f-excel/google_sheets_import)**  
  *(All 19 sheets exported as clean `.csv` files for individual sheet import)*

---

## 1. How to Use in Google Sheets

This spreadsheet template is 100% compatible with Google Sheets and requires **no macros or add-ons**.

### Step 1: Upload to Google Drive
1. Go to [Google Drive](https://drive.google.com).
2. Click **New** → **File upload** (or drag and drop `construction_estimate_template.xlsx` into your Drive folder).
3. Once uploaded, double-click the file and select **Open with Google Sheets** at the top of the preview window.
4. Google Sheets will automatically convert the file while preserving all 19 tabs, cell formatting, uppercase formulas, professional fonts (`Segoe UI`), and visual charts.

### Step 2: Working with Protected Cells in Google Sheets
- In Excel, all formula and header cells are locked and protected with the password **`premium`**.
- In Google Sheets, worksheet protection is managed through Google Account permissions:
  - If you want to restrict editing of formula cells when sharing with team members or clients, click **Data → Protect sheets and ranges** in Google Sheets and set permissions to "Only you".
  - Editable user input cells (white background with gray border) remain freely editable by default.

---

## 2. How to Use in Microsoft Excel

1. Open `construction_estimate_template.xlsx` directly in Microsoft Excel (Windows, macOS, or Excel Online).
2. All 19 sheets are ready for use immediately.
3. If you ever need to edit a formula cell or table header, go to **Review → Unprotect Sheet** and enter the password: **`premium`**.

---

## 3. Sheet Tab Reference (19 Sheets Total)

### Essential & Recommended Tabs
1. **Instructions**: Full user guide, color legend, and step-by-step workflow instructions.
2. **Project Info**: Global setup sheet for Project Name, Client Info, Contractor Details, and global percentage rates (Sales Tax, Discount, Overhead, Contingency, and default markups).
3. **Estimate**: Master Construction Estimate & Contractor Bid Template compiling line items from Materials, Labor, Equipment, and Subcontractors, applying Overhead and Contingency allowances, and computing the **Grand Total Contract Bid Price**.
4. **Materials**: Itemized Materials Cost Sheet with SKU, material description, supplier/vendor, quantity, unit cost, and tax/freight allowance.
5. **Labor**: Itemized Labor Cost Sheet tracking craft roles, crew size, regular hours/wages, and overtime hours/rates.
6. **Equipment**: Itemized Equipment Rental & Special Tooling sheet tracking daily/weekly rates, usage, and delivery fees.
7. **Subcontractors**: Licensed Subcontractor package quote sheet with contractor markup amounts.
8. **Markup & Margin Calculator**:
   - Separate markups for Materials, Labor, Equipment, and Subcontractors.
   - **Markup vs. Profit Margin Comparison Table** showing equivalent margin percentages.
   - **Target Profit Margin Reverse Calculator** to compute the required selling price for any desired profit margin %.
9. **Job Costing & Actuals**: Field expenditure tracking sheet comparing Estimated vs. Actual costs with variance percentages and status indicators (`UNDER BUDGET` / `OVER BUDGET`).
10. **Summary Dashboard**: Executive KPI cards, contract summary table, and a cell-based **visual dashboard** comparing Estimated Budget vs. Actual Incurred Costs by Category (no drawings, 100%% Excel-safe).
11. **Client Proposal**: Printable, A4/Letter invoice-style proposal page with scope of work, pricing summary, and formal client acceptance signature block.
12. **Terms & Conditions**: Standard contractor legal agreement covering validity, payment schedule, change orders, site access, and 1-year workmanship warranty.

### Optional Advanced Feature Tabs
13. **Material Price DB**: Master material inventory reference database with SKUs, suppliers, and standard prices.
14. **Labor Rate DB**: Master trade wage database calculating fully burdened labor costs and standard billing rates.
15. **Estimate Versions**: Side-by-side proposal comparison table for Good / Better / Best packages (**Option 1: Standard / Base**, **Option 2: Upgraded / Preferred**, **Option 3: Premium / Luxury**).
16. **Change Order Tracker**: Log for tracking project scope additions/deductions with markup calculation, schedule impact, and client approval status.
17. **Payment Schedule**: Milestone payment schedule (20% deposit, rough-in, drywall, finishes, punchlist) linking to the Contract Total.
18. **Invoice Summary**: AIA-style progress billing invoice generator linking contract amount, change orders, work completed, and retainage withheld.
19. **Project Timeline**: 12-week Gantt-style schedule tracker with phase status and visual timeline bars.

---

## 4. Regenerating or Customizing the Spreadsheet

The complete Python source code that generates this spreadsheet is located at:
- **`examples/construction_estimate_template_generator.py`**

To generate a new copy of the template at any time, run:
```bash
PYTHONPATH=. python3 examples/construction_estimate_template_generator.py construction_estimate_template.xlsx
```
