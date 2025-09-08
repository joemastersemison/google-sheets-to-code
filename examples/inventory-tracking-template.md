# Inventory Tracking Google Sheets Template

## Overview
This template demonstrates inventory management with reorder points, stock calculations, ABC classification, and automated alerts.

## Google Sheets Setup Instructions

### Step 1: Create the Spreadsheet
1. Go to [Google Sheets](https://sheets.google.com)
2. Create a new spreadsheet
3. Name it "Inventory Tracking Example"
4. Create the following sheets (tabs):
   - Products
   - Transactions
   - Settings
   - CurrentInventory
   - Alerts
   - Reports

### Step 2: Set up "Products" Sheet

| Cell | A (SKU) | B (Name) | C (Unit Cost) | D (Lead Time) | E (Safety Stock) | F (Annual Demand) | G (Supplier) | H (Category) |
|------|---------|----------|---------------|---------------|------------------|-------------------|--------------|--------------|
| 1 | SKU | Product Name | Unit Cost | Lead Time (days) | Safety Stock | Annual Demand | Supplier | Category |
| 2 | SKU001 | Widget Alpha | 25.50 | 7 | 50 | 2400 | Supplier A | Electronics |
| 3 | SKU002 | Widget Beta | 45.00 | 14 | 30 | 1800 | Supplier B | Electronics |
| 4 | SKU003 | Widget Gamma | 12.75 | 3 | 100 | 5000 | Supplier A | Accessories |
| 5 | SKU004 | Widget Delta | 89.99 | 21 | 20 | 600 | Supplier C | Premium |
| 6 | SKU005 | Widget Epsilon | 5.25 | 5 | 200 | 10000 | Supplier B | Consumables |
| 7 | SKU006 | Widget Zeta | 150.00 | 30 | 10 | 200 | Supplier D | Premium |
| 8 | SKU007 | Widget Eta | 35.00 | 10 | 40 | 1500 | Supplier A | Electronics |
| 9 | SKU008 | Widget Theta | 18.50 | 7 | 75 | 3000 | Supplier C | Accessories |
| 10 | SKU009 | Widget Iota | 62.00 | 14 | 25 | 900 | Supplier B | Electronics |
| 11 | SKU010 | Widget Kappa | 8.99 | 3 | 150 | 7500 | Supplier A | Consumables |

### Step 3: Set up "Transactions" Sheet

| Cell | A (Date) | B (SKU) | C (Type) | D (Quantity) | E (Unit Price) | F (Reference) | G (Notes) |
|------|----------|---------|----------|--------------|----------------|---------------|-----------|
| 1 | Date | Product SKU | Transaction Type | Quantity | Unit Price | Reference Number | Notes |
| 2 | 1/1/2024 | SKU001 | IN | 200 | 25.50 | PO-2024-001 | Initial stock |
| 3 | 1/1/2024 | SKU002 | IN | 150 | 45.00 | PO-2024-001 | Initial stock |
| 4 | 1/1/2024 | SKU003 | IN | 500 | 12.75 | PO-2024-001 | Initial stock |
| 5 | 1/2/2024 | SKU001 | OUT | 15 | 35.00 | SO-2024-001 | Customer order |
| 6 | 1/3/2024 | SKU002 | OUT | 8 | 65.00 | SO-2024-002 | Customer order |
| 7 | 1/4/2024 | SKU003 | OUT | 45 | 18.00 | SO-2024-003 | Customer order |
| 8 | 1/5/2024 | SKU001 | OUT | 20 | 35.00 | SO-2024-004 | Customer order |
| 9 | 1/6/2024 | SKU004 | IN | 50 | 89.99 | PO-2024-002 | Restock |
| 10 | 1/7/2024 | SKU005 | IN | 1000 | 5.25 | PO-2024-002 | Restock |
| 11 | 1/8/2024 | SKU001 | OUT | 12 | 35.00 | SO-2024-005 | Customer order |
| 12 | 1/9/2024 | SKU003 | OUT | 60 | 18.00 | SO-2024-006 | Customer order |
| 13 | 1/10/2024 | SKU002 | OUT | 10 | 65.00 | SO-2024-007 | Customer order |
| 14 | 1/11/2024 | SKU005 | OUT | 150 | 8.00 | SO-2024-008 | Bulk order |
| 15 | 1/12/2024 | SKU001 | IN | 100 | 25.50 | PO-2024-003 | Restock |
| 16 | 1/13/2024 | SKU004 | OUT | 5 | 125.00 | SO-2024-009 | Customer order |
| 17 | 1/14/2024 | SKU006 | IN | 20 | 150.00 | PO-2024-004 | Initial stock |
| 18 | 1/15/2024 | SKU007 | IN | 80 | 35.00 | PO-2024-004 | Initial stock |
| 19 | 1/16/2024 | SKU003 | OUT | 75 | 18.00 | SO-2024-010 | Large order |
| 20 | 1/17/2024 | SKU001 | OUT | 18 | 35.00 | SO-2024-011 | Customer order |

*Continue with more transaction history*

### Step 4: Set up "Settings" Sheet

| Cell | Label | Value | Description |
|------|-------|-------|-------------|
| A1 | **Inventory Settings** | | Header |
| A2 | Lead Time Multiplier | | Label |
| B2 | 1.5 | | Safety factor for lead time |
| A3 | Ordering Cost per Order | | Label |
| B3 | 50 | | Fixed cost per order |
| A4 | Holding Cost Rate (annual %) | | Label |
| B4 | 0.20 | | 20% annual holding cost |
| A5 | Service Level Target | | Label |
| B5 | 0.95 | | 95% service level |
| A6 | Review Period (days) | | Label |
| B6 | 7 | | Weekly review |
| A7 | Min Order Quantity | | Label |
| B7 | 10 | | Minimum order size |
| A8 | Max Order Quantity | | Label |
| B8 | 1000 | | Maximum order size |
| A9 | Critical Stock Level (%) | | Label |
| B9 | 0.25 | | 25% of reorder point |
| A10 | Overstock Threshold (days) | | Label |
| B10 | 90 | | 90 days of supply |

### Step 5: Set up "CurrentInventory" Sheet

| Cell | Label | Formula | Description |
|------|-------|---------|-------------|
| A1 | SKU | Current Stock | Reorder Point | Status | Stock Value | Days of Supply | Turnover | ABC Class | Risk | EOQ | Avg Daily Usage | Coverage |
| A2 | SKU001 | | | | | | | | | | | |
| B2 | | `=SUMIFS(Transactions!D:D, Transactions!B:B, A2, Transactions!C:C, "IN") - SUMIFS(Transactions!D:D, Transactions!B:B, A2, Transactions!C:C, "OUT")` | Current stock level |
| C2 | | `=VLOOKUP(A2, Products!A:E, 4, FALSE) * Settings!B2 * (VLOOKUP(A2, Products!A:F, 6, FALSE)/365) + VLOOKUP(A2, Products!A:E, 5, FALSE)` | Reorder point |
| D2 | | `=IF(B2<=C2*Settings!B9, "CRITICAL", IF(B2<=C2, "REORDER", "OK"))` | Stock status |
| E2 | | `=B2 * VLOOKUP(A2, Products!A:C, 3, FALSE)` | Inventory value |
| F2 | | `=IF(SUMIFS(Transactions!D:D, Transactions!B:B, A2, Transactions!C:C, "OUT")/30>0, B2/(SUMIFS(Transactions!D:D, Transactions!B:B, A2, Transactions!C:C, "OUT")/30), 999)` | Days of supply |
| G2 | | `=IF(B2>0, SUMIFS(Transactions!D:D, Transactions!B:B, A2, Transactions!C:C, "OUT")/B2*365/30, 0)` | Turnover rate |
| H2 | | `=IF(E2/SUM(E:E)>0.7, "A", IF(E2/SUM(E:E)>0.2, "B", "C"))` | ABC classification |
| I2 | | `=IF(B2<C2*0.5, "HIGH", IF(B2<C2, "MEDIUM", "LOW"))` | Stockout risk |
| J2 | | `=SQRT(2*VLOOKUP(A2, Products!A:F, 6, FALSE)*Settings!B3/(VLOOKUP(A2, Products!A:C, 3, FALSE)*Settings!B4))` | Economic order qty |
| K2 | | `=SUMIFS(Transactions!D:D, Transactions!B:B, A2, Transactions!C:C, "OUT")/30` | Average daily usage |
| L2 | | `=IF(K2>0, B2/K2, 999)` | Stock coverage days |

*Repeat formulas for SKU002 through SKU010 in rows 3-11*

### Step 6: Set up "Alerts" Sheet

| Cell | Label | Formula | Description |
|------|-------|---------|-------------|
| A1 | Alert Type | Product SKU | Message | Priority | Action Required | Date |
| A2 | Stock Level | | | | | |
| B2 | | `=IF(COUNTIF(CurrentInventory!D:D, "CRITICAL")>0, INDEX(CurrentInventory!A:A, MATCH("CRITICAL", CurrentInventory!D:D, 0)), "")` | SKU with critical stock |
| C2 | | `=IF(B2<>"", "Critical stock level - immediate reorder required", "")` | Alert message |
| D2 | | `=IF(B2<>"", "HIGH", "")` | Priority level |
| E2 | | `=IF(B2<>"", CONCATENATE("Order ", VLOOKUP(B2, CurrentInventory!A:J, 10, FALSE), " units"), "")` | Action |
| F2 | | `=IF(B2<>"", TODAY(), "")` | Alert date |
| A3 | Reorder Point | | | | | |
| B3 | | `=IF(COUNTIF(CurrentInventory!D:D, "REORDER")>0, INDEX(CurrentInventory!A:A, MATCH("REORDER", CurrentInventory!D:D, 0)), "")` | SKU at reorder point |
| C3 | | `=IF(B3<>"", "Stock at reorder point - place order", "")` | Alert message |
| D3 | | `=IF(B3<>"", "MEDIUM", "")` | Priority level |
| E3 | | `=IF(B3<>"", CONCATENATE("Order ", VLOOKUP(B3, CurrentInventory!A:J, 10, FALSE), " units"), "")` | Action |
| F3 | | `=IF(B3<>"", TODAY(), "")` | Alert date |
| A4 | Overstock | | | | | |
| B4 | | `=IF(COUNTIF(CurrentInventory!F:F, ">="&Settings!B10)>0, INDEX(CurrentInventory!A:A, MATCH(TRUE, CurrentInventory!F:F>=Settings!B10, 0)), "")` | Overstocked SKU |
| C4 | | `=IF(B4<>"", "Excess inventory - review for reduction", "")` | Alert message |
| D4 | | `=IF(B4<>"", "LOW", "")` | Priority level |
| E4 | | `=IF(B4<>"", "Review inventory levels", "")` | Action |
| F4 | | `=IF(B4<>"", TODAY(), "")` | Alert date |
| A5 | Slow Moving | | | | | |
| B5 | | `=IF(COUNTIF(CurrentInventory!G:G, "<2")>0, INDEX(CurrentInventory!A:A, MATCH(TRUE, CurrentInventory!G:G<2, 0)), "")` | Slow moving SKU |
| C5 | | `=IF(B5<>"", "Low turnover rate - consider promotion", "")` | Alert message |
| D5 | | `=IF(B5<>"", "LOW", "")` | Priority level |
| E5 | | `=IF(B5<>"", "Review pricing/promotion strategy", "")` | Action |
| F5 | | `=IF(B5<>"", TODAY(), "")` | Alert date |

### Step 7: Set up "Reports" Sheet

| Cell | Label | Formula | Description |
|------|-------|---------|-------------|
| A1 | **Inventory Dashboard** | | Header |
| A2 | Total Inventory Value | | Label |
| B2 | `=SUM(CurrentInventory!E:E)` | Total value |
| A3 | Number of SKUs | | Label |
| B3 | `=COUNTA(CurrentInventory!A:A)-1` | SKU count |
| A4 | Items Below Reorder Point | | Label |
| B4 | `=COUNTIF(CurrentInventory!D:D, "REORDER") + COUNTIF(CurrentInventory!D:D, "CRITICAL")` | Count below ROP |
| A5 | Items Out of Stock | | Label |
| B5 | `=COUNTIF(CurrentInventory!B:B, 0)` | Out of stock count |
| A6 | Critical Stock Items | | Label |
| B6 | `=COUNTIF(CurrentInventory!D:D, "CRITICAL")` | Critical items |
| | | | |
| A8 | **Performance Metrics** | | Header |
| A9 | Average Turnover Rate | | Label |
| B9 | `=AVERAGE(CurrentInventory!G:G)` | Avg turnover |
| A10 | Fill Rate (%) | | Label |
| B10 | `=(1-B5/B3)*100` | Fill rate percentage |
| A11 | Service Level (%) | | Label |
| B11 | `=COUNTIF(CurrentInventory!D:D, "OK")/B3*100` | Service level |
| A12 | Stock Accuracy (%) | | Label |
| B12 | `=95` | Placeholder for physical count accuracy |
| | | | |
| A14 | **ABC Analysis Summary** | | Header |
| A15 | Class | Count | Value | % of Total |
| A16 | A | `=COUNTIF(CurrentInventory!H:H, "A")` | `=SUMIF(CurrentInventory!H:H, "A", CurrentInventory!E:E)` | `=C16/B2*100` |
| A17 | B | `=COUNTIF(CurrentInventory!H:H, "B")` | `=SUMIF(CurrentInventory!H:H, "B", CurrentInventory!E:E)` | `=C17/B2*100` |
| A18 | C | `=COUNTIF(CurrentInventory!H:H, "C")` | `=SUMIF(CurrentInventory!H:H, "C", CurrentInventory!E:E)` | `=C18/B2*100` |
| | | | |
| A20 | **Top Moving Items** | | Header |
| A21 | Rank | SKU | Turnover Rate | Stock Status |
| A22 | 1 | `=INDEX(CurrentInventory!A:A, MATCH(LARGE(CurrentInventory!G:G, 1), CurrentInventory!G:G, 0))` | `=LARGE(CurrentInventory!G:G, 1)` | `=VLOOKUP(B22, CurrentInventory!A:D, 4, FALSE)` |
| A23 | 2 | `=INDEX(CurrentInventory!A:A, MATCH(LARGE(CurrentInventory!G:G, 2), CurrentInventory!G:G, 0))` | `=LARGE(CurrentInventory!G:G, 2)` | `=VLOOKUP(B23, CurrentInventory!A:D, 4, FALSE)` |
| A24 | 3 | `=INDEX(CurrentInventory!A:A, MATCH(LARGE(CurrentInventory!G:G, 3), CurrentInventory!G:G, 0))` | `=LARGE(CurrentInventory!G:G, 3)` | `=VLOOKUP(B24, CurrentInventory!A:D, 4, FALSE)` |
| | | | |
| A26 | **Order Recommendations** | | Header |
| A27 | SKU | Current Stock | Reorder Point | Suggested Order | Supplier |
| A28 | | `=IF(COUNTIF(CurrentInventory!D:D, "REORDER")>0, INDEX(CurrentInventory!A:A, MATCH("REORDER", CurrentInventory!D:D, 0)), "")` | | | |
| B28 | | `=IF(A28<>"", VLOOKUP(A28, CurrentInventory!A:B, 2, FALSE), "")` | | | |
| C28 | | `=IF(A28<>"", VLOOKUP(A28, CurrentInventory!A:C, 3, FALSE), "")` | | | |
| D28 | | `=IF(A28<>"", VLOOKUP(A28, CurrentInventory!A:J, 10, FALSE), "")` | | | |
| E28 | | `=IF(A28<>"", VLOOKUP(A28, Products!A:G, 7, FALSE), "")` | | | |

## How to Use This Template

1. **Create the Google Sheet** with the structure above
2. **Add transaction history** to simulate real inventory movements
3. **Adjust settings** based on your business requirements
4. **Share the sheet** (read-only) to get a public URL
5. **Update the example JSON** file with your sheet's URL
6. **Run the converter**:
   ```bash
   npm run cli -- convert --config examples/inventory-tracking.json
   ```

## Expected Output

The converter will generate TypeScript or Python code that:
- Calculates current stock levels from transaction history
- Determines reorder points based on lead time and safety stock
- Implements ABC classification for inventory prioritization
- Generates alerts for critical stock levels
- Calculates key metrics (turnover, fill rate, days of supply)
- Provides order recommendations with EOQ calculations

## Key Features Demonstrated

1. **Inventory Calculations**: Stock levels, reorder points, EOQ
2. **Conditional Logic**: IF statements for status and alerts
3. **Lookup Functions**: VLOOKUP for product information
4. **Aggregation**: SUMIF, COUNTIF for transaction analysis
5. **Advanced Analysis**: ABC classification, turnover rates
6. **Cross-sheet Dependencies**: Complex references across 6 sheets

## Sample Public Sheet URL

Once you create your sheet, it will have a URL like:
```
https://docs.google.com/spreadsheets/d/1PqR789StU012VwX345YzA/edit#gid=0
```

Make sure to:
1. Click "Share" → "Anyone with the link can view"
2. Copy the URL
3. Update the `inventory-tracking.json` file with this URL