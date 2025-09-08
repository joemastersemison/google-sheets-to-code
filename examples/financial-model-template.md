# Financial Model Google Sheets Template

## Overview
This template demonstrates financial calculations using PMT, FV, PV, RATE, NPV, and IRR functions.

## Google Sheets Setup Instructions

### Step 1: Create the Spreadsheet
1. Go to [Google Sheets](https://sheets.google.com)
2. Create a new spreadsheet
3. Name it "Financial Model Example"
4. Create the following sheets (tabs):
   - LoanParameters
   - CashFlows
   - LoanCalculations
   - InvestmentAnalysis

### Step 2: Set up "LoanParameters" Sheet

| Cell | Label | Value | Description |
|------|-------|-------|-------------|
| A1 | **Loan Parameters** | | Header |
| A2 | Annual Interest Rate | | Label |
| B2 | 0.05 | | 5% annual rate |
| A3 | Loan Term (Years) | | Label |
| B3 | 30 | | 30-year mortgage |
| A4 | Loan Principal | | Label |
| B4 | 500000 | | $500,000 loan |
| | | | |
| A6 | **Investment Parameters** | | Header |
| A7 | Investment Annual Rate | | Label |
| B7 | 0.08 | | 8% annual return |
| A8 | Investment Period (Years) | | Label |
| B8 | 20 | | 20-year investment |
| A9 | Monthly Contribution | | Label |
| B9 | 1000 | | $1,000/month |
| A10 | Initial Investment | | Label |
| B10 | 10000 | | $10,000 initial |
| | | | |
| A12 | **NPV/IRR Parameters** | | Header |
| A13 | Discount Rate | | Label |
| B13 | 0.10 | | 10% discount rate |
| A14 | Number of Periods | | Label |
| B14 | 10 | | 10 periods |
| A15 | Payment Amount | | Label |
| B15 | 5000 | | $5,000 payment |
| A16 | Present Value | | Label |
| B16 | 40000 | | $40,000 PV |
| A17 | Future Value | | Label |
| B17 | 0 | | $0 FV |

### Step 3: Set up "CashFlows" Sheet

| Cell | Label | Value | Description |
|------|-------|-------|-------------|
| A1 | Period | Cash Flow (NPV) | Cash Flow (IRR) |
| A2 | 0 | | -100000 |
| B2 | | 0 | Initial investment for IRR |
| C2 | | -100000 | |
| A3 | 1 | | |
| B3 | | 15000 | Year 1 cash flow |
| C3 | | 15000 | |
| A4 | 2 | | |
| B4 | | 20000 | Year 2 cash flow |
| C4 | | 20000 | |
| A5 | 3 | | |
| B5 | | 25000 | Year 3 cash flow |
| C5 | | 25000 | |
| A6 | 4 | | |
| B6 | | 30000 | Year 4 cash flow |
| C6 | | 30000 | |
| A7 | 5 | | |
| B7 | | 35000 | Year 5 cash flow |
| C7 | | 35000 | |
| A8 | 6 | | |
| B8 | | 40000 | Year 6 cash flow |
| C8 | | 40000 | |
| A9 | 7 | | |
| B9 | | 45000 | Year 7 cash flow |
| C9 | | 45000 | |
| A10 | 8 | | |
| B10 | | 50000 | Year 8 cash flow |
| C10 | | 50000 | |
| A11 | 9 | | |
| B11 | | 55000 | Year 9 cash flow |
| C11 | | 55000 | |
| A12 | 10 | | |
| B12 | | 60000 | Year 10 cash flow |
| C12 | | 60000 | |

### Step 4: Set up "LoanCalculations" Sheet

| Cell | Label | Formula | Description |
|------|-------|---------|-------------|
| A1 | **Loan Analysis** | | Header |
| A2 | Monthly Payment | | Label |
| B2 | `=PMT(LoanParameters!B2/12, LoanParameters!B3*12, -LoanParameters!B4)` | Monthly payment calculation |
| A3 | Total Amount Paid | | Label |
| B3 | `=B2*LoanParameters!B3*12` | Total over loan term |
| A4 | Total Interest Paid | | Label |
| B4 | `=B3-LoanParameters!B4` | Interest portion |
| A5 | Effective Annual Rate | | Label |
| B5 | `=RATE(LoanParameters!B14, -LoanParameters!B15, LoanParameters!B16, -LoanParameters!B17)*12` | Effective rate calculation |
| | | | |
| A7 | **Amortization Schedule** | | Header |
| A8 | Month | Principal | Interest | Balance |
| A9 | 1 | `=B2-D9` | `=LoanParameters!B4*LoanParameters!B2/12` | `=LoanParameters!B4-C9` |
| A10 | 2 | `=B2-D10` | `=E9*LoanParameters!B2/12` | `=E9-C10` |
| | | | |
| A15 | **Loan Comparison** | | Header |
| A16 | Scenario | Rate | Payment | Total Interest |
| A17 | Current | `=LoanParameters!B2` | `=B2` | `=B4` |
| A18 | Rate + 1% | `=LoanParameters!B2+0.01` | `=PMT(B18/12, LoanParameters!B3*12, -LoanParameters!B4)` | `=C18*LoanParameters!B3*12-LoanParameters!B4` |
| A19 | Rate - 1% | `=LoanParameters!B2-0.01` | `=PMT(B19/12, LoanParameters!B3*12, -LoanParameters!B4)` | `=C19*LoanParameters!B3*12-LoanParameters!B4` |

### Step 5: Set up "InvestmentAnalysis" Sheet

| Cell | Label | Formula | Description |
|------|-------|---------|-------------|
| A1 | **Investment Growth** | | Header |
| A2 | Future Value (Lump Sum) | | Label |
| B2 | `=FV(LoanParameters!B7, LoanParameters!B8, 0, -LoanParameters!B10)` | FV of initial investment |
| A3 | Future Value (With Contributions) | | Label |
| B3 | `=FV(LoanParameters!B7/12, LoanParameters!B8*12, -LoanParameters!B9, -LoanParameters!B10)` | FV with monthly contributions |
| A4 | Present Value Needed | | Label |
| B4 | `=PV(LoanParameters!B7, LoanParameters!B8, 0, -1000000)` | PV needed for $1M goal |
| A5 | Monthly Contribution Needed | | Label |
| B5 | `=PMT(LoanParameters!B7/12, LoanParameters!B8*12, -LoanParameters!B10, 1000000)` | Monthly amount for $1M |
| | | | |
| A7 | **Project Analysis** | | Header |
| A8 | Net Present Value | | Label |
| B8 | `=NPV(LoanParameters!B13, CashFlows!B3:B12)` | NPV of cash flows |
| A9 | Internal Rate of Return | | Label |
| B9 | `=IRR(CashFlows!C2:C12)` | IRR calculation |
| A10 | Payback Period (Years) | | Label |
| B10 | `=MATCH(TRUE, SUBTOTAL(9,OFFSET(CashFlows!C2,0,0,ROW(CashFlows!C3:C12)-1,1))>0, 0)` | Years to break even |
| A11 | Profitability Index | | Label |
| B11 | `=1 + B8/ABS(CashFlows!C2)` | PI calculation |
| | | | |
| A13 | **Sensitivity Analysis** | | Header |
| A14 | Discount Rate | NPV | Decision |
| A15 | 5% | `=NPV(0.05, CashFlows!B3:B12)` | `=IF(B15>0, "Accept", "Reject")` |
| A16 | 10% | `=NPV(0.10, CashFlows!B3:B12)` | `=IF(B16>0, "Accept", "Reject")` |
| A17 | 15% | `=NPV(0.15, CashFlows!B3:B12)` | `=IF(B17>0, "Accept", "Reject")` |
| A18 | 20% | `=NPV(0.20, CashFlows!B3:B12)` | `=IF(B18>0, "Accept", "Reject")` |

## How to Use This Template

1. **Create the Google Sheet** with the structure above
2. **Share the sheet** (read-only) to get a public URL
3. **Update the example JSON** file with your sheet's URL
4. **Run the converter**:
   ```bash
   npm run cli -- convert --config examples/financial-model.json
   ```

## Expected Output

The converter will generate TypeScript or Python code that:
- Implements all financial calculations (PMT, FV, PV, RATE, NPV, IRR)
- Handles loan amortization logic
- Performs investment analysis
- Includes sensitivity analysis functions
- Maintains calculation dependencies

## Key Features Demonstrated

1. **Financial Functions**: PMT, FV, PV, RATE, NPV, IRR
2. **Cross-sheet References**: Formulas reference data from other sheets
3. **Complex Calculations**: Amortization schedules, sensitivity analysis
4. **Conditional Logic**: IF statements for decision making
5. **Array Operations**: Working with ranges of cash flows

## Sample Public Sheet URL

Once you create your sheet, it will have a URL like:
```
https://docs.google.com/spreadsheets/d/1AbCdEfGhIjKlMnOpQrStUvWxYz/edit#gid=0
```

Make sure to:
1. Click "Share" → "Anyone with the link can view"
2. Copy the URL
3. Update the `financial-model.json` file with this URL