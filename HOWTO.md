# How to Validate Google Sheets to Code Conversions

This guide walks you through creating a Google Sheet with comprehensive test cases and validating that the generated code produces identical results.

## Table of Contents
1. [Prerequisites](#prerequisites)
2. [Setting Up Your Google Sheet](#setting-up-your-google-sheet)
3. [Creating Validation Test Cases](#creating-validation-test-cases)
4. [Configuring the Converter](#configuring-the-converter)
5. [Running Conversion with Validation](#running-conversion-with-validation)
6. [Interpreting Results](#interpreting-results)
7. [Troubleshooting](#troubleshooting)

## Prerequisites

1. **Google Sheets API Access** - Follow the setup instructions:
   ```bash
   npm run cli -- setup
   ```
   Save your `credentials.json` in the project root.

2. **Node.js and npm** installed

3. **Clone and install the project**:
   ```bash
   git clone <repository-url>
   cd google-sheets-to-code
   npm install
   ```

## Setting Up Your Google Sheet

### 1. Create a New Google Sheet

Create a new Google Sheet with the following structure:

#### Sheet 1: "Input" (Input Tab)
This sheet contains the input data that will be provided to your calculations.

```
   A          B          C
1  Quantity   Price      Discount
2  10         25.50      0.1
3  5          100.00     0.15
4  20         15.75      0.05
5  8          50.00      0.2
```

#### Sheet 2: "Calculations" (Output Tab)
This sheet contains formulas that reference the Input sheet.

```
   A                B                      C                        D
1  Subtotal         Discount Amount        Total                    Tax (8%)
2  =Input!A2*Input!B2  =A2*Input!C2          =A2-B2                   =C2*0.08
3  =Input!A3*Input!B3  =A3*Input!C3          =A3-B3                   =C3*0.08
4  =Input!A4*Input!B4  =A4*Input!C4          =A4-B4                   =C4*0.08
5  =Input!A5*Input!B5  =A5*Input!C5          =A5-B5                   =C5*0.08
```

#### Sheet 3: "Validation" (Output Tab)
This sheet tests various formula types for comprehensive validation.

### 2. Add 10+ Different Validation Test Cases

Create these formulas in the "Validation" sheet to test different scenarios:

```
   A                           B                               C
1  Test Case                   Formula                         Description
2  Basic Math                  =Input!A2+Input!B2              Addition
3  Complex Math                 =POWER(Input!A2,2)+SQRT(Input!B2) Power and Square Root
4  Conditional                  =IF(Input!C2>0.1,"High","Low") IF statement
5  Nested IF                    =IF(Input!A2>10,IF(Input!B2>50,"Premium","Standard"),"Basic") Nested conditions
6  VLOOKUP                      =VLOOKUP(Input!A2,Input!A2:C5,3,FALSE) Lookup function
7  Statistical                  =AVERAGE(Input!B2:B5)           Average calculation
8  Count Functions              =COUNTIF(Input!C2:C5,">0.1")    Conditional counting
9  Text Functions               =CONCATENATE("Order: ",Input!A2," units") String operations
10 Date Functions               =TODAY()                        Current date
11 Financial                    =PMT(0.05/12,60,-Input!B2*100) Payment calculation
12 Array Formula                =SUM(Input!A2:A5*Input!B2:B5)  Array multiplication
13 Round Functions              =ROUND(Input!B2*Input!C2,2)    Rounding
14 Min/Max                      =MAX(Input!B2:B5)               Maximum value
15 Logical Operations           =AND(Input!A2>5,Input!B2<100)  AND operation
```

### 3. Share the Sheet (for Service Account)

If using a service account:
1. Click "Share" in your Google Sheet
2. Add the service account email (found in your `credentials.json` as `client_email`)
3. Give it "Viewer" access

## Configuring the Converter

### 1. Create a Configuration File

Create `validation-config.json`:

```json
{
  "spreadsheetUrl": "https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit",
  "inputTabs": ["Input"],
  "outputTabs": ["Calculations", "Validation"],
  "outputLanguage": "typescript"
}
```

### 2. Create Input Data for Validation

Create `validation-input.json` with different test data:

```json
{
  "Input": {
    "A2": 15,
    "B2": 30.00,
    "C2": 0.12,
    "A3": 7,
    "B3": 85.50,
    "C3": 0.18,
    "A4": 25,
    "B4": 20.00,
    "C4": 0.08,
    "A5": 12,
    "B5": 45.75,
    "C5": 0.25
  }
}
```

## Running Conversion with Validation

### Method 1: Convert and Validate in One Command

```bash
npm run cli -- convert \
  --config validation-config.json \
  --output-file generated-sheet.ts \
  --validate \
  --validation-input validation-input.json \
  --validation-tolerance 1e-10 \
  --verbose
```

### Method 2: Convert First, Then Validate Separately

#### Step 1: Generate the code
```bash
npm run cli -- convert \
  --config validation-config.json \
  --output-file generated-sheet.ts \
  --verbose
```

#### Step 2: Run validation
```bash
npm run cli -- validate \
  --config validation-config.json \
  --generated-file generated-sheet.ts \
  --input-data validation-input.json \
  --tolerance 1e-10 \
  --output validation-report.md \
  --verbose
```

### Method 3: Multiple Snapshots (for Time-Based Functions)

For sheets with time-based functions (TODAY, NOW, RAND), use multiple snapshots:

```bash
npm run cli -- validate \
  --config validation-config.json \
  --generated-file generated-sheet.ts \
  --input-data validation-input.json \
  --snapshots 3 \
  --delay 2000 \
  --output validation-report-snapshots.md \
  --verbose
```

This will:
1. Fetch actual values from Google Sheets
2. Wait 2 seconds
3. Fetch again (3 times total)
4. Validate generated code against each snapshot

## Creating Comprehensive Test Scenarios

### Test Scenario 1: Financial Calculations
```javascript
// In Google Sheets "Financial" tab
A1: "Principal"     B1: 10000
A2: "Rate"          B2: 0.05
A3: "Years"         B3: 5
A4: "Monthly Pmt"   B4: =PMT(B2/12,B3*12,-B1)
A5: "Total Paid"    B5: =B4*B3*12
A6: "Interest"      B6: =B5+B1
```

### Test Scenario 2: Data Validation Rules
```javascript
// Add validation rules to your sheet
A1: =IF(ISBLANK(Input!A2),"Missing","Present")
A2: =IF(ISNUMBER(Input!B2),"Valid Price","Invalid")
A3: =IF(AND(Input!C2>=0,Input!C2<=1),"Valid Discount","Invalid")
```

### Test Scenario 3: Cross-Sheet References
```javascript
// Create formulas that reference multiple sheets
=SUMIF(Input!A:A,">10",Input!B:B)
=AVERAGEIF(Calculations!D:D,">0")
=COUNTIFS(Input!A:A,">5",Input!B:B,"<100")
```

## Interpreting Results

### Successful Validation Output
```
✅ VALIDATION PASSED
Accuracy: 100.00% (45/45 cells)
```

### Failed Validation Output
```
❌ VALIDATION FAILED
Accuracy: 95.56% (43/45 cells)

First 10 mismatches:
  - Validation!B10: expected 44330.98, got 44330.97
  - Validation!B11: expected 0.00833333, got 0.008333
```

### Understanding Validation Report

The generated `validation-report.md` contains:
- **Summary**: Overall pass/fail status
- **Accuracy**: Percentage of cells that match
- **Mismatches**: Detailed list of cells with different values
- **Tolerance**: Applied numeric tolerance for comparisons

## Validation Options Explained

| Option | Description | Default | Example |
|--------|-------------|---------|---------|
| `--validate` | Enable validation after conversion | false | `--validate` |
| `--validation-input` | JSON file with input data | none | `--validation-input data.json` |
| `--validation-tolerance` | Numeric tolerance for float comparison | 1e-10 | `--validation-tolerance 0.001` |
| `--validation-snapshots` | Number of snapshots to take | 1 | `--validation-snapshots 5` |
| `--validation-delay` | Delay between snapshots (ms) | 1000 | `--validation-delay 2000` |

## Troubleshooting

### Common Issues and Solutions

#### 1. Validation Fails Due to Rounding
**Problem**: Small differences in floating-point calculations
**Solution**: Increase tolerance
```bash
--validation-tolerance 0.0001
```

#### 2. Time-Based Functions Don't Match
**Problem**: TODAY() or NOW() values change between fetch and execution
**Solution**: Use multiple snapshots or mock these functions in tests

#### 3. Missing Cell References
**Problem**: "Cell not found in validation data"
**Solution**: Ensure all referenced cells exist in your Google Sheet

#### 4. Authentication Errors
**Problem**: "Failed to authenticate"
**Solution**: 
- Check `credentials.json` exists
- For service accounts, ensure sheet is shared with the service account email
- For OAuth, ensure you complete the browser authentication

#### 5. Formula Not Supported
**Problem**: "Failed to parse formula"
**Solution**: Check [supported formulas list](README.md#supported-formulas) or simplify complex formulas

### Debugging Tips

1. **Use Verbose Mode**: Add `--verbose` to see detailed progress
2. **Check Generated Code**: Review the generated TypeScript/Python file
3. **Test with Simple Data First**: Start with basic formulas before complex ones
4. **Validate Incrementally**: Test one sheet at a time

## Best Practices

1. **Organize Test Cases**: Group similar formulas together
2. **Use Named Ranges**: Makes formulas more readable and maintainable
3. **Document Complex Formulas**: Add comments in adjacent cells
4. **Test Edge Cases**: Include zero, negative, and empty values
5. **Version Control**: Keep your config and test data in git
6. **Regular Validation**: Run validation after any formula changes

## Example Full Workflow

```bash
# 1. Setup credentials
npm run cli -- setup

# 2. Create your Google Sheet with test cases (as shown above)

# 3. Create configuration
cat > my-validation-config.json << EOF
{
  "spreadsheetUrl": "https://docs.google.com/spreadsheets/d/YOUR_ID/edit",
  "inputTabs": ["Input"],
  "outputTabs": ["Calculations", "Validation"],
  "outputLanguage": "typescript"
}
EOF

# 4. Create test input data
cat > test-input.json << EOF
{
  "Input": {
    "A2": 100, "B2": 25.50, "C2": 0.1,
    "A3": 50,  "B3": 75.00, "C3": 0.15,
    "A4": 75,  "B4": 30.25, "C4": 0.05,
    "A5": 25,  "B5": 99.99, "C5": 0.2
  }
}
EOF

# 5. Run conversion with validation
npm run cli -- convert \
  --config my-validation-config.json \
  --output-file my-sheet.ts \
  --validate \
  --validation-input test-input.json \
  --verbose

# 6. Check the results
cat validation-report.md
```

## Advanced Validation Scenarios

### Validating Python Output
```bash
npm run cli -- convert \
  --config validation-config.json \
  --output-file generated-sheet.py \
  --language python \
  --validate \
  --validation-input validation-input.json
```

### Continuous Validation in CI/CD
```yaml
# .github/workflows/validate.yml
name: Validate Conversions
on: [push]
jobs:
  validate:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v2
      - uses: actions/setup-node@v2
      - run: npm install
      - run: |
          npm run cli -- validate \
            --config validation-config.json \
            --generated-file generated-sheet.ts \
            --input-data test-input.json \
            --tolerance 1e-10
```

## Getting Help

- Check the [README](README.md) for general usage
- Review [TODO.md](TODO.md) for feature status
- Report issues on GitHub Issues
- See example configurations in the `examples/` directory

Remember: The validation system helps ensure your generated code produces the same results as your Google Sheets. Always validate after making changes to complex formulas!