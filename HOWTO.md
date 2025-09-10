# How to Convert Your Existing Google Sheet to Code

This guide helps you convert your existing complex Google Sheet into executable TypeScript or Python code that produces identical results.

## Table of Contents
1. [Prerequisites](#prerequisites)
2. [Quick Start](#quick-start)
3. [Preparing Your Existing Sheet](#preparing-your-existing-sheet)
4. [Running the Conversion](#running-the-conversion)
5. [Comprehensive Validation Guide](#comprehensive-validation-guide)
6. [Troubleshooting Common Issues](#troubleshooting-common-issues)
7. [Advanced Usage](#advanced-usage)

## Prerequisites

1. **Node.js and npm** installed on your system

2. **Clone and install this project**:
   ```bash
   git clone <repository-url>
   cd google-sheets-to-code
   npm install
   ```

3. **Google Sheets API Access** - Set up authentication:
   ```bash
   npm run cli -- setup
   ```
   Follow the instructions to either:
   - Use a Service Account (recommended for automation)
   - Use OAuth2 (for personal use)
   
   Save your `credentials.json` in the project root.

4. **Share your existing sheet** (if using Service Account):
   - Open your Google Sheet
   - Click "Share"
   - Add the service account email (found in `credentials.json` as `client_email`)
   - Give it "Viewer" access

## Quick Start

If you want to convert your sheet immediately:

```bash
# Convert your existing sheet to TypeScript
npm run cli -- convert \
  --url "https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit" \
  --input-tabs "Data,Settings" \
  --output-tabs "Calculations,Reports" \
  --output-file my-spreadsheet.ts
```

This will generate code that:
- Reads data from "Data" and "Settings" tabs
- Calculates all formulas from "Calculations" and "Reports" tabs
- Returns the same results as your Google Sheet

## Preparing Your Existing Sheet

### Understanding Input vs Output Tabs

Before conversion, identify which tabs in your sheet are:
- **Input tabs**: Contain raw data (numbers, text, dates) that feed into calculations
- **Output tabs**: Contain formulas that reference input tabs or other cells

Example structure of a typical financial model:
```
Input Tabs:
- "Assumptions" - interest rates, growth rates, etc.
- "Historical Data" - past performance data
- "Scenarios" - different scenario parameters

Output Tabs:
- "Projections" - formulas calculating future values
- "Summary" - formulas aggregating results
- "Dashboard" - formulas for key metrics
```

### Checking Formula Compatibility

Review the formulas in your output tabs. Currently supported functions include:
- Math: `SUM`, `AVERAGE`, `MIN`, `MAX`, `ROUND`, `POWER`, `SQRT`, `ABS`
- Logic: `IF`, `AND`, `OR`, `NOT`, `IFERROR`, `IFNA`
- Lookup: `VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH`
- Text: `CONCATENATE`, `UPPER`, `LOWER`, `TRIM`, `LEN`
- Date: `TODAY`, `NOW`, `DATE`, `YEAR`, `MONTH`, `DAY`
- Statistical: `COUNT`, `COUNTA`, `COUNTIF`, `SUMIF`, `AVERAGEIF`
- Financial: `PMT`, `FV`, `PV`, `NPV`, `IRR`

See the full list in [README.md](README.md#supported-formulas).

### Handling Named Ranges

If your sheet uses named ranges, they will be automatically detected and used in the generated code. No special preparation needed!

## Running the Conversion

### Option 1: Using Command Line Arguments

```bash
npm run cli -- convert \
  --url "YOUR_SHEET_URL" \
  --input-tabs "Input1,Input2,Input3" \
  --output-tabs "Output1,Output2" \
  --language typescript \
  --output-file generated-code.ts \
  --verbose
```

### Option 2: Using a Configuration File (Recommended)

Create a `config.json` file:

```json
{
  "spreadsheetUrl": "https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit",
  "inputTabs": ["Assumptions", "Data", "Settings"],
  "outputTabs": ["Calculations", "Summary", "Reports"],
  "outputLanguage": "typescript"
}
```

Then run:
```bash
npm run cli -- convert --config config.json --output-file my-model.ts
```

### Option 3: For Python Output

```bash
npm run cli -- convert \
  --config config.json \
  --language python \
  --output-file my-model.py
```

## Comprehensive Validation Guide

Validation is crucial to ensure your generated code produces exactly the same results as your Google Sheet. The validation process works by:
1. Fetching actual values from your Google Sheet
2. Running the generated code with test inputs
3. Comparing every cell value between the sheet and the code output

### Understanding Validation Workflow

The validation system supports two main workflows:

#### Workflow 1: Validate During Conversion (Integrated)
This runs validation immediately after generating code:
```bash
npm run cli -- convert \
  --config config.json \
  --output-file my-model.ts \
  --validate \
  --validation-input test-input.json \
  --validation-tolerance 1e-10
```

#### Workflow 2: Separate Validation (Recommended for Testing)
First generate the code, then validate separately:
```bash
# Step 1: Generate code
npm run cli -- convert \
  --config config.json \
  --output-file my-model.ts

# Step 2: Validate against the original sheet
npm run cli -- validate \
  --config config.json \
  --generated-file my-model.ts \
  --input-data test-input.json \
  --tolerance 1e-10 \
  --output validation-report.md
```

### Creating Validation Test Cases

#### Basic Test Input File
Create `test-input.json` with values for your input cells:

```json
{
  "Assumptions": {
    "B2": 0.05,
    "B3": 0.03,
    "B4": 1000000,
    "B5": "2024-01-01"
  },
  "Data": {
    "A2": "Product A",
    "B2": 100,
    "C2": 25.50,
    "D2": true,
    "A3": "Product B",
    "B3": 200,
    "C3": 35.75,
    "D3": false
  }
}
```

#### Multiple Test Scenarios
Create different input files for various scenarios:

```bash
# Scenario 1: Best case
npm run cli -- validate \
  --config config.json \
  --generated-file my-model.ts \
  --input-data best-case-input.json \
  --output best-case-report.md

# Scenario 2: Worst case  
npm run cli -- validate \
  --config config.json \
  --generated-file my-model.ts \
  --input-data worst-case-input.json \
  --output worst-case-report.md

# Scenario 3: Average case
npm run cli -- validate \
  --config config.json \
  --generated-file my-model.ts \
  --input-data average-case-input.json \
  --output average-case-report.md
```

### Validating Against Different Sheet Versions

A powerful feature is validating your generated code against different versions or copies of your sheet. This is useful for:
- Testing against sheets with different data
- Validating against backup copies
- Comparing results across different scenarios

#### Method 1: Using Different Sheet URLs for Validation

Create separate configuration files for each validation scenario:

`config-production.json`:
```json
{
  "spreadsheetUrl": "https://docs.google.com/spreadsheets/d/PRODUCTION_SHEET_ID/edit",
  "inputTabs": ["Inputs"],
  "outputTabs": ["Outputs"],
  "outputLanguage": "typescript"
}
```

`config-test.json`:
```json
{
  "spreadsheetUrl": "https://docs.google.com/spreadsheets/d/TEST_SHEET_ID/edit",
  "inputTabs": ["Inputs"],
  "outputTabs": ["Outputs"],
  "outputLanguage": "typescript"
}
```

Then validate against different sheets:
```bash
# Generate code from production sheet
npm run cli -- convert \
  --config config-production.json \
  --output-file production-model.ts

# Validate against test sheet with different data
npm run cli -- validate \
  --config config-test.json \
  --generated-file production-model.ts \
  --input-data test-data.json \
  --output test-validation-report.md
```

#### Method 2: Creating Validation Test Sheets

1. **Make a copy of your original sheet**:
   - In Google Sheets: File → Make a copy
   - Name it "MySheet - Validation Test 1"

2. **Modify the test data in the input tabs**:
   - Change values to test edge cases
   - Add extreme values (0, negative, very large)
   - Test with empty cells

3. **Create a validation config for the test sheet**:

`validation-test-config.json`:
```json
{
  "spreadsheetUrl": "https://docs.google.com/spreadsheets/d/TEST_COPY_ID/edit",
  "inputTabs": ["Inputs"],
  "outputTabs": ["Outputs"],
  "outputLanguage": "typescript"
}
```

4. **Run validation against the test sheet**:
```bash
# Validate your generated code against the test sheet
npm run cli -- validate \
  --config validation-test-config.json \
  --generated-file my-model.ts \
  --input-data edge-case-inputs.json \
  --tolerance 1e-10 \
  --output edge-case-validation.md \
  --verbose
```

### Advanced Validation Techniques

#### Capturing and Storing Validation Data

The `capture-validation` command allows you to capture snapshots of your sheet's current state and store them for later validation. This is extremely useful for:
- Building a test suite with multiple scenarios
- Testing against historical data
- Validating across different environments
- Creating regression tests

##### Capturing Validation Snapshots

```bash
# Capture current state of the sheet defined in config.json
npm run cli -- capture-validation \
  --config config.json \
  --verbose

# Capture from a different sheet URL (but using same tab structure)
npm run cli -- capture-validation \
  --config config.json \
  --url "https://docs.google.com/spreadsheets/d/DIFFERENT_SHEET_ID/edit" \
  --name "test-scenario-1"

# Capture multiple scenarios
npm run cli -- capture-validation \
  --config config.json \
  --url "https://docs.google.com/spreadsheets/d/SCENARIO_1_ID/edit" \
  --name "best-case"

npm run cli -- capture-validation \
  --config config.json \
  --url "https://docs.google.com/spreadsheets/d/SCENARIO_2_ID/edit" \
  --name "worst-case"

npm run cli -- capture-validation \
  --config config.json \
  --url "https://docs.google.com/spreadsheets/d/SCENARIO_3_ID/edit" \
  --name "average-case"
```

Captured data is stored in `.validation/${configName}/${name-or-uuid}.json` with:
- Input tab values (automatically used as test inputs)
- Output tab values (used for validation)
- Timestamp and source URL for reference

##### Using Stored Validation Snapshots

When you run the `validate` command, it automatically checks for stored snapshots:

```bash
# This will automatically find and use ALL stored snapshots in .validation/config/
npm run cli -- validate \
  --config config.json \
  --generated-file my-model.ts \
  --tolerance 1e-10 \
  --verbose
```

The validation will:
1. Check `.validation/${configName}/` directory for snapshots
2. Load each snapshot with its input and output data
3. Run your generated code with each snapshot's inputs
4. Compare outputs against each snapshot's expected values
5. Report results for all snapshots

##### Building a Comprehensive Test Suite

```bash
#!/bin/bash
# capture-test-suite.sh

# Configuration file to use
CONFIG="production-config.json"

# Capture baseline from production
npm run cli -- capture-validation \
  --config $CONFIG \
  --name "baseline-2024-01"

# Capture edge cases from test sheets
npm run cli -- capture-validation \
  --config $CONFIG \
  --url "https://docs.google.com/spreadsheets/d/EMPTY_VALUES_ID/edit" \
  --name "empty-values"

npm run cli -- capture-validation \
  --config $CONFIG \
  --url "https://docs.google.com/spreadsheets/d/MAX_VALUES_ID/edit" \
  --name "maximum-values"

npm run cli -- capture-validation \
  --config $CONFIG \
  --url "https://docs.google.com/spreadsheets/d/NEGATIVE_ID/edit" \
  --name "negative-values"

npm run cli -- capture-validation \
  --config $CONFIG \
  --url "https://docs.google.com/spreadsheets/d/DECIMAL_ID/edit" \
  --name "decimal-precision"

echo "Test suite captured in .validation/production-config/"
ls -la .validation/production-config/
```

##### Continuous Validation Workflow

```bash
# Daily validation script
#!/bin/bash

# Step 1: Generate fresh code from production sheet
npm run cli -- convert \
  --config production.json \
  --output-file generated-model.ts

# Step 2: Capture today's production data
npm run cli -- capture-validation \
  --config production.json \
  --name "production-$(date +%Y%m%d)"

# Step 3: Validate against ALL stored test cases
npm run cli -- validate \
  --config production.json \
  --generated-file generated-model.ts \
  --output "reports/validation-$(date +%Y%m%d).md"

# Step 4: Check validation results
if [ $? -eq 0 ]; then
  echo "✅ All validations passed!"
else
  echo "❌ Validation failed - check reports for details"
  exit 1
fi
```

#### Validating with Multiple Snapshots (Live Fetching)
For sheets with volatile functions (TODAY, NOW, RAND), take multiple live snapshots:

```bash
npm run cli -- validate \
  --config config.json \
  --generated-file my-model.ts \
  --input-data test-input.json \
  --snapshots 5 \
  --delay 2000 \
  --output time-based-validation.md
```

This will:
1. Fetch values from Google Sheets
2. Wait 2 seconds
3. Fetch again (5 times total)
4. Validate generated code against each snapshot
5. Report any inconsistencies

#### Batch Validation Script
Create a script to validate multiple scenarios:

`validate-all.sh`:
```bash
#!/bin/bash

# Generate the code once
npm run cli -- convert \
  --config production-config.json \
  --output-file generated-model.ts

# Validate against multiple test scenarios
echo "Validating Scenario 1: Empty inputs..."
npm run cli -- validate \
  --config test-config-1.json \
  --generated-file generated-model.ts \
  --input-data empty-inputs.json \
  --output reports/empty-validation.md

echo "Validating Scenario 2: Maximum values..."
npm run cli -- validate \
  --config test-config-2.json \
  --generated-file generated-model.ts \
  --input-data max-inputs.json \
  --output reports/max-validation.md

echo "Validating Scenario 3: Negative values..."
npm run cli -- validate \
  --config test-config-3.json \
  --generated-file generated-model.ts \
  --input-data negative-inputs.json \
  --output reports/negative-validation.md

# Summarize results
echo "Validation Complete. Check reports/ directory for details."
```

### Understanding Validation Output

#### Successful Validation
```
🔍 Validating generated code against actual data...
📊 Fetching validation data from Google Sheets...
✅ Sheet "Calculations" values fetched: 245 cells
✅ Sheet "Summary" values fetched: 89 cells

✅ VALIDATION PASSED
Accuracy: 100.00% (334/334 cells)
```

#### Failed Validation with Details
```
🔍 Validating generated code against actual data...
📊 Fetching validation data from Google Sheets...
✅ Sheet "Calculations" values fetched: 245 cells
✅ Sheet "Summary" values fetched: 89 cells

❌ VALIDATION FAILED
Accuracy: 98.80% (330/334 cells)

First 10 mismatches:
  - Calculations!B15: expected 1234567.89, got 1234567.88
  - Calculations!C20: expected 0.05, got 0.0499999
  - Summary!D5: expected "Total: $1,000", got "Total: $1000"
  - Summary!E10: expected null, got 0
```

#### Validation Report File
The generated markdown report includes:
- **Summary statistics**: Total cells validated, accuracy percentage
- **Detailed mismatches**: Every cell that doesn't match, with expected vs actual values
- **Tolerance information**: The numeric tolerance used for comparisons
- **Timestamp**: When the validation was performed
- **Configuration**: Which sheets and tabs were validated

### Validation Options Reference

| Option | Description | Default | Example |
|--------|-------------|---------|---------|
| `--config` | Configuration file with sheet URL and tabs | Required | `--config config.json` |
| `--generated-file` | Path to generated code file to validate | Required | `--generated-file model.ts` |
| `--input-data` | JSON file with input cell values | {} | `--input-data test.json` |
| `--tolerance` | Numeric tolerance for float comparison | 1e-10 | `--tolerance 0.0001` |
| `--snapshots` | Number of validation snapshots | 1 | `--snapshots 5` |
| `--delay` | Delay between snapshots (ms) | 1000 | `--delay 2000` |
| `--output` | Output validation report to file | Console | `--output report.md` |
| `--verbose` | Show detailed progress | false | `--verbose` |

### Validation Best Practices

1. **Create a validation test suite** with multiple input scenarios:
   - Empty/null values
   - Boundary values (min/max)
   - Typical values
   - Error-inducing values

2. **Use different sheets for different test cases**:
   - Production sheet (read-only)
   - Test sheet with edge cases
   - Validation sheet with known outputs

3. **Automate validation in CI/CD**:
   ```yaml
   # .github/workflows/validate.yml
   name: Validate Sheet Conversion
   on: [push, pull_request]
   jobs:
     validate:
       runs-on: ubuntu-latest
       steps:
         - uses: actions/checkout@v2
         - uses: actions/setup-node@v2
         - run: npm install
         - run: |
             # Generate code from production sheet
             npm run cli -- convert \
               --config config.json \
               --output-file generated.ts
             
             # Validate against multiple test scenarios
             npm run cli -- validate \
               --config test-config.json \
               --generated-file generated.ts \
               --input-data test-cases/case1.json \
               --output reports/case1.md
             
             npm run cli -- validate \
               --config test-config.json \
               --generated-file generated.ts \
               --input-data test-cases/case2.json \
               --output reports/case2.md
   ```

4. **Monitor validation accuracy over time**:
   - Save validation reports with timestamps
   - Track accuracy trends
   - Alert on accuracy drops below threshold

## Troubleshooting Common Issues

### "Formula contains unsupported function: XLOOKUP"
**Solution**: The converter doesn't support XLOOKUP yet. Consider replacing with VLOOKUP or INDEX/MATCH.

### "Authentication failed"
**Solution**: 
- Ensure `credentials.json` exists in the project root
- If using service account, verify the sheet is shared with the service account email
- Try re-running `npm run cli -- setup`

### Validation shows small differences in decimal values
**Solution**: Increase the tolerance for floating-point comparisons:
```bash
--validation-tolerance 0.0001  # Instead of default 1e-10
```

### "Sheet not found: SheetName"
**Solution**: Check the exact spelling and case of your sheet names. Use quotes if they contain spaces:
```bash
--input-tabs "Sheet 1,Sheet 2"
```

### Formulas with volatile functions (TODAY, NOW, RAND) fail validation
**Solution**: Use multiple snapshots for time-based validation:
```bash
npm run cli -- validate \
  --config config.json \
  --generated-file my-model.ts \
  --snapshots 3 \
  --delay 2000
```

### Generated code doesn't compile
**Solution**: 
- Check for unsupported formulas in your sheet
- Ensure all referenced sheets are included in either inputTabs or outputTabs
- Review the generated code for syntax errors

## Advanced Usage

### Handling Complex Sheets

For sheets with 10+ tabs and thousands of formulas:

1. **Start with a subset**: Test with a few tabs first
2. **Use verbose mode**: Add `--verbose` to see detailed progress
3. **Validate incrementally**: Validate one output tab at a time

### Optimizing Performance

For large sheets with many formulas:

```bash
# Generate optimized code with dependency analysis
npm run cli -- convert \
  --config config.json \
  --output-file optimized-model.ts \
  --optimize
```

### Continuous Integration

Add to your CI/CD pipeline to ensure changes don't break the conversion:

```yaml
# .github/workflows/validate-sheet.yml
name: Validate Sheet Conversion
on: [push]
jobs:
  validate:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v2
      - uses: actions/setup-node@v2
      - run: npm install
      - run: |
          npm run cli -- convert \
            --config config.json \
            --output-file generated.ts \
            --validate \
            --validation-input test-data.json
```

### Working with Protected Sheets

If your sheet has protected ranges or requires specific permissions:
1. Ensure the service account has the necessary permissions
2. Use OAuth2 authentication if you need user-level access
3. Consider making a copy of the sheet for conversion purposes

### Handling External Data Sources

If your sheet uses `IMPORTRANGE`, `IMPORTDATA`, or other external data:
1. Create local copies of external data in input tabs
2. Replace `IMPORTRANGE` formulas with direct references
3. Or use the `--fetch-external` flag (if supported)

## Best Practices

1. **Version Control**: Keep your config files and test data in git
2. **Document Dependencies**: Note which tabs depend on which inputs
3. **Test Edge Cases**: Include zero, negative, and maximum values in test data
4. **Regular Validation**: Re-validate after any significant sheet changes
5. **Modular Approach**: Break complex sheets into logical groups of tabs

## Examples

### Example 1: Financial Model with Validation
```bash
# Generate code from production sheet
npm run cli -- convert \
  --url "https://docs.google.com/spreadsheets/d/PROD_ID/edit" \
  --input-tabs "Assumptions,MarketData,Scenarios" \
  --output-tabs "Projections,Valuation,Sensitivity" \
  --output-file financial-model.ts

# Validate against test sheet with edge cases
npm run cli -- validate \
  --config test-sheet-config.json \
  --generated-file financial-model.ts \
  --input-data edge-cases.json \
  --tolerance 0.01 \
  --output validation-report.md
```

### Example 2: Sales Dashboard with Multiple Validations
```bash
#!/bin/bash
# validate-sales-dashboard.sh

# Generate the dashboard code
npm run cli -- convert \
  --config sales-config.json \
  --output-file sales-dashboard.py

# Validate against Q1 data
npm run cli -- validate \
  --config q1-test-config.json \
  --generated-file sales-dashboard.py \
  --input-data q1-data.json \
  --output reports/q1-validation.md

# Validate against Q2 data
npm run cli -- validate \
  --config q2-test-config.json \
  --generated-file sales-dashboard.py \
  --input-data q2-data.json \
  --output reports/q2-validation.md
```

### Example 3: Inventory System with Time-Based Validation
```bash
npm run cli -- convert \
  --config inventory-config.json \
  --output-file inventory-system.ts

# Validate with multiple snapshots for NOW() functions
npm run cli -- validate \
  --config inventory-test-config.json \
  --generated-file inventory-system.ts \
  --input-data current-inventory.json \
  --snapshots 5 \
  --delay 3000 \
  --tolerance 0.001 \
  --output time-validation-report.md
```

## Getting Help

- **Documentation**: Check [README.md](README.md) for detailed feature documentation
- **Supported Functions**: See the complete list of supported formulas
- **Examples**: Look in the `examples/` directory for sample configurations
- **Issues**: Report problems on GitHub Issues
- **Feature Status**: Check [TODO.md](TODO.md) for upcoming features

## Next Steps

After successfully converting and validating your sheet:

1. **Set up automated validation** to run on schedule
2. **Create a test suite** with multiple validation scenarios
3. **Monitor accuracy metrics** over time
4. **Integrate validation into your CI/CD pipeline**
5. **Document validation procedures** for your team

Remember: Validation is your safety net. Always validate your generated code against known good data before using it in production!