# Google Sheets to Code - Example Templates

This directory contains example configurations and templates for common use cases. Each example includes:
1. A JSON configuration file
2. A detailed template (`.md` file) with instructions for creating the Google Sheet
3. Sample data and formulas to demonstrate various features
4. Ready-to-use XLSX files that can be uploaded to Google Sheets
5. A generate_xlsx.py script that can be used to generate the XLSX files

## Available Examples

### 1. Financial Model (`financial-model`)
**Features demonstrated:**
- Financial functions: PMT, FV, PV, RATE, NPV, IRR
- Loan amortization calculations
- Investment analysis
- Sensitivity analysis
- Cross-sheet references

**Files:**
- `financial-model.json` - Configuration file
- `financial-model-template.md` - Detailed sheet creation instructions
- `financial-model.xlsx` - Ready-to-use XLSX file
- `generate_xlsx.py` - Script to generate XLSX files

**Usage:**
```bash
# After creating your sheet from the template, or using the sheet URL provided in the JSON, provided it still works
npm run cli -- convert --config examples/financial-model.json
```

### 2. Data Analysis (`data-analysis`)
**Features demonstrated:**
- Statistical functions: AVERAGE, MEDIAN, STDEV, PERCENTILE
- Conditional aggregation: COUNTIF, SUMIF, AVERAGEIF
- Ranking and sorting: RANK, LARGE, SMALL
- Outlier detection and Z-scores
- Regional and product analysis
- Top performers identification

**Files:**
- `data-analysis.json` - Configuration file
- `data-analysis-template.md` - Detailed sheet creation instructions
- `data-analysis.xlsx` - Ready-to-use XLSX file
- `generate_xlsx.py` - Script to generate XLSX files

**Usage:**
```bash
# After creating your sheet from the template, or using the sheet URL provided in the JSON, provided it still works
npm run cli -- convert --config examples/data-analysis.json
```

### 3. Inventory Tracking (`inventory-tracking`)
**Features demonstrated:**
- Stock level calculations from transaction history
- Reorder point determination
- ABC classification
- Economic Order Quantity (EOQ)
- Automated alerts for critical stock
- Performance metrics and KPIs
- Multi-sheet dependencies

**Files:**
- `inventory-tracking.json` - Configuration file
- `inventory-tracking-template.md` - Detailed sheet creation instructions
- `inventory-tracking.xlsx` - Ready-to-use XLSX file
- `generate_xlsx.py` - Script to generate XLSX files

**Usage:**
```bash
# After creating your sheet from the template, or using the sheet URL provided in the JSON, provided it still works
npm run cli -- convert --config examples/inventory-tracking.json
```

## How to Use These Examples

### Quick Start with XLSX Files
The fastest way to get started is to use the pre-built XLSX files:

1. **Upload to Google Sheets**:
   - Go to [Google Sheets](https://sheets.google.com)
   - Click "File" → "Import" → "Upload"
   - Select the XLSX file (e.g., `financial-model.xlsx`)
   - Choose "Replace spreadsheet" and click "Import data"

2. **Share the sheet** and get the URL (see Step 2 below)

3. **Update the JSON** configuration (see Step 3 below)

### Manual Creation from Templates
If you prefer to build the sheets manually:

1. Choose an example that matches your use case
2. Open the corresponding `-template.md` file
3. Follow the detailed instructions to create the Google Sheet
4. Add the sample data as shown in the template

### Step 2: Share the Sheet
1. In Google Sheets, click "Share" button
2. Click "Change to anyone with the link"
3. Set permission to "Viewer"
4. Copy the share link

### Step 3: Update the Configuration
1. Open the corresponding `.json` file
2. Replace `YOUR_SHEET_ID_HERE` with your actual sheet ID from the URL
3. The URL format is: `https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit`

### Step 4: Run the Converter
```bash
# For financial model (TypeScript output)
npm run cli -- convert --config examples/financial-model.json

# For data analysis (Python output)
npm run cli -- convert --config examples/data-analysis.json

# For inventory tracking (TypeScript output)
npm run cli -- convert --config examples/inventory-tracking.json

# With watch mode for development
npm run cli -- convert --config examples/financial-model.json --watch
```

### Step 5: Use the Generated Code
The converter will create a file with all your spreadsheet logic converted to code:
- TypeScript files can be run directly with `node` after compilation
- Python files can be run with `python3`
- Both include CLI support for testing with different inputs

## Creating Your Own Examples

To create a custom example:

1. **Design your spreadsheet** with clear input and output tabs
2. **Create a configuration file** specifying:
   - `spreadsheetUrl`: Your Google Sheets URL
   - `inputTabs`: Array of input sheet names
   - `outputTabs`: Array of output sheet names
   - `outputLanguage`: "typescript" or "python"
3. **Document your formulas** for reference
4. **Test the conversion** with the CLI tool

## Common Patterns

### Financial Calculations
```excel
=PMT(rate/12, years*12, -principal)
=FV(rate, periods, payment, present_value)
=NPV(discount_rate, cash_flows_range)
```

### Statistical Analysis
```excel
=AVERAGE(data_range)
=STDEV(data_range)
=PERCENTILE(data_range, 0.75)
=COUNTIF(range, ">100")
```

### Inventory Management
```excel
=SUMIFS(quantity, sku_column, sku, type_column, "IN") - SUMIFS(quantity, sku_column, sku, type_column, "OUT")
=IF(current_stock <= reorder_point, "REORDER", "OK")
```

## Tips for Best Results

1. **Organize your sheets clearly** - Separate input data from calculations
2. **Use named ranges** where possible - They convert to readable variable names
3. **Document complex formulas** - Add notes in adjacent cells
4. **Test with sample data** - Ensure formulas work before conversion
5. **Keep formulas simple** - Break complex logic into multiple cells

## Troubleshooting

If you encounter issues:

1. **Authentication errors** - Ensure credentials.json is properly configured
2. **Sheet not found** - Check that the sheet is shared publicly
3. **Formula not supported** - Check the README for supported functions
4. **Circular dependencies** - The tool will warn but still generate code with error handling

## Regenerating XLSX Files

If you need to regenerate or modify the XLSX files:

```bash
cd examples
python3 generate_xlsx.py
```

This will create/update:
- `financial-model.xlsx` - Loan calculations with financial functions
- `data-analysis.xlsx` - Statistical analysis with 100 rows of sample data
- `inventory-tracking.xlsx` - Complete inventory management system

The script requires `openpyxl`:
```bash
pip install openpyxl
```

## Contributing Examples

We welcome new example contributions! To add an example:

1. Create a well-documented Google Sheet
2. Write a detailed template markdown file
3. Create a configuration JSON file
4. Generate an XLSX file (update `generate_xlsx.py`)
5. Test the conversion thoroughly
6. Submit a pull request with all files

## Questions or Issues?

If you have questions about these examples or encounter any issues:
- Check the main [README](../README.md) for general documentation
- Review the template files for detailed instructions
- Open an issue on GitHub with specific error messages
- Include your configuration and any error output