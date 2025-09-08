# Google Sheets to Code Converter

[![CI/CD Pipeline](https://github.com/joemastersemison/google-sheets-to-code/actions/workflows/test.yml/badge.svg)](https://github.com/joemastersemison/google-sheets-to-code/actions/workflows/test.yml)

Convert Google Sheets formulas and logic into executable TypeScript or Python code.

## Features

- 🔍 **Formula Parser**: Supports complex Google Sheets formulas including functions, operators, and cell references
- 📊 **Dependency Analysis**: Automatically determines calculation order and detects circular references
- 🎯 **Code Generation**: Generates clean, readable TypeScript or Python code
- 🔧 **CLI Interface**: Easy-to-use command-line tool
- ⚡ **Performance**: Optimized for large spreadsheets with hundreds of formulas

## Installation

### Node.js/TypeScript Version

```bash
npm install
```

### Python Package (for generated Python code)

```bash
# Install from source
pip install -e .

# Or build and install
python -m build
pip install dist/google_sheets_to_code-1.0.0-py3-none-any.whl
```

## Development Setup

### Pre-commit Hook

This project uses a pre-commit hook to automatically format code and check for linting errors. The hook will:

1. **Automatically fix** formatting and linting issues on staged files
2. **Block commits** if there are errors that can't be automatically fixed
3. **Run TypeScript type checking** to ensure type safety

The pre-commit hook is automatically installed when you run `npm install`.

### GitHub Actions

This project includes comprehensive GitHub Actions workflows:

- **CI/CD Pipeline** (`.github/workflows/test.yml`): Complete pipeline with code quality checks, comprehensive testing, security audits, and build verification
- **Auto-format PR** (`.github/workflows/format-pr.yml`): Allows maintainers to comment `/format` on PRs to auto-format code

All workflows run automatically on pushes and pull requests to ensure code quality.

## Setup

1. **Install dependencies**:
   ```bash
   npm install
   ```

2. **Setup Google Sheets API Credentials**:
   ```bash
   npm run cli -- setup
   ```

3. **Create credentials.json**:

   **Option A: Service Account (Recommended for CLI/automation)**
   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Enable Google Sheets API
   - Create a service account and download the JSON key
   - Save as `credentials.json` in project root
   - Share your Google Sheet with the service account email

   **Option B: OAuth2 (For interactive use)**
   - Create OAuth2 credentials in Google Cloud Console
   - Download and save as `credentials.json`
   - On first run, a browser will open for authorization
   - Token will be saved for future use

   **Note**: OAuth2 requires browser interaction on first run. For headless/CI environments, use a service account.

## Usage

### Command Line

```bash
# Basic conversion
npm run cli -- convert \
  --url "https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit" \
  --input-tabs "Input,Parameters" \
  --output-tabs "Results,Summary" \
  --language typescript \
  --output-file spreadsheet-logic.ts

# Using configuration file
npm run cli -- convert --config config.json

# Watch mode - automatically regenerate on changes
npm run cli -- convert --config config.json --watch
npm run cli -- convert --config config.json --watch --watch-interval 60

# Validate configuration
npm run cli -- validate --config config.json
```

### Configuration File

Create a `config.json` file:

```json
{
  "spreadsheetUrl": "https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit",
  "inputTabs": ["Input", "Parameters"],
  "outputTabs": ["Results", "Summary"],
  "outputLanguage": "typescript"
}
```

### Programmatic Usage

```typescript
import { SheetToCodeConverter } from './src/index.js';

const converter = new SheetToCodeConverter({
  spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit',
  inputTabs: ['Input'],
  outputTabs: ['Output'],
  outputLanguage: 'typescript'
});

const code = await converter.convert();
console.log(code);
```

### Example Configurations

The `examples/` directory contains complete examples with configurations, templates, and ready-to-use XLSX files:

- **`financial-model`**: Loan calculations with PMT, FV, PV, NPV, IRR functions
  - JSON configuration file
  - Excel file ready for Google Sheets import
  - Full documentation in template markdown
  
- **`data-analysis`**: Statistical analysis with 100 rows of sample data
  - Aggregation functions (AVERAGE, MEDIAN, STDEV, PERCENTILE)
  - Conditional analysis (COUNTIF, SUMIF, AVERAGEIF)
  - Outlier detection and Z-scores
  
- **`inventory-tracking`**: Complete inventory management system
  - Stock level calculations from transaction history
  - Reorder point determination with safety stock
  - ABC classification and EOQ calculations
  - Automated alerts for critical stock levels

#### Quick Start with Examples

1. **Upload XLSX to Google Sheets**:
   ```bash
   # The examples/ directory includes ready-to-use Excel files
   # Upload financial-model.xlsx, data-analysis.xlsx, or inventory-tracking.xlsx
   # to Google Sheets via File > Import
   ```

2. **Run the converter**:
   ```bash
   # Convert financial model to TypeScript
   npm run cli -- convert --config examples/financial-model.json
   
   # Convert data analysis to Python
   npm run cli -- convert --config examples/data-analysis.json
   
   # Convert inventory tracking with watch mode
   npm run cli -- convert --config examples/inventory-tracking.json --watch
   ```

3. **Regenerate XLSX files** (if needed):
   ```bash
   cd examples
   python3 generate_xlsx.py
   ```

## Supported Features

### Formulas
- ✅ Arithmetic operators (`+`, `-`, `*`, `/`, `^`)
- ✅ Comparison operators (`=`, `<>`, `<`, `>`, `<=`, `>=`)
- ✅ String concatenation (`&`)
- ✅ Cell references (`A1`, `$A$1`, `Sheet1!A1`)
- ✅ Range references (`A1:B10`, `Sheet1!A1:B10`, `A:A`, `D:D`)
- ✅ Named ranges (automatically resolved from Google Sheets)
- ✅ Cross-sheet references

### Functions
- ✅ Math: `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `ROUND`, `ABS`, `SQRT`
- ✅ Financial: `PMT`, `FV`, `PV`, `RATE`, `NPV`, `IRR`
- ✅ Logic: `IF`, `AND`, `OR`, `NOT`, `IFS`
- ✅ Text: `CONCATENATE`, `LEN`, `UPPER`, `LOWER`, `TRIM`, `LEFT`, `RIGHT`, `MID`
- ✅ Lookup: `VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH`, `INDIRECT`
- ✅ Date: `TODAY`, `NOW`, `DATE`
- ✅ Statistical: `STDEV`, `COUNTIF`, `SUMIF`, `SUMIFS`, `RANK`, `SMALL`, `LARGE`, `MEDIAN`, `PERCENTILE`
- ✅ Array literals: `{1,2,3;4,5,6}`

## Generated Code Examples

### TypeScript Output
```typescript
export interface SpreadsheetInput {
  Input: {
    A1?: number | string;
    B1?: number | string;
  };
}

export interface SpreadsheetOutput {
  Results: {
    A1: number | string;
    B1: number | string;
  };
}

export function calculateSpreadsheet(input: SpreadsheetInput): SpreadsheetOutput {
  const cells: Record<string, any> = {};
  
  // Initialize input values
  cells['Input!A1'] = input.Input.A1 ?? 0;
  cells['Input!B1'] = input.Input.B1 ?? 0;
  
  // Calculate formulas
  cells['Results!A1'] = (cells['Input!A1'] + cells['Input!B1']);
  cells['Results!B1'] = sum(cells['Input!A1'], cells['Input!B1']);
  
  return {
    Results: {
      A1: cells['Results!A1'],
      B1: cells['Results!B1'],
    },
  };
}
```

### Python Output
```python
def calculate_spreadsheet(input_data: Dict[str, Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    cells = {}
    
    # Initialize input values
    input_sheet = input_data.get('Input', {})
    cells['Input!A1'] = input_sheet.get('A1', 0)
    cells['Input!B1'] = input_sheet.get('B1', 0)
    
    # Calculate formulas
    cells['Results!A1'] = (cells.get('Input!A1') + cells.get('Input!B1'))
    cells['Results!B1'] = sum_values(cells.get('Input!A1'), cells.get('Input!B1'))
    
    output = {}
    output['Results'] = {}
    output['Results']['A1'] = cells.get('Results!A1')
    output['Results']['B1'] = cells.get('Results!B1')
    
    return output
```

## Development

```bash
# Install dependencies
npm install

# Run tests
npm test

# Build project
npm run build

# Run linter
npm run lint

# Type check
npm run typecheck

# Auto-fix formatting and linting
npm run format:fix && npm run lint:fix

# Run all checks
npm run typecheck && npm run lint && npm run format && npm test
```

### AI Assistant Instructions

If you're using Claude or another AI assistant to help with this project, see [CLAUDE.md](./CLAUDE.md) for specific instructions on maintaining code quality.

## Architecture

- **`src/parsers/`**: Formula lexer and parser using Chevrotain
- **`src/utils/`**: Google Sheets API client and dependency analyzer
- **`src/generators/`**: TypeScript and Python code generators
- **`src/cli/`**: Command-line interface
- **`src/types/`**: Type definitions

## Limitations

- Complex array formulas may require manual review
- Some Google Sheets specific functions don't have direct equivalents
- Formatting and visual elements are not preserved
- External data connections need special handling

## Contributing

We welcome contributions! Please ensure your code meets our quality standards:

### Code Quality Requirements

- ✅ **Linting**: All code must pass `npm run lint`
- ✅ **Formatting**: Code must be formatted with `npm run format:fix`
- ✅ **Type Safety**: No TypeScript errors (`npm run typecheck`)
- ✅ **Tests**: All tests must pass (`npm test`)
- ✅ **Pre-commit**: The pre-commit hook will automatically check these

### Development Process

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes
4. Ensure all quality checks pass:
   ```bash
   npm run lint
   npm run format:fix
   npm run typecheck
   npm test
   ```
5. Commit your changes (pre-commit hook will run automatically)
6. Push to your branch (`git push origin feature/amazing-feature`)
7. Open a Pull Request

### Maintainer Commands

Maintainers can use these special PR comments:
- `/format` - Auto-format the code in the PR

## License

MIT License - see LICENSE file for details.
