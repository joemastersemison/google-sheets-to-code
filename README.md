# Google Sheets to Code Converter

[![CI/CD Pipeline](https://github.com/YOUR_USERNAME/google-sheets-to-code/actions/workflows/test.yml/badge.svg)](https://github.com/YOUR_USERNAME/google-sheets-to-code/actions/workflows/test.yml)

Convert Google Sheets formulas and logic into executable TypeScript or Python code.

## Features

- 🔍 **Formula Parser**: Supports complex Google Sheets formulas including functions, operators, and cell references
- 📊 **Dependency Analysis**: Automatically determines calculation order and detects circular references
- 🎯 **Code Generation**: Generates clean, readable TypeScript or Python code
- 🔧 **CLI Interface**: Easy-to-use command-line tool
- ⚡ **Performance**: Optimized for large spreadsheets with hundreds of formulas

## Installation

```bash
npm install
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

## Supported Features

### Formulas
- ✅ Arithmetic operators (`+`, `-`, `*`, `/`, `^`)
- ✅ Comparison operators (`=`, `<>`, `<`, `>`, `<=`, `>=`)
- ✅ String concatenation (`&`)
- ✅ Cell references (`A1`, `$A$1`, `Sheet1!A1`)
- ✅ Range references (`A1:B10`, `Sheet1!A1:B10`)

### Functions
- ✅ Math: `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `ROUND`, `ABS`, `SQRT`
- ✅ Logic: `IF`, `AND`, `OR`, `NOT`
- ✅ Text: `CONCATENATE`, `LEN`, `UPPER`, `LOWER`, `TRIM`
- ✅ Lookup: `VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH`
- ✅ Date: `TODAY`, `NOW`, `DATE`
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
```

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