# Google Sheets to Code Converter

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

## Setup

1. **Google Sheets API Credentials**:
   ```bash
   npm run cli setup
   ```

2. **Create credentials.json**:
   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Enable Google Sheets API
   - Create service account credentials
   - Download and save as `credentials.json`

## Usage

### Command Line

```bash
# Basic conversion
npm run cli convert \
  --url "https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit" \
  --input-tabs "Input,Parameters" \
  --output-tabs "Results,Summary" \
  --language typescript \
  --output-file spreadsheet-logic.ts

# Using configuration file
npm run cli convert --config config.json

# Validate configuration
npm run cli validate --config config.json
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

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests
5. Submit a pull request

## License

MIT License - see LICENSE file for details.