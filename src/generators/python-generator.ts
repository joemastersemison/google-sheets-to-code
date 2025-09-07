import type { DependencyNode, ParsedFormula, Sheet } from "../types/index.js";
import { DependencyAnalyzer } from "../utils/dependency-analyzer.js";

export class PythonGenerator {
  private indentLevel = 0;
  private indentString = "    ";

  generate(
    sheets: Map<string, Sheet>,
    dependencyGraph: Map<string, DependencyNode>,
    inputTabs: string[],
    outputTabs: string[]
  ): string {
    const code: string[] = [];

    // Add imports
    code.push("from typing import Dict, Any, List, Union");
    code.push("import math");
    code.push("from datetime import datetime");
    code.push("");

    // Generate the main calculation function
    code.push(
      this.generateCalculationFunction(
        sheets,
        dependencyGraph,
        inputTabs,
        outputTabs
      )
    );
    code.push("");

    // Add helper functions
    code.push(this.generateHelperFunctions());

    return code.join("\n");
  }

  private generateCalculationFunction(
    sheets: Map<string, Sheet>,
    dependencyGraph: Map<string, DependencyNode>,
    inputTabs: string[],
    outputTabs: string[]
  ): string {
    const lines: string[] = [];

    lines.push(
      "def calculate_spreadsheet(input_data: Dict[str, Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:"
    );
    this.indentLevel++;

    // Docstring
    lines.push(this.indent('"""'));
    lines.push(
      this.indent("Calculate spreadsheet formulas based on input data.")
    );
    lines.push(this.indent(""));
    lines.push(this.indent("Args:"));
    lines.push(
      this.indent(
        "    input_data: Dictionary with sheet names as keys and cell values as nested dict"
      )
    );
    lines.push(this.indent(""));
    lines.push(this.indent("Returns:"));
    lines.push(this.indent("    Dictionary with calculated output values"));
    lines.push(this.indent('"""'));

    // Initialize cells dictionary
    lines.push(this.indent("cells = {}"));
    lines.push("");

    // Copy input values
    lines.push(this.indent("# Initialize input values"));
    for (const tabName of inputTabs) {
      const sheet = sheets.get(tabName);
      if (sheet) {
        const sanitizedTabName = this.sanitizePropertyName(tabName);
        lines.push(
          this.indent(`input_sheet = input_data.get('${sanitizedTabName}', {})`)
        );

        for (const [cellRef, cell] of sheet.cells) {
          if (!cell.formula) {
            const sanitizedCellRef = this.sanitizePropertyName(cellRef);
            const cellKey = `${tabName}!${cellRef}`;
            const defaultValue = this.formatPythonValue(cell.value);
            lines.push(
              this.indent(
                `cells['${cellKey}'] = input_sheet.get('${sanitizedCellRef}', ${defaultValue})`
              )
            );
          }
        }
        lines.push("");
      }
    }

    // Calculate formulas in dependency order
    lines.push(this.indent("# Calculate formulas"));
    const analyzer = new DependencyAnalyzer();
    analyzer.dependencies = dependencyGraph;
    const calculationOrder = analyzer.getCalculationOrder();

    for (const nodeId of calculationOrder) {
      const node = dependencyGraph.get(nodeId);
      if (node?.formula) {
        const formulaCode = this.generateFormulaCode(
          node.formula,
          node.sheetName
        );
        lines.push(this.indent(`cells['${nodeId}'] = ${formulaCode}`));
      }
    }
    lines.push("");

    // Build output dictionary
    lines.push(this.indent("# Build output"));
    lines.push(this.indent("output = {}"));

    for (const tabName of outputTabs) {
      const sheet = sheets.get(tabName);
      if (sheet) {
        const sanitizedTabName = this.sanitizePropertyName(tabName);
        lines.push(this.indent(`output['${sanitizedTabName}'] = {}`));

        for (const [cellRef, _cell] of sheet.cells) {
          const sanitizedCellRef = this.sanitizePropertyName(cellRef);
          const cellKey = `${tabName}!${cellRef}`;
          lines.push(
            this.indent(
              `output['${sanitizedTabName}']['${sanitizedCellRef}'] = cells.get('${cellKey}')`
            )
          );
        }
        lines.push("");
      }
    }

    lines.push(this.indent("return output"));
    this.indentLevel--;

    return lines.join("\n");
  }

  private generateFormulaCode(
    formula: ParsedFormula,
    currentSheet: string
  ): string {
    switch (formula.type) {
      case "literal":
        return this.formatPythonValue(formula.value);

      case "reference":
        return this.generateReferenceCode(formula.value, currentSheet);

      case "operator":
        return this.generateOperatorCode(formula, currentSheet);

      case "function":
        return this.generateFunctionCode(formula, currentSheet);

      default:
        throw new Error(`Unknown formula type: ${formula.type}`);
    }
  }

  private generateReferenceCode(ref: string, currentSheet: string): string {
    if (ref.includes(":")) {
      // Range reference
      const fullRef = ref.includes("!") ? ref : `${currentSheet}!${ref}`;
      return `get_range('${fullRef}', cells)`;
    } else {
      // Cell reference
      const fullRef = ref.includes("!") ? ref : `${currentSheet}!${ref}`;
      return `cells.get('${fullRef}')`;
    }
  }

  private generateOperatorCode(
    formula: ParsedFormula,
    currentSheet: string
  ): string {
    const operator = formula.value;
    const children = formula.children || [];

    if (children.length === 1) {
      // Unary operator
      const operand = this.generateFormulaCode(children[0], currentSheet);
      switch (operator) {
        case "-":
          return `-${operand}`;
        case "+":
          return `+${operand}`;
        case "%":
          return `(${operand} / 100)`;
        default:
          throw new Error(`Unknown unary operator: ${operator}`);
      }
    } else if (children.length === 2) {
      // Binary operator
      const left = this.generateFormulaCode(children[0], currentSheet);
      const right = this.generateFormulaCode(children[1], currentSheet);

      switch (operator) {
        case "+":
          return `(${left} + ${right})`;
        case "-":
          return `(${left} - ${right})`;
        case "*":
          return `(${left} * ${right})`;
        case "/":
          return `safe_divide(${left}, ${right})`;
        case "^":
          return `(${left} ** ${right})`;
        case "&":
          return `str(${left}) + str(${right})`;
        case "=":
          return `(${left} == ${right})`;
        case "<>":
          return `(${left} != ${right})`;
        case "<":
          return `(${left} < ${right})`;
        case ">":
          return `(${left} > ${right})`;
        case "<=":
          return `(${left} <= ${right})`;
        case ">=":
          return `(${left} >= ${right})`;
        default:
          throw new Error(`Unknown binary operator: ${operator}`);
      }
    } else {
      throw new Error(`Invalid operator with ${children.length} operands`);
    }
  }

  private generateFunctionCode(
    formula: ParsedFormula,
    currentSheet: string
  ): string {
    const functionName = formula.value.toUpperCase();
    const args = (formula.children || []).map((child) =>
      this.generateFormulaCode(child, currentSheet)
    );

    // Map spreadsheet functions to Python implementations
    switch (functionName) {
      case "SUM":
        return `sum_values(${args.join(", ")})`;
      case "AVERAGE":
        return `average_values(${args.join(", ")})`;
      case "MIN":
        return `min_values(${args.join(", ")})`;
      case "MAX":
        return `max_values(${args.join(", ")})`;
      case "COUNT":
        return `count_values(${args.join(", ")})`;
      case "IF":
        return `(${args[1]} if ${args[0]} else ${args[2] || "False"})`;
      case "CONCATENATE":
        return `concatenate_values(${args.join(", ")})`;
      case "VLOOKUP":
        return `vlookup(${args.join(", ")})`;
      case "ROUND":
        return `round(${args[0]}, ${args[1] || "0"})`;
      case "ABS":
        return `abs(${args[0]})`;
      case "SQRT":
        return `math.sqrt(${args[0]})`;
      case "LEN":
        return `len(str(${args[0]}))`;
      case "UPPER":
        return `str(${args[0]}).upper()`;
      case "LOWER":
        return `str(${args[0]}).lower()`;
      case "TRIM":
        return `str(${args[0]}).strip()`;
      case "TODAY":
        return `datetime.now().strftime('%Y-%m-%d')`;
      case "NOW":
        return `datetime.now().isoformat()`;
      default:
        // For unknown functions, generate a generic call
        return `${functionName.toLowerCase()}(${args.join(", ")})`;
    }
  }

  private generateHelperFunctions(): string {
    return `# Helper functions
def flatten_values(*args):
    """Flatten nested lists/values into a single list."""
    result = []
    for item in args:
        if isinstance(item, list):
            result.extend(flatten_values(*item))
        else:
            result.append(item)
    return result


def sum_values(*args):
    """Sum all numeric values."""
    values = flatten_values(*args)
    return sum(float(v) for v in values if v is not None and str(v).strip() != '')


def average_values(*args):
    """Calculate average of numeric values."""
    values = flatten_values(*args)
    numeric_values = [float(v) for v in values if v is not None and str(v).strip() != '']
    return sum(numeric_values) / len(numeric_values) if numeric_values else 0


def min_values(*args):
    """Find minimum of numeric values."""
    values = flatten_values(*args)
    numeric_values = [float(v) for v in values if v is not None and str(v).strip() != '']
    return min(numeric_values) if numeric_values else 0


def max_values(*args):
    """Find maximum of numeric values."""
    values = flatten_values(*args)
    numeric_values = [float(v) for v in values if v is not None and str(v).strip() != '']
    return max(numeric_values) if numeric_values else 0


def count_values(*args):
    """Count non-empty values."""
    values = flatten_values(*args)
    return sum(1 for v in values if v is not None and str(v).strip() != '')


def concatenate_values(*args):
    """Concatenate values as strings."""
    return ''.join(str(arg) for arg in args)


def safe_divide(numerator, denominator):
    """Safely divide two numbers, returning error for division by zero."""
    try:
        return numerator / denominator if denominator != 0 else '#DIV/0!'
    except:
        return '#DIV/0!'


def vlookup(lookup_value, table_array, col_index, exact_match=True):
    """VLOOKUP implementation."""
    for row in table_array:
        if isinstance(row, list) and len(row) >= col_index:
            if exact_match:
                if row[0] == lookup_value:
                    return row[col_index - 1]
            else:
                if row[0] >= lookup_value:
                    return row[col_index - 1]
    return '#N/A'


def get_range(range_ref: str, cells: dict) -> list:
    """Extract values from a range reference."""
    import re
    
    # Parse range reference and return list of values
    sheet_name, range_part = range_ref.split('!')
    start_cell, end_cell = range_part.split(':')
    
    def parse_cell(cell_ref: str) -> tuple:
        """Parse column and row from cell reference."""
        match = re.match(r'^([A-Z]+)(d+)$', cell_ref)
        if not match:
            raise ValueError(f'Invalid cell reference: {cell_ref}')
        return match.group(1), int(match.group(2))
    
    def column_to_number(col: str) -> int:
        """Convert column letters to number (A=1, B=2, ..., Z=26, AA=27, etc.)"""
        num = 0
        for char in col:
            num = num * 26 + (ord(char) - 64)
        return num
    
    def number_to_column(num: int) -> str:
        """Convert number to column letters."""
        col = ''
        while num > 0:
            num -= 1
            col = chr(65 + (num % 26)) + col
            num = num // 26
        return col
    
    start_col, start_row = parse_cell(start_cell)
    end_col, end_row = parse_cell(end_cell)
    
    start_col_num = column_to_number(start_col)
    end_col_num = column_to_number(end_col)
    
    result = []
    
    # Iterate through the range and collect values
    for row in range(start_row, end_row + 1):
        row_values = []
        for col_num in range(start_col_num, end_col_num + 1):
            col = number_to_column(col_num)
            cell_key = f'{sheet_name}!{col}{row}'
            row_values.append(cells.get(cell_key))
        
        # For single column ranges, extend with individual values
        # For multi-column ranges, append row arrays
        if start_col_num == end_col_num:
            result.extend(row_values)
        else:
            result.append(row_values)
    
    return result`;
  }

  private sanitizePropertyName(name: string): string {
    // Convert to valid Python identifier
    // Replace invalid characters with underscores
    let sanitized = name.replace(/[^a-zA-Z0-9_]/g, "_");

    // Ensure it doesn't start with a number
    if (/^\d/.test(sanitized)) {
      sanitized = `_${sanitized}`;
    }

    return sanitized;
  }

  private formatPythonValue(value: string | number | boolean | null): string {
    if (typeof value === "string") {
      // Check if it's a numeric string
      if (!isNaN(Number(value)) && value.trim() !== "") {
        return value;
      }
      if (value === "TRUE") {
        return "True";
      }
      if (value === "FALSE") {
        return "False";
      }
      // Escape quotes in string
      const escaped = value.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
      return `'${escaped}'`;
    }
    if (value === null || value === undefined) {
      return "None";
    }
    return String(value);
  }

  private indent(text: string): string {
    return this.indentString.repeat(this.indentLevel) + text;
  }
}
