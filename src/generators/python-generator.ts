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
    analyzer.detectCircularDependencies();
    const calculationOrder = analyzer.getCalculationOrder();

    // Add warning comment if circular dependencies were detected
    if (analyzer.circularDependencies.size > 0) {
      lines.push(
        this.indent(
          "# ⚠️  Warning: Circular dependencies detected in the following cells:"
        )
      );
      lines.push(
        this.indent(`# ${Array.from(analyzer.circularDependencies).join(", ")}`)
      );
      lines.push(
        this.indent("# These cells will be calculated with #REF! error values")
      );
      lines.push("");
    }

    for (const nodeId of calculationOrder) {
      const node = dependencyGraph.get(nodeId);
      if (node?.formula) {
        // Check if this cell is part of a circular dependency
        if (analyzer.circularDependencies.has(nodeId)) {
          lines.push(this.indent(`# Circular dependency: ${nodeId}`));
          lines.push(
            this.indent(
              `cells['${nodeId}'] = '#REF!'  # Circular reference error`
            )
          );
        } else {
          const formulaCode = this.generateFormulaCode(
            node.formula,
            node.sheetName
          );
          lines.push(this.indent(`cells['${nodeId}'] = ${formulaCode}`));
        }
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
    lines.push("");
    lines.push("");

    // Add CLI execution code
    lines.push("# CLI execution");
    lines.push("if __name__ == '__main__':");
    lines.push("    import sys");
    lines.push("    import json");
    lines.push("");
    lines.push("    # Default input values from the spreadsheet");
    lines.push("    default_input = {");

    for (const tabName of inputTabs) {
      const sheet = sheets.get(tabName);
      if (sheet) {
        lines.push(`        '${tabName}': {`);
        for (const [cellRef, cell] of sheet.cells) {
          if (!cell.formula) {
            const value = this.formatPythonValue(cell.value);
            lines.push(`            '${cellRef}': ${value},`);
          }
        }
        lines.push("        },");
      }
    }

    lines.push("    }");
    lines.push("");
    lines.push("    # Parse command line arguments");
    lines.push("    input_data = dict(default_input)");
    lines.push("    ");
    lines.push("    # Parse --input flag for JSON input");
    lines.push("    if '--input' in sys.argv:");
    lines.push("        try:");
    lines.push("            input_index = sys.argv.index('--input')");
    lines.push("            if input_index + 1 < len(sys.argv):");
    lines.push(
      "                custom_input = json.loads(sys.argv[input_index + 1])"
    );
    lines.push("                for key in custom_input:");
    lines.push("                    if key in input_data:");
    lines.push(
      "                        input_data[key].update(custom_input[key])"
    );
    lines.push("                    else:");
    lines.push("                        input_data[key] = custom_input[key]");
    lines.push("        except (json.JSONDecodeError, IndexError) as e:");
    lines.push(
      "            print(f'Error parsing input JSON: {e}', file=sys.stderr)"
    );
    lines.push("            sys.exit(1)");
    lines.push("");
    lines.push("    # Calculate output");
    lines.push("    output = calculate_spreadsheet(input_data)");
    lines.push("");
    lines.push("    # Display results");
    lines.push("    print('\\n📊 Spreadsheet Calculation Results:')");
    lines.push("    print('=' * 37)");
    lines.push("    print()");
    lines.push("");
    lines.push("    # Display input values");
    lines.push("    print('📥 Input Values:')");

    for (const tabName of inputTabs) {
      lines.push(`    print('  ${tabName}:')`);
      lines.push(
        `    for key, value in input_data.get('${tabName}', {}).items():`
      );
      lines.push("        print(f'    {key}: {value}')");
    }

    lines.push("");
    lines.push("    # Display output values");
    lines.push("    print('\\n📤 Output Values:')");

    for (const tabName of outputTabs) {
      lines.push(`    print('  ${tabName}:')`);
      lines.push(`    for key, value in output.get('${tabName}', {}).items():`);
      lines.push("        print(f'    {key}: {value}')");
    }

    lines.push("");
    lines.push("    # Output JSON format if --json flag is present");
    lines.push("    if '--json' in sys.argv:");
    lines.push("        print('\\n📋 JSON Output:')");
    lines.push("        print(json.dumps(output, indent=2))");

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
    // Normalize the reference to handle quoted sheet names
    let normalizedRef = ref;

    // If reference includes sheet name (has !)
    if (ref.includes("!")) {
      // Handle quoted sheet names: 'My Sheet'!A1
      if (ref.startsWith("'")) {
        const exclamationIndex = ref.indexOf("!");
        const sheetPart = ref.substring(0, exclamationIndex);
        const cellPart = ref.substring(exclamationIndex + 1);

        // Remove surrounding quotes and unescape doubled quotes
        const sheetName = sheetPart.slice(1, -1).replace(/''/g, "'");
        normalizedRef = `${sheetName}!${cellPart}`;
      }
    }

    if (normalizedRef.includes(":")) {
      // Range reference
      const fullRef = normalizedRef.includes("!")
        ? normalizedRef
        : `${currentSheet}!${normalizedRef}`;
      return `get_range('${fullRef}', cells)`;
    } else {
      // Cell reference
      const fullRef = normalizedRef.includes("!")
        ? normalizedRef
        : `${currentSheet}!${normalizedRef}`;
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
      // Math functions
      case "SUM":
        return `sum_values(${args.join(", ")})`;
      case "ABS":
        return `abs(${args[0]})`;
      case "SQRT":
        return `math.sqrt(${args[0]})`;
      case "ROUND":
        return `round(${args[0]}, ${args[1] || "0"})`;
      case "EXP":
        return `math.exp(${args[0]})`;
      case "LN":
        return `math.log(${args[0]})`;
      case "LOG":
        if (args.length > 1) {
          return `(math.log(${args[0]}) / math.log(${args[1]}))`;
        }
        return `math.log10(${args[0]})`;
      case "TRUNC":
        return `math.trunc(${args[0]})`;

      // Statistical functions
      case "AVERAGE":
        return `average_values(${args.join(", ")})`;
      case "MIN":
        return `min_values(${args.join(", ")})`;
      case "MAX":
        return `max_values(${args.join(", ")})`;
      case "COUNT":
        return `count_values(${args.join(", ")})`;
      case "COUNTIF":
        return `countif(${args.join(", ")})`;
      case "STDEV":
        return `stdev(${args.join(", ")})`;
      case "CHIINV":
        return `chiinv(${args.join(", ")})`;
      case "FINV":
        return `finv(${args.join(", ")})`;
      case "T.INV":
      case "TINV":
        return `tinv(${args.join(", ")})`;
      case "NORMSDIST":
        return `normsdist(${args[0]})`;
      case "NORMSINV":
        return `normsinv(${args[0]})`;

      // Logical functions
      case "IF":
        return `(${args[1]} if ${args[0]} else ${args.length > 2 ? args[2] : "False"})`;
      case "AND":
        return `(${args.join(" and ")})`;
      case "OR":
        return `(${args.join(" or ")})`;

      // Information functions
      case "ISNUMBER":
        return `isinstance(${args[0]}, (int, float)) and not isinstance(${args[0]}, bool)`;
      case "ISBLANK":
        return `(${args[0]} is None or ${args[0]} == '')`;
      case "ISTEXT":
        return `isinstance(${args[0]}, str)`;
      case "ISNA":
        return `(${args[0]} == '#N/A')`;
      case "NA":
        return `'#N/A'`;

      // Lookup functions
      case "VLOOKUP":
        return `vlookup(${args.join(", ")})`;
      case "MATCH":
        return `match(${args.join(", ")})`;
      case "INDIRECT":
        return `indirect(${args.join(", ")})`;
      case "ROW":
        if (args.length > 0) {
          return `get_row(${args[0]})`;
        }
        return `get_current_row()`;

      // Array functions
      case "SORT":
        return `sort_array(${args.join(", ")})`;
      case "UNIQUE":
        return `unique(${args.join(", ")})`;
      case "RANK":
        return `rank(${args.join(", ")})`;
      case "SMALL":
        return `small(${args.join(", ")})`;

      // Text functions
      case "CONCATENATE":
        return `concatenate_values(${args.join(", ")})`;
      case "LEN":
        return `len(str(${args[0]}))`;
      case "UPPER":
        return `str(${args[0]}).upper()`;
      case "LOWER":
        return `str(${args[0]}).lower()`;
      case "TRIM":
        return `str(${args[0]}).strip()`;

      // Date functions
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


# Statistical functions
def countif(range_values, criterion):
    """Count cells that meet a criterion."""
    count = 0
    criterion_str = str(criterion)
    values = flatten_values(range_values)
    
    for value in values:
        if criterion_str.startswith('>='):
            try:
                compare_value = float(criterion_str[2:])
                if isinstance(value, (int, float)) and not isinstance(value, bool):
                    if value >= compare_value:
                        count += 1
            except (ValueError, TypeError):
                pass
        elif criterion_str.startswith('<='):
            try:
                compare_value = float(criterion_str[2:])
                if isinstance(value, (int, float)) and not isinstance(value, bool):
                    if value <= compare_value:
                        count += 1
            except (ValueError, TypeError):
                pass
        elif criterion_str.startswith('<>') or criterion_str.startswith('!='):
            compare_value = criterion_str[2:]
            if value != compare_value:
                count += 1
        elif criterion_str.startswith('>'):
            try:
                compare_value = float(criterion_str[1:])
                if isinstance(value, (int, float)) and not isinstance(value, bool):
                    if value > compare_value:
                        count += 1
            except (ValueError, TypeError):
                pass
        elif criterion_str.startswith('<'):
            try:
                compare_value = float(criterion_str[1:])
                if isinstance(value, (int, float)) and not isinstance(value, bool):
                    if value < compare_value:
                        count += 1
            except (ValueError, TypeError):
                pass
        elif criterion_str.startswith('='):
            # Handle explicit equals
            compare_value = criterion_str[1:]
            if str(value) == compare_value:
                count += 1
        else:
            # Direct equality comparison
            if value == criterion:
                count += 1
    return count


def stdev(*args):
    """Calculate sample standard deviation."""
    import math
    values = flatten_values(*args)
    numeric_values = [float(v) for v in values if v is not None and str(v).strip() != '']
    n = len(numeric_values)
    if n < 2:
        return 0
    mean = sum(numeric_values) / n
    variance = sum((x - mean) ** 2 for x in numeric_values) / (n - 1)
    return math.sqrt(variance)


# Statistical distribution functions
def normsdist(z):
    """Cumulative standard normal distribution."""
    import math
    a1 = 0.254829592
    a2 = -0.284496736
    a3 = 1.421413741
    a4 = -1.453152027
    a5 = 1.061405429
    p = 0.3275911
    sign = 1 if z >= 0 else -1
    z = abs(z) / math.sqrt(2)
    t = 1 / (1 + p * z)
    y = 1 - ((((a5 * t + a4) * t + a3) * t + a2) * t + a1) * t * math.exp(-z * z)
    return 0.5 * (1 + sign * y)


def normsinv(p):
    """Inverse cumulative standard normal distribution."""
    import math
    if p <= 0 or p >= 1:
        return float('nan')
    
    a = [2.50662823884, -18.61500062529, 41.39119773534, -25.44106049637]
    b = [-8.47351093090, 23.08336743743, -21.06224101826, 3.13082909833]
    c = [0.3374754822726147, 0.9761690190917186, 0.1607979714918209,
         0.0276438810333863, 0.0038405729373609, 0.0003951896511919,
         0.0000321767881768, 0.0000002888167364, 0.0000003960315187]
    
    y = p - 0.5
    if abs(y) < 0.42:
        z = y * y
        return y * (((a[3] * z + a[2]) * z + a[1]) * z + a[0]) / \
                   ((((b[3] * z + b[2]) * z + b[1]) * z + b[0]) * z + 1)
    else:
        z = math.log(-math.log(1 - p)) if y > 0 else math.log(-math.log(p))
        x = c[0]
        for i in range(1, 9):
            x = x * z + c[i]
        return x if y > 0 else -x


def chiinv(p, df):
    """Chi-square inverse (simplified)."""
    import math
    # Simplified implementation
    return df * (1 - 2 / (9 * df) + normsinv(1 - p) * math.sqrt(2 / (9 * df))) ** 3


def finv(p, df1, df2):
    """F-distribution inverse (simplified)."""
    import math
    # Simplified implementation
    return (df2 / df1) * (math.exp(2 * normsinv(1 - p) * math.sqrt(2 / (df1 + df2 - 2))) - 1)


def tinv(p, df):
    """T-distribution inverse (simplified)."""
    import math
    z = normsinv(1 - p / 2)
    return z * math.sqrt(1 + z * z / (2 * df))


# Lookup and reference functions
def match(lookup_value, lookup_array, match_type=1):
    """Find position of value in array."""
    for i, value in enumerate(lookup_array):
        if match_type == 0:
            # Exact match
            if value == lookup_value:
                return i + 1
        elif match_type == 1:
            # Largest value less than or equal to lookup_value
            if value > lookup_value:
                return i
            if i == len(lookup_array) - 1:
                return i + 1
        elif match_type == -1:
            # Smallest value greater than or equal to lookup_value
            if value <= lookup_value:
                return i + 1
    return -1


def indirect(ref):
    """Get value from indirect reference."""
    # This would need access to the cells dictionary
    return cells.get(ref, '#REF!')


def get_row(cell_ref):
    """Extract row number from cell reference."""
    import re
    match = re.search(r'\\d+$', cell_ref)
    return int(match.group()) if match else 0


def get_current_row():
    """Get current row being evaluated."""
    # This would need context about the current cell
    return 1


# Array functions
def sort_array(array, sort_column=1, ascending=True):
    """Sort array by specified column."""
    result = list(array)
    if result and isinstance(result[0], list):
        result.sort(key=lambda x: x[sort_column - 1], reverse=not ascending)
    else:
        result.sort(reverse=not ascending)
    return result


def unique(array):
    """Return unique values from array."""
    seen = set()
    result = []
    for item in array:
        item_key = str(item) if not isinstance(item, list) else str(tuple(item))
        if item_key not in seen:
            seen.add(item_key)
            result.append(item)
    return result


def rank(value, array, order=0):
    """Get rank of value in array."""
    sorted_array = sorted(array, reverse=(order == 0))
    try:
        return sorted_array.index(value) + 1
    except ValueError:
        return 0


def small(array, k):
    """Get k-th smallest value from array."""
    numeric_values = [v for v in flatten_values(array) if isinstance(v, (int, float))]
    sorted_values = sorted(numeric_values)
    return sorted_values[k - 1] if k <= len(sorted_values) else 0


def get_range(range_ref: str, cells: dict) -> list:
    """Extract values from a range reference."""
    import re
    
    # Parse range reference and return list of values
    sheet_name, range_part = range_ref.split('!')
    start_cell, end_cell = range_part.split(':')
    
    def parse_cell(cell_ref: str) -> tuple:
        """Parse column and row from cell reference (handles both relative A1 and absolute $A$1)."""
        match = re.match(r'^\\$?([A-Z]+)\\$?(\\d+)$', cell_ref)
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
      if (!Number.isNaN(Number(value)) && value.trim() !== "") {
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
