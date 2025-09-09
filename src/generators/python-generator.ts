import type { DependencyNode, ParsedFormula, Sheet } from "../types/index.js";
import { DependencyAnalyzer } from "../utils/dependency-analyzer.js";

export class PythonGenerator {
  private indentLevel = 0;
  private indentString = "    ";
  private validationCode?: string;

  setValidationCode(code: string): void {
    this.validationCode = code;
  }

  generate(
    sheets: Map<string, Sheet>,
    dependencyGraph: Map<string, DependencyNode>,
    inputTabs: string[],
    outputTabs: string[],
    includeValidation = false
  ): string {
    const code: string[] = [];

    // Add imports
    code.push("from typing import Dict, Any, List, Union");
    code.push("import math");
    code.push("from datetime import datetime");
    code.push("");

    // Add helper functions FIRST (before the main function that uses them)
    code.push(this.generateHelperFunctions());
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

    // Add validation code if provided
    if (includeValidation && this.validationCode) {
      code.push("");
      code.push(this.validationCode);
    }

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
          // Use safe_add if either operand contains a function that might return errors
          if (
            left.includes("vlookup") ||
            right.includes("vlookup") ||
            left.includes("match") ||
            right.includes("match") ||
            left.includes("index") ||
            right.includes("index") ||
            left.includes("#")
          ) {
            return `safe_add(${left}, ${right})`;
          }
          return `(${left} + ${right})`;
        case "-":
          return `(${left} - ${right})`;
        case "*":
          // Use safe_multiply if either operand contains cell references or functions that might return errors
          if (
            left.includes("cells.get") ||
            right.includes("cells.get") ||
            left.includes("vlookup") ||
            right.includes("vlookup") ||
            left.includes("match") ||
            right.includes("match") ||
            left.includes("index") ||
            right.includes("index") ||
            left.includes("#")
          ) {
            return `safe_multiply(${left}, ${right})`;
          }
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
          // Use safe comparison if operands might contain errors
          if (
            left.includes("cells.get") ||
            right.includes("cells.get") ||
            left.includes("safe_") ||
            right.includes("safe_") ||
            left.includes("vlookup") ||
            right.includes("vlookup") ||
            left.includes("#")
          ) {
            return `safe_less(${left}, ${right})`;
          }
          return `(${left} < ${right})`;
        case ">":
          // Use safe comparison if operands might contain errors
          if (
            left.includes("cells.get") ||
            right.includes("cells.get") ||
            left.includes("safe_") ||
            right.includes("safe_") ||
            left.includes("vlookup") ||
            right.includes("vlookup") ||
            left.includes("#")
          ) {
            return `safe_greater(${left}, ${right})`;
          }
          return `(${left} > ${right})`;
        case "<=":
          // Use safe comparison if operands might contain errors
          if (
            left.includes("cells.get") ||
            right.includes("cells.get") ||
            left.includes("safe_") ||
            right.includes("safe_") ||
            left.includes("vlookup") ||
            right.includes("vlookup") ||
            left.includes("#")
          ) {
            return `safe_less_equal(${left}, ${right})`;
          }
          return `(${left} <= ${right})`;
        case ">=":
          // Use safe comparison if operands might contain errors
          if (
            left.includes("cells.get") ||
            right.includes("cells.get") ||
            left.includes("safe_") ||
            right.includes("safe_") ||
            left.includes("vlookup") ||
            right.includes("vlookup") ||
            left.includes("#")
          ) {
            return `safe_greater_equal(${left}, ${right})`;
          }
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
        return `safe_abs(${args[0]})`;
      case "SQRT":
        return `safe_sqrt(${args[0]})`;
      case "ROUND":
        return `round(${args[0]}, ${args[1] || "0"})`;
      case "EXP":
        return `safe_exp(${args[0]})`;
      case "LN":
        return `safe_log(${args[0]})`;
      case "LOG":
        if (args.length > 1) {
          return `safe_log(${args[0]}, ${args[1]})`;
        }
        return `safe_log10(${args[0]})`;
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
      case "COUNTA":
        return `counta(${args.join(", ")})`;
      case "COUNTIF":
        return `countif(${args.join(", ")})`;
      case "SUMIF":
        return `sumif(${args.join(", ")})`;
      case "SUMIFS":
        return `sumifs(${args.join(", ")})`;
      case "AVERAGEIF":
        return `averageif(${args.join(", ")})`;
      case "STDEV":
        return `stdev(${args.join(", ")})`;
      case "VAR":
      case "VARIANCE":
        return `var(${args.join(", ")})`;
      case "MEDIAN":
        return `median(${args.join(", ")})`;
      case "PERCENTILE":
        return `percentile(${args.join(", ")})`;
      case "LARGE":
        return `large(${args.join(", ")})`;
      case "SUMPRODUCT":
        return `sumproduct(${args.join(", ")})`;
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
      case "INDEX":
        return `index(${args.join(", ")})`;
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

      // Financial functions
      case "PMT":
        return `pmt(${args.join(", ")})`;
      case "FV":
        return `fv(${args.join(", ")})`;
      case "PV":
        return `pv(${args.join(", ")})`;
      case "RATE":
        return `rate(${args.join(", ")})`;
      case "NPV":
        // NPV needs to unpack range arguments (all args after the rate)
        if (args.length > 1) {
          const rate = args[0];
          const cashflows = args
            .slice(1)
            .map((arg) => (arg.includes("get_range") ? `*${arg}` : arg));
          return `npv(${rate}, ${cashflows.join(", ")})`;
        }
        return `npv(${args.join(", ")})`;
      case "IRR":
        // IRR expects a list as first argument, not unpacked values
        // Keep range as-is (it returns a list)
        if (args.length === 1) {
          return `irr(${args[0]})`;
        } else if (args.length === 2) {
          // IRR with guess parameter
          return `irr(${args[0]}, ${args[1]})`;
        }
        return `irr(${args.join(", ")})`;

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
    numeric_values = []
    for v in values:
        if v is not None and str(v).strip() != '':
            # Skip error values
            if isinstance(v, str) and v.startswith('#'):
                continue
            try:
                numeric_values.append(float(v))
            except (TypeError, ValueError):
                continue
    return sum(numeric_values)


def average_values(*args):
    """Calculate average of numeric values."""
    values = flatten_values(*args)
    numeric_values = []
    for v in values:
        if v is not None and str(v).strip() != '':
            # Skip error values
            if isinstance(v, str) and v.startswith('#'):
                continue
            try:
                numeric_values.append(float(v))
            except (TypeError, ValueError):
                continue
    return sum(numeric_values) / len(numeric_values) if numeric_values else 0


def min_values(*args):
    """Find minimum of numeric values."""
    values = flatten_values(*args)
    numeric_values = []
    for v in values:
        if v is not None and str(v).strip() != '':
            # Skip error values
            if isinstance(v, str) and v.startswith('#'):
                continue
            try:
                numeric_values.append(float(v))
            except (TypeError, ValueError):
                continue
    return min(numeric_values) if numeric_values else 0


def max_values(*args):
    """Find maximum of numeric values."""
    values = flatten_values(*args)
    numeric_values = []
    for v in values:
        if v is not None and str(v).strip() != '':
            # Skip error values
            if isinstance(v, str) and v.startswith('#'):
                continue
            try:
                numeric_values.append(float(v))
            except (TypeError, ValueError):
                continue
    return max(numeric_values) if numeric_values else 0


def count_values(*args):
    """Count non-empty values."""
    values = flatten_values(*args)
    count = 0
    for v in values:
        if v is not None and str(v).strip() != '':
            # Count error values as non-empty
            count += 1
    return count


def concatenate_values(*args):
    """Concatenate values as strings."""
    return ''.join(str(arg) for arg in args)


def safe_divide(numerator, denominator):
    """Safely divide two numbers, returning error for division by zero."""
    try:
        return numerator / denominator if denominator != 0 else '#DIV/0!'
    except:
        return '#DIV/0!'


def safe_multiply(a, b):
    """Safely multiply two values, handling non-numeric values."""
    try:
        # Handle None values
        if a is None or b is None:
            return 0
        # Handle error values
        if a == '#N/A' or b == '#N/A':
            return '#N/A'
        if a == '#DIV/0!' or b == '#DIV/0!':
            return '#DIV/0!'
        if a == '#VALUE!' or b == '#VALUE!':
            return '#VALUE!'
        # Convert to float and multiply
        return float(a) * float(b)
    except (TypeError, ValueError):
        return '#VALUE!'


def safe_add(a, b):
    """Safely add two values, handling non-numeric values."""
    try:
        # Handle None values
        if a is None or b is None:
            return 0
        # Handle error values
        if a == '#N/A' or b == '#N/A':
            return '#N/A'
        if a == '#DIV/0!' or b == '#DIV/0!':
            return '#DIV/0!'
        if a == '#VALUE!' or b == '#VALUE!':
            return '#VALUE!'
        # Convert to float and add
        return float(a) + float(b)
    except (TypeError, ValueError):
        return '#VALUE!'


def safe_less_equal(a, b):
    """Safely compare two values with <=, handling non-numeric values."""
    try:
        # Handle None values
        if a is None:
            a = 0
        if b is None:
            b = 0
        # Handle error values - error values are always False in comparisons
        if isinstance(a, str) and a.startswith('#'):
            return False
        if isinstance(b, str) and b.startswith('#'):
            return False
        # Convert to float and compare
        return float(a) <= float(b)
    except (TypeError, ValueError):
        return False


def safe_less(a, b):
    """Safely compare two values with <, handling non-numeric values."""
    try:
        # Handle None values
        if a is None:
            a = 0
        if b is None:
            b = 0
        # Handle error values
        if isinstance(a, str) and a.startswith('#'):
            return False
        if isinstance(b, str) and b.startswith('#'):
            return False
        # Convert to float and compare
        return float(a) < float(b)
    except (TypeError, ValueError):
        return False


def safe_greater_equal(a, b):
    """Safely compare two values with >=, handling non-numeric values."""
    try:
        # Handle None values
        if a is None:
            a = 0
        if b is None:
            b = 0
        # Handle error values
        if isinstance(a, str) and a.startswith('#'):
            return False
        if isinstance(b, str) and b.startswith('#'):
            return False
        # Convert to float and compare
        return float(a) >= float(b)
    except (TypeError, ValueError):
        return False


def safe_greater(a, b):
    """Safely compare two values with >, handling non-numeric values."""
    try:
        # Handle None values
        if a is None:
            a = 0
        if b is None:
            b = 0
        # Handle error values
        if isinstance(a, str) and a.startswith('#'):
            return False
        if isinstance(b, str) and b.startswith('#'):
            return False
        # Convert to float and compare
        return float(a) > float(b)
    except (TypeError, ValueError):
        return False


def safe_sqrt(value):
    """Safely calculate square root, handling non-numeric values."""
    try:
        # Handle None values
        if value is None:
            return 0
        # Handle error values
        if isinstance(value, str) and value.startswith('#'):
            return value  # Propagate the error
        # Convert to float and calculate sqrt
        num_value = float(value)
        if num_value < 0:
            return '#NUM!'  # Negative number error
        return math.sqrt(num_value)
    except (TypeError, ValueError):
        return '#VALUE!'


def safe_exp(value):
    """Safely calculate exponential, handling non-numeric values."""
    try:
        if value is None:
            return 1  # e^0 = 1
        if isinstance(value, str) and value.startswith('#'):
            return value  # Propagate the error
        return math.exp(float(value))
    except (TypeError, ValueError, OverflowError):
        return '#VALUE!'


def safe_log(value, base=None):
    """Safely calculate logarithm, handling non-numeric values."""
    try:
        if value is None or value == 0:
            return '#NUM!'  # Log of 0 is undefined
        if isinstance(value, str) and value.startswith('#'):
            return value  # Propagate the error
        num_value = float(value)
        if num_value <= 0:
            return '#NUM!'  # Log of negative number
        if base is not None:
            if isinstance(base, str) and base.startswith('#'):
                return base  # Propagate the error
            base_value = float(base)
            if base_value <= 0 or base_value == 1:
                return '#NUM!'
            return math.log(num_value) / math.log(base_value)
        return math.log(num_value)
    except (TypeError, ValueError):
        return '#VALUE!'


def safe_log10(value):
    """Safely calculate base-10 logarithm, handling non-numeric values."""
    try:
        if value is None or value == 0:
            return '#NUM!'
        if isinstance(value, str) and value.startswith('#'):
            return value  # Propagate the error
        num_value = float(value)
        if num_value <= 0:
            return '#NUM!'
        return math.log10(num_value)
    except (TypeError, ValueError):
        return '#VALUE!'


def safe_abs(value):
    """Safely calculate absolute value, handling non-numeric values."""
    try:
        if value is None:
            return 0
        if isinstance(value, str) and value.startswith('#'):
            return value  # Propagate the error
        return abs(float(value))
    except (TypeError, ValueError):
        return '#VALUE!'


def vlookup(lookup_value, table_array, col_index, exact_match=True):
    """VLOOKUP implementation."""
    # Handle None lookup value
    if lookup_value is None:
        return '#N/A'
    
    for row in table_array:
        if isinstance(row, list) and len(row) >= col_index:
            # Skip rows where the first value is None
            if row[0] is None:
                continue
            
            if exact_match:
                if row[0] == lookup_value:
                    return row[col_index - 1]
            else:
                # For non-exact match, both values must be comparable types
                try:
                    # Try numeric comparison first
                    if isinstance(row[0], (int, float)) and isinstance(lookup_value, (int, float)):
                        if row[0] >= lookup_value:
                            return row[col_index - 1]
                    # Fall back to string comparison
                    elif str(row[0]) >= str(lookup_value):
                        return row[col_index - 1]
                except (TypeError, ValueError):
                    # Skip incomparable values
                    continue
    return '#N/A'


# Statistical functions
def counta(range_values):
    """Count non-empty cells."""
    count = 0
    values = flatten_values(range_values)
    for value in values:
        if value is not None and str(value).strip() != '':
            count += 1
    return count


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


def median(values):
    """Calculate median of numeric values."""
    numeric_values = [float(v) for v in flatten_values(values) if v is not None and str(v).strip() != '']
    if not numeric_values:
        return 0
    sorted_values = sorted(numeric_values)
    n = len(sorted_values)
    if n % 2 == 0:
        return (sorted_values[n//2 - 1] + sorted_values[n//2]) / 2
    else:
        return sorted_values[n//2]


def var(*args):
    """Calculate sample variance."""
    values = flatten_values(*args)
    numeric_values = [float(v) for v in values if v is not None and str(v).strip() != '']
    n = len(numeric_values)
    if n < 2:
        return 0
    mean = sum(numeric_values) / n
    return sum((x - mean) ** 2 for x in numeric_values) / (n - 1)


def stdev(*args):
    """Calculate sample standard deviation."""
    import math
    variance = var(*args)
    return math.sqrt(variance) if variance > 0 else 0


def percentile(values, k):
    """Calculate k-th percentile (0-1 scale)."""
    numeric_values = [float(v) for v in flatten_values(values) if v is not None and str(v).strip() != '']
    if not numeric_values:
        return 0
    sorted_values = sorted(numeric_values)
    n = len(sorted_values)
    rank = k * (n - 1)
    lower = int(rank)
    upper = lower + 1
    if upper >= n:
        return sorted_values[-1]
    weight = rank - lower
    return sorted_values[lower] * (1 - weight) + sorted_values[upper] * weight


def sumif(range_values, criterion, sum_range=None):
    """Sum cells that meet a criterion."""
    range_list = flatten_values(range_values)
    sum_list = flatten_values(sum_range) if sum_range is not None else range_list
    
    # Ensure both lists are the same length
    if len(range_list) != len(sum_list):
        return 0
    
    total = 0
    criterion_str = str(criterion)
    
    for i, value in enumerate(range_list):
        match = False
        
        if criterion_str.startswith('>='):
            try:
                compare_value = float(criterion_str[2:])
                if isinstance(value, (int, float)) and value >= compare_value:
                    match = True
            except:
                pass
        elif criterion_str.startswith('<='):
            try:
                compare_value = float(criterion_str[2:])
                if isinstance(value, (int, float)) and value <= compare_value:
                    match = True
            except:
                pass
        elif criterion_str.startswith('<>') or criterion_str.startswith('!='):
            compare_value = criterion_str[2:]
            if str(value) != compare_value:
                match = True
        elif criterion_str.startswith('>'):
            try:
                compare_value = float(criterion_str[1:])
                if isinstance(value, (int, float)) and value > compare_value:
                    match = True
            except:
                pass
        elif criterion_str.startswith('<'):
            try:
                compare_value = float(criterion_str[1:])
                if isinstance(value, (int, float)) and value < compare_value:
                    match = True
            except:
                pass
        else:
            # Direct equality comparison
            if str(value) == str(criterion):
                match = True
        
        if match:
            try:
                total += float(sum_list[i])
            except:
                pass
    
    return total


def sumifs(sum_range, *criteria):
    """Sum cells that meet multiple criteria."""
    sum_values = flatten_values(sum_range)
    
    # Criteria comes in pairs: range1, criterion1, range2, criterion2, etc.
    criteria_ranges = []
    criteria_values = []
    
    for i in range(0, len(criteria), 2):
        if i + 1 < len(criteria):
            criteria_ranges.append(flatten_values(criteria[i]))
            criteria_values.append(criteria[i + 1])
    
    total = 0
    
    for i in range(len(sum_values)):
        all_match = True
        
        for j in range(len(criteria_ranges)):
            range_list = criteria_ranges[j]
            criterion = criteria_values[j]
            criterion_str = str(criterion)
            
            if i >= len(range_list):
                all_match = False
                break
            
            value = range_list[i]
            match = False
            
            if criterion_str.startswith('>='):
                try:
                    compare_value = float(criterion_str[2:])
                    if isinstance(value, (int, float)) and value >= compare_value:
                        match = True
                except:
                    pass
            elif criterion_str.startswith('<='):
                try:
                    compare_value = float(criterion_str[2:])
                    if isinstance(value, (int, float)) and value <= compare_value:
                        match = True
                except:
                    pass
            elif criterion_str.startswith('<>') or criterion_str.startswith('!='):
                compare_value = criterion_str[2:]
                if value != compare_value:
                    match = True
            elif criterion_str.startswith('>'):
                try:
                    compare_value = float(criterion_str[1:])
                    if isinstance(value, (int, float)) and value > compare_value:
                        match = True
                except:
                    pass
            elif criterion_str.startswith('<'):
                try:
                    compare_value = float(criterion_str[1:])
                    if isinstance(value, (int, float)) and value < compare_value:
                        match = True
                except:
                    pass
            else:
                if value == criterion:
                    match = True
            
            if not match:
                all_match = False
                break
        
        if all_match:
            try:
                total += float(sum_values[i])
            except:
                pass
    
    return total


def averageif(range_values, criterion, average_range=None):
    """Average cells that meet a criterion."""
    range_list = flatten_values(range_values)
    avg_list = flatten_values(average_range) if average_range is not None else range_list
    
    # Ensure both lists are the same length
    if len(range_list) != len(avg_list):
        return 0
    
    total = 0
    count = 0
    criterion_str = str(criterion)
    
    for i, value in enumerate(range_list):
        match = False
        
        if criterion_str.startswith('>='):
            try:
                compare_value = float(criterion_str[2:])
                if isinstance(value, (int, float)) and value >= compare_value:
                    match = True
            except:
                pass
        elif criterion_str.startswith('<='):
            try:
                compare_value = float(criterion_str[2:])
                if isinstance(value, (int, float)) and value <= compare_value:
                    match = True
            except:
                pass
        elif criterion_str.startswith('<>') or criterion_str.startswith('!='):
            compare_value = criterion_str[2:]
            if str(value) != compare_value:
                match = True
        elif criterion_str.startswith('>'):
            try:
                compare_value = float(criterion_str[1:])
                if isinstance(value, (int, float)) and value > compare_value:
                    match = True
            except:
                pass
        elif criterion_str.startswith('<'):
            try:
                compare_value = float(criterion_str[1:])
                if isinstance(value, (int, float)) and value < compare_value:
                    match = True
            except:
                pass
        else:
            # Direct equality comparison
            if str(value) == str(criterion):
                match = True
        
        if match:
            try:
                total += float(avg_list[i])
                count += 1
            except:
                pass
    
    return total / count if count > 0 else 0


def large(array, k):
    """Get k-th largest value from array."""
    numeric_values = [float(v) for v in flatten_values(array) if v is not None and str(v).strip() != '']
    if not numeric_values or k > len(numeric_values):
        return 0
    sorted_values = sorted(numeric_values, reverse=True)
    return sorted_values[k - 1] if k > 0 else 0


def index(array, row_num, col_num=1):
    """Return value at specified position in array."""
    if isinstance(array, list):
        if row_num > 0 and row_num <= len(array):
            row = array[row_num - 1]
            if isinstance(row, list):
                if col_num > 0 and col_num <= len(row):
                    return row[col_num - 1]
            else:
                return row
    return '#REF!'


def match(lookup_value, lookup_array, match_type=1):
    """Find position of value in array."""
    lookup_list = flatten_values(lookup_array)
    
    if match_type == 0:
        # Exact match
        try:
            return lookup_list.index(lookup_value) + 1
        except ValueError:
            return '#N/A'
    elif match_type == 1:
        # Largest value less than or equal to lookup_value
        for i, value in enumerate(lookup_list):
            if value > lookup_value:
                return i if i > 0 else '#N/A'
        return len(lookup_list)
    else:
        # Smallest value greater than or equal to lookup_value
        for i, value in enumerate(lookup_list):
            if value >= lookup_value:
                return i + 1
        return '#N/A'


def sumproduct(*arrays):
    """Calculate sum of products of corresponding array elements."""
    if not arrays:
        return 0
    
    # Flatten all arrays
    flattened = [flatten_values(arr) for arr in arrays]
    
    # Find minimum length
    min_length = min(len(arr) for arr in flattened)
    
    total = 0
    for i in range(min_length):
        product = 1
        for arr in flattened:
            try:
                product *= float(arr[i])
            except:
                product = 0
                break
        total += product
    
    return total


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
        return '#NUM!'
    
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
        # Check for column-only reference (e.g., "D" or "$D")
        column_only_match = re.match(r'^\\$?([A-Z]+)$', cell_ref)
        if column_only_match:
            # For column-only references, return the column with row 1
            return column_only_match.group(1), 1
        
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
    
    # Handle column-only ranges (e.g., "D:D")
    is_column_only_range = start_cell.replace('$', '').isalpha() and end_cell.replace('$', '').isalpha()
    
    start_col, start_row_default = parse_cell(start_cell)
    end_col, end_row_default = parse_cell(end_cell)
    
    # For column-only ranges, use rows 1-1000
    start_row = 1 if is_column_only_range else start_row_default
    end_row = 1000 if is_column_only_range else end_row_default
    
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
    
    return result


# Financial functions
def pmt(rate, nper, pv, fv=0, type=0):
    """Calculate payment for a loan."""
    if rate == 0:
        return -(pv + fv) / nper
    pvif = (1 + rate) ** nper
    pmt_factor = (1 + rate) if type == 1 else 1
    return -(pv * pvif + fv) / (pmt_factor * (pvif - 1) / rate)


def fv(rate, nper, pmt, pv=0, type=0):
    """Calculate future value of an investment."""
    if rate == 0:
        return -pv - pmt * nper
    pvif = (1 + rate) ** nper
    pmt_factor = (1 + rate) if type == 1 else 1
    return -pv * pvif - pmt * pmt_factor * (pvif - 1) / rate


def pv(rate, nper, pmt, fv=0, type=0):
    """Calculate present value of an investment."""
    if rate == 0:
        return -pmt * nper - fv
    pvif = (1 + rate) ** nper
    pmt_factor = (1 + rate) if type == 1 else 1
    return -(pmt * pmt_factor * (pvif - 1) / rate + fv) / pvif


def rate(nper, pmt, pv, fv=0, type=0, guess=0.1):
    """Calculate interest rate using Newton-Raphson method."""
    max_iterations = 100
    tolerance = 1e-6
    
    rate = guess
    
    for i in range(max_iterations):
        if rate == 0:
            f = pv + pmt * nper + fv
            df = nper * (nper - 1) * pmt / 2
        else:
            pvif = (1 + rate) ** nper
            pmt_factor = (1 + rate) if type == 1 else 1
            
            f = pv * pvif + pmt * pmt_factor * (pvif - 1) / rate + fv
            
            dpvif = nper * (1 + rate) ** (nper - 1)
            df = pv * dpvif + pmt * pmt_factor * (dpvif * rate - (pvif - 1)) / (rate * rate)
            
            if type == 1:
                df += pmt * (pvif - 1) / rate
        
        new_rate = rate - f / df
        
        if abs(new_rate - rate) < tolerance:
            return new_rate
        
        rate = new_rate
    
    # If no convergence, return NaN
    return '#NUM!'


def npv(rate, *cashflows):
    """Calculate net present value."""
    if rate == 0:
        # When rate is 0, NPV is undefined (division by zero)
        return '#NUM!'
    npv_value = 0
    for i, cashflow in enumerate(cashflows):
        npv_value += cashflow / ((1 + rate) ** (i + 1))
    return npv_value


def irr(cashflows, guess=0.1):
    """Calculate internal rate of return using Newton-Raphson method."""
    max_iterations = 100
    tolerance = 1e-6
    
    rate = guess
    
    for i in range(max_iterations):
        npv_value = 0
        dnpv = 0
        
        for j, cashflow in enumerate(cashflows):
            divisor = (1 + rate) ** j
            npv_value += cashflow / divisor
            dnpv -= j * cashflow / ((1 + rate) ** (j + 1))
        
        new_rate = rate - npv_value / dnpv
        
        if abs(new_rate - rate) < tolerance:
            return new_rate
        
        rate = new_rate
    
    # If no convergence, return NaN
    return '#NUM!'


def nper(rate, pmt, pv, fv=0, type=0):
    """Calculate number of periods."""
    if rate == 0:
        return -(pv + fv) / pmt
    pmt_factor = (1 + rate) if type == 1 else 1
    import math
    log1 = math.log((pmt * pmt_factor / rate - fv) / (pv + pmt * pmt_factor / rate))
    log2 = math.log(1 + rate)
    return log1 / log2


def ipmt(rate, per, nper, pv, fv=0, type=0):
    """Calculate interest payment for a specific period."""
    payment = pmt(rate, nper, pv, fv, type)
    balance = pv
    
    for i in range(1, int(per)):
        interest_payment = balance * rate
        principal_payment = payment - interest_payment
        balance += principal_payment
    
    return balance * rate


def ppmt(rate, per, nper, pv, fv=0, type=0):
    """Calculate principal payment for a specific period."""
    payment = pmt(rate, nper, pv, fv, type)
    ipmt_val = ipmt(rate, per, nper, pv, fv, type)
    return payment - ipmt_val`;
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
