import type { DependencyNode, ParsedFormula, Sheet } from "../types/index.js";
import { DependencyAnalyzer } from "../utils/dependency-analyzer.js";

export class TypeScriptGenerator {
  private indentLevel = 0;
  private indentString = "  ";

  generate(
    sheets: Map<string, Sheet>,
    dependencyGraph: Map<string, DependencyNode>,
    inputTabs: string[],
    outputTabs: string[]
  ): string {
    const code: string[] = [];

    // Generate interface definitions
    code.push(this.generateInterfaces(sheets, inputTabs, outputTabs));
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

    return code.join("\n");
  }

  private generateInterfaces(
    sheets: Map<string, Sheet>,
    inputTabs: string[],
    outputTabs: string[]
  ): string {
    const lines: string[] = [];

    // Input interface
    lines.push("export interface SpreadsheetInput {");
    this.indentLevel++;

    for (const tabName of inputTabs) {
      const sheet = sheets.get(tabName);
      if (sheet) {
        lines.push(this.indent(`${this.sanitizePropertyName(tabName)}: {`));
        this.indentLevel++;

        // Find all cells without formulas (input cells)
        for (const [cellRef, cell] of sheet.cells) {
          if (!cell.formula) {
            const propertyName = this.sanitizePropertyName(cellRef);
            lines.push(this.indent(`${propertyName}?: number | string;`));
          }
        }

        this.indentLevel--;
        lines.push(this.indent("};"));
      }
    }

    this.indentLevel--;
    lines.push("}");
    lines.push("");

    // Output interface
    lines.push("export interface SpreadsheetOutput {");
    this.indentLevel++;

    for (const tabName of outputTabs) {
      const sheet = sheets.get(tabName);
      if (sheet) {
        lines.push(this.indent(`${this.sanitizePropertyName(tabName)}: {`));
        this.indentLevel++;

        for (const [cellRef, _cell] of sheet.cells) {
          const propertyName = this.sanitizePropertyName(cellRef);
          lines.push(this.indent(`${propertyName}: number | string;`));
        }

        this.indentLevel--;
        lines.push(this.indent("};"));
      }
    }

    this.indentLevel--;
    lines.push("}");

    return lines.join("\n");
  }

  private generateCalculationFunction(
    sheets: Map<string, Sheet>,
    dependencyGraph: Map<string, DependencyNode>,
    inputTabs: string[],
    outputTabs: string[]
  ): string {
    const lines: string[] = [];

    lines.push(
      "export function calculateSpreadsheet(input: SpreadsheetInput): SpreadsheetOutput {"
    );
    this.indentLevel++;

    // Initialize cells object to store all values
    lines.push(this.indent("const cells: Record<string, any> = {};"));
    lines.push("");

    // Copy input values
    lines.push(this.indent("// Initialize input values"));
    for (const tabName of inputTabs) {
      const sheet = sheets.get(tabName);
      if (sheet) {
        for (const [cellRef, cell] of sheet.cells) {
          if (!cell.formula) {
            const inputPath = `input.${this.sanitizePropertyName(tabName)}.${this.sanitizePropertyName(cellRef)}`;
            const cellKey = `${tabName}!${cellRef}`;
            lines.push(
              this.indent(
                `cells['${cellKey}'] = ${inputPath} ?? ${this.formatValue(cell.value)};`
              )
            );
          }
        }
      }
    }
    lines.push("");

    // Calculate formulas in dependency order
    lines.push(this.indent("// Calculate formulas"));
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
        lines.push(this.indent(`cells['${nodeId}'] = ${formulaCode};`));
      }
    }
    lines.push("");

    // Build output object
    lines.push(this.indent("// Build output"));
    lines.push(this.indent("return {"));
    this.indentLevel++;

    for (const tabName of outputTabs) {
      const sheet = sheets.get(tabName);
      if (sheet) {
        lines.push(this.indent(`${this.sanitizePropertyName(tabName)}: {`));
        this.indentLevel++;

        for (const [cellRef, _cell] of sheet.cells) {
          const cellKey = `${tabName}!${cellRef}`;
          lines.push(
            this.indent(
              `${this.sanitizePropertyName(cellRef)}: cells['${cellKey}'],`
            )
          );
        }

        this.indentLevel--;
        lines.push(this.indent("},"));
      }
    }

    this.indentLevel--;
    lines.push(this.indent("};"));

    this.indentLevel--;
    lines.push("}");

    // Add helper functions
    lines.push("");
    lines.push(this.generateHelperFunctions());

    return lines.join("\n");
  }

  private generateFormulaCode(
    formula: ParsedFormula,
    currentSheet: string
  ): string {
    switch (formula.type) {
      case "literal":
        return this.formatValue(formula.value);

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
      return `getRange('${fullRef}', cells)`;
    } else {
      // Cell reference
      const fullRef = ref.includes("!") ? ref : `${currentSheet}!${ref}`;
      return `cells['${fullRef}']`;
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
          return `(${left} / ${right})`;
        case "^":
          return `Math.pow(${left}, ${right})`;
        case "&":
          return `String(${left}) + String(${right})`;
        case "=":
          return `(${left} === ${right})`;
        case "<>":
          return `(${left} !== ${right})`;
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

    // Map spreadsheet functions to JavaScript implementations
    switch (functionName) {
      case "SUM":
        return `sum(${args.join(", ")})`;
      case "AVERAGE":
        return `average(${args.join(", ")})`;
      case "MIN":
        return `min(${args.join(", ")})`;
      case "MAX":
        return `max(${args.join(", ")})`;
      case "COUNT":
        return `count(${args.join(", ")})`;
      case "IF":
        return `(${args[0]} ? ${args[1]} : ${args[2] || "false"})`;
      case "CONCATENATE":
        return `concatenate(${args.join(", ")})`;
      case "VLOOKUP":
        return `vlookup(${args.join(", ")})`;
      case "ROUND":
        return `round(${args[0]}, ${args[1] || "0"})`;
      case "ABS":
        return `Math.abs(${args[0]})`;
      case "SQRT":
        return `Math.sqrt(${args[0]})`;
      case "LEN":
        return `String(${args[0]}).length`;
      case "UPPER":
        return `String(${args[0]}).toUpperCase()`;
      case "LOWER":
        return `String(${args[0]}).toLowerCase()`;
      case "TRIM":
        return `String(${args[0]}).trim()`;
      case "TODAY":
        return `new Date().toISOString().split('T')[0]`;
      case "NOW":
        return `new Date().toISOString()`;
      default:
        // For unknown functions, generate a generic call
        return `${functionName.toLowerCase()}(${args.join(", ")})`;
    }
  }

  private generateHelperFunctions(): string {
    return `// Helper functions
function sum(...args: any[]): number {
  return args.flat(Infinity).reduce((a, b) => Number(a) + Number(b), 0);
}

function average(...args: any[]): number {
  const values = args.flat(Infinity).map(Number);
  return sum(...values) / values.length;
}

function min(...args: any[]): number {
  return Math.min(...args.flat(Infinity).map(Number));
}

function max(...args: any[]): number {
  return Math.max(...args.flat(Infinity).map(Number));
}

function count(...args: any[]): number {
  return args.flat(Infinity).filter(v => v != null && v !== '').length;
}

function concatenate(...args: any[]): string {
  return args.map(String).join('');
}

function round(value: number, digits: number = 0): number {
  const multiplier = Math.pow(10, digits);
  return Math.round(value * multiplier) / multiplier;
}

function vlookup(lookupValue: any, range: any[], colIndex: number, exactMatch: boolean = true): any {
  for (const row of range) {
    if (exactMatch ? row[0] === lookupValue : row[0] >= lookupValue) {
      return row[colIndex - 1];
    }
  }
  return '#N/A';
}

function getRange(rangeRef: string, cells: Record<string, any>): any[] {
  // Parse range reference and return array of values
  const [sheetName, range] = rangeRef.split('!');
  const [start, end] = range.split(':');
  
  // Simple implementation - would need enhancement for production
  const result: any[] = [];
  // TODO: Implement proper range extraction
  return result;
}`;
  }

  private sanitizePropertyName(name: string): string {
    // Convert to valid JavaScript property name
    return name.replace(/[^a-zA-Z0-9_]/g, "_");
  }

  private formatValue(value: string | number | boolean | null): string {
    if (typeof value === "string") {
      if (value === "TRUE" || value === "FALSE") {
        return value.toLowerCase();
      }
      return `"${value.replace(/"/g, '\\"')}"`;
    }
    return String(value);
  }

  private indent(text: string): string {
    return this.indentString.repeat(this.indentLevel) + text;
  }
}
