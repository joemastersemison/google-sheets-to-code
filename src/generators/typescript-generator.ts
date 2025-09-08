/** biome-ignore-all lint/suspicious/noTemplateCurlyInString: this file generates code */
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

    // Also initialize static values from output tabs (non-formula cells)
    lines.push(this.indent("// Initialize static values from output sheets"));
    for (const tabName of outputTabs) {
      const sheet = sheets.get(tabName);
      if (sheet) {
        let hasStaticValues = false;
        for (const [cellRef, cell] of sheet.cells) {
          if (
            !cell.formula &&
            cell.value !== null &&
            cell.value !== undefined &&
            cell.value !== ""
          ) {
            const cellKey = `${tabName}!${cellRef}`;
            hasStaticValues = true;
            lines.push(
              this.indent(
                `cells['${cellKey}'] = ${this.formatValue(cell.value)};`
              )
            );
          }
        }
        if (hasStaticValues) {
          lines.push("");
        }
      }
    }
    lines.push("");

    // Calculate formulas in dependency order
    lines.push(this.indent("// Calculate formulas"));
    const analyzer = new DependencyAnalyzer();
    analyzer.dependencies = dependencyGraph;
    analyzer.detectCircularDependencies();
    const calculationOrder = analyzer.getCalculationOrder();

    // Add warning comment if circular dependencies were detected
    if (analyzer.circularDependencies.size > 0) {
      lines.push(
        this.indent(
          "// ⚠️  Warning: Circular dependencies detected in the following cells:"
        )
      );
      lines.push(
        this.indent(
          `// ${Array.from(analyzer.circularDependencies).join(", ")}`
        )
      );
      lines.push(
        this.indent("// These cells will be calculated with #REF! error values")
      );
      lines.push("");
    }

    for (const nodeId of calculationOrder) {
      const node = dependencyGraph.get(nodeId);
      if (node?.formula) {
        // Check if this cell is part of a circular dependency
        if (analyzer.circularDependencies.has(nodeId)) {
          lines.push(this.indent(`// Circular dependency: ${nodeId}`));
          lines.push(
            this.indent(
              `cells['${nodeId}'] = '#REF!'; // Circular reference error`
            )
          );
        } else {
          const formulaCode = this.generateFormulaCode(
            node.formula,
            node.sheetName
          );
          lines.push(this.indent(`cells['${nodeId}'] = ${formulaCode};`));
        }
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

    // Add CLI execution code
    lines.push("");
    lines.push("// CLI execution");
    lines.push(
      "// Check if running as main module (works for both CommonJS and ES modules)"
    );
    lines.push(
      "const isMainModule = typeof require !== 'undefined' && require.main === module ||"
    );
    lines.push(
      "  typeof import.meta !== 'undefined' && import.meta.url === `file://${process.argv[1]}`;"
    );
    lines.push("");
    lines.push("if (isMainModule) {");
    lines.push("  // Default input values from the spreadsheet");
    lines.push("  const defaultInput: SpreadsheetInput = {");

    for (const tabName of inputTabs) {
      const sheet = sheets.get(tabName);
      if (sheet) {
        lines.push(`    ${this.sanitizePropertyName(tabName)}: {`);
        for (const [cellRef, cell] of sheet.cells) {
          if (!cell.formula) {
            const value = this.formatValue(cell.value);
            lines.push(`      ${cellRef}: ${value},`);
          }
        }
        lines.push("    },");
      }
    }

    lines.push("  };");
    lines.push("");
    lines.push("  // Parse command line arguments");
    lines.push("  const args = process.argv.slice(2);");
    lines.push("  let input = { ...defaultInput };");
    lines.push("  ");
    lines.push("  // Parse --input flag for JSON input");
    lines.push("  const inputIndex = args.indexOf('--input');");
    lines.push("  if (inputIndex !== -1 && args[inputIndex + 1]) {");
    lines.push("    try {");
    lines.push("      const customInput = JSON.parse(args[inputIndex + 1]);");
    lines.push("      input = { ...input, ...customInput };");
    lines.push("    } catch (e) {");
    lines.push("      console.error('Error parsing input JSON:', e);");
    lines.push("      process.exit(1);");
    lines.push("    }");
    lines.push("  }");
    lines.push("");
    lines.push("  // Calculate output");
    lines.push("  const output = calculateSpreadsheet(input);");
    lines.push("");
    lines.push("  // Display results");
    lines.push("  console.log('\\n📊 Spreadsheet Calculation Results:');");
    lines.push("  console.log('=====================================\\n');");
    lines.push("");
    lines.push("  // Display input values");
    lines.push("  console.log('📥 Input Values:');");

    for (const tabName of inputTabs) {
      lines.push(`  console.log('  ${tabName}:');`);
      lines.push(
        `  for (const [key, value] of Object.entries(input.${this.sanitizePropertyName(tabName)} || {})) {`
      );
      lines.push("    console.log(`    ${key}: ${value}`);");
      lines.push("  }");
    }

    lines.push("");
    lines.push("  // Display output values");
    lines.push("  console.log('\\n📤 Output Values:');");

    for (const tabName of outputTabs) {
      lines.push(`  console.log('  ${tabName}:');`);
      lines.push(
        `  for (const [key, value] of Object.entries(output.${this.sanitizePropertyName(tabName)} || {})) {`
      );
      lines.push("    console.log(`    ${key}: ${value}`);");
      lines.push("  }");
    }

    lines.push("");
    lines.push("  // Output JSON format if --json flag is present");
    lines.push("  if (args.includes('--json')) {");
    lines.push("    console.log('\\n📋 JSON Output:');");
    lines.push("    console.log(JSON.stringify(output, null, 2));");
    lines.push("  }");
    lines.push("}");

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
      return `getRange('${fullRef}', cells)`;
    } else {
      // Cell reference
      const fullRef = normalizedRef.includes("!")
        ? normalizedRef
        : `${currentSheet}!${normalizedRef}`;
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
          return `safeDivide(${left}, ${right})`;
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
      // Math functions
      case "SUM":
        return `sum(${args.join(", ")})`;
      case "ABS":
        return `Math.abs(${args[0]})`;
      case "SQRT":
        return `Math.sqrt(${args[0]})`;
      case "ROUND":
        return `round(${args[0]}, ${args[1] || "0"})`;
      case "EXP":
        return `Math.exp(${args[0]})`;
      case "LN":
        return `Math.log(${args[0]})`;
      case "LOG":
        if (args.length > 1) {
          return `(Math.log(${args[0]}) / Math.log(${args[1]}))`;
        }
        return `(Math.log(${args[0]}) / Math.LN10)`;
      case "TRUNC":
        return `Math.trunc(${args[0]})`;

      // Statistical functions
      case "AVERAGE":
        return `average(${args.join(", ")})`;
      case "MIN":
        return `min(${args.join(", ")})`;
      case "MAX":
        return `max(${args.join(", ")})`;
      case "COUNT":
        return `count(${args.join(", ")})`;
      case "COUNTIF":
        return `countif(${args.join(", ")})`;
      case "COUNTA":
        return `counta(${args.join(", ")})`;
      case "SUMIF":
        return `sumif(${args.join(", ")})`;
      case "SUMIFS":
        return `sumifs(${args.join(", ")})`;
      case "STDEV":
        return `stdev(${args.join(", ")})`;
      case "VAR":
      case "VARIANCE":
        return `variance(${args.join(", ")})`;
      case "MEDIAN":
        return `median(${args.join(", ")})`;
      case "PERCENTILE":
        return `percentile(${args.join(", ")})`;
      case "LARGE":
        return `large(${args.join(", ")})`;
      case "AVERAGEIF":
        return `averageif(${args.join(", ")})`;
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
        return `(${args[0]} ? ${args[1]} : ${args.length > 2 ? args[2] : "false"})`;
      case "AND":
        return `(${args.join(" && ")})`;
      case "OR":
        return `(${args.join(" || ")})`;

      // Information functions
      case "ISNUMBER":
        return `(typeof ${args[0]} === 'number' && !isNaN(${args[0]}))`;
      case "ISBLANK":
        return `(${args[0]} === null || ${args[0]} === undefined || ${args[0]} === '')`;
      case "ISTEXT":
        return `(typeof ${args[0]} === 'string')`;
      case "ISNA":
        return `(${args[0]} === '#N/A')`;
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
        return `indirect(${args.join(", ")}, cells)`;
      case "ROW":
        if (args.length > 0) {
          return `getRow(${args[0]})`;
        }
        return `getCurrentRow()`;

      // Array functions
      case "SORT":
        return `sort(${args.join(", ")})`;
      case "UNIQUE":
        return `unique(${args.join(", ")})`;
      case "RANK":
        return `rank(${args.join(", ")})`;
      case "SMALL":
        return `small(${args.join(", ")})`;

      // Text functions
      case "CONCATENATE":
        return `concatenate(${args.join(", ")})`;
      case "LEN":
        return `String(${args[0]}).length`;
      case "UPPER":
        return `String(${args[0]}).toUpperCase()`;
      case "LOWER":
        return `String(${args[0]}).toLowerCase()`;
      case "TRIM":
        return `String(${args[0]}).trim()`;

      // Date functions
      case "TODAY":
        return `new Date().toISOString().split('T')[0]`;
      case "NOW":
        return `new Date().toISOString()`;

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
        return `npv(${args.join(", ")})`;
      case "IRR":
        return `irr(${args.join(", ")})`;

      default:
        // For unknown functions, generate a generic call
        return `${functionName.toLowerCase()}(${args.join(", ")})`;
    }
  }

  private generateHelperFunctions(): string {
    return `// Helper functions
function sum(...args: any[]): number {
  return args.flat(2).reduce((acc, val) => {
    const num = typeof val === 'number' ? val : Number(val);
    return !isNaN(num) ? acc + num : acc;
  }, 0);
}

function safeDivide(numerator: number, denominator: number): number {
  return denominator === 0 ? 0 : numerator / denominator;
}

function safeDivideWithError(numerator: number, denominator: number): number | string {
  return denominator === 0 ? '#DIV/0!' : numerator / denominator;
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

// Statistical functions
function countif(range: any[], criterion: any): number {
  let count = 0;
  const criterionStr = String(criterion);
  for (const value of range.flat()) {
    if (criterionStr.startsWith('>=')) {
      const compareValue = parseFloat(criterionStr.slice(2));
      if (!isNaN(compareValue) && typeof value === 'number' && !isNaN(value)) {
        if (value >= compareValue) count++;
      }
    } else if (criterionStr.startsWith('<=')) {
      const compareValue = parseFloat(criterionStr.slice(2));
      if (!isNaN(compareValue) && typeof value === 'number' && !isNaN(value)) {
        if (value <= compareValue) count++;
      }
    } else if (criterionStr.startsWith('<>') || criterionStr.startsWith('!=')) {
      const compareValue = criterionStr.slice(criterionStr.startsWith('<>') ? 2 : 2);
      if (value != compareValue) count++;
    } else if (criterionStr.startsWith('>')) {
      const compareValue = parseFloat(criterionStr.slice(1));
      if (!isNaN(compareValue) && typeof value === 'number' && !isNaN(value)) {
        if (value > compareValue) count++;
      }
    } else if (criterionStr.startsWith('<')) {
      const compareValue = parseFloat(criterionStr.slice(1));
      if (!isNaN(compareValue) && typeof value === 'number' && !isNaN(value)) {
        if (value < compareValue) count++;
      }
    } else if (criterionStr.startsWith('=')) {
      // Handle explicit equals
      const compareValue = criterionStr.slice(1);
      if (value == compareValue) count++;
    } else {
      // Direct equality comparison
      if (value == criterion) count++;
    }
  }
  return count;
}

function counta(...args: any[]): number {
  const values = args.flat(Infinity);
  let count = 0;
  for (const value of values) {
    if (value !== null && value !== undefined && value !== '') {
      count++;
    }
  }
  return count;
}

function sumif(range: any[], criterion: any, sumRange?: any[]): number {
  const rangeValues = range.flat();
  const sumValues = sumRange ? sumRange.flat() : rangeValues;
  
  if (rangeValues.length !== sumValues.length) {
    return 0; // Range size mismatch
  }
  
  let sum = 0;
  const criterionStr = String(criterion);
  
  for (let i = 0; i < rangeValues.length; i++) {
    const value = rangeValues[i];
    let matches = false;
    
    if (criterionStr.startsWith('>=')) {
      const compareValue = parseFloat(criterionStr.slice(2));
      if (!isNaN(compareValue) && typeof value === 'number' && !isNaN(value)) {
        matches = value >= compareValue;
      }
    } else if (criterionStr.startsWith('<=')) {
      const compareValue = parseFloat(criterionStr.slice(2));
      if (!isNaN(compareValue) && typeof value === 'number' && !isNaN(value)) {
        matches = value <= compareValue;
      }
    } else if (criterionStr.startsWith('<>') || criterionStr.startsWith('!=')) {
      const compareValue = criterionStr.slice(criterionStr.startsWith('<>') ? 2 : 2);
      matches = value != compareValue;
    } else if (criterionStr.startsWith('>')) {
      const compareValue = parseFloat(criterionStr.slice(1));
      if (!isNaN(compareValue) && typeof value === 'number' && !isNaN(value)) {
        matches = value > compareValue;
      }
    } else if (criterionStr.startsWith('<')) {
      const compareValue = parseFloat(criterionStr.slice(1));
      if (!isNaN(compareValue) && typeof value === 'number' && !isNaN(value)) {
        matches = value < compareValue;
      }
    } else {
      matches = value == criterion;
    }
    
    if (matches) {
      const sumValue = Number(sumValues[i]);
      if (!isNaN(sumValue)) {
        sum += sumValue;
      }
    }
  }
  
  return sum;
}

function sumifs(sumRange: any[], ...criteria: any[]): number {
  const sumValues = sumRange.flat();
  
  // Criteria comes in pairs: range1, criterion1, range2, criterion2, etc.
  const criteriaRanges: any[][] = [];
  const criteriaValues: any[] = [];
  
  for (let i = 0; i < criteria.length; i += 2) {
    if (i + 1 < criteria.length) {
      criteriaRanges.push(criteria[i].flat ? criteria[i].flat() : [criteria[i]]);
      criteriaValues.push(criteria[i + 1]);
    }
  }
  
  let sum = 0;
  
  for (let i = 0; i < sumValues.length; i++) {
    let allMatch = true;
    
    for (let j = 0; j < criteriaRanges.length; j++) {
      const range = criteriaRanges[j];
      const criterion = criteriaValues[j];
      const criterionStr = String(criterion);
      
      if (i >= range.length) {
        allMatch = false;
        break;
      }
      
      const value = range[i];
      let matches = false;
      
      if (criterionStr.startsWith('>=')) {
        const compareValue = parseFloat(criterionStr.slice(2));
        if (!isNaN(compareValue) && typeof value === 'number' && !isNaN(value)) {
          matches = value >= compareValue;
        }
      } else if (criterionStr.startsWith('<=')) {
        const compareValue = parseFloat(criterionStr.slice(2));
        if (!isNaN(compareValue) && typeof value === 'number' && !isNaN(value)) {
          matches = value <= compareValue;
        }
      } else if (criterionStr.startsWith('<>') || criterionStr.startsWith('!=')) {
        const compareValue = criterionStr.slice(criterionStr.startsWith('<>') ? 2 : 2);
        matches = value != compareValue;
      } else if (criterionStr.startsWith('>')) {
        const compareValue = parseFloat(criterionStr.slice(1));
        if (!isNaN(compareValue) && typeof value === 'number' && !isNaN(value)) {
          matches = value > compareValue;
        }
      } else if (criterionStr.startsWith('<')) {
        const compareValue = parseFloat(criterionStr.slice(1));
        if (!isNaN(compareValue) && typeof value === 'number' && !isNaN(value)) {
          matches = value < compareValue;
        }
      } else {
        matches = value == criterion;
      }
      
      if (!matches) {
        allMatch = false;
        break;
      }
    }
    
    if (allMatch) {
      const sumValue = Number(sumValues[i]);
      if (!isNaN(sumValue)) {
        sum += sumValue;
      }
    }
  }
  
  return sum;
}

function index(array: any[], row: number, column?: number): any {
  const flatArray = array.flat ? array.flat() : array;
  
  // If column is specified, treat as 2D array
  if (column !== undefined && column > 0) {
    // This is simplified - in real sheets, would need to know actual dimensions
    return flatArray[(row - 1) * column + (column - 1)] || '#REF!';
  }
  
  // Single dimension index
  return flatArray[row - 1] || '#REF!';
}

function stdev(...args: any[]): number {
  const values = args.flat(Infinity).filter(v => typeof v === 'number' && !isNaN(v));
  const n = values.length;
  if (n < 2) return 0;
  const mean = values.reduce((a, b) => a + b, 0) / n;
  const variance = values.reduce((acc, val) => acc + Math.pow(val - mean, 2), 0) / (n - 1);
  return Math.sqrt(variance);
}

function variance(...args: any[]): number {
  const values = args.flat(Infinity).filter(v => typeof v === 'number' && !isNaN(v));
  const n = values.length;
  if (n < 2) return 0;
  const mean = values.reduce((a, b) => a + b, 0) / n;
  return values.reduce((acc, val) => acc + Math.pow(val - mean, 2), 0) / (n - 1);
}

function median(...args: any[]): number {
  const values = args.flat(Infinity).filter(v => typeof v === 'number' && !isNaN(v));
  if (values.length === 0) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 !== 0 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
}

function percentile(range: any[], k: number): number {
  const values = range.flat(Infinity).filter(v => typeof v === 'number' && !isNaN(v));
  if (values.length === 0) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  const pos = (sorted.length - 1) * k;
  const base = Math.floor(pos);
  const rest = pos - base;
  if (sorted[base + 1] !== undefined) {
    return sorted[base] + rest * (sorted[base + 1] - sorted[base]);
  }
  return sorted[base];
}

function large(range: any[], k: number): number {
  const values = range.flat(Infinity).filter(v => typeof v === 'number' && !isNaN(v));
  if (values.length === 0 || k > values.length || k < 1) return 0;
  const sorted = [...values].sort((a, b) => b - a);
  return sorted[Math.floor(k) - 1];
}

function averageif(range: any[], criterion: any, averageRange?: any[]): number {
  const rangeValues = range.flat();
  const avgValues = averageRange ? averageRange.flat() : rangeValues;
  
  if (rangeValues.length !== avgValues.length) {
    return 0; // Range size mismatch
  }
  
  let sum = 0;
  let count = 0;
  const criterionStr = String(criterion);
  
  for (let i = 0; i < rangeValues.length; i++) {
    const value = rangeValues[i];
    let matches = false;
    
    if (criterionStr.startsWith('>=')) {
      const compareValue = parseFloat(criterionStr.slice(2));
      if (!isNaN(compareValue) && typeof value === 'number' && !isNaN(value)) {
        matches = value >= compareValue;
      }
    } else if (criterionStr.startsWith('<=')) {
      const compareValue = parseFloat(criterionStr.slice(2));
      if (!isNaN(compareValue) && typeof value === 'number' && !isNaN(value)) {
        matches = value <= compareValue;
      }
    } else if (criterionStr.startsWith('<>') || criterionStr.startsWith('!=')) {
      const compareValue = criterionStr.slice(criterionStr.startsWith('<>') ? 2 : 2);
      matches = value != compareValue;
    } else if (criterionStr.startsWith('>')) {
      const compareValue = parseFloat(criterionStr.slice(1));
      if (!isNaN(compareValue) && typeof value === 'number' && !isNaN(value)) {
        matches = value > compareValue;
      }
    } else if (criterionStr.startsWith('<')) {
      const compareValue = parseFloat(criterionStr.slice(1));
      if (!isNaN(compareValue) && typeof value === 'number' && !isNaN(value)) {
        matches = value < compareValue;
      }
    } else {
      matches = value == criterion;
    }
    
    if (matches) {
      const avgValue = Number(avgValues[i]);
      if (!isNaN(avgValue)) {
        sum += avgValue;
        count++;
      }
    }
  }
  
  return count > 0 ? sum / count : 0;
}

function sumproduct(...arrays: any[]): number {
  if (arrays.length === 0) return 0;
  
  // Flatten all arrays
  const flattened = arrays.map(arr => arr.flat ? arr.flat(Infinity) : [arr]);
  
  // Find minimum length
  const minLength = Math.min(...flattened.map(arr => arr.length));
  
  let total = 0;
  for (let i = 0; i < minLength; i++) {
    let product = 1;
    for (const arr of flattened) {
      const val = Number(arr[i]);
      if (!isNaN(val)) {
        product *= val;
      } else {
        product = 0;
        break;
      }
    }
    total += product;
  }
  
  return total;
}

// Statistical distribution functions (simplified implementations)
function normsdist(z: number): number {
  // Cumulative standard normal distribution
  const a1 = 0.254829592;
  const a2 = -0.284496736;
  const a3 = 1.421413741;
  const a4 = -1.453152027;
  const a5 = 1.061405429;
  const p = 0.3275911;
  const sign = z < 0 ? -1 : 1;
  z = Math.abs(z) / Math.sqrt(2);
  const t = 1 / (1 + p * z);
  const y = 1 - ((((a5 * t + a4) * t + a3) * t + a2) * t + a1) * t * Math.exp(-z * z);
  return 0.5 * (1 + sign * y);
}

function normsinv(p: number): number {
  // Inverse cumulative standard normal distribution (approximation)
  if (p <= 0 || p >= 1) return NaN;
  const a = [2.50662823884, -18.61500062529, 41.39119773534, -25.44106049637];
  const b = [-8.47351093090, 23.08336743743, -21.06224101826, 3.13082909833];
  const c = [0.3374754822726147, 0.9761690190917186, 0.1607979714918209,
            0.0276438810333863, 0.0038405729373609, 0.0003951896511919,
            0.0000321767881768, 0.0000002888167364, 0.0000003960315187];
  
  const y = p - 0.5;
  if (Math.abs(y) < 0.42) {
    const z = y * y;
    return y * (((a[3] * z + a[2]) * z + a[1]) * z + a[0]) /
               ((((b[3] * z + b[2]) * z + b[1]) * z + b[0]) * z + 1);
  } else {
    let z = y > 0 ? Math.log(-Math.log(1 - p)) : Math.log(-Math.log(p));
    let x = c[0];
    for (let i = 1; i < 9; i++) {
      x = x * z + c[i];
    }
    return y > 0 ? x : -x;
  }
}

function chiinv(p: number, df: number): number {
  // Chi-square inverse (simplified - would need more complex implementation)
  // This is a placeholder - real implementation would use iterative methods
  return df * (1 - 2 / (9 * df) + normsinv(1 - p) * Math.sqrt(2 / (9 * df))) ** 3;
}

function finv(p: number, df1: number, df2: number): number {
  // F-distribution inverse (simplified - would need more complex implementation)
  // This is a placeholder - real implementation would use iterative methods
  return (df2 / df1) * (Math.exp(2 * normsinv(1 - p) * Math.sqrt(2 / (df1 + df2 - 2))) - 1);
}

function tinv(p: number, df: number): number {
  // T-distribution inverse (simplified - would need more complex implementation)
  // This is a placeholder - real implementation would use iterative methods
  const z = normsinv(1 - p / 2);
  return z * Math.sqrt(1 + z * z / (2 * df));
}

// Lookup and reference functions
function match(lookupValue: any, lookupArray: any[], matchType: number = 1): number {
  for (let i = 0; i < lookupArray.length; i++) {
    if (matchType === 0) {
      // Exact match
      if (lookupArray[i] === lookupValue) return i + 1;
    } else if (matchType === 1) {
      // Largest value less than or equal to lookupValue
      if (lookupArray[i] > lookupValue) return i;
      if (i === lookupArray.length - 1) return i + 1;
    } else if (matchType === -1) {
      // Smallest value greater than or equal to lookupValue
      if (lookupArray[i] <= lookupValue) return i + 1;
    }
  }
  return -1; // Not found
}

function indirect(ref: string, cells: Record<string, any>): any {
  // Access the cells object to get the referenced value
  return cells[ref] || '#REF!';
}

function getRow(cellRef: string): number {
  const match = cellRef.match(/\\d+$/);
  return match ? parseInt(match[0]) : 0;
}

function getCurrentRow(): number {
  // This would need context about the current cell being evaluated
  return 1;
}

// Array functions
function sort(array: any[], sortColumn: number = 1, ascending: boolean = true): any[] {
  const result = [...array];
  result.sort((a, b) => {
    const aVal = Array.isArray(a) ? a[sortColumn - 1] : a;
    const bVal = Array.isArray(b) ? b[sortColumn - 1] : b;
    const comparison = aVal < bVal ? -1 : aVal > bVal ? 1 : 0;
    return ascending ? comparison : -comparison;
  });
  return result;
}

function unique(array: any[]): any[] {
  const seen = new Set();
  return array.filter(item => {
    const key = JSON.stringify(item);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function rank(value: number, array: number[], order: number = 0): number {
  const sorted = [...array].sort((a, b) => order ? a - b : b - a);
  return sorted.indexOf(value) + 1;
}

function small(array: number[], k: number): number {
  const sorted = [...array].filter(v => typeof v === 'number' && !isNaN(v)).sort((a, b) => a - b);
  return sorted[k - 1] || 0;
}

function getRange(rangeRef: string, cells: Record<string, any>): any[] {
  // Parse range reference and return array of values
  const [sheetName, range] = rangeRef.split('!');
  const [start, end] = range.split(':');
  
  // Parse column and row from cell references (handles both relative A1 and absolute $A$1)
  const parseCell = (cellRef: string): [string, number] => {
    // Check for column-only reference (e.g., "A" or "$A")
    const columnOnlyMatch = cellRef.match(/^\\$?([A-Z]+)$/);
    if (columnOnlyMatch) {
      // For column-only references, return the column with row 1
      return [columnOnlyMatch[1], 1];
    }
    
    const match = cellRef.match(/^\\$?([A-Z]+)\\$?(\\d+)$/);
    if (!match) throw new Error('Invalid cell reference: ' + cellRef);
    return [match[1], parseInt(match[2])];
  };
  
  // Convert column letters to number (A=1, B=2, ..., Z=26, AA=27, etc.)
  const columnToNumber = (col: string): number => {
    let num = 0;
    for (let i = 0; i < col.length; i++) {
      num = num * 26 + (col.charCodeAt(i) - 64);
    }
    return num;
  };
  
  // Convert number to column letters
  const numberToColumn = (num: number): string => {
    let col = '';
    while (num > 0) {
      num--;
      col = String.fromCharCode(65 + (num % 26)) + col;
      num = Math.floor(num / 26);
    }
    return col;
  };
  
  // Handle column-only ranges (e.g., "D:D")
  const isColumnOnlyRange = start.match(/^\\$?[A-Z]+$/) && end.match(/^\\$?[A-Z]+$/);
  
  const [startCol, startRowDefault] = parseCell(start);
  const [endCol, endRowDefault] = parseCell(end);
  
  // For column-only ranges, use rows 1-1000
  const startRow = isColumnOnlyRange ? 1 : startRowDefault;
  const endRow = isColumnOnlyRange ? 1000 : endRowDefault;
  
  const startColNum = columnToNumber(startCol);
  const endColNum = columnToNumber(endCol);
  
  const result: any[] = [];
  
  // Iterate through the range and collect values
  for (let row = startRow; row <= endRow; row++) {
    const rowValues: any[] = [];
    for (let colNum = startColNum; colNum <= endColNum; colNum++) {
      const col = numberToColumn(colNum);
      const cellKey = sheetName + '!' + col + row;
      rowValues.push(cells[cellKey] ?? null);
    }
    // For single column ranges, push individual values
    // For multi-column ranges, push row arrays
    if (startColNum === endColNum) {
      result.push(...rowValues);
    } else {
      // Skip rows that are completely empty (all null values)
      const hasValue = rowValues.some(v => v !== null);
      if (hasValue || row <= 100) { // Include first 100 rows even if empty for consistency
        result.push(rowValues);
      }
    }
  }
  
  return result;
}

// Financial functions
function pmt(rate: number, nper: number, pv: number, fv: number = 0, type: number = 0): number {
  // PMT formula: PMT = (rate * PV) / (1 - (1 + rate)^(-nper))
  // When rate is 0
  if (rate === 0) {
    return -(pv + fv) / nper;
  }
  
  const pvif = Math.pow(1 + rate, nper);
  const payment = rate * (pv * pvif + fv) / (pvif - 1);
  
  // Adjust for payment at beginning of period
  if (type === 1) {
    return payment / (1 + rate);
  }
  
  return payment;
}

function fv(rate: number, nper: number, pmt: number, pv: number = 0, type: number = 0): number {
  // FV formula: FV = -PV * (1 + rate)^nper - PMT * (((1 + rate)^nper - 1) / rate)
  // When rate is 0
  if (rate === 0) {
    return -(pv + pmt * nper);
  }
  
  const pvif = Math.pow(1 + rate, nper);
  let fvValue = -pv * pvif;
  
  // Calculate payment portion
  const pmtMultiplier = (pvif - 1) / rate;
  if (type === 1) {
    // Payment at beginning of period
    fvValue -= pmt * pmtMultiplier * (1 + rate);
  } else {
    // Payment at end of period
    fvValue -= pmt * pmtMultiplier;
  }
  
  return fvValue;
}

function pv(rate: number, nper: number, pmt: number, fv: number = 0, type: number = 0): number {
  // PV formula: PV = -PMT * ((1 - (1 + rate)^(-nper)) / rate) - FV / (1 + rate)^nper
  // When rate is 0
  if (rate === 0) {
    return -(fv + pmt * nper);
  }
  
  const pvif = Math.pow(1 + rate, nper);
  let pvValue = -fv / pvif;
  
  // Calculate payment portion
  const pmtMultiplier = (1 - 1 / pvif) / rate;
  if (type === 1) {
    // Payment at beginning of period
    pvValue -= pmt * pmtMultiplier * (1 + rate);
  } else {
    // Payment at end of period
    pvValue -= pmt * pmtMultiplier;
  }
  
  return pvValue;
}

function rate(nper: number, pmt: number, pv: number, fv: number = 0, type: number = 0, guess: number = 0.1): number {
  // Use Newton-Raphson method to find rate
  const maxIterations = 100;
  const tolerance = 1e-6;
  
  let rate = guess;
  
  for (let i = 0; i < maxIterations; i++) {
    let f: number;
    let df: number;
    
    if (rate === 0) {
      f = pv + pmt * nper + fv;
      df = nper * (nper - 1) * pmt / 2;
    } else {
      const pvif = Math.pow(1 + rate, nper);
      const pmtFactor = type === 1 ? (1 + rate) : 1;
      
      f = pv * pvif + pmt * pmtFactor * (pvif - 1) / rate + fv;
      
      const dpvif = nper * Math.pow(1 + rate, nper - 1);
      df = pv * dpvif + pmt * pmtFactor * (dpvif * rate - (pvif - 1)) / (rate * rate);
      
      if (type === 1) {
        df += pmt * (pvif - 1) / rate;
      }
    }
    
    const newRate = rate - f / df;
    
    if (Math.abs(newRate - rate) < tolerance) {
      return newRate;
    }
    
    rate = newRate;
  }
  
  // If no convergence, return error
  return NaN;
}

function npv(rate: number, ...cashflows: any[]): number {
  // NPV formula: NPV = Σ(cashflow / (1 + rate)^period)
  // Handle arrays and flatten them
  const flatCashflows = cashflows.flat(Infinity).filter(v => v !== null && v !== undefined).map(Number);
  let npvValue = 0;
  
  for (let i = 0; i < flatCashflows.length; i++) {
    npvValue += flatCashflows[i] / Math.pow(1 + rate, i + 1);
  }
  
  return npvValue;
}

function irr(...cashflows: any[]): number {
  // Handle arrays and flatten them
  const flatCashflows = cashflows.flat(Infinity).filter(v => v !== null && v !== undefined).map(Number);
  const guess = 0.1;
  // Use Newton-Raphson method to find IRR
  const maxIterations = 100;
  const tolerance = 1e-6;
  
  let rate = guess;
  
  for (let i = 0; i < maxIterations; i++) {
    let npvValue = 0;
    let dnpv = 0;
    
    for (let j = 0; j < flatCashflows.length; j++) {
      const divisor = Math.pow(1 + rate, j);
      npvValue += flatCashflows[j] / divisor;
      dnpv -= j * flatCashflows[j] / Math.pow(1 + rate, j + 1);
    }
    
    const newRate = rate - npvValue / dnpv;
    
    if (Math.abs(newRate - rate) < tolerance) {
      return newRate;
    }
    
    rate = newRate;
  }
  
  // If no convergence, return error
  return NaN;
}`;
  }

  private sanitizePropertyName(name: string): string {
    // Convert to valid JavaScript property name
    return name.replace(/[^a-zA-Z0-9_]/g, "_");
  }

  private formatValue(value: string | number | boolean | null): string {
    if (typeof value === "string") {
      // Check if it's a numeric string
      if (!Number.isNaN(Number(value)) && value.trim() !== "") {
        return value;
      }
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
