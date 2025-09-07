import { TypeScriptGenerator } from "../generators/typescript-generator.js";
import type {
  Cell,
  DependencyNode,
  ParsedFormula,
  Sheet,
} from "../types/index.js";

describe("TypeScriptGenerator", () => {
  let generator: TypeScriptGenerator;

  beforeEach(() => {
    generator = new TypeScriptGenerator();
  });

  function createSheet(
    name: string,
    cells: Record<string, Partial<Cell>>
  ): Sheet {
    const cellMap = new Map<string, Cell>();

    for (const [cellRef, cellData] of Object.entries(cells)) {
      const [column, row] = [
        cellRef.match(/[A-Z]+/)?.[0] || "A",
        cellRef.match(/\d+/)?.[0] || "1",
      ];
      cellMap.set(cellRef, {
        row: parseInt(row, 10),
        column,
        value: cellData.value || 0,
        formula: cellData.formula,
        parsedFormula: cellData.parsedFormula,
        ...cellData,
      });
    }

    return {
      name,
      cells: cellMap,
      range: {
        startRow: 1,
        endRow: 10,
        startColumn: "A",
        endColumn: "Z",
      },
    };
  }

  function createParsedFormula(
    type: string,
    value: string,
    children?: ParsedFormula[]
  ): ParsedFormula {
    return {
      type: type as "function" | "reference" | "literal" | "operator",
      value,
      children,
    };
  }

  function createDependencyNode(
    cellRef: string,
    sheetName: string,
    dependencies: string[],
    formula?: ParsedFormula
  ): DependencyNode {
    return {
      cellRef,
      sheetName,
      dependencies: new Set(dependencies),
      formula,
    };
  }

  describe("generate", () => {
    it("should generate basic TypeScript interfaces", () => {
      const sheets = new Map([
        [
          "Input",
          createSheet("Input", {
            A1: { value: 10 },
            B1: { value: 20 },
          }),
        ],
        [
          "Output",
          createSheet("Output", {
            A1: {
              formula: "=Input!A1+Input!B1",
              parsedFormula: createParsedFormula("operator", "+", [
                createParsedFormula("reference", "Input!A1"),
                createParsedFormula("reference", "Input!B1"),
              ]),
            },
          }),
        ],
      ]);

      const dependencyGraph = new Map([
        [
          "Output!A1",
          createDependencyNode(
            "A1",
            "Output",
            ["Input!A1", "Input!B1"],
            createParsedFormula("operator", "+", [
              createParsedFormula("reference", "Input!A1"),
              createParsedFormula("reference", "Input!B1"),
            ])
          ),
        ],
      ]);

      const code = generator.generate(
        sheets,
        dependencyGraph,
        ["Input"],
        ["Output"]
      );

      expect(code).toContain("export interface SpreadsheetInput");
      expect(code).toContain("export interface SpreadsheetOutput");
      expect(code).toContain("export function calculateSpreadsheet");
      expect(code).toContain("Input: {");
      expect(code).toContain("A1?: number | string;");
      expect(code).toContain("B1?: number | string;");
      expect(code).toContain("Output: {");
      expect(code).toContain("A1: number | string;");
    });

    it("should generate calculation function with correct order", () => {
      const sheets = new Map([
        [
          "Data",
          createSheet("Data", {
            A1: { value: 100 },
            B1: {
              formula: "=A1*2",
              parsedFormula: createParsedFormula("operator", "*", [
                createParsedFormula("reference", "A1"),
                createParsedFormula("literal", "2"),
              ]),
            },
          }),
        ],
      ]);

      const dependencyGraph = new Map([
        [
          "Data!B1",
          createDependencyNode(
            "B1",
            "Data",
            ["Data!A1"],
            createParsedFormula("operator", "*", [
              createParsedFormula("reference", "A1"),
              createParsedFormula("literal", "2"),
            ])
          ),
        ],
      ]);

      const code = generator.generate(
        sheets,
        dependencyGraph,
        ["Data"],
        ["Data"]
      );

      expect(code).toContain("cells['Data!A1'] = input.Data.A1 ?? 100;");
      expect(code).toContain("cells['Data!B1'] = (cells['Data!A1'] * 2);");
    });
  });

  describe("formula code generation", () => {
    let sheets: Map<string, Sheet>;
    let dependencyGraph: Map<string, DependencyNode>;

    beforeEach(() => {
      sheets = new Map([["Test", createSheet("Test", {})]]);
      dependencyGraph = new Map();
    });

    it("should generate arithmetic operations", () => {
      const formula = createParsedFormula("operator", "+", [
        createParsedFormula("literal", "5"),
        createParsedFormula("literal", "3"),
      ]);

      dependencyGraph.set(
        "Test!A1",
        createDependencyNode("A1", "Test", [], formula)
      );

      const code = generator.generate(sheets, dependencyGraph, [], ["Test"]);

      expect(code).toContain("(5 + 3)");
    });

    it("should generate function calls", () => {
      const formula = createParsedFormula("function", "SUM", [
        createParsedFormula("reference", "A1:A10"),
      ]);

      dependencyGraph.set(
        "Test!B1",
        createDependencyNode("B1", "Test", [], formula)
      );

      const code = generator.generate(sheets, dependencyGraph, [], ["Test"]);

      expect(code).toContain("sum(getRange(");
    });

    it("should generate IF statements", () => {
      const formula = createParsedFormula("function", "IF", [
        createParsedFormula("operator", ">", [
          createParsedFormula("reference", "A1"),
          createParsedFormula("literal", "0"),
        ]),
        createParsedFormula("literal", "Positive"),
        createParsedFormula("literal", "Negative"),
      ]);

      dependencyGraph.set(
        "Test!C1",
        createDependencyNode("C1", "Test", [], formula)
      );

      const code = generator.generate(sheets, dependencyGraph, [], ["Test"]);

      expect(code).toContain(
        '(cells[\'Test!A1\'] > 0) ? "Positive" : "Negative"'
      );
    });

    it("should generate string concatenation", () => {
      const formula = createParsedFormula("operator", "&", [
        createParsedFormula("literal", "Hello"),
        createParsedFormula("literal", "World"),
      ]);

      dependencyGraph.set(
        "Test!D1",
        createDependencyNode("D1", "Test", [], formula)
      );

      const code = generator.generate(sheets, dependencyGraph, [], ["Test"]);

      expect(code).toContain('String("Hello") + String("World")');
    });

    it("should handle nested functions", () => {
      const formula = createParsedFormula("function", "ROUND", [
        createParsedFormula("function", "AVERAGE", [
          createParsedFormula("reference", "A1:A10"),
        ]),
        createParsedFormula("literal", "2"),
      ]);

      dependencyGraph.set(
        "Test!E1",
        createDependencyNode("E1", "Test", [], formula)
      );

      const code = generator.generate(sheets, dependencyGraph, [], ["Test"]);

      expect(code).toContain("round(average(");
    });

    it("should handle percentage operations", () => {
      const formula = createParsedFormula("operator", "%", [
        createParsedFormula("literal", "50"),
      ]);

      dependencyGraph.set(
        "Test!F1",
        createDependencyNode("F1", "Test", [], formula)
      );

      const code = generator.generate(sheets, dependencyGraph, [], ["Test"]);

      expect(code).toContain("(50 / 100)");
    });
  });

  describe("helper functions generation", () => {
    it("should include helper functions", () => {
      const sheets = new Map([["Test", createSheet("Test", {})]]);
      const dependencyGraph = new Map();

      const code = generator.generate(sheets, dependencyGraph, [], []);

      expect(code).toContain("function sum(...args: any[]): number");
      expect(code).toContain("function average(...args: any[]): number");
      expect(code).toContain("function min(...args: any[]): number");
      expect(code).toContain("function max(...args: any[]): number");
      expect(code).toContain("function count(...args: any[]): number");
      expect(code).toContain("function concatenate(...args: any[]): string");
      expect(code).toContain(
        "function round(value: number, digits: number = 0): number"
      );
      expect(code).toContain("function vlookup");
      expect(code).toContain("function getRange");
    });
  });

  describe("value formatting", () => {
    it("should format string values correctly", () => {
      const sheets = new Map([
        [
          "Input",
          createSheet("Input", {
            A1: { value: 'Hello "World"' },
          }),
        ],
      ]);

      const code = generator.generate(sheets, new Map(), ["Input"], []);

      expect(code).toContain('input.Input.A1 ?? "Hello \\"World\\""');
    });

    it("should format boolean values correctly", () => {
      const sheets = new Map([
        [
          "Input",
          createSheet("Input", {
            A1: { value: "TRUE" },
            B1: { value: "FALSE" },
          }),
        ],
      ]);

      const code = generator.generate(sheets, new Map(), ["Input"], []);

      expect(code).toContain("input.Input.A1 ?? true");
      expect(code).toContain("input.Input.B1 ?? false");
    });

    it("should format numeric values correctly", () => {
      const sheets = new Map([
        [
          "Input",
          createSheet("Input", {
            A1: { value: 42 },
            B1: { value: 3.14 },
          }),
        ],
      ]);

      const code = generator.generate(sheets, new Map(), ["Input"], []);

      expect(code).toContain("input.Input.A1 ?? 42");
      expect(code).toContain("input.Input.B1 ?? 3.14");
    });
  });

  describe("property name sanitization", () => {
    it("should sanitize invalid property names", () => {
      const sheets = new Map([
        [
          "My Sheet",
          createSheet("My Sheet", {
            A1: { value: 10 },
          }),
        ],
      ]);

      const code = generator.generate(sheets, new Map(), ["My Sheet"], []);

      expect(code).toContain("My_Sheet: {");
    });
  });
});
