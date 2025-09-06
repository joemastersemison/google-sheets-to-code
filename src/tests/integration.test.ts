import { PythonGenerator } from "../generators/python-generator.js";
import { TypeScriptGenerator } from "../generators/typescript-generator.js";
import { FormulaParser } from "../parsers/formula-parser.js";
import type { Cell, Sheet } from "../types/index.js";
import { DependencyAnalyzer } from "../utils/dependency-analyzer.js";

describe("Integration Tests", () => {
  describe("End-to-End Formula Processing", () => {
    it("should parse and analyze a complete spreadsheet workflow", () => {
      const parser = new FormulaParser();
      const _analyzer = new DependencyAnalyzer();

      // Test complex formula parsing
      const formulas = [
        "=A1+B1",
        "=SUM(A1:A10)",
        "=IF(C1>0,D1*0.1,0)",
        "=VLOOKUP(E1,Sheet2!A:B,2,FALSE)",
      ];

      formulas.forEach((formula) => {
        expect(() => parser.parse(formula)).not.toThrow();
      });
    });

    it("should handle real-world financial calculation scenario", () => {
      const parser = new FormulaParser();

      // Financial formulas
      const financialFormulas = [
        "=A1*B1", // Price * Quantity
        "=SUM(C1:C10)", // Total Sales
        "=D1*0.08", // Tax calculation
        "=E1-F1", // Net profit
        "=ROUND(G1/H1,2)", // Ratio calculation
      ];

      financialFormulas.forEach((formula) => {
        const ast = parser.parse(formula);
        expect(ast).toBeDefined();
        expect(ast.type).toBeDefined();
        expect(ast.value).toBeDefined();
      });
    });
  });

  describe("Code Generation Integration", () => {
    function createMockSheets(): Map<string, Sheet> {
      const inputSheet: Sheet = {
        name: "Input",
        cells: new Map([
          ["A1", { row: 1, column: "A", value: 100 }],
          ["B1", { row: 1, column: "B", value: 0.08 }],
          ["C1", { row: 1, column: "C", value: "Product A" }],
        ]),
        range: { startRow: 1, endRow: 1, startColumn: "A", endColumn: "C" },
      };

      const calculationSheet: Sheet = {
        name: "Calculations",
        cells: new Map([
          [
            "A1",
            {
              row: 1,
              column: "A",
              value: "=Input!A1*Input!B1",
              formula: "=Input!A1*Input!B1",
              parsedFormula: {
                type: "operator",
                value: "*",
                children: [
                  { type: "reference", value: "Input!A1" },
                  { type: "reference", value: "Input!B1" },
                ],
              },
            },
          ],
          [
            "B1",
            {
              row: 1,
              column: "B",
              value: "=ROUND(A1,2)",
              formula: "=ROUND(A1,2)",
              parsedFormula: {
                type: "function",
                value: "ROUND",
                children: [
                  { type: "reference", value: "A1" },
                  { type: "literal", value: "2" },
                ],
              },
            },
          ],
        ]),
        range: { startRow: 1, endRow: 1, startColumn: "A", endColumn: "B" },
      };

      return new Map([
        ["Input", inputSheet],
        ["Calculations", calculationSheet],
      ]);
    }

    it("should generate valid TypeScript code", () => {
      const sheets = createMockSheets();
      const analyzer = new DependencyAnalyzer();
      const dependencyGraph = analyzer.buildDependencyGraph(sheets);
      const generator = new TypeScriptGenerator();

      const code = generator.generate(
        sheets,
        dependencyGraph,
        ["Input"],
        ["Calculations"]
      );

      // Verify structure
      expect(code).toContain("export interface SpreadsheetInput");
      expect(code).toContain("export interface SpreadsheetOutput");
      expect(code).toContain("export function calculateSpreadsheet");

      // Verify input interface
      expect(code).toContain("Input: {");
      expect(code).toContain("A1?: number | string;");
      expect(code).toContain("B1?: number | string;");
      expect(code).toContain("C1?: number | string;");

      // Verify output interface
      expect(code).toContain("Calculations: {");
      expect(code).toContain("A1: number | string;");
      expect(code).toContain("B1: number | string;");

      // Verify calculations
      expect(code).toContain("cells['Input!A1'] = input.Input.A1 ?? 100;");
      expect(code).toContain("cells['Input!B1'] = input.Input.B1 ?? 0.08;");
      expect(code).toContain(
        "cells['Input!C1'] = input.Input.C1 ?? \"Product A\";"
      );

      // Verify helper functions are included
      expect(code).toContain("function sum(");
      expect(code).toContain("function round(");
    });

    it("should generate valid Python code", () => {
      const sheets = createMockSheets();
      const analyzer = new DependencyAnalyzer();
      const dependencyGraph = analyzer.buildDependencyGraph(sheets);
      const generator = new PythonGenerator();

      const code = generator.generate(
        sheets,
        dependencyGraph,
        ["Input"],
        ["Calculations"]
      );

      // Verify structure
      expect(code).toContain("from typing import Dict, Any, List, Union");
      expect(code).toContain("import math");
      expect(code).toContain(
        "def calculate_spreadsheet(input_data: Dict[str, Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:"
      );

      // Verify input handling
      expect(code).toContain("input_sheet = input_data.get('Input', {})");
      expect(code).toContain("cells['Input!A1'] = input_sheet.get('A1', 100)");
      expect(code).toContain("cells['Input!B1'] = input_sheet.get('B1', 0.08)");
      expect(code).toContain(
        "cells['Input!C1'] = input_sheet.get('C1', 'Product A')"
      );

      // Verify output building
      expect(code).toContain("output = {}");
      expect(code).toContain("output['Calculations'] = {}");

      // Verify helper functions
      expect(code).toContain("def sum_values(");
      expect(code).toContain("def safe_divide(");
      expect(code).toContain("def round(");
    });

    it("should generate executable TypeScript code", () => {
      const sheets = createMockSheets();
      const analyzer = new DependencyAnalyzer();
      const dependencyGraph = analyzer.buildDependencyGraph(sheets);
      const generator = new TypeScriptGenerator();

      const code = generator.generate(
        sheets,
        dependencyGraph,
        ["Input"],
        ["Calculations"]
      );

      // The generated code should be syntactically valid TypeScript
      expect(code).not.toContain("undefined");
      expect(code).not.toContain("null");

      // Should have proper structure
      const lines = code.split("\n");
      const exportInterfaceLines = lines.filter((line) =>
        line.startsWith("export interface")
      );
      const exportFunctionLines = lines.filter((line) =>
        line.startsWith("export function")
      );

      expect(exportInterfaceLines.length).toBe(2); // Input and Output interfaces
      expect(exportFunctionLines.length).toBe(1); // calculateSpreadsheet function
    });

    it("should handle complex nested formulas", () => {
      const parser = new FormulaParser();
      const complexFormula =
        "=IF(SUM(A1:A10)>AVERAGE(B1:B10),MAX(C1:C10)*0.1,MIN(D1:D10))";

      expect(() => {
        const ast = parser.parse(complexFormula);
        expect(ast.type).toBe("function");
        expect(ast.value).toBe("IF");
        expect(ast.children).toHaveLength(3);
      }).not.toThrow();
    });

    it("should maintain calculation order in generated code", () => {
      const sheets = new Map<string, Sheet>([
        [
          "Test",
          {
            name: "Test",
            cells: new Map([
              ["A1", { row: 1, column: "A", value: 10 }],
              [
                "B1",
                {
                  row: 1,
                  column: "B",
                  value: null,
                  formula: "=A1*2",
                  parsedFormula: {
                    type: "operator",
                    value: "*",
                    children: [
                      { type: "reference", value: "A1" },
                      { type: "literal", value: "2" },
                    ],
                  },
                },
              ],
              [
                "C1",
                {
                  row: 1,
                  column: "C",
                  value: null,
                  formula: "=B1+5",
                  parsedFormula: {
                    type: "operator",
                    value: "+",
                    children: [
                      { type: "reference", value: "B1" },
                      { type: "literal", value: "5" },
                    ],
                  },
                },
              ],
            ]),
            range: { startRow: 1, endRow: 1, startColumn: "A", endColumn: "C" },
          },
        ],
      ]);

      const analyzer = new DependencyAnalyzer();
      const dependencyGraph = analyzer.buildDependencyGraph(sheets);
      const generator = new TypeScriptGenerator();

      const code = generator.generate(
        sheets,
        dependencyGraph,
        ["Test"],
        ["Test"]
      );

      // B1 should be calculated before C1
      const b1Index = code.indexOf("cells['Test!B1'] =");
      const c1Index = code.indexOf("cells['Test!C1'] =");

      expect(b1Index).toBeLessThan(c1Index);
      expect(b1Index).toBeGreaterThan(-1);
      expect(c1Index).toBeGreaterThan(-1);
    });
  });

  describe("Error Handling Integration", () => {
    it("should handle circular dependencies gracefully", () => {
      const sheets = new Map<string, Sheet>([
        [
          "Test",
          {
            name: "Test",
            cells: new Map([
              [
                "A1",
                {
                  row: 1,
                  column: "A",
                  value: null,
                  formula: "=B1+1",
                  parsedFormula: {
                    type: "operator",
                    value: "+",
                    children: [
                      { type: "reference", value: "B1" },
                      { type: "literal", value: "1" },
                    ],
                  },
                },
              ],
              [
                "B1",
                {
                  row: 1,
                  column: "B",
                  value: null,
                  formula: "=A1+1",
                  parsedFormula: {
                    type: "operator",
                    value: "+",
                    children: [
                      { type: "reference", value: "A1" },
                      { type: "literal", value: "1" },
                    ],
                  },
                },
              ],
            ]),
            range: { startRow: 1, endRow: 1, startColumn: "A", endColumn: "B" },
          },
        ],
      ]);

      const analyzer = new DependencyAnalyzer();

      expect(() => {
        analyzer.buildDependencyGraph(sheets);
      }).toThrow(/circular dependency/i);
    });

    it("should handle malformed formulas gracefully", () => {
      const parser = new FormulaParser();

      const malformedFormulas = [
        "=A1+", // Missing operand
        "=SUM(", // Unclosed function
        "=(A1+B1", // Unclosed parentheses
        "A1+B1", // Missing equals sign
      ];

      malformedFormulas.forEach((formula) => {
        expect(() => parser.parse(formula)).toThrow();
      });
    });
  });

  describe("Performance Integration", () => {
    it("should handle large dependency graphs efficiently", () => {
      const sheets = new Map<string, Sheet>();
      const cells = new Map<string, Cell>();

      // Create a large chain of dependencies A1 -> B1 -> C1 -> ... -> Z1
      for (let i = 0; i < 26; i++) {
        const currentColumn = String.fromCharCode(65 + i); // A, B, C, ...
        const cellRef = `${currentColumn}1`;

        if (i === 0) {
          // First cell has a value
          cells.set(cellRef, {
            row: 1,
            column: currentColumn,
            value: 100,
          });
        } else {
          // Other cells reference the previous cell
          const prevColumn = String.fromCharCode(65 + i - 1);
          const prevRef = `${prevColumn}1`;

          cells.set(cellRef, {
            row: 1,
            column: currentColumn,
            value: null,
            formula: `=${prevRef}*2`,
            parsedFormula: {
              type: "operator",
              value: "*",
              children: [
                { type: "reference", value: prevRef },
                { type: "literal", value: "2" },
              ],
            },
          });
        }
      }

      sheets.set("Large", {
        name: "Large",
        cells,
        range: { startRow: 1, endRow: 1, startColumn: "A", endColumn: "Z" },
      });

      const analyzer = new DependencyAnalyzer();

      const startTime = performance.now();
      const dependencyGraph = analyzer.buildDependencyGraph(sheets);
      const order = analyzer.getCalculationOrder();
      const endTime = performance.now();

      // Should complete within reasonable time (< 100ms for this size)
      expect(endTime - startTime).toBeLessThan(100);

      // Should have correct number of dependencies
      expect(dependencyGraph.size).toBe(25); // A1 has no formula, B1-Z1 have formulas
      expect(order.length).toBe(25);

      // Should maintain correct order
      for (let i = 1; i < order.length; i++) {
        const current = order[i];
        const previous = order[i - 1];

        // Each cell should come after its dependency
        const currentColumn = current.split("!")[1][0];
        const previousColumn = previous.split("!")[1][0];

        expect(currentColumn.charCodeAt(0)).toBeGreaterThan(
          previousColumn.charCodeAt(0)
        );
      }
    });
  });
});
