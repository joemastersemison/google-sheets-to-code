import { PythonGenerator } from '../generators/python-generator.js';
import { Sheet, Cell, ParsedFormula, DependencyNode } from '../types/index.js';

describe('PythonGenerator', () => {
  let generator: PythonGenerator;

  beforeEach(() => {
    generator = new PythonGenerator();
  });

  function createSheet(name: string, cells: Record<string, Partial<Cell>>): Sheet {
    const cellMap = new Map<string, Cell>();
    
    for (const [cellRef, cellData] of Object.entries(cells)) {
      const [column, row] = [cellRef.match(/[A-Z]+/)?.[0] || 'A', cellRef.match(/\d+/)?.[0] || '1'];
      cellMap.set(cellRef, {
        row: parseInt(row),
        column,
        value: cellData.value || 0,
        formula: cellData.formula,
        parsedFormula: cellData.parsedFormula,
        ...cellData
      });
    }

    return {
      name,
      cells: cellMap,
      range: {
        startRow: 1,
        endRow: 10,
        startColumn: 'A',
        endColumn: 'Z'
      }
    };
  }

  function createParsedFormula(type: string, value: string, children?: ParsedFormula[]): ParsedFormula {
    return { type: type as any, value, children };
  }

  function createDependencyNode(cellRef: string, sheetName: string, dependencies: string[], formula?: ParsedFormula): DependencyNode {
    return {
      cellRef,
      sheetName,
      dependencies: new Set(dependencies),
      formula
    };
  }

  describe('generate', () => {
    it('should generate basic Python function with imports', () => {
      const sheets = new Map([
        ['Input', createSheet('Input', {
          'A1': { value: 10 },
          'B1': { value: 20 }
        })],
        ['Output', createSheet('Output', {
          'A1': { 
            formula: '=Input!A1+Input!B1',
            parsedFormula: createParsedFormula('operator', '+', [
              createParsedFormula('reference', 'Input!A1'),
              createParsedFormula('reference', 'Input!B1')
            ])
          }
        })]
      ]);

      const dependencyGraph = new Map([
        ['Output!A1', createDependencyNode('A1', 'Output', ['Input!A1', 'Input!B1'], 
          createParsedFormula('operator', '+', [
            createParsedFormula('reference', 'Input!A1'),
            createParsedFormula('reference', 'Input!B1')
          ])
        )]
      ]);

      const code = generator.generate(sheets, dependencyGraph, ['Input'], ['Output']);
      
      expect(code).toContain('from typing import Dict, Any, List, Union');
      expect(code).toContain('import math');
      expect(code).toContain('from datetime import datetime');
      expect(code).toContain('def calculate_spreadsheet(input_data: Dict[str, Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:');
      expect(code).toContain('"""');
      expect(code).toContain('cells = {}');
    });

    it('should generate input initialization correctly', () => {
      const sheets = new Map([
        ['Data', createSheet('Data', {
          'A1': { value: 100 },
          'B1': { value: 'Hello' }
        })]
      ]);

      const code = generator.generate(sheets, new Map(), ['Data'], []);
      
      expect(code).toContain("input_sheet = input_data.get('Data', {})");
      expect(code).toContain("cells['Data!A1'] = input_sheet.get('A1', 100)");
      expect(code).toContain("cells['Data!B1'] = input_sheet.get('B1', 'Hello')");
    });

    it('should generate output building correctly', () => {
      const sheets = new Map([
        ['Results', createSheet('Results', {
          'A1': { value: 42 },
          'B1': { value: 3.14 }
        })]
      ]);

      const code = generator.generate(sheets, new Map(), [], ['Results']);
      
      expect(code).toContain("output = {}");
      expect(code).toContain("output['Results'] = {}");
      expect(code).toContain("output['Results']['A1'] = cells.get('Results!A1')");
      expect(code).toContain("output['Results']['B1'] = cells.get('Results!B1')");
      expect(code).toContain('return output');
    });
  });

  describe('formula code generation', () => {
    let sheets: Map<string, Sheet>;
    let dependencyGraph: Map<string, DependencyNode>;

    beforeEach(() => {
      sheets = new Map([
        ['Test', createSheet('Test', {})]
      ]);
      dependencyGraph = new Map();
    });

    it('should generate arithmetic operations', () => {
      const formula = createParsedFormula('operator', '+', [
        createParsedFormula('literal', '5'),
        createParsedFormula('literal', '3')
      ]);

      dependencyGraph.set('Test!A1', createDependencyNode('A1', 'Test', [], formula));
      
      const code = generator.generate(sheets, dependencyGraph, [], ['Test']);
      
      expect(code).toContain('(5 + 3)');
    });

    it('should generate safe division', () => {
      const formula = createParsedFormula('operator', '/', [
        createParsedFormula('literal', '10'),
        createParsedFormula('literal', '2')
      ]);

      dependencyGraph.set('Test!A1', createDependencyNode('A1', 'Test', [], formula));
      
      const code = generator.generate(sheets, dependencyGraph, [], ['Test']);
      
      expect(code).toContain('safe_divide(10, 2)');
    });

    it('should generate power operations', () => {
      const formula = createParsedFormula('operator', '^', [
        createParsedFormula('literal', '2'),
        createParsedFormula('literal', '3')
      ]);

      dependencyGraph.set('Test!A1', createDependencyNode('A1', 'Test', [], formula));
      
      const code = generator.generate(sheets, dependencyGraph, [], ['Test']);
      
      expect(code).toContain('(2 ** 3)');
    });

    it('should generate function calls', () => {
      const formula = createParsedFormula('function', 'SUM', [
        createParsedFormula('reference', 'A1:A10')
      ]);

      dependencyGraph.set('Test!B1', createDependencyNode('B1', 'Test', [], formula));
      
      const code = generator.generate(sheets, dependencyGraph, [], ['Test']);
      
      expect(code).toContain('sum_values(get_range(');
    });

    it('should generate IF statements', () => {
      const formula = createParsedFormula('function', 'IF', [
        createParsedFormula('operator', '>', [
          createParsedFormula('reference', 'A1'),
          createParsedFormula('literal', '0')
        ]),
        createParsedFormula('literal', 'Positive'),
        createParsedFormula('literal', 'Negative')
      ]);

      dependencyGraph.set('Test!C1', createDependencyNode('C1', 'Test', [], formula));
      
      const code = generator.generate(sheets, dependencyGraph, [], ['Test']);
      
      expect(code).toContain("('Positive' if (cells.get('Test!A1') > 0) else 'Negative')");
    });

    it('should generate string concatenation', () => {
      const formula = createParsedFormula('operator', '&', [
        createParsedFormula('literal', 'Hello'),
        createParsedFormula('literal', 'World')
      ]);

      dependencyGraph.set('Test!D1', createDependencyNode('D1', 'Test', [], formula));
      
      const code = generator.generate(sheets, dependencyGraph, [], ['Test']);
      
      expect(code).toContain("str('Hello') + str('World')");
    });

    it('should handle math functions', () => {
      const formula = createParsedFormula('function', 'SQRT', [
        createParsedFormula('literal', '16')
      ]);

      dependencyGraph.set('Test!E1', createDependencyNode('E1', 'Test', [], formula));
      
      const code = generator.generate(sheets, dependencyGraph, [], ['Test']);
      
      expect(code).toContain('math.sqrt(16)');
    });

    it('should handle string functions', () => {
      const formula = createParsedFormula('function', 'LEN', [
        createParsedFormula('literal', 'Hello')
      ]);

      dependencyGraph.set('Test!F1', createDependencyNode('F1', 'Test', [], formula));
      
      const code = generator.generate(sheets, dependencyGraph, [], ['Test']);
      
      expect(code).toContain("len(str('Hello'))");
    });

    it('should handle date functions', () => {
      const formula = createParsedFormula('function', 'TODAY', []);

      dependencyGraph.set('Test!G1', createDependencyNode('G1', 'Test', [], formula));
      
      const code = generator.generate(sheets, dependencyGraph, [], ['Test']);
      
      expect(code).toContain("datetime.now().strftime('%Y-%m-%d')");
    });
  });

  describe('helper functions generation', () => {
    it('should include all helper functions', () => {
      const sheets = new Map([
        ['Test', createSheet('Test', {})]
      ]);
      const dependencyGraph = new Map();

      const code = generator.generate(sheets, dependencyGraph, [], []);
      
      expect(code).toContain('def flatten_values(*args):');
      expect(code).toContain('def sum_values(*args):');
      expect(code).toContain('def average_values(*args):');
      expect(code).toContain('def min_values(*args):');
      expect(code).toContain('def max_values(*args):');
      expect(code).toContain('def count_values(*args):');
      expect(code).toContain('def concatenate_values(*args):');
      expect(code).toContain('def safe_divide(numerator, denominator):');
      expect(code).toContain('def vlookup(');
      expect(code).toContain('def get_range(range_ref: str, cells: dict) -> list:');
    });

    it('should include proper docstrings', () => {
      const sheets = new Map([
        ['Test', createSheet('Test', {})]
      ]);
      const dependencyGraph = new Map();

      const code = generator.generate(sheets, dependencyGraph, [], []);
      
      expect(code).toContain('"""Flatten nested lists/values into a single list."""');
      expect(code).toContain('"""Sum all numeric values."""');
      expect(code).toContain('"""Calculate average of numeric values."""');
      expect(code).toContain('"""Safely divide two numbers, returning error for division by zero."""');
    });
  });

  describe('value formatting', () => {
    it('should format string values correctly', () => {
      const sheets = new Map([
        ['Input', createSheet('Input', {
          'A1': { value: "Hello 'World'" }
        })]
      ]);

      const code = generator.generate(sheets, new Map(), ['Input'], []);
      
      expect(code).toContain("input_sheet.get('A1', 'Hello \\'World\\'')");
    });

    it('should format boolean values correctly', () => {
      const sheets = new Map([
        ['Input', createSheet('Input', {
          'A1': { value: 'TRUE' },
          'B1': { value: 'FALSE' }
        })]
      ]);

      const code = generator.generate(sheets, new Map(), ['Input'], []);
      
      expect(code).toContain("input_sheet.get('A1', True)");
      expect(code).toContain("input_sheet.get('B1', False)");
    });

    it('should format null values correctly', () => {
      const sheets = new Map([
        ['Input', createSheet('Input', {
          'A1': { value: null }
        })]
      ]);

      const code = generator.generate(sheets, new Map(), ['Input'], []);
      
      expect(code).toContain("input_sheet.get('A1', None)");
    });

    it('should format numeric values correctly', () => {
      const sheets = new Map([
        ['Input', createSheet('Input', {
          'A1': { value: 42 },
          'B1': { value: 3.14 }
        })]
      ]);

      const code = generator.generate(sheets, new Map(), ['Input'], []);
      
      expect(code).toContain("input_sheet.get('A1', 42)");
      expect(code).toContain("input_sheet.get('B1', 3.14)");
    });
  });

  describe('property name sanitization', () => {
    it('should sanitize invalid property names', () => {
      const sheets = new Map([
        ['My-Sheet', createSheet('My-Sheet', {
          'A1': { value: 10 }
        })]
      ]);

      const code = generator.generate(sheets, new Map(), ['My-Sheet'], []);
      
      expect(code).toContain("input_sheet = input_data.get('My_Sheet', {})");
    });

    it('should handle names starting with numbers', () => {
      const sheets = new Map([
        ['2023Data', createSheet('2023Data', {
          'A1': { value: 10 }
        })]
      ]);

      const code = generator.generate(sheets, new Map(), ['2023Data'], []);
      
      expect(code).toContain("input_sheet = input_data.get('_2023Data', {})");
    });
  });

  describe('cell reference generation', () => {
    it('should generate cell references with .get()', () => {
      const formula = createParsedFormula('reference', 'A1');
      const dependencyGraph = new Map([
        ['Test!B1', createDependencyNode('B1', 'Test', ['Test!A1'], formula)]
      ]);

      const sheets = new Map([
        ['Test', createSheet('Test', {})]
      ]);

      const code = generator.generate(sheets, dependencyGraph, [], ['Test']);
      
      expect(code).toContain("cells.get('Test!A1')");
    });

    it('should generate range references', () => {
      const formula = createParsedFormula('reference', 'A1:A10');
      const dependencyGraph = new Map([
        ['Test!B1', createDependencyNode('B1', 'Test', [], formula)]
      ]);

      const sheets = new Map([
        ['Test', createSheet('Test', {})]
      ]);

      const code = generator.generate(sheets, dependencyGraph, [], ['Test']);
      
      expect(code).toContain("get_range('Test!A1:A10', cells)");
    });
  });
});