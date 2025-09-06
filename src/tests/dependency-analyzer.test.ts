import { DependencyAnalyzer } from '../utils/dependency-analyzer.js';
import { Sheet, Cell, ParsedFormula } from '../types/index.js';

describe('DependencyAnalyzer', () => {
  let analyzer: DependencyAnalyzer;

  beforeEach(() => {
    analyzer = new DependencyAnalyzer();
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

  describe('buildDependencyGraph', () => {
    it('should build simple dependency graph', () => {
      const sheets = new Map([
        ['Sheet1', createSheet('Sheet1', {
          'A1': { value: 10 },
          'B1': { 
            formula: '=A1*2', 
            parsedFormula: createParsedFormula('operator', '*', [
              createParsedFormula('reference', 'A1'),
              createParsedFormula('literal', '2')
            ])
          }
        })]
      ]);

      const graph = analyzer.buildDependencyGraph(sheets);
      
      expect(graph.size).toBe(1);
      expect(graph.has('Sheet1!B1')).toBe(true);
      
      const node = graph.get('Sheet1!B1');
      expect(node?.dependencies).toEqual(new Set(['Sheet1!A1']));
    });

    it('should handle cross-sheet references', () => {
      const sheets = new Map([
        ['Input', createSheet('Input', {
          'A1': { value: 100 }
        })],
        ['Calc', createSheet('Calc', {
          'A1': {
            formula: '=Input!A1*0.1',
            parsedFormula: createParsedFormula('operator', '*', [
              createParsedFormula('reference', 'Input!A1'),
              createParsedFormula('literal', '0.1')
            ])
          }
        })]
      ]);

      const graph = analyzer.buildDependencyGraph(sheets);
      
      const node = graph.get('Calc!A1');
      expect(node?.dependencies).toEqual(new Set(['Input!A1']));
    });

    it('should handle range references', () => {
      const sheets = new Map([
        ['Sheet1', createSheet('Sheet1', {
          'A1': { value: 10 },
          'A2': { value: 20 },
          'A3': { value: 30 },
          'B1': {
            formula: '=SUM(A1:A3)',
            parsedFormula: createParsedFormula('function', 'SUM', [
              createParsedFormula('reference', 'A1:A3')
            ])
          }
        })]
      ]);

      const graph = analyzer.buildDependencyGraph(sheets);
      
      const node = graph.get('Sheet1!B1');
      expect(node?.dependencies).toEqual(new Set(['Sheet1!A1:A3']));
    });
  });

  describe('getCalculationOrder', () => {
    it('should return correct calculation order for simple chain', () => {
      const sheets = new Map([
        ['Sheet1', createSheet('Sheet1', {
          'A1': { value: 10 },
          'B1': {
            formula: '=A1*2',
            parsedFormula: createParsedFormula('operator', '*', [
              createParsedFormula('reference', 'A1'),
              createParsedFormula('literal', '2')
            ])
          },
          'C1': {
            formula: '=B1+5',
            parsedFormula: createParsedFormula('operator', '+', [
              createParsedFormula('reference', 'B1'),
              createParsedFormula('literal', '5')
            ])
          }
        })]
      ]);

      analyzer.buildDependencyGraph(sheets);
      const order = analyzer.getCalculationOrder();
      
      expect(order).toEqual(['Sheet1!B1', 'Sheet1!C1']);
    });

    it('should handle multiple independent chains', () => {
      const sheets = new Map([
        ['Sheet1', createSheet('Sheet1', {
          'A1': { value: 10 },
          'A2': { value: 20 },
          'B1': {
            formula: '=A1*2',
            parsedFormula: createParsedFormula('operator', '*', [
              createParsedFormula('reference', 'A1'),
              createParsedFormula('literal', '2')
            ])
          },
          'B2': {
            formula: '=A2*3',
            parsedFormula: createParsedFormula('operator', '*', [
              createParsedFormula('reference', 'A2'),
              createParsedFormula('literal', '3')
            ])
          }
        })]
      ]);

      analyzer.buildDependencyGraph(sheets);
      const order = analyzer.getCalculationOrder();
      
      expect(order).toHaveLength(2);
      expect(order).toContain('Sheet1!B1');
      expect(order).toContain('Sheet1!B2');
    });

    it('should handle complex dependency trees', () => {
      const sheets = new Map([
        ['Sheet1', createSheet('Sheet1', {
          'A1': { value: 10 },
          'A2': { value: 20 },
          'B1': {
            formula: '=A1*2',
            parsedFormula: createParsedFormula('operator', '*', [
              createParsedFormula('reference', 'A1'),
              createParsedFormula('literal', '2')
            ])
          },
          'B2': {
            formula: '=A2*3',
            parsedFormula: createParsedFormula('operator', '*', [
              createParsedFormula('reference', 'A2'),
              createParsedFormula('literal', '3')
            ])
          },
          'C1': {
            formula: '=B1+B2',
            parsedFormula: createParsedFormula('operator', '+', [
              createParsedFormula('reference', 'B1'),
              createParsedFormula('reference', 'B2')
            ])
          }
        })]
      ]);

      analyzer.buildDependencyGraph(sheets);
      const order = analyzer.getCalculationOrder();
      
      expect(order).toHaveLength(3);
      expect(order.indexOf('Sheet1!B1')).toBeLessThan(order.indexOf('Sheet1!C1'));
      expect(order.indexOf('Sheet1!B2')).toBeLessThan(order.indexOf('Sheet1!C1'));
    });
  });

  describe('circular dependency detection', () => {
    it('should detect simple circular dependency', () => {
      const sheets = new Map([
        ['Sheet1', createSheet('Sheet1', {
          'A1': {
            formula: '=B1+1',
            parsedFormula: createParsedFormula('operator', '+', [
              createParsedFormula('reference', 'B1'),
              createParsedFormula('literal', '1')
            ])
          },
          'B1': {
            formula: '=A1+1',
            parsedFormula: createParsedFormula('operator', '+', [
              createParsedFormula('reference', 'A1'),
              createParsedFormula('literal', '1')
            ])
          }
        })]
      ]);

      expect(() => analyzer.buildDependencyGraph(sheets)).toThrow(/circular dependency/i);
    });

    it('should detect complex circular dependency', () => {
      const sheets = new Map([
        ['Sheet1', createSheet('Sheet1', {
          'A1': {
            formula: '=B1+1',
            parsedFormula: createParsedFormula('operator', '+', [
              createParsedFormula('reference', 'B1'),
              createParsedFormula('literal', '1')
            ])
          },
          'B1': {
            formula: '=C1+1',
            parsedFormula: createParsedFormula('operator', '+', [
              createParsedFormula('reference', 'C1'),
              createParsedFormula('literal', '1')
            ])
          },
          'C1': {
            formula: '=A1+1',
            parsedFormula: createParsedFormula('operator', '+', [
              createParsedFormula('reference', 'A1'),
              createParsedFormula('literal', '1')
            ])
          }
        })]
      ]);

      expect(() => analyzer.buildDependencyGraph(sheets)).toThrow(/circular dependency/i);
    });
  });

  describe('getDependents', () => {
    beforeEach(() => {
      const sheets = new Map([
        ['Sheet1', createSheet('Sheet1', {
          'A1': { value: 10 },
          'B1': {
            formula: '=A1*2',
            parsedFormula: createParsedFormula('operator', '*', [
              createParsedFormula('reference', 'A1'),
              createParsedFormula('literal', '2')
            ])
          },
          'C1': {
            formula: '=A1+5',
            parsedFormula: createParsedFormula('operator', '+', [
              createParsedFormula('reference', 'A1'),
              createParsedFormula('literal', '5')
            ])
          }
        })]
      ]);

      analyzer.buildDependencyGraph(sheets);
    });

    it('should return cells that depend on given cell', () => {
      const dependents = analyzer.getDependents('Sheet1!A1');
      
      expect(dependents).toEqual(new Set(['Sheet1!B1', 'Sheet1!C1']));
    });

    it('should return empty set for cells with no dependents', () => {
      const dependents = analyzer.getDependents('Sheet1!B1');
      
      expect(dependents).toEqual(new Set());
    });
  });

  describe('getTransitiveDependencies', () => {
    beforeEach(() => {
      const sheets = new Map([
        ['Sheet1', createSheet('Sheet1', {
          'A1': { value: 10 },
          'B1': {
            formula: '=A1*2',
            parsedFormula: createParsedFormula('operator', '*', [
              createParsedFormula('reference', 'A1'),
              createParsedFormula('literal', '2')
            ])
          },
          'C1': {
            formula: '=B1+5',
            parsedFormula: createParsedFormula('operator', '+', [
              createParsedFormula('reference', 'B1'),
              createParsedFormula('literal', '5')
            ])
          }
        })]
      ]);

      analyzer.buildDependencyGraph(sheets);
    });

    it('should return all transitive dependencies', () => {
      const deps = analyzer.getTransitiveDependencies('Sheet1!C1');
      
      expect(deps).toEqual(new Set(['Sheet1!B1', 'Sheet1!A1']));
    });
  });

  describe('getTransitiveDependents', () => {
    beforeEach(() => {
      const sheets = new Map([
        ['Sheet1', createSheet('Sheet1', {
          'A1': { value: 10 },
          'B1': {
            formula: '=A1*2',
            parsedFormula: createParsedFormula('operator', '*', [
              createParsedFormula('reference', 'A1'),
              createParsedFormula('literal', '2')
            ])
          },
          'C1': {
            formula: '=B1+5',
            parsedFormula: createParsedFormula('operator', '+', [
              createParsedFormula('reference', 'B1'),
              createParsedFormula('literal', '5')
            ])
          }
        })]
      ]);

      analyzer.buildDependencyGraph(sheets);
    });

    it('should return all transitive dependents', () => {
      const dependents = analyzer.getTransitiveDependents('Sheet1!A1');
      
      expect(dependents).toEqual(new Set(['Sheet1!B1', 'Sheet1!C1']));
    });
  });
});