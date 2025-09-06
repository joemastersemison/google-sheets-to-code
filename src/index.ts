import { SheetConfig } from './types/index.js';
import { GoogleSheetsReader } from './utils/sheets-reader.js';
import { FormulaParser } from './parsers/formula-parser.js';
import { DependencyAnalyzer } from './utils/dependency-analyzer.js';
import { TypeScriptGenerator } from './generators/typescript-generator.js';
import { PythonGenerator } from './generators/python-generator.js';

export class SheetToCodeConverter {
  constructor(private config: SheetConfig) {}

  async convert(): Promise<string> {
    const reader = new GoogleSheetsReader();
    const sheets = await reader.readSheets(this.config.spreadsheetUrl, [
      ...this.config.inputTabs,
      ...this.config.outputTabs
    ]);

    const parser = new FormulaParser();
    const parsedSheets = new Map();
    
    for (const [sheetName, sheet] of sheets) {
      const parsedCells = new Map();
      for (const [cellRef, cell] of sheet.cells) {
        if (cell.formula) {
          parsedCells.set(cellRef, {
            ...cell,
            parsedFormula: parser.parse(cell.formula)
          });
        } else {
          parsedCells.set(cellRef, cell);
        }
      }
      parsedSheets.set(sheetName, { ...sheet, cells: parsedCells });
    }

    const analyzer = new DependencyAnalyzer();
    const dependencyGraph = analyzer.buildDependencyGraph(parsedSheets);

    const generator = this.config.outputLanguage === 'typescript' 
      ? new TypeScriptGenerator()
      : new PythonGenerator();

    return generator.generate(
      parsedSheets,
      dependencyGraph,
      this.config.inputTabs,
      this.config.outputTabs
    );
  }
}