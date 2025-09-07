export interface SheetConfig {
  spreadsheetUrl: string;
  inputTabs: string[];
  outputTabs: string[];
  outputLanguage: "typescript" | "python";
}

export interface Cell {
  row: number;
  column: string;
  value: string | number | boolean | null;
  formula?: string;
  formattedValue?: string;
  parsedFormula?: ParsedFormula;
}

export interface Sheet {
  name: string;
  cells: Map<string, Cell>;
  range: {
    startRow: number;
    endRow: number;
    startColumn: string;
    endColumn: string;
  };
}

export interface ParsedFormula {
  type: "function" | "reference" | "literal" | "operator";
  value: string;
  children?: ParsedFormula[];
}

export interface DependencyNode {
  cellRef: string;
  sheetName: string;
  dependencies: Set<string>;
  formula?: ParsedFormula;
}
