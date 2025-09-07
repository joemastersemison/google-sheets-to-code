import { PythonGenerator } from "./generators/python-generator.js";
import { TypeScriptGenerator } from "./generators/typescript-generator.js";
import { FormulaParser } from "./parsers/formula-parser.js";
import type { SheetConfig } from "./types/index.js";
import { DependencyAnalyzer } from "./utils/dependency-analyzer.js";
import { GoogleSheetsReader } from "./utils/sheets-reader.js";

export class SheetToCodeConverter {
  constructor(
    private config: SheetConfig,
    private verbose = false
  ) {}

  async convert(): Promise<string> {
    if (this.verbose) {
      console.log("🚀 Starting Google Sheets to Code conversion...");
      console.log(`📋 Input tabs: ${this.config.inputTabs.join(", ")}`);
      console.log(`📋 Output tabs: ${this.config.outputTabs.join(", ")}`);
    }

    const reader = new GoogleSheetsReader();

    // First, fetch the initially requested sheets
    const { sheets, namedRanges } = await reader.readSheets(
      this.config.spreadsheetUrl,
      [...this.config.inputTabs, ...this.config.outputTabs],
      this.verbose
    );

    // Find all sheet references in formulas
    const referencedSheets = new Set<string>();
    for (const [_sheetName, sheet] of sheets) {
      for (const [_cellRef, cell] of sheet.cells) {
        if (cell.formula) {
          // Look for sheet references in formula (Sheet!Cell pattern)
          const sheetRefs = cell.formula.match(/([A-Za-z0-9_]+)!/g);
          if (sheetRefs) {
            sheetRefs.forEach((ref) => {
              const sheetName = ref.replace("!", "");
              if (!sheets.has(sheetName)) {
                referencedSheets.add(sheetName);
              }
            });
          }
        }
      }
    }

    // Fetch any referenced sheets that weren't initially loaded
    if (referencedSheets.size > 0) {
      if (this.verbose) {
        console.log(
          `📊 Found references to additional sheets: ${Array.from(referencedSheets).join(", ")}`
        );
        console.log("📄 Fetching referenced sheets...");
      }

      const additionalData = await reader.readSheets(
        this.config.spreadsheetUrl,
        Array.from(referencedSheets),
        this.verbose
      );

      // Merge the additional sheets
      for (const [name, sheet] of additionalData.sheets) {
        sheets.set(name, sheet);
      }

      // Merge any additional named ranges (though they should be the same)
      for (const [name, range] of additionalData.namedRanges) {
        if (!namedRanges.has(name)) {
          namedRanges.set(name, range);
        }
      }
    }

    if (this.verbose) console.log("🔍 Parsing formulas...");
    const parser = new FormulaParser();
    const parsedSheets = new Map();
    let totalFormulas = 0;
    let parsedFormulas = 0;

    for (const [sheetName, sheet] of sheets) {
      const parsedCells = new Map();
      for (const [cellRef, cell] of sheet.cells) {
        if (cell.formula) {
          totalFormulas++;
          try {
            // Replace named ranges in formula before parsing
            let expandedFormula = cell.formula;
            for (const [name, range] of namedRanges) {
              // Replace named range with actual cell reference
              // Use word boundaries to avoid partial matches
              const regex = new RegExp(`\\b${name}\\b`, "g");
              expandedFormula = expandedFormula.replace(regex, range);
            }

            if (expandedFormula !== cell.formula && this.verbose) {
              console.log(
                `  📝 Expanded named ranges in ${sheetName}!${cellRef}: ${cell.formula} -> ${expandedFormula}`
              );
            }

            parsedCells.set(cellRef, {
              ...cell,
              formula: expandedFormula, // Store expanded formula
              originalFormula: cell.formula, // Keep original for reference
              parsedFormula: parser.parse(expandedFormula),
            });
            parsedFormulas++;
          } catch (error) {
            console.warn(
              `⚠️  Failed to parse formula in ${sheetName}!${cellRef}: ${cell.formula}. Error: ${error instanceof Error ? error.message : String(error)}`
            );
            parsedCells.set(cellRef, cell);
          }
        } else {
          parsedCells.set(cellRef, cell);
        }
      }
      parsedSheets.set(sheetName, { ...sheet, cells: parsedCells });
    }

    if (this.verbose) {
      console.log(
        `✅ Parsed ${parsedFormulas}/${totalFormulas} formulas successfully`
      );
    }

    if (this.verbose) console.log("🔗 Building dependency graph...");
    const analyzer = new DependencyAnalyzer();
    const dependencyGraph = analyzer.buildDependencyGraph(parsedSheets);
    if (this.verbose) {
      console.log(`✅ Dependency graph built: ${dependencyGraph.size} nodes`);
    }

    if (this.verbose) {
      console.log(`🏭 Generating ${this.config.outputLanguage} code...`);
    }
    const generator =
      this.config.outputLanguage === "typescript"
        ? new TypeScriptGenerator()
        : new PythonGenerator();

    const code = generator.generate(
      parsedSheets,
      dependencyGraph,
      this.config.inputTabs,
      this.config.outputTabs
    );

    if (this.verbose) {
      console.log(`✅ Code generation complete: ${code.length} characters`);
    }

    return code;
  }
}
