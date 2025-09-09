import { PythonGenerator } from "./generators/python-generator.js";
import { TypeScriptGenerator } from "./generators/typescript-generator.js";
import { FormulaParser } from "./parsers/formula-parser.js";
import type { SheetConfig } from "./types/index.js";
import type { DataValidationRule } from "./types/validation.js";
import { DependencyAnalyzer } from "./utils/dependency-analyzer.js";
import { GoogleSheetsReader } from "./utils/sheets-reader.js";
import {
  ValidationComparator,
  type ValidationOptions,
  type ValidationResult,
} from "./utils/validation-comparator.js";
import {
  type ValidationData,
  ValidationDataFetcher,
} from "./utils/validation-data-fetcher.js";
import { ValidationEngine } from "./utils/validation-engine.js";

export class SheetToCodeConverter {
  private validationEngine: ValidationEngine;
  // private testGenerator: TestGenerator; // TODO: Implement test generator
  private generateTests: boolean;

  constructor(
    private config: SheetConfig,
    private verbose = false,
    private validationRules?: DataValidationRule[],
    generateTests = false
  ) {
    this.validationEngine = new ValidationEngine();
    // this.testGenerator = new TestGenerator(); // TODO: Implement test generator
    this.generateTests = generateTests;
    if (validationRules) {
      // Group rules by sheet
      const rulesBySheet = new Map<string, DataValidationRule[]>();
      for (const rule of validationRules) {
        const sheetName = this.extractSheetFromRange(rule.range);
        if (!rulesBySheet.has(sheetName)) {
          rulesBySheet.set(sheetName, []);
        }
        rulesBySheet.get(sheetName)?.push(rule);
      }

      // Add rules to validation engine
      for (const [sheetName, rules] of rulesBySheet) {
        this.validationEngine.addRules(sheetName, rules);
      }
    }
  }

  private extractSheetFromRange(range: string): string {
    const sheetMatch = range.match(/^([^!]+)!/);
    return sheetMatch ? sheetMatch[1] : "Sheet1";
  }

  async convert(): Promise<{
    code: string;
    tests?: string;
    validationData?: ValidationData;
  }> {
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
          const formula = cell.formula; // Store in variable to avoid TypeScript issues
          // Look for sheet references in formula
          // Matches both unquoted (Sheet1!A1) and quoted ('My Sheet'!A1) patterns
          // Also handles escaped quotes within sheet names ('John''s Sheet'!A1)
          const quotedSheetPattern = /'(?:[^']|'')+'/g;
          const unquotedSheetPattern = /[A-Za-z0-9_]+(?=!)/g;

          // First, extract quoted sheet names
          // Use matchAll to get actual positions of matches
          const quotedMatches = [...formula.matchAll(quotedSheetPattern)];
          const processedQuotedRanges: Array<[number, number]> = [];

          for (const match of quotedMatches) {
            // Check if this quoted string is followed by !
            const matchIndex = match.index;
            if (matchIndex === undefined) continue; // Skip if no index (shouldn't happen with matchAll)
            const matchText = match[0];
            if (formula[matchIndex + matchText.length] === "!") {
              // Remove surrounding quotes and unescape doubled quotes
              const sheetName = matchText.slice(1, -1).replace(/''/g, "'");
              if (!sheets.has(sheetName)) {
                referencedSheets.add(sheetName);
              }
              // Track this range as processed
              processedQuotedRanges.push([
                matchIndex,
                matchIndex + matchText.length + 1,
              ]);
            }
          }

          // Then, extract unquoted sheet names (but avoid matching already processed quoted ones)
          // Check if a position is within any processed quoted range
          const isInQuotedRange = (index: number): boolean => {
            return processedQuotedRanges.some(
              ([start, end]) => index >= start && index < end
            );
          };

          // Use matchAll for unquoted patterns too
          const unquotedMatches = [...formula.matchAll(unquotedSheetPattern)];
          for (const match of unquotedMatches) {
            const matchIndex = match.index;
            if (matchIndex === undefined) continue; // Skip if no index (shouldn't happen with matchAll)
            // Skip if this match is within a quoted sheet reference we already processed
            if (!isInQuotedRange(matchIndex)) {
              const sheetName = match[0];
              if (!sheets.has(sheetName)) {
                referencedSheets.add(sheetName);
              }
            }
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

    // Generate validation code if rules are present
    let includeValidation = false;
    if (this.validationRules && this.validationRules.length > 0) {
      const validationCode = this.validationEngine.generateValidationCode(
        this.config.outputLanguage
      );
      generator.setValidationCode(validationCode);
      includeValidation = true;
      if (this.verbose) {
        console.log(
          `📋 Including ${this.validationRules.length} validation rules`
        );
      }
    }

    const code = generator.generate(
      parsedSheets,
      dependencyGraph,
      this.config.inputTabs,
      this.config.outputTabs,
      includeValidation
    );

    if (this.verbose) {
      console.log(`✅ Code generation complete: ${code.length} characters`);
    }

    // Generate tests if requested
    let tests: string | undefined;
    if (this.generateTests) {
      if (this.verbose) {
        console.log("🧪 Test generation requested but not yet implemented");
      }
      // TODO: Implement test generation
      // tests =
      //   this.config.outputLanguage === "typescript"
      //     ? this.testGenerator.generateTypeScriptTests(
      //         parsedSheets,
      //         dependencyGraph,
      //         this.config.inputTabs,
      //         this.config.outputTabs
      //       )
      //     : this.testGenerator.generatePythonTests(
      //         parsedSheets,
      //         dependencyGraph,
      //         this.config.inputTabs,
      //         this.config.outputTabs
      //       );
      // if (this.verbose) {
      //   console.log(`✅ Test generation complete: ${tests.length} characters`);
      // }
    }

    return { code, tests, validationData: undefined };
  }

  /**
   * Fetches actual data from Google Sheets for validation
   */
  async fetchValidationData(verbose = false): Promise<ValidationData> {
    const fetcher = new ValidationDataFetcher();
    const allSheets = [...this.config.inputTabs, ...this.config.outputTabs];

    if (verbose) {
      console.log("📊 Fetching validation data from Google Sheets...");
    }

    const validationData = await fetcher.fetchValidationData(
      this.config.spreadsheetUrl,
      allSheets,
      verbose
    );

    if (verbose) {
      console.log(
        `✅ Validation data fetched at ${validationData.timestamp.toISOString()}`
      );
    }

    return validationData;
  }

  /**
   * Fetches validation data multiple times (useful for time-based functions)
   */
  async fetchValidationDataMultipleTimes(
    count = 3,
    delayMs = 1000,
    verbose = false
  ): Promise<ValidationData[]> {
    const fetcher = new ValidationDataFetcher();
    const allSheets = [...this.config.inputTabs, ...this.config.outputTabs];

    if (verbose) {
      console.log(
        `📊 Fetching validation data ${count} times with ${delayMs}ms delay...`
      );
    }

    const results = await fetcher.fetchValidationDataMultipleTimes(
      this.config.spreadsheetUrl,
      allSheets,
      count,
      delayMs,
      verbose
    );

    if (verbose) {
      console.log(`✅ Fetched ${results.length} validation snapshots`);
    }

    return results;
  }

  /**
   * Validates generated code against actual Google Sheets data
   */
  async validateGeneratedCode(
    generatedFilePath: string,
    validationData: ValidationData,
    inputData: Record<string, any>,
    options?: ValidationOptions
  ): Promise<ValidationResult> {
    const comparator = new ValidationComparator();

    if (this.verbose || options?.verbose) {
      console.log("🔍 Validating generated code against actual data...");
    }

    const result = await comparator.compareWithGeneratedCode(
      validationData,
      generatedFilePath,
      inputData,
      this.config.outputTabs,
      options
    );

    if (this.verbose || options?.verbose) {
      const status = result.passed ? "✅ PASSED" : "❌ FAILED";
      console.log(`Validation ${status}`);
      console.log(
        `Accuracy: ${((result.matchingCells / result.totalCells) * 100).toFixed(2)}%`
      );
    }

    return result;
  }

  /**
   * Performs a complete conversion with validation
   */
  async convertAndValidate(
    outputFilePath: string,
    inputData: Record<string, any>,
    validationOptions?: ValidationOptions
  ): Promise<{
    code: string;
    validationResult: ValidationResult;
    validationData: ValidationData;
  }> {
    // First, fetch validation data
    const validationData = await this.fetchValidationData(this.verbose);

    // Generate the code
    const { code } = await this.convert();

    // Write the code to file
    const fs = await import("node:fs");
    await fs.promises.writeFile(outputFilePath, code);

    if (this.verbose) {
      console.log(`📝 Generated code written to: ${outputFilePath}`);
    }

    // Validate the generated code
    const validationResult = await this.validateGeneratedCode(
      outputFilePath,
      validationData,
      inputData,
      validationOptions
    );

    return {
      code,
      validationResult,
      validationData,
    };
  }
}
