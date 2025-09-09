import { spawn } from "node:child_process";
import { promises as fs, writeFileSync } from "node:fs";
import path from "node:path";
import { promisify } from "node:util";
import type { ValidationData } from "./validation-data-fetcher.js";

/**
 * Safely executes a command using spawn to prevent shell injection
 */
function execSafe(
  command: string,
  args: string[]
): Promise<{ stdout: string; stderr: string }> {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { stdio: "pipe" });
    let stdout = "";
    let stderr = "";

    child.stdout?.on("data", (data) => {
      stdout += data.toString();
    });

    child.stderr?.on("data", (data) => {
      stderr += data.toString();
    });

    child.on("close", (code) => {
      if (code === 0) {
        resolve({ stdout, stderr });
      } else {
        reject(new Error(`Command failed with exit code ${code}: ${stderr}`));
      }
    });

    child.on("error", (error) => {
      reject(error);
    });
  });
}

export interface ValidationResult {
  passed: boolean;
  totalCells: number;
  matchingCells: number;
  mismatchedCells: Array<{
    sheet: string;
    cell: string;
    expected: any;
    actual: any;
    difference?: number;
  }>;
  errors: string[];
}

export interface ValidationOptions {
  tolerance?: number; // Numeric tolerance for floating point comparisons
  ignoreEmptyCells?: boolean; // Whether to ignore empty cells in comparison
  compareFormattedValues?: boolean; // Whether to compare formatted strings instead of raw values
  verbose?: boolean;
}

export class ValidationComparator {
  /**
   * Compares actual Google Sheets data with generated code output
   * @param validationData The actual data from Google Sheets
   * @param generatedFilePath Path to the generated code file
   * @param inputData Input data to pass to the generated code
   * @param outputTabs The output tabs to validate
   * @param options Validation options
   * @returns Validation result with comparison details
   */
  async compareWithGeneratedCode(
    validationData: ValidationData,
    generatedFilePath: string,
    inputData: Record<string, any>,
    outputTabs: string[],
    options: ValidationOptions = {}
  ): Promise<ValidationResult> {
    const {
      tolerance = 1e-10,
      ignoreEmptyCells = true,
      verbose = false,
    } = options;

    const result: ValidationResult = {
      passed: true,
      totalCells: 0,
      matchingCells: 0,
      mismatchedCells: [],
      errors: [],
    };

    try {
      // Execute the generated code with input data
      if (verbose) {
        console.log(`🚀 Executing generated code: ${generatedFilePath}`);
      }

      const generatedOutput = await this.executeGeneratedCode(
        generatedFilePath,
        inputData,
        verbose
      );

      // Compare each output tab
      for (const tabName of outputTabs) {
        const actualSheetData = validationData.sheets.get(tabName);
        const generatedSheetData = generatedOutput[tabName];

        if (!actualSheetData) {
          result.errors.push(`Sheet "${tabName}" not found in validation data`);
          result.passed = false;
          continue;
        }

        if (!generatedSheetData) {
          result.errors.push(
            `Sheet "${tabName}" not found in generated output`
          );
          result.passed = false;
          continue;
        }

        // Compare cell by cell
        for (const [cellRef, actualValue] of actualSheetData) {
          // Skip empty cells if configured
          if (
            ignoreEmptyCells &&
            (actualValue === "" ||
              actualValue === null ||
              actualValue === undefined)
          ) {
            continue;
          }

          result.totalCells++;

          const generatedValue = generatedSheetData[cellRef];

          if (this.valuesMatch(actualValue, generatedValue, tolerance)) {
            result.matchingCells++;
          } else {
            result.mismatchedCells.push({
              sheet: tabName,
              cell: cellRef,
              expected: actualValue,
              actual: generatedValue,
              difference:
                typeof actualValue === "number" &&
                typeof generatedValue === "number"
                  ? Math.abs(actualValue - generatedValue)
                  : undefined,
            });
            result.passed = false;
          }
        }
      }

      if (verbose && result.mismatchedCells.length > 0) {
        console.log(`\n❌ Found ${result.mismatchedCells.length} mismatches:`);
        for (const mismatch of result.mismatchedCells.slice(0, 10)) {
          console.log(
            `  ${mismatch.sheet}!${mismatch.cell}: expected ${mismatch.expected}, got ${mismatch.actual}`
          );
        }
        if (result.mismatchedCells.length > 10) {
          console.log(
            `  ... and ${result.mismatchedCells.length - 10} more mismatches`
          );
        }
      }

      if (verbose) {
        const accuracy = (result.matchingCells / result.totalCells) * 100;
        console.log(
          `\n📊 Validation Results: ${result.matchingCells}/${result.totalCells} cells match (${accuracy.toFixed(2)}%)`
        );
      }
    } catch (error) {
      result.errors.push(
        `Error during validation: ${error instanceof Error ? error.message : String(error)}`
      );
      result.passed = false;
    }

    return result;
  }

  /**
   * Executes generated code and returns the output
   */
  private async executeGeneratedCode(
    filePath: string,
    inputData: Record<string, any>,
    verbose: boolean
  ): Promise<Record<string, any>> {
    const ext = path.extname(filePath);
    const inputJson = JSON.stringify(inputData);

    // Create a temporary file for input data
    const tempInputFile = path.join(
      path.dirname(filePath),
      `temp_input_${Date.now()}.json`
    );
    await fs.writeFile(tempInputFile, inputJson);

    try {
      let execCommand: string;
      let execArgs: string[];
      
      if (ext === ".ts") {
        // TypeScript - compile and run
        const jsPath = filePath.replace(/\.ts$/, ".js");
        await execSafe("npx", [
          "tsc",
          filePath,
          "--target",
          "es2020",
          "--module",
          "commonjs"
        ]);
        execCommand = "node";
        execArgs = [jsPath, "--input", tempInputFile];
      } else if (ext === ".py") {
        // Python
        execCommand = "python3";
        execArgs = [filePath, "--input", tempInputFile];
      } else if (ext === ".js") {
        // JavaScript
        execCommand = "node";
        execArgs = [filePath, "--input", tempInputFile];
      } else {
        throw new Error(`Unsupported file extension: ${ext}`);
      }

      if (verbose) {
        console.log(`⚙️  Executing: ${execCommand} ${execArgs.join(" ")}`);
      }

      const { stdout, stderr } = await execSafe(execCommand, execArgs);

      if (stderr && verbose) {
        console.warn(`⚠️  Stderr output: ${stderr}`);
      }

      // Parse the JSON output
      const output = JSON.parse(stdout);
      return output;
    } finally {
      // Clean up temp file
      await fs.unlink(tempInputFile).catch(() => {});
    }
  }

  /**
   * Checks if two values match within tolerance
   */
  private valuesMatch(actual: any, generated: any, tolerance: number): boolean {
    // Handle null/undefined
    if (actual === null || actual === undefined) {
      return generated === null || generated === undefined || generated === "";
    }

    // Handle numbers with tolerance
    if (typeof actual === "number" && typeof generated === "number") {
      return Math.abs(actual - generated) <= tolerance;
    }

    // Handle booleans
    if (typeof actual === "boolean" || typeof generated === "boolean") {
      return Boolean(actual) === Boolean(generated);
    }

    // Handle strings
    if (typeof actual === "string" || typeof generated === "string") {
      // Try to compare as numbers if both can be parsed
      const actualNum = Number(actual);
      const generatedNum = Number(generated);
      if (!Number.isNaN(actualNum) && !Number.isNaN(generatedNum)) {
        return Math.abs(actualNum - generatedNum) <= tolerance;
      }

      // Otherwise compare as strings
      return String(actual).trim() === String(generated).trim();
    }

    // Direct comparison for other types
    return actual === generated;
  }

  /**
   * Generates a validation report in markdown format
   */
  generateReport(results: ValidationResult[], outputPath?: string): string {
    let report = "# Google Sheets to Code Validation Report\n\n";
    report += `Generated: ${new Date().toISOString()}\n\n`;

    let totalPassed = 0;
    let totalFailed = 0;

    for (let i = 0; i < results.length; i++) {
      const result = results[i];
      const status = result.passed ? "✅ PASSED" : "❌ FAILED";

      report += `## Test ${i + 1}: ${status}\n\n`;
      report += `- Total cells validated: ${result.totalCells}\n`;
      report += `- Matching cells: ${result.matchingCells}\n`;
      report += `- Accuracy: ${((result.matchingCells / result.totalCells) * 100).toFixed(2)}%\n\n`;

      if (result.errors.length > 0) {
        report += "### Errors:\n";
        for (const error of result.errors) {
          report += `- ${error}\n`;
        }
        report += "\n";
      }

      if (result.mismatchedCells.length > 0) {
        report += "### Mismatched Cells (first 20):\n\n";
        report += "| Sheet | Cell | Expected | Actual | Difference |\n";
        report += "|-------|------|----------|--------|------------|\n";

        for (const mismatch of result.mismatchedCells.slice(0, 20)) {
          const diff =
            mismatch.difference !== undefined
              ? mismatch.difference.toExponential(2)
              : "N/A";
          report += `| ${mismatch.sheet} | ${mismatch.cell} | ${mismatch.expected} | ${mismatch.actual} | ${diff} |\n`;
        }

        if (result.mismatchedCells.length > 20) {
          report += `\n_...and ${result.mismatchedCells.length - 20} more mismatches_\n`;
        }
        report += "\n";
      }

      if (result.passed) {
        totalPassed++;
      } else {
        totalFailed++;
      }
    }

    report += `## Summary\n\n`;
    report += `- Tests Passed: ${totalPassed}\n`;
    report += `- Tests Failed: ${totalFailed}\n`;
    report += `- Success Rate: ${((totalPassed / results.length) * 100).toFixed(2)}%\n`;

    if (outputPath) {
      writeFileSync(outputPath, report);
    }

    return report;
  }
}
