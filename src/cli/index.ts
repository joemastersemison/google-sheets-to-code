#!/usr/bin/env node

import { execSync } from "node:child_process";
import { existsSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { Command } from "commander";
import { SheetToCodeConverter } from "../index.js";
import type { SheetConfig } from "../types/index.js";
import type { DataValidationRule } from "../types/validation.js";
import { parseValidationRules } from "../utils/validation-engine.js";

interface ConvertOptions {
  url: string;
  inputTabs: string;
  outputTabs: string;
  language: string;
  outputFile?: string;
  config?: string;
  credentials: string;
  verbose?: boolean;
  watch?: boolean;
  watchInterval?: string;
  validationRules?: string;
  generateTests?: boolean;
  validate?: boolean;
  validationInput?: string;
  validationTolerance?: string;
  validationSnapshots?: string;
  validationDelay?: string;
}

interface ValidateOptions {
  config: string;
  generatedFile: string;
  inputData?: string;
  tolerance?: string;
  snapshots?: string;
  delay?: string;
  verbose?: boolean;
  output?: string;
}

const program = new Command();

program
  .name("sheets-to-code")
  .description("Convert Google Sheets to TypeScript or Python code")
  .version("1.0.0");

program
  .command("convert")
  .description("Convert a Google Sheets document to code")
  .option("-u, --url <url>", "Google Sheets URL")
  .option("-i, --input-tabs <tabs>", "Comma-separated list of input tab names")
  .option(
    "-o, --output-tabs <tabs>",
    "Comma-separated list of output tab names"
  )
  .option(
    "-l, --language <language>",
    "Output language (typescript|python)",
    "typescript"
  )
  .option("-f, --output-file <file>", "Output file path")
  .option("--config <file>", "Path to configuration file")
  .option(
    "--credentials <file>",
    "Path to Google API credentials JSON file",
    "./credentials.json"
  )
  .option("--verbose", "Enable verbose output")
  .option("--watch", "Watch for changes and regenerate code automatically")
  .option(
    "--watch-interval <seconds>",
    "Interval between checks when watching (in seconds)",
    "30"
  )
  .option(
    "--validation-rules <file>",
    "Path to JSON file containing data validation rules"
  )
  .option("--generate-tests", "Generate unit tests for the output code")
  .option(
    "--validate",
    "Validate generated code against actual Google Sheets data"
  )
  .option(
    "--validation-input <file>",
    "Path to JSON file containing input data for validation"
  )
  .option(
    "--validation-tolerance <number>",
    "Numeric tolerance for floating point comparisons",
    "1e-10"
  )
  .option(
    "--validation-snapshots <number>",
    "Number of validation snapshots to take",
    "1"
  )
  .option(
    "--validation-delay <ms>",
    "Delay in milliseconds between validation snapshots",
    "1000"
  )
  .action(async (options: ConvertOptions) => {
    try {
      if (options.watch) {
        await watchAndConvert(options);
      } else {
        await convertSheet(options);
      }
    } catch (error) {
      console.error("Error:", (error as Error).message);
      process.exit(1);
    }
  });

program
  .command("setup")
  .description("Setup Google Sheets API credentials")
  .action(() => {
    console.log("Google Sheets API Setup");
    console.log("=====================");
    console.log();
    console.log("Option 1: Service Account (Recommended for automation)");
    console.log("-------------------------------------------------------");
    console.log(
      "1. Go to the Google Cloud Console: https://console.cloud.google.com/"
    );
    console.log("2. Create a new project or select an existing one");
    console.log("3. Enable the Google Sheets API:");
    console.log("   - Go to APIs & Services > Library");
    console.log("   - Search for 'Google Sheets API' and enable it");
    console.log("4. Create a Service Account:");
    console.log("   - Go to APIs & Services > Credentials");
    console.log("   - Click 'Create Credentials' > 'Service Account'");
    console.log("   - Fill in the service account details");
    console.log("5. Create and download the key:");
    console.log("   - Click on the created service account");
    console.log("   - Go to 'Keys' tab > 'Add Key' > 'Create new key'");
    console.log("   - Select 'JSON' format");
    console.log(
      '6. Save the downloaded file as "credentials.json" in your project root'
    );
    console.log("7. Share your Google Sheet with the service account email");
    console.log("   (found in the JSON file as 'client_email')");
    console.log();
    console.log("Option 2: OAuth2 (For user authentication)");
    console.log("------------------------------------------");
    console.log("1. Follow steps 1-3 from Option 1");
    console.log("2. Create OAuth2 credentials:");
    console.log("   - Go to APIs & Services > Credentials");
    console.log("   - Click 'Create Credentials' > 'OAuth client ID'");
    console.log("   - Application type: 'Desktop app'");
    console.log("3. Download the credentials JSON file");
    console.log('4. Save it as "credentials.json" in your project root');
    console.log();
    console.log("For detailed instructions, visit:");
    console.log("https://developers.google.com/sheets/api/quickstart/nodejs");
  });

program
  .command("validate")
  .description("Validate generated code against actual Google Sheets data")
  .requiredOption("-c, --config <file>", "Path to configuration file")
  .requiredOption("-g, --generated-file <file>", "Path to generated code file")
  .option("-i, --input-data <file>", "Path to JSON file containing input data")
  .option(
    "-t, --tolerance <number>",
    "Numeric tolerance for floating point comparisons",
    "1e-10"
  )
  .option(
    "-s, --snapshots <number>",
    "Number of validation snapshots to take",
    "1"
  )
  .option("-d, --delay <ms>", "Delay in milliseconds between snapshots", "1000")
  .option("-o, --output <file>", "Output validation report to file")
  .option("--verbose", "Enable verbose output")
  .action(async (options: ValidateOptions) => {
    try {
      await validateGeneratedCode(options);
    } catch (error) {
      console.error("Validation error:", (error as Error).message);
      process.exit(1);
    }
  });

program
  .command("validate-config")
  .description("Validate a configuration file")
  .requiredOption("-c, --config <file>", "Path to configuration file")
  .action(async (options: { config: string }) => {
    try {
      const config = loadConfig(options.config);
      console.log("Configuration is valid:");
      console.log(JSON.stringify(config, null, 2));
    } catch (error) {
      console.error("Configuration error:", (error as Error).message);
      process.exit(1);
    }
  });

async function convertSheet(options: ConvertOptions) {
  if (options.verbose) {
    console.log("Starting Google Sheets conversion...");
  }

  let config: SheetConfig;

  if (options.config) {
    config = loadConfig(options.config);
    // Override language if provided via command line
    if (options.language) {
      config.outputLanguage = options.language as "typescript" | "python";
    }
  } else {
    // Validate that required options are provided when not using config
    if (!options.url) {
      console.error("Error: Either --config or --url is required");
      process.exit(1);
    }
    if (!options.inputTabs) {
      console.error("Error: Either --config or --input-tabs is required");
      process.exit(1);
    }
    if (!options.outputTabs) {
      console.error("Error: Either --config or --output-tabs is required");
      process.exit(1);
    }

    config = {
      spreadsheetUrl: options.url,
      inputTabs: options.inputTabs.split(",").map((tab: string) => tab.trim()),
      outputTabs: options.outputTabs
        .split(",")
        .map((tab: string) => tab.trim()),
      outputLanguage: options.language as "typescript" | "python",
    };
  }

  // Validate configuration
  validateConfig(config);

  // Check credentials file
  if (!existsSync(options.credentials)) {
    throw new Error(`Credentials file not found: ${options.credentials}`);
  }

  if (options.verbose) {
    console.log("Configuration:", JSON.stringify(config, null, 2));
    console.log("Credentials file:", options.credentials);
  }

  // Load validation rules if provided
  let validationRules: DataValidationRule[] | undefined;
  if (options.validationRules) {
    if (!existsSync(options.validationRules)) {
      throw new Error(
        `Validation rules file not found: ${options.validationRules}`
      );
    }
    const rulesData = JSON.parse(readFileSync(options.validationRules, "utf8"));
    validationRules = parseValidationRules(rulesData);

    if (options.verbose) {
      console.log(`Loaded ${validationRules.length} validation rules`);
    }
  }

  // Convert the sheet (enable verbose mode by default to show progress)
  const converter = new SheetToCodeConverter(
    config,
    true,
    validationRules,
    options.generateTests
  );

  const { code: generatedCode, tests } = await converter.convert();

  // Determine output file
  let outputFile = options.outputFile;
  if (!outputFile) {
    const extension = config.outputLanguage === "typescript" ? ".ts" : ".py";
    outputFile = `generated-spreadsheet${extension}`;
  }

  // Write output
  writeFileSync(outputFile, generatedCode, "utf8");
  console.log(`✅ Successfully generated code: ${outputFile}`);

  // Write tests if generated
  if (tests && options.generateTests) {
    const testExtension =
      config.outputLanguage === "typescript" ? ".test.ts" : "_test.py";
    const testFile = outputFile.replace(/\.(ts|py)$/, testExtension);
    writeFileSync(testFile, tests, "utf8");
    console.log(`🧪 Successfully generated tests: ${testFile}`);
  }

  // If TypeScript, also compile to JavaScript
  if (config.outputLanguage === "typescript") {
    try {
      console.log("🔨 Compiling TypeScript to JavaScript...");

      // Generate a simple tsconfig if it doesn't exist in the output directory
      const outputDir =
        outputFile.substring(0, outputFile.lastIndexOf("/")) || ".";
      const tsconfigPath = `${outputDir}/tsconfig.generated.json`;

      const tsconfig = {
        compilerOptions: {
          target: "ES2020",
          module: "ES2020",
          lib: ["ES2020"],
          outDir: outputDir,
          rootDir: outputDir,
          strict: true,
          esModuleInterop: true,
          skipLibCheck: true,
          forceConsistentCasingInFileNames: true,
          declaration: true,
          declarationMap: true,
          moduleResolution: "node",
        },
        files: [outputFile],
      };

      writeFileSync(tsconfigPath, JSON.stringify(tsconfig, null, 2), "utf8");

      // Compile the TypeScript file
      try {
        execSync(`npx tsc --project ${tsconfigPath}`, { stdio: "inherit" });
      } catch (error: unknown) {
        console.warn(
          `⚠️  TypeScript compilation failed: ${(error as Error).message}`
        );
        // Don't exit, just warn and continue
      }

      // Clean up temporary tsconfig
      if (existsSync(tsconfigPath)) {
        execSync(`rm ${tsconfigPath}`, { stdio: "pipe" });
      }

      const jsFile = outputFile.replace(/\.ts$/, ".js");
      const dtsFile = outputFile.replace(/\.ts$/, ".d.ts");

      console.log(`✅ Successfully compiled to JavaScript: ${jsFile}`);
      console.log(`✅ Type definitions generated: ${dtsFile}`);
    } catch (error) {
      console.warn(
        `⚠️  Could not compile TypeScript to JavaScript: ${(error as Error).message}`
      );
      console.log(`💡 You can manually compile it with: npx tsc ${outputFile}`);
    }
  }

  if (options.verbose) {
    console.log(`Output language: ${config.outputLanguage}`);
    console.log(`Input tabs: ${config.inputTabs.join(", ")}`);
    console.log(`Output tabs: ${config.outputTabs.join(", ")}`);
    console.log(`Code length: ${generatedCode.length} characters`);
  }

  // Perform validation if requested
  if (options.validate) {
    console.log("\n🔍 Starting validation process...");

    // Load input data for validation
    let inputData: Record<string, any> = {};
    if (options.validationInput) {
      if (!existsSync(options.validationInput)) {
        console.warn(
          `⚠️  Validation input file not found: ${options.validationInput}`
        );
        console.log("Using empty input data for validation");
      } else {
        const inputContent = readFileSync(options.validationInput, "utf8");
        inputData = JSON.parse(inputContent);
      }
    }

    // Perform validation
    const snapshots = Number.parseInt(options.validationSnapshots || "1", 10);
    const delay = Number.parseInt(options.validationDelay || "1000", 10);
    const tolerance = Number.parseFloat(options.validationTolerance || "1e-10");

    if (snapshots > 1) {
      const validationResults =
        await converter.fetchValidationDataMultipleTimes(
          snapshots,
          delay,
          options.verbose
        );

      console.log(
        `\n📸 Collected ${validationResults.length} validation snapshots`
      );

      for (let i = 0; i < validationResults.length; i++) {
        console.log(
          `\n🧪 Validating snapshot ${i + 1}/${validationResults.length}...`
        );

        const result = await converter.validateGeneratedCode(
          outputFile,
          validationResults[i],
          inputData,
          {
            tolerance,
            verbose: options.verbose,
            ignoreEmptyCells: true,
          }
        );

        const status = result.passed ? "✅ PASSED" : "❌ FAILED";
        console.log(`Snapshot ${i + 1}: ${status}`);
        console.log(
          `Accuracy: ${((result.matchingCells / result.totalCells) * 100).toFixed(2)}% (${result.matchingCells}/${result.totalCells} cells)`
        );

        if (!result.passed && result.mismatchedCells.length > 0) {
          console.log("First 5 mismatches:");
          for (const mismatch of result.mismatchedCells.slice(0, 5)) {
            console.log(
              `  - ${mismatch.sheet}!${mismatch.cell}: expected ${mismatch.expected}, got ${mismatch.actual}`
            );
          }
        }
      }
    } else {
      const validationData = await converter.fetchValidationData(
        options.verbose
      );

      const result = await converter.validateGeneratedCode(
        outputFile,
        validationData,
        inputData,
        {
          tolerance,
          verbose: options.verbose,
          ignoreEmptyCells: true,
        }
      );

      const status = result.passed
        ? "✅ VALIDATION PASSED"
        : "❌ VALIDATION FAILED";
      console.log(`\n${status}`);
      console.log(
        `Accuracy: ${((result.matchingCells / result.totalCells) * 100).toFixed(2)}% (${result.matchingCells}/${result.totalCells} cells)`
      );

      if (!result.passed) {
        if (result.errors.length > 0) {
          console.log("\nErrors:");
          for (const error of result.errors) {
            console.log(`  - ${error}`);
          }
        }

        if (result.mismatchedCells.length > 0) {
          console.log("\nFirst 10 mismatches:");
          for (const mismatch of result.mismatchedCells.slice(0, 10)) {
            console.log(
              `  - ${mismatch.sheet}!${mismatch.cell}: expected ${mismatch.expected}, got ${mismatch.actual}`
            );
          }
          if (result.mismatchedCells.length > 10) {
            console.log(
              `  ... and ${result.mismatchedCells.length - 10} more mismatches`
            );
          }
        }
      }
    }
  }
}

function loadConfig(configPath: string): SheetConfig {
  if (!existsSync(configPath)) {
    throw new Error(`Configuration file not found: ${configPath}`);
  }

  try {
    const configContent = readFileSync(configPath, "utf8");
    const config = JSON.parse(configContent);

    return {
      spreadsheetUrl: config.spreadsheetUrl,
      inputTabs: config.inputTabs || [],
      outputTabs: config.outputTabs || [],
      outputLanguage: config.outputLanguage || "typescript",
    };
  } catch (error) {
    throw new Error(`Invalid configuration file: ${(error as Error).message}`);
  }
}

function validateConfig(config: SheetConfig) {
  if (!config.spreadsheetUrl) {
    throw new Error("Missing spreadsheet URL");
  }

  if (!config.spreadsheetUrl.includes("docs.google.com/spreadsheets")) {
    throw new Error("Invalid Google Sheets URL format");
  }

  if (!config.inputTabs || config.inputTabs.length === 0) {
    throw new Error("At least one input tab must be specified");
  }

  if (!config.outputTabs || config.outputTabs.length === 0) {
    throw new Error("At least one output tab must be specified");
  }

  if (!["typescript", "python"].includes(config.outputLanguage)) {
    throw new Error('Output language must be "typescript" or "python"');
  }

  // Check for overlap between input and output tabs
  const inputSet = new Set(config.inputTabs);
  const outputSet = new Set(config.outputTabs);
  const overlap = [...inputSet].filter((tab) => outputSet.has(tab));

  if (overlap.length > 0) {
    console.warn(
      `Warning: Tabs specified as both input and output: ${overlap.join(", ")}`
    );
  }
}

async function watchAndConvert(options: ConvertOptions) {
  const interval = Number.parseInt(options.watchInterval || "30", 10) * 1000;

  // Warn about potential rate limits for short intervals
  if (interval < 60000) {
    // Less than 1 minute
    console.warn(
      "⚠️  Short intervals may hit Google API rate limits. Consider 60+ second intervals for production."
    );
  }

  console.log("👁️  Watch mode enabled");
  console.log(`📊 Checking for changes every ${interval / 1000} seconds`);
  console.log("Press Ctrl+C to stop watching\n");

  // Initial conversion
  await convertSheet(options);

  let isConverting = false;

  // Set up periodic check
  const checkForChanges = async () => {
    if (isConverting) {
      console.log("⏳ Previous conversion still in progress, skipping...");
      return;
    }

    try {
      isConverting = true;

      console.log(
        `\n🔍 Checking for updates... (${new Date().toLocaleTimeString()})`
      );

      // Perform the conversion
      await convertSheet(options);
    } catch (error) {
      console.error(`❌ Error during conversion: ${(error as Error).message}`);
    } finally {
      isConverting = false;
    }
  };

  // Set up the interval
  const intervalId = setInterval(checkForChanges, interval);

  // Handle graceful shutdown
  process.on("SIGINT", () => {
    console.log("\n\n👋 Stopping watch mode...");
    clearInterval(intervalId);
    process.exit(0);
  });

  // Keep the process running
  process.stdin.resume();
}

async function validateGeneratedCode(options: ValidateOptions) {
  // Load configuration
  const config = loadConfig(options.config);
  validateConfig(config);

  // Load input data
  let inputData: Record<string, any> = {};
  if (options.inputData) {
    if (!existsSync(options.inputData)) {
      throw new Error(`Input data file not found: ${options.inputData}`);
    }
    const inputContent = readFileSync(options.inputData, "utf8");
    inputData = JSON.parse(inputContent);
  }

  // Check generated file exists
  if (!existsSync(options.generatedFile)) {
    throw new Error(`Generated file not found: ${options.generatedFile}`);
  }

  // Create converter
  const converter = new SheetToCodeConverter(config, options.verbose || false);

  // Parse validation options
  const snapshots = Number.parseInt(options.snapshots || "1", 10);
  const delay = Number.parseInt(options.delay || "1000", 10);
  const tolerance = Number.parseFloat(options.tolerance || "1e-10");

  // Use absolute import or handle import errors
  let comparator;
  try {
    const { ValidationComparator } = await import(
      path.resolve(path.dirname(new URL(import.meta.url).pathname), "../utils/validation-comparator.js")
    );
    comparator = new ValidationComparator();
  } catch (error) {
    throw new Error(`Failed to load validation comparator: ${error.message}`);
  }

  const allResults: any[] = [];

  if (snapshots > 1) {
    console.log(
      `📸 Fetching ${snapshots} validation snapshots with ${delay}ms delay...`
    );

    const validationSnapshots =
      await converter.fetchValidationDataMultipleTimes(
        snapshots,
        delay,
        options.verbose
      );

    for (let i = 0; i < validationSnapshots.length; i++) {
      console.log(
        `\n🧪 Validating snapshot ${i + 1}/${validationSnapshots.length}...`
      );

      const result = await converter.validateGeneratedCode(
        options.generatedFile,
        validationSnapshots[i],
        inputData,
        {
          tolerance,
          verbose: options.verbose,
          ignoreEmptyCells: true,
        }
      );

      allResults.push(result);

      const status = result.passed ? "✅ PASSED" : "❌ FAILED";
      console.log(`Snapshot ${i + 1}: ${status}`);
      console.log(
        `Accuracy: ${((result.matchingCells / result.totalCells) * 100).toFixed(2)}%`
      );
    }
  } else {
    console.log("📊 Fetching validation data from Google Sheets...");

    const validationData = await converter.fetchValidationData(options.verbose);

    console.log("🧪 Validating generated code...");

    const result = await converter.validateGeneratedCode(
      options.generatedFile,
      validationData,
      inputData,
      {
        tolerance,
        verbose: options.verbose,
        ignoreEmptyCells: true,
      }
    );

    allResults.push(result);

    const status = result.passed
      ? "✅ VALIDATION PASSED"
      : "❌ VALIDATION FAILED";
    console.log(`\n${status}`);
    console.log(
      `Accuracy: ${((result.matchingCells / result.totalCells) * 100).toFixed(2)}% (${result.matchingCells}/${result.totalCells} cells)`
    );

    if (!result.passed && result.mismatchedCells.length > 0) {
      console.log("\nFirst 10 mismatches:");
      for (const mismatch of result.mismatchedCells.slice(0, 10)) {
        console.log(
          `  - ${mismatch.sheet}!${mismatch.cell}: expected ${mismatch.expected}, got ${mismatch.actual}`
        );
      }
    }
  }

  // Generate report if requested
  if (options.output) {
    const report = comparator.generateReport(allResults, options.output);
    console.log(`\n📄 Validation report saved to: ${options.output}`);
  }
}

program.parse();
