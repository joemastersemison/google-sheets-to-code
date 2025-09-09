#!/usr/bin/env node

import { execSync } from "node:child_process";
import { existsSync, readFileSync, writeFileSync } from "node:fs";
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
}

interface ValidateOptions {
  config: string;
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
  .description("Validate a configuration file")
  .requiredOption("-c, --config <file>", "Path to configuration file")
  .action(async (options: ValidateOptions) => {
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

program.parse();
