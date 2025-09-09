import { google } from "googleapis";
import { GoogleSheetsReader } from "./sheets-reader.js";

// Type for cell values that can come from Google Sheets
export type CellValue = string | number | boolean | Date | null | undefined;

export interface ValidationData {
  timestamp: Date;
  sheets: Map<string, Map<string, CellValue>>;
}

export class ValidationDataFetcher {
  private sheetsReader: GoogleSheetsReader;

  constructor() {
    this.sheetsReader = new GoogleSheetsReader();
  }

  /**
   * Fetches actual cell values from Google Sheets for validation purposes
   * @param spreadsheetUrl The URL of the Google Sheet
   * @param sheetNames The sheets to fetch data from
   * @param verbose Whether to output verbose logs
   * @returns Map of sheet names to cell references to actual values
   */
  async fetchValidationData(
    spreadsheetUrl: string,
    sheetNames: string[],
    verbose = false
  ): Promise<ValidationData> {
    // Authenticate if needed
    await this.sheetsReader.authenticate();

    const spreadsheetId = this.extractSpreadsheetId(spreadsheetUrl);
    if (verbose)
      console.log(`📊 Fetching validation data from: ${spreadsheetId}`);

    // Use the authenticated client from GoogleSheetsReader
    const auth = this.sheetsReader.getAuth();
    if (!auth) {
      throw new Error("Authentication failed - no auth client available");
    }

    const sheets = google.sheets({ version: "v4", auth });
    const result = new Map<string, Map<string, CellValue>>();

    for (const sheetName of sheetNames) {
      if (verbose) console.log(`📄 Fetching values from "${sheetName}"...`);

      // Fetch with UNFORMATTED_VALUE to get raw numeric values
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `${sheetName}!A1:ZZ1000`,
        valueRenderOption: "UNFORMATTED_VALUE",
      });

      const rows = response.data.values || [];
      const cellValues = new Map<string, CellValue>();

      for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
        const row = rows[rowIndex];

        for (let colIndex = 0; colIndex < row.length; colIndex++) {
          const cellValue = row[colIndex];
          const column = this.indexToColumn(colIndex);
          const cellRef = `${column}${rowIndex + 1}`;

          // Store the actual value (could be number, string, boolean, etc.)
          cellValues.set(cellRef, cellValue);
        }
      }

      result.set(sheetName, cellValues);

      if (verbose) {
        console.log(
          `✅ Sheet "${sheetName}" values fetched: ${cellValues.size} cells`
        );
      }
    }

    return {
      timestamp: new Date(),
      sheets: result,
    };
  }

  /**
   * Fetches validation data multiple times with a delay between fetches
   * Useful for testing time-based functions like NOW() or TODAY()
   * @param spreadsheetUrl The URL of the Google Sheet
   * @param sheetNames The sheets to fetch data from
   * @param count Number of times to fetch
   * @param delayMs Delay in milliseconds between fetches
   * @param verbose Whether to output verbose logs
   * @returns Array of validation data snapshots
   */
  async fetchValidationDataMultipleTimes(
    spreadsheetUrl: string,
    sheetNames: string[],
    count: number,
    delayMs: number,
    verbose = false
  ): Promise<ValidationData[]> {
    const results: ValidationData[] = [];

    for (let i = 0; i < count; i++) {
      if (verbose) {
        console.log(`📸 Fetching snapshot ${i + 1}/${count}...`);
      }

      const data = await this.fetchValidationData(
        spreadsheetUrl,
        sheetNames,
        verbose
      );
      results.push(data);

      if (i < count - 1) {
        if (verbose) {
          console.log(`⏳ Waiting ${delayMs}ms before next fetch...`);
        }
        await new Promise((resolve) => setTimeout(resolve, delayMs));
      }
    }

    return results;
  }

  private extractSpreadsheetId(url: string): string {
    const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (!match) {
      throw new Error("Invalid Google Sheets URL");
    }
    return match[1];
  }

  private indexToColumn(index: number): string {
    let column = "";
    while (index >= 0) {
      column = String.fromCharCode(65 + (index % 26)) + column;
      index = Math.floor(index / 26) - 1;
    }
    return column;
  }
}
