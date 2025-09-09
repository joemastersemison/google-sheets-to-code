import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { authenticate } from "@google-cloud/local-auth";
import { GoogleAuth } from "google-auth-library";
import { google } from "googleapis";
import type { Cell, Sheet } from "../types/index.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

export class GoogleSheetsReader {
  // biome-ignore lint/suspicious/noExplicitAny: Google Auth can return various client types
  private auth: any;

  // Getter to allow other classes to reuse the authenticated client
  getAuth() {
    return this.auth;
  }

  async authenticate() {
    const SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"];
    const CREDENTIALS_PATH = path.join(process.cwd(), "credentials.json");

    if (!existsSync(CREDENTIALS_PATH)) {
      throw new Error(
        `Credentials file not found at ${CREDENTIALS_PATH}. Please run 'npm run cli -- setup' for instructions.`
      );
    }

    // Read the credentials file to determine the type
    const credentialsContent = JSON.parse(
      readFileSync(CREDENTIALS_PATH, "utf8")
    );

    if (credentialsContent.type === "service_account") {
      // Service account authentication
      const auth = new GoogleAuth({
        keyFilename: CREDENTIALS_PATH,
        scopes: SCOPES,
      });
      this.auth = await auth.getClient();
    } else if (credentialsContent.installed || credentialsContent.web) {
      // OAuth2 authentication
      console.log(
        "🌐 OAuth2 authentication detected. Opening browser for authorization..."
      );
      console.log("⏳ This may take a moment on first run...");

      try {
        this.auth = await authenticate({
          scopes: SCOPES,
          keyfilePath: CREDENTIALS_PATH,
        });
      } catch (error) {
        console.error("\n❌ OAuth2 authentication failed.");
        console.error("This might be because:");
        console.error("1. You're running in a non-interactive environment");
        console.error("2. The browser couldn't be opened");
        console.error("3. The authentication was cancelled\n");
        console.error(
          "💡 Tip: Consider using a service account for automation."
        );
        console.error("   Run 'npm run cli -- setup' for instructions.\n");
        throw error;
      }
    } else {
      throw new Error(
        "Invalid credentials file. Expected either a service account key or OAuth2 credentials."
      );
    }
  }

  async readSheets(
    spreadsheetUrl: string,
    sheetNames: string[],
    verbose = false
  ): Promise<{ sheets: Map<string, Sheet>; namedRanges: Map<string, string> }> {
    if (!this.auth) {
      if (verbose) console.log("🔐 Authenticating with Google Sheets API...");
      await this.authenticate();
      if (verbose) console.log("✅ Authentication successful");
    }

    const spreadsheetId = this.extractSpreadsheetId(spreadsheetUrl);
    if (verbose) console.log(`📊 Reading spreadsheet: ${spreadsheetId}`);

    const sheets = google.sheets({ version: "v4", auth: this.auth });

    // Fetch spreadsheet metadata to get named ranges
    const metadataResponse = await sheets.spreadsheets.get({
      spreadsheetId,
      fields: "namedRanges,sheets(properties(title,sheetId))",
    });

    // Process named ranges
    const namedRanges = new Map<string, string>();
    if (metadataResponse.data.namedRanges) {
      if (verbose)
        console.log(
          `📝 Processing ${metadataResponse.data.namedRanges.length} named ranges...`
        );

      for (const namedRange of metadataResponse.data.namedRanges) {
        const name = namedRange.name || "";
        const range = namedRange.range;

        if (range) {
          // Convert range to A1 notation
          const sheetId = range.sheetId;
          const sheetName =
            metadataResponse.data.sheets?.find(
              (s) => s.properties?.sheetId === sheetId
            )?.properties?.title || "";

          let a1Notation = "";
          if (
            range.startColumnIndex !== undefined &&
            range.startColumnIndex !== null &&
            range.endColumnIndex !== undefined &&
            range.endColumnIndex !== null &&
            range.startRowIndex !== undefined &&
            range.endRowIndex !== undefined
          ) {
            const startCol = this.indexToColumn(range.startColumnIndex);
            const endCol = this.indexToColumn(range.endColumnIndex - 1);
            const startRow = (range.startRowIndex || 0) + 1;
            const endRow = range.endRowIndex || 1;

            if (startCol === endCol && startRow === endRow) {
              // Single cell
              a1Notation = `${sheetName}!${startCol}${startRow}`;
            } else {
              // Range
              a1Notation = `${sheetName}!${startCol}${startRow}:${endCol}${endRow}`;
            }
          }

          namedRanges.set(name, a1Notation);
          if (verbose)
            console.log(`  ✅ Named range "${name}" -> ${a1Notation}`);
        }
      }
    }

    const result = new Map<string, Sheet>();

    for (const sheetName of sheetNames) {
      if (verbose) console.log(`📄 Fetching sheet "${sheetName}"...`);

      const response = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `${sheetName}!A1:ZZ1000`,
        valueRenderOption: "FORMULA",
      });

      const formattedResponse = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `${sheetName}!A1:ZZ1000`,
        valueRenderOption: "FORMATTED_VALUE",
      });

      const rows = response.data.values || [];
      const formattedRows = formattedResponse.data.values || [];
      const cells = new Map<string, Cell>();

      for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
        const row = rows[rowIndex];
        const formattedRow = formattedRows[rowIndex] || [];

        for (let colIndex = 0; colIndex < row.length; colIndex++) {
          const cellValue = row[colIndex];
          const formattedValue = formattedRow[colIndex];
          const column = this.indexToColumn(colIndex);
          const cellRef = `${column}${rowIndex + 1}`;

          const cell: Cell = {
            row: rowIndex + 1,
            column,
            value: cellValue,
            formattedValue,
          };

          if (typeof cellValue === "string" && cellValue.startsWith("=")) {
            cell.formula = cellValue;
          }

          cells.set(cellRef, cell);
        }
      }

      result.set(sheetName, {
        name: sheetName,
        cells,
        range: {
          startRow: 1,
          endRow: rows.length,
          startColumn: "A",
          endColumn: this.indexToColumn(
            Math.max(...rows.map((r) => r.length - 1))
          ),
        },
      });

      if (verbose) {
        console.log(
          `✅ Sheet "${sheetName}" loaded: ${cells.size} cells with data`
        );
      }
    }

    if (verbose) console.log(`📦 Total sheets loaded: ${result.size}`);
    return { sheets: result, namedRanges };
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
