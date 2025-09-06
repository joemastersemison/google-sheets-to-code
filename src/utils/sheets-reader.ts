import { google } from 'googleapis';
import { authenticate } from '@google-cloud/local-auth';
import { Sheet, Cell } from '../types/index.js';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

export class GoogleSheetsReader {
  private auth: any;

  async authenticate() {
    const SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly'];
    const TOKEN_PATH = path.join(process.cwd(), 'token.json');
    const CREDENTIALS_PATH = path.join(process.cwd(), 'credentials.json');

    this.auth = await authenticate({
      scopes: SCOPES,
      keyfilePath: CREDENTIALS_PATH,
    });
  }

  async readSheets(spreadsheetUrl: string, sheetNames: string[]): Promise<Map<string, Sheet>> {
    if (!this.auth) {
      await this.authenticate();
    }

    const spreadsheetId = this.extractSpreadsheetId(spreadsheetUrl);
    const sheets = google.sheets({ version: 'v4', auth: this.auth });
    const result = new Map<string, Sheet>();

    for (const sheetName of sheetNames) {
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `${sheetName}!A1:ZZ1000`,
        valueRenderOption: 'FORMULA',
      });

      const formattedResponse = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `${sheetName}!A1:ZZ1000`,
        valueRenderOption: 'FORMATTED_VALUE',
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

          if (typeof cellValue === 'string' && cellValue.startsWith('=')) {
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
          startColumn: 'A',
          endColumn: this.indexToColumn(Math.max(...rows.map(r => r.length - 1))),
        },
      });
    }

    return result;
  }

  private extractSpreadsheetId(url: string): string {
    const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (!match) {
      throw new Error('Invalid Google Sheets URL');
    }
    return match[1];
  }

  private indexToColumn(index: number): string {
    let column = '';
    while (index >= 0) {
      column = String.fromCharCode(65 + (index % 26)) + column;
      index = Math.floor(index / 26) - 1;
    }
    return column;
  }
}