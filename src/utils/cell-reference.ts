export function parseA1Notation(ref: string): { col: number; row: number } {
  const match = ref.match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    throw new Error(`Invalid A1 notation: ${ref}`);
  }

  const colStr = match[1];
  const rowStr = match[2];

  let col = 0;
  for (let i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 64);
  }
  col--; // Convert to 0-based index

  const row = parseInt(rowStr, 10) - 1; // Convert to 0-based index

  return { col, row };
}

export function toA1Notation(col: number, row: number): string {
  let colRef = "";
  let colNum = col + 1;

  while (colNum > 0) {
    colNum--;
    colRef = String.fromCharCode(65 + (colNum % 26)) + colRef;
    colNum = Math.floor(colNum / 26);
  }

  return `${colRef}${row + 1}`;
}
