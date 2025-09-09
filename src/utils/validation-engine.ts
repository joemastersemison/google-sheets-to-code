import type { Cell } from "../types/index.js";
import type {
  DataValidationRule,
  SheetValidationRules,
  ValidationCondition,
  ValidationError,
  ValidationResult,
  ValidationWarning,
} from "../types/validation.js";
import { parseA1Notation } from "./cell-reference.js";

export class ValidationEngine {
  private rules: Map<string, DataValidationRule[]> = new Map();

  addRules(sheetName: string, rules: DataValidationRule[]): void {
    this.rules.set(sheetName, rules);
  }

  validateSheet(sheetName: string, cells: Cell[][]): ValidationResult {
    const rules = this.rules.get(sheetName) || [];
    const errors: ValidationError[] = [];
    const warnings: ValidationWarning[] = [];

    for (const rule of rules) {
      const affectedCells = this.getCellsInRange(rule.range, cells);

      for (const { cell, value } of affectedCells) {
        const validationResult = this.validateValue(value, rule.condition);

        if (!validationResult.isValid) {
          const error: ValidationError = {
            cell,
            value,
            rule,
            message: this.getErrorMessage(rule, value, validationResult.reason),
          };

          if (rule.strict) {
            errors.push(error);
          } else {
            warnings.push({
              ...error,
              message: `Warning: ${error.message}`,
            });
          }
        }
      }
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings,
    };
  }

  private getCellsInRange(
    range: string,
    cells: Cell[][]
  ): Array<{ cell: string; value: any }> {
    const result: Array<{ cell: string; value: any }> = [];
    const { startCol, startRow, endCol, endRow } = this.parseRange(range);

    for (let row = startRow; row <= endRow && row < cells.length; row++) {
      for (
        let col = startCol;
        col <= endCol && col < (cells[row]?.length || 0);
        col++
      ) {
        const cellRef = this.getCellReference(col, row);
        const value = cells[row]?.[col]?.value;
        result.push({ cell: cellRef, value });
      }
    }

    return result;
  }

  private parseRange(range: string): {
    startCol: number;
    startRow: number;
    endCol: number;
    endRow: number;
  } {
    const parts = range.split(":");

    if (parts.length === 1) {
      const { col, row } = parseA1Notation(parts[0]);
      return {
        startCol: col,
        startRow: row,
        endCol: col,
        endRow: row,
      };
    }

    const start = parseA1Notation(parts[0]);
    const end = parseA1Notation(parts[1]);

    return {
      startCol: Math.min(start.col, end.col),
      startRow: Math.min(start.row, end.row),
      endCol: Math.max(start.col, end.col),
      endRow: Math.max(start.row, end.row),
    };
  }

  private getCellReference(col: number, row: number): string {
    let colRef = "";
    let colNum = col;

    while (colNum >= 0) {
      colRef = String.fromCharCode(65 + (colNum % 26)) + colRef;
      colNum = Math.floor(colNum / 26) - 1;
      if (colNum < 0) break;
    }

    return `${colRef}${row + 1}`;
  }

  private validateValue(
    value: any,
    condition: ValidationCondition
  ): { isValid: boolean; reason?: string } {
    switch (condition.type) {
      case "NUMBER_BETWEEN":
        return this.validateNumberBetween(value, condition);
      case "NUMBER_GREATER":
        return this.validateNumberGreater(value, condition);
      case "NUMBER_GREATER_EQUAL":
        return this.validateNumberGreaterEqual(value, condition);
      case "NUMBER_LESS":
        return this.validateNumberLess(value, condition);
      case "NUMBER_LESS_EQUAL":
        return this.validateNumberLessEqual(value, condition);
      case "NUMBER_EQUAL":
        return this.validateNumberEqual(value, condition);
      case "NUMBER_NOT_EQUAL":
        return this.validateNumberNotEqual(value, condition);
      case "NUMBER_NOT_BETWEEN":
        return this.validateNumberNotBetween(value, condition);
      case "TEXT_CONTAINS":
        return this.validateTextContains(value, condition);
      case "TEXT_NOT_CONTAINS":
        return this.validateTextNotContains(value, condition);
      case "TEXT_EQUALS":
        return this.validateTextEquals(value, condition);
      case "TEXT_IS_VALID_EMAIL":
        return this.validateEmail(value);
      case "TEXT_IS_VALID_URL":
        return this.validateUrl(value);
      case "DATE_BETWEEN":
        return this.validateDateBetween(value, condition);
      case "DATE_AFTER":
        return this.validateDateAfter(value, condition);
      case "DATE_BEFORE":
        return this.validateDateBefore(value, condition);
      case "DATE_ON_OR_AFTER":
        return this.validateDateOnOrAfter(value, condition);
      case "DATE_ON_OR_BEFORE":
        return this.validateDateOnOrBefore(value, condition);
      case "DATE_EQUAL":
        return this.validateDateEqual(value, condition);
      case "DATE_NOT_EQUAL":
        return this.validateDateNotEqual(value, condition);
      case "DATE_IS_VALID":
        return this.validateDateIsValid(value);
      case "ONE_OF_LIST":
        return this.validateOneOfList(value, condition);
      case "CHECKBOX":
        return this.validateCheckbox(value);
      case "LENGTH_BETWEEN":
        return this.validateLengthBetween(value, condition);
      case "LENGTH_GREATER":
        return this.validateLengthGreater(value, condition);
      case "LENGTH_LESS":
        return this.validateLengthLess(value, condition);
      default:
        return { isValid: true };
    }
  }

  private validateNumberBetween(
    value: any,
    condition: ValidationCondition
  ): { isValid: boolean; reason?: string } {
    const num = Number(value);
    if (isNaN(num)) {
      return { isValid: false, reason: "Value is not a number" };
    }
    const min = Number(condition.value1);
    const max = Number(condition.value2);
    const isValid = num >= min && num <= max;
    return {
      isValid,
      reason: isValid ? undefined : `Value must be between ${min} and ${max}`,
    };
  }

  private validateNumberGreater(
    value: any,
    condition: ValidationCondition
  ): { isValid: boolean; reason?: string } {
    const num = Number(value);
    if (isNaN(num)) {
      return { isValid: false, reason: "Value is not a number" };
    }
    const threshold = Number(condition.value1);
    const isValid = num > threshold;
    return {
      isValid,
      reason: isValid ? undefined : `Value must be greater than ${threshold}`,
    };
  }

  private validateNumberGreaterEqual(
    value: any,
    condition: ValidationCondition
  ): { isValid: boolean; reason?: string } {
    const num = Number(value);
    if (isNaN(num)) {
      return { isValid: false, reason: "Value is not a number" };
    }
    const threshold = Number(condition.value1);
    const isValid = num >= threshold;
    return {
      isValid,
      reason: isValid
        ? undefined
        : `Value must be greater than or equal to ${threshold}`,
    };
  }

  private validateNumberLess(
    value: any,
    condition: ValidationCondition
  ): { isValid: boolean; reason?: string } {
    const num = Number(value);
    if (isNaN(num)) {
      return { isValid: false, reason: "Value is not a number" };
    }
    const threshold = Number(condition.value1);
    const isValid = num < threshold;
    return {
      isValid,
      reason: isValid ? undefined : `Value must be less than ${threshold}`,
    };
  }

  private validateNumberLessEqual(
    value: any,
    condition: ValidationCondition
  ): { isValid: boolean; reason?: string } {
    const num = Number(value);
    if (isNaN(num)) {
      return { isValid: false, reason: "Value is not a number" };
    }
    const threshold = Number(condition.value1);
    const isValid = num <= threshold;
    return {
      isValid,
      reason: isValid
        ? undefined
        : `Value must be less than or equal to ${threshold}`,
    };
  }

  private validateNumberEqual(
    value: any,
    condition: ValidationCondition
  ): { isValid: boolean; reason?: string } {
    const num = Number(value);
    if (isNaN(num)) {
      return { isValid: false, reason: "Value is not a number" };
    }
    const target = Number(condition.value1);
    const isValid = num === target;
    return {
      isValid,
      reason: isValid ? undefined : `Value must be equal to ${target}`,
    };
  }

  private validateNumberNotEqual(
    value: any,
    condition: ValidationCondition
  ): { isValid: boolean; reason?: string } {
    const num = Number(value);
    if (isNaN(num)) {
      return { isValid: false, reason: "Value is not a number" };
    }
    const target = Number(condition.value1);
    const isValid = num !== target;
    return {
      isValid,
      reason: isValid ? undefined : `Value must not be equal to ${target}`,
    };
  }

  private validateNumberNotBetween(
    value: any,
    condition: ValidationCondition
  ): { isValid: boolean; reason?: string } {
    const num = Number(value);
    if (isNaN(num)) {
      return { isValid: false, reason: "Value is not a number" };
    }
    const min = Number(condition.value1);
    const max = Number(condition.value2);
    const isValid = num < min || num > max;
    return {
      isValid,
      reason: isValid
        ? undefined
        : `Value must not be between ${min} and ${max}`,
    };
  }

  private validateTextContains(
    value: any,
    condition: ValidationCondition
  ): { isValid: boolean; reason?: string } {
    const text = String(value);
    const searchText = String(condition.value1);
    const isValid = text.includes(searchText);
    return {
      isValid,
      reason: isValid ? undefined : `Text must contain "${searchText}"`,
    };
  }

  private validateTextNotContains(
    value: any,
    condition: ValidationCondition
  ): { isValid: boolean; reason?: string } {
    const text = String(value);
    const searchText = String(condition.value1);
    const isValid = !text.includes(searchText);
    return {
      isValid,
      reason: isValid ? undefined : `Text must not contain "${searchText}"`,
    };
  }

  private validateTextEquals(
    value: any,
    condition: ValidationCondition
  ): { isValid: boolean; reason?: string } {
    const text = String(value);
    const target = String(condition.value1);
    const isValid = text === target;
    return {
      isValid,
      reason: isValid ? undefined : `Text must be equal to "${target}"`,
    };
  }

  private validateEmail(value: any): { isValid: boolean; reason?: string } {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    const isValid = emailRegex.test(String(value));
    return {
      isValid,
      reason: isValid ? undefined : "Value must be a valid email address",
    };
  }

  private validateUrl(value: any): { isValid: boolean; reason?: string } {
    try {
      new URL(String(value));
      return { isValid: true };
    } catch {
      return { isValid: false, reason: "Value must be a valid URL" };
    }
  }

  private validateDateBetween(
    value: any,
    condition: ValidationCondition
  ): { isValid: boolean; reason?: string } {
    const date = this.parseDate(value);
    if (!date) {
      return { isValid: false, reason: "Value is not a valid date" };
    }
    const startDate = this.parseDate(condition.value1);
    const endDate = this.parseDate(condition.value2);
    if (!startDate || !endDate) {
      return { isValid: false, reason: "Invalid date range specified" };
    }
    const isValid = date >= startDate && date <= endDate;
    return {
      isValid,
      reason: isValid
        ? undefined
        : `Date must be between ${startDate.toISOString()} and ${endDate.toISOString()}`,
    };
  }

  private validateDateAfter(
    value: any,
    condition: ValidationCondition
  ): { isValid: boolean; reason?: string } {
    const date = this.parseDate(value);
    if (!date) {
      return { isValid: false, reason: "Value is not a valid date" };
    }
    const afterDate = this.parseDate(condition.value1);
    if (!afterDate) {
      return { isValid: false, reason: "Invalid comparison date specified" };
    }
    const isValid = date > afterDate;
    return {
      isValid,
      reason: isValid
        ? undefined
        : `Date must be after ${afterDate.toISOString()}`,
    };
  }

  private validateDateBefore(
    value: any,
    condition: ValidationCondition
  ): { isValid: boolean; reason?: string } {
    const date = this.parseDate(value);
    if (!date) {
      return { isValid: false, reason: "Value is not a valid date" };
    }
    const beforeDate = this.parseDate(condition.value1);
    if (!beforeDate) {
      return { isValid: false, reason: "Invalid comparison date specified" };
    }
    const isValid = date < beforeDate;
    return {
      isValid,
      reason: isValid
        ? undefined
        : `Date must be before ${beforeDate.toISOString()}`,
    };
  }

  private validateDateOnOrAfter(
    value: any,
    condition: ValidationCondition
  ): { isValid: boolean; reason?: string } {
    const date = this.parseDate(value);
    if (!date) {
      return { isValid: false, reason: "Value is not a valid date" };
    }
    const afterDate = this.parseDate(condition.value1);
    if (!afterDate) {
      return { isValid: false, reason: "Invalid comparison date specified" };
    }
    const isValid = date >= afterDate;
    return {
      isValid,
      reason: isValid
        ? undefined
        : `Date must be on or after ${afterDate.toISOString()}`,
    };
  }

  private validateDateOnOrBefore(
    value: any,
    condition: ValidationCondition
  ): { isValid: boolean; reason?: string } {
    const date = this.parseDate(value);
    if (!date) {
      return { isValid: false, reason: "Value is not a valid date" };
    }
    const beforeDate = this.parseDate(condition.value1);
    if (!beforeDate) {
      return { isValid: false, reason: "Invalid comparison date specified" };
    }
    const isValid = date <= beforeDate;
    return {
      isValid,
      reason: isValid
        ? undefined
        : `Date must be on or before ${beforeDate.toISOString()}`,
    };
  }

  private validateDateEqual(
    value: any,
    condition: ValidationCondition
  ): { isValid: boolean; reason?: string } {
    const date = this.parseDate(value);
    if (!date) {
      return { isValid: false, reason: "Value is not a valid date" };
    }
    const targetDate = this.parseDate(condition.value1);
    if (!targetDate) {
      return { isValid: false, reason: "Invalid comparison date specified" };
    }
    const isValid = date.getTime() === targetDate.getTime();
    return {
      isValid,
      reason: isValid
        ? undefined
        : `Date must be equal to ${targetDate.toISOString()}`,
    };
  }

  private validateDateNotEqual(
    value: any,
    condition: ValidationCondition
  ): { isValid: boolean; reason?: string } {
    const date = this.parseDate(value);
    if (!date) {
      return { isValid: false, reason: "Value is not a valid date" };
    }
    const targetDate = this.parseDate(condition.value1);
    if (!targetDate) {
      return { isValid: false, reason: "Invalid comparison date specified" };
    }
    const isValid = date.getTime() !== targetDate.getTime();
    return {
      isValid,
      reason: isValid
        ? undefined
        : `Date must not be equal to ${targetDate.toISOString()}`,
    };
  }

  private validateDateIsValid(value: any): {
    isValid: boolean;
    reason?: string;
  } {
    const date = this.parseDate(value);
    return {
      isValid: date !== null,
      reason: date ? undefined : "Value must be a valid date",
    };
  }

  private validateOneOfList(
    value: any,
    condition: ValidationCondition
  ): { isValid: boolean; reason?: string } {
    const values = condition.values || [];
    const isValid = values.includes(value);
    return {
      isValid,
      reason: isValid
        ? undefined
        : `Value must be one of: ${values.join(", ")}`,
    };
  }

  private validateCheckbox(value: any): { isValid: boolean; reason?: string } {
    const isValid =
      value === true ||
      value === false ||
      value === "TRUE" ||
      value === "FALSE";
    return {
      isValid,
      reason: isValid ? undefined : "Value must be a boolean (TRUE/FALSE)",
    };
  }

  private validateLengthBetween(
    value: any,
    condition: ValidationCondition
  ): { isValid: boolean; reason?: string } {
    const text = String(value);
    const min = Number(condition.value1);
    const max = Number(condition.value2);
    const isValid = text.length >= min && text.length <= max;
    return {
      isValid,
      reason: isValid
        ? undefined
        : `Text length must be between ${min} and ${max} characters`,
    };
  }

  private validateLengthGreater(
    value: any,
    condition: ValidationCondition
  ): { isValid: boolean; reason?: string } {
    const text = String(value);
    const min = Number(condition.value1);
    const isValid = text.length > min;
    return {
      isValid,
      reason: isValid
        ? undefined
        : `Text length must be greater than ${min} characters`,
    };
  }

  private validateLengthLess(
    value: any,
    condition: ValidationCondition
  ): { isValid: boolean; reason?: string } {
    const text = String(value);
    const max = Number(condition.value1);
    const isValid = text.length < max;
    return {
      isValid,
      reason: isValid
        ? undefined
        : `Text length must be less than ${max} characters`,
    };
  }

  private parseDate(value: any): Date | null {
    if (value instanceof Date) {
      return value;
    }

    const date = new Date(value);
    return isNaN(date.getTime()) ? null : date;
  }

  private getErrorMessage(
    rule: DataValidationRule,
    value: any,
    reason?: string
  ): string {
    if (rule.errorMessage?.message) {
      return rule.errorMessage.message;
    }

    return (
      reason ||
      `Value "${value}" violates validation rule for range ${rule.range}`
    );
  }

  generateValidationCode(language: "typescript" | "python"): string {
    if (language === "typescript") {
      return this.generateTypeScriptValidation();
    } else {
      return this.generatePythonValidation();
    }
  }

  private generateTypeScriptValidation(): string {
    const rulesArray: SheetValidationRules[] = [];

    for (const [sheetName, rules] of this.rules.entries()) {
      rulesArray.push({ sheetName, rules });
    }

    return `
// Data Validation Rules
export const validationRules = ${JSON.stringify(rulesArray, null, 2)};

export function validateData(sheetName: string, data: any[][]): { isValid: boolean; errors: any[] } {
  const sheetRules = validationRules.find(r => r.sheetName === sheetName);
  if (!sheetRules) {
    return { isValid: true, errors: [] };
  }

  const errors: any[] = [];
  
  for (const rule of sheetRules.rules) {
    // Validation logic would be implemented here
    // This is a placeholder for the actual validation
  }

  return { isValid: errors.length === 0, errors };
}
`;
  }

  private generatePythonValidation(): string {
    const rulesArray: SheetValidationRules[] = [];

    for (const [sheetName, rules] of this.rules.entries()) {
      rulesArray.push({ sheetName, rules });
    }

    return `
# Data Validation Rules
validation_rules = ${JSON.stringify(rulesArray, null, 2).replace(/true/g, "True").replace(/false/g, "False")}

def validate_data(sheet_name: str, data: list) -> dict:
    """Validate data against defined rules."""
    sheet_rules = next((r for r in validation_rules if r['sheetName'] == sheet_name), None)
    if not sheet_rules:
        return {'isValid': True, 'errors': []}
    
    errors = []
    
    for rule in sheet_rules['rules']:
        # Validation logic would be implemented here
        # This is a placeholder for the actual validation
        pass
    
    return {'isValid': len(errors) == 0, 'errors': errors}
`;
  }
}

export function parseValidationRules(rulesData: any): DataValidationRule[] {
  if (!Array.isArray(rulesData)) {
    return [];
  }

  return rulesData.map((rule) => ({
    range: rule.range || "A1",
    condition: rule.condition || { type: "TEXT_CONTAINS" as const },
    inputMessage: rule.inputMessage,
    errorMessage: rule.errorMessage,
    strict: rule.strict !== false,
    showErrorMessage: rule.showErrorMessage !== false,
    showInputMessage: rule.showInputMessage !== false,
  }));
}
