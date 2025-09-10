export type ValidationType =
  | "NUMBER_BETWEEN"
  | "NUMBER_GREATER"
  | "NUMBER_GREATER_EQUAL"
  | "NUMBER_LESS"
  | "NUMBER_LESS_EQUAL"
  | "NUMBER_EQUAL"
  | "NUMBER_NOT_EQUAL"
  | "NUMBER_NOT_BETWEEN"
  | "TEXT_CONTAINS"
  | "TEXT_NOT_CONTAINS"
  | "TEXT_EQUALS"
  | "TEXT_IS_VALID_EMAIL"
  | "TEXT_IS_VALID_URL"
  | "DATE_BETWEEN"
  | "DATE_AFTER"
  | "DATE_BEFORE"
  | "DATE_ON_OR_AFTER"
  | "DATE_ON_OR_BEFORE"
  | "DATE_EQUAL"
  | "DATE_NOT_EQUAL"
  | "DATE_IS_VALID"
  | "ONE_OF_LIST"
  | "ONE_OF_RANGE"
  | "CUSTOM_FORMULA"
  | "CHECKBOX"
  | "LENGTH_BETWEEN"
  | "LENGTH_GREATER"
  | "LENGTH_LESS";

export interface ValidationCondition {
  type: ValidationType;
  operator?:
    | "BETWEEN"
    | "NOT_BETWEEN"
    | "EQUAL"
    | "NOT_EQUAL"
    | "GREATER"
    | "GREATER_EQUAL"
    | "LESS"
    | "LESS_EQUAL";
  value1?: string | number | Date;
  value2?: string | number | Date;
  values?: (string | number)[];
  formula?: string;
  showDropdown?: boolean;
}

export interface DataValidationRule {
  range: string;
  condition: ValidationCondition;
  inputMessage?: {
    title?: string;
    message?: string;
  };
  errorMessage?: {
    title?: string;
    message?: string;
  };
  strict: boolean;
  showErrorMessage: boolean;
  showInputMessage: boolean;
}

// Type for cell values
export type CellValue = string | number | boolean | Date | null | undefined;

export interface ValidationResult {
  isValid: boolean;
  errors: ValidationError[];
  warnings: ValidationWarning[];
}

export interface ValidationError {
  cell: string;
  value: CellValue;
  rule: DataValidationRule;
  message: string;
}

export interface ValidationWarning {
  cell: string;
  value: CellValue;
  rule: DataValidationRule;
  message: string;
}

export interface SheetValidationRules {
  sheetName: string;
  rules: DataValidationRule[];
}
