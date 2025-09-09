import { describe, expect, it } from "@jest/globals";
import type { Cell } from "../types/index.js";
import type { DataValidationRule } from "../types/validation.js";
import {
  parseValidationRules,
  ValidationEngine,
} from "../utils/validation-engine.js";

describe("ValidationEngine", () => {
  describe("parseValidationRules", () => {
    it("should parse validation rules from JSON", () => {
      const rulesData = [
        {
          range: "A1:A10",
          condition: {
            type: "NUMBER_BETWEEN",
            value1: 0,
            value2: 100,
          },
          strict: true,
          showErrorMessage: true,
          showInputMessage: false,
        },
      ];

      const rules = parseValidationRules(rulesData);
      expect(rules).toHaveLength(1);
      expect(rules[0].range).toBe("A1:A10");
      expect(rules[0].condition.type).toBe("NUMBER_BETWEEN");
      expect(rules[0].strict).toBe(true);
    });

    it("should handle empty or invalid data", () => {
      expect(parseValidationRules(null)).toEqual([]);
      expect(parseValidationRules(undefined)).toEqual([]);
      expect(parseValidationRules([])).toEqual([]);
    });
  });

  describe("validateSheet", () => {
    let engine: ValidationEngine;

    beforeEach(() => {
      engine = new ValidationEngine();
    });

    it("should validate numbers between range", () => {
      const rule: DataValidationRule = {
        range: "A1:A3",
        condition: {
          type: "NUMBER_BETWEEN",
          value1: 0,
          value2: 100,
        },
        strict: true,
        showErrorMessage: true,
        showInputMessage: false,
      };

      engine.addRules("Sheet1", [rule]);

      const cells: Cell[][] = [
        [
          { value: 50, row: 0, column: "A" },
          { value: "test", row: 0, column: "B" },
        ],
        [
          { value: 150, row: 1, column: "A" },
          { value: "test", row: 1, column: "B" },
        ],
        [
          { value: -10, row: 2, column: "A" },
          { value: "test", row: 2, column: "B" },
        ],
      ];

      const result = engine.validateSheet("Sheet1", cells);
      expect(result.isValid).toBe(false);
      expect(result.errors).toHaveLength(2); // 150 and -10 are out of range
      expect(result.warnings).toHaveLength(0);
    });

    it("should validate email addresses", () => {
      const rule: DataValidationRule = {
        range: "B1:B2",
        condition: {
          type: "TEXT_IS_VALID_EMAIL",
        },
        strict: true,
        showErrorMessage: true,
        showInputMessage: false,
      };

      engine.addRules("Sheet1", [rule]);

      const cells: Cell[][] = [
        [
          { value: "test", row: 0, column: "A" },
          { value: "user@example.com", row: 0, column: "B" },
        ],
        [
          { value: "test", row: 1, column: "A" },
          { value: "invalid-email", row: 1, column: "B" },
        ],
      ];

      const result = engine.validateSheet("Sheet1", cells);
      expect(result.isValid).toBe(false);
      expect(result.errors).toHaveLength(1); // 'invalid-email' is not valid
    });

    it("should validate one of list", () => {
      const rule: DataValidationRule = {
        range: "A1:A3",
        condition: {
          type: "ONE_OF_LIST",
          values: ["Apple", "Banana", "Orange"],
        },
        strict: true,
        showErrorMessage: true,
        showInputMessage: false,
      };

      engine.addRules("Sheet1", [rule]);

      const cells: Cell[][] = [
        [{ value: "Apple", row: 0, column: "A" }],
        [{ value: "Grape", row: 1, column: "A" }],
        [{ value: "Banana", row: 2, column: "A" }],
      ];

      const result = engine.validateSheet("Sheet1", cells);
      expect(result.isValid).toBe(false);
      expect(result.errors).toHaveLength(1); // 'Grape' is not in the list
    });

    it("should handle non-strict rules as warnings", () => {
      const rule: DataValidationRule = {
        range: "A1",
        condition: {
          type: "NUMBER_GREATER",
          value1: 10,
        },
        strict: false,
        showErrorMessage: true,
        showInputMessage: false,
      };

      engine.addRules("Sheet1", [rule]);

      const cells: Cell[][] = [[{ value: 5, row: 0, column: "A" }]];

      const result = engine.validateSheet("Sheet1", cells);
      expect(result.isValid).toBe(true); // Non-strict rules don't make validation fail
      expect(result.errors).toHaveLength(0);
      expect(result.warnings).toHaveLength(1);
    });

    it("should validate text length", () => {
      const rule: DataValidationRule = {
        range: "A1:A2",
        condition: {
          type: "LENGTH_BETWEEN",
          value1: 5,
          value2: 10,
        },
        strict: true,
        showErrorMessage: true,
        showInputMessage: false,
      };

      engine.addRules("Sheet1", [rule]);

      const cells: Cell[][] = [
        [{ value: "hello", row: 0, column: "A" }], // length 5, valid
        [{ value: "this is too long", row: 1, column: "A" }], // length > 10, invalid
      ];

      const result = engine.validateSheet("Sheet1", cells);
      expect(result.isValid).toBe(false);
      expect(result.errors).toHaveLength(1);
    });

    it("should validate dates", () => {
      const rule: DataValidationRule = {
        range: "A1:A3",
        condition: {
          type: "DATE_AFTER",
          value1: new Date("2024-01-01"),
        },
        strict: true,
        showErrorMessage: true,
        showInputMessage: false,
      };

      engine.addRules("Sheet1", [rule]);

      const cells: Cell[][] = [
        [{ value: "2024-06-01", row: 0, column: "A" }], // valid
        [{ value: "2023-12-01", row: 1, column: "A" }], // invalid
        [{ value: "not a date", row: 2, column: "A" }], // invalid
      ];

      const result = engine.validateSheet("Sheet1", cells);
      expect(result.isValid).toBe(false);
      expect(result.errors).toHaveLength(2);
    });

    it("should generate TypeScript validation code", () => {
      const rule: DataValidationRule = {
        range: "A1",
        condition: {
          type: "NUMBER_BETWEEN",
          value1: 0,
          value2: 100,
        },
        strict: true,
        showErrorMessage: true,
        showInputMessage: false,
      };

      engine.addRules("Sheet1", [rule]);
      const code = engine.generateValidationCode("typescript");

      expect(code).toContain("validationRules");
      expect(code).toContain("Sheet1");
      expect(code).toContain("NUMBER_BETWEEN");
      expect(code).toContain("validateData");
    });

    it("should generate Python validation code", () => {
      const rule: DataValidationRule = {
        range: "A1",
        condition: {
          type: "TEXT_EQUALS",
          value1: "test",
        },
        strict: false,
        showErrorMessage: true,
        showInputMessage: false,
      };

      engine.addRules("Sheet1", [rule]);
      const code = engine.generateValidationCode("python");

      expect(code).toContain("validation_rules");
      expect(code).toContain("Sheet1");
      expect(code).toContain("TEXT_EQUALS");
      expect(code).toContain("def validate_data");
    });
  });

  describe("individual validation functions", () => {
    let engine: ValidationEngine;

    beforeEach(() => {
      engine = new ValidationEngine();
    });

    it("should validate URLs", () => {
      const rule: DataValidationRule = {
        range: "A1:A3",
        condition: {
          type: "TEXT_IS_VALID_URL",
        },
        strict: true,
        showErrorMessage: true,
        showInputMessage: false,
      };

      engine.addRules("Sheet1", [rule]);

      const cells: Cell[][] = [
        [{ value: "https://example.com", row: 0, column: "A" }], // valid
        [{ value: "not a url", row: 1, column: "A" }], // invalid
        [{ value: "ftp://files.example.com", row: 2, column: "A" }], // valid
      ];

      const result = engine.validateSheet("Sheet1", cells);
      expect(result.errors).toHaveLength(1);
    });

    it("should validate checkboxes", () => {
      const rule: DataValidationRule = {
        range: "A1:A4",
        condition: {
          type: "CHECKBOX",
        },
        strict: true,
        showErrorMessage: true,
        showInputMessage: false,
      };

      engine.addRules("Sheet1", [rule]);

      const cells: Cell[][] = [
        [{ value: true, row: 0, column: "A" }], // valid
        [{ value: false, row: 1, column: "A" }], // valid
        [{ value: "TRUE", row: 2, column: "A" }], // valid
        [{ value: "yes", row: 3, column: "A" }], // invalid
      ];

      const result = engine.validateSheet("Sheet1", cells);
      expect(result.errors).toHaveLength(1);
    });
  });
});
