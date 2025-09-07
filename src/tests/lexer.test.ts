import { FormulaLexer } from "../parsers/lexer.js";

describe("FormulaLexer", () => {
  describe("tokenization", () => {
    it("should tokenize numbers", () => {
      const result = FormulaLexer.tokenize("42");

      expect(result.errors).toHaveLength(0);
      expect(result.tokens).toHaveLength(1);
      expect(result.tokens[0].tokenType.name).toBe("Number");
      expect(result.tokens[0].image).toBe("42");
    });

    it("should tokenize decimal numbers", () => {
      const result = FormulaLexer.tokenize("3.14");

      expect(result.errors).toHaveLength(0);
      expect(result.tokens).toHaveLength(1);
      expect(result.tokens[0].tokenType.name).toBe("Number");
      expect(result.tokens[0].image).toBe("3.14");
    });

    it("should tokenize scientific notation", () => {
      const result = FormulaLexer.tokenize("1.23e-4");

      expect(result.errors).toHaveLength(0);
      expect(result.tokens).toHaveLength(1);
      expect(result.tokens[0].tokenType.name).toBe("Number");
      expect(result.tokens[0].image).toBe("1.23e-4");
    });

    it("should tokenize negative numbers", () => {
      const result = FormulaLexer.tokenize("-42");

      expect(result.errors).toHaveLength(0);
      expect(result.tokens).toHaveLength(2);
      expect(result.tokens[0].tokenType.name).toBe("Minus");
      expect(result.tokens[1].tokenType.name).toBe("Number");
    });

    it("should tokenize strings", () => {
      const result = FormulaLexer.tokenize('"Hello World"');

      expect(result.errors).toHaveLength(0);
      expect(result.tokens).toHaveLength(1);
      expect(result.tokens[0].tokenType.name).toBe("String");
      expect(result.tokens[0].image).toBe('"Hello World"');
    });

    it("should tokenize strings with escaped quotes", () => {
      const result = FormulaLexer.tokenize('"Hello \\"World\\""');

      expect(result.errors).toHaveLength(0);
      expect(result.tokens).toHaveLength(1);
      expect(result.tokens[0].tokenType.name).toBe("String");
      expect(result.tokens[0].image).toBe('"Hello \\"World\\""');
    });

    it("should tokenize boolean literals", () => {
      const trueResult = FormulaLexer.tokenize("TRUE");
      expect(trueResult.tokens[0].tokenType.name).toBe("True");

      const falseResult = FormulaLexer.tokenize("FALSE");
      expect(falseResult.tokens[0].tokenType.name).toBe("False");

      const lowerTrueResult = FormulaLexer.tokenize("true");
      expect(lowerTrueResult.tokens[0].tokenType.name).toBe("True");
    });
  });

  describe("cell references", () => {
    it("should tokenize simple cell references", () => {
      const result = FormulaLexer.tokenize("A1");

      expect(result.errors).toHaveLength(0);
      expect(result.tokens).toHaveLength(1);
      expect(result.tokens[0].tokenType.name).toBe("CellReference");
      expect(result.tokens[0].image).toBe("A1");
    });

    it("should tokenize absolute cell references", () => {
      const result = FormulaLexer.tokenize("$A$1");

      expect(result.errors).toHaveLength(0);
      expect(result.tokens[0].tokenType.name).toBe("CellReference");
      expect(result.tokens[0].image).toBe("$A$1");
    });

    it("should tokenize mixed absolute references", () => {
      const result1 = FormulaLexer.tokenize("$A1");
      expect(result1.tokens[0].tokenType.name).toBe("CellReference");

      const result2 = FormulaLexer.tokenize("A$1");
      expect(result2.tokens[0].tokenType.name).toBe("CellReference");
    });

    it("should tokenize multi-letter column references", () => {
      const result = FormulaLexer.tokenize("AA100");

      expect(result.tokens[0].tokenType.name).toBe("CellReference");
      expect(result.tokens[0].image).toBe("AA100");
    });

    it("should tokenize range references", () => {
      const result = FormulaLexer.tokenize("A1:B10");

      expect(result.errors).toHaveLength(0);
      expect(result.tokens).toHaveLength(1);
      expect(result.tokens[0].tokenType.name).toBe("RangeReference");
      expect(result.tokens[0].image).toBe("A1:B10");
    });

    it("should tokenize absolute range references", () => {
      const result = FormulaLexer.tokenize("$A$1:$B$10");

      expect(result.tokens[0].tokenType.name).toBe("RangeReference");
      expect(result.tokens[0].image).toBe("$A$1:$B$10");
    });

    it("should tokenize sheet references", () => {
      const result = FormulaLexer.tokenize("Sheet1!A1");

      expect(result.errors).toHaveLength(0);
      expect(result.tokens).toHaveLength(2);
      expect(result.tokens[0].tokenType.name).toBe("SheetReference");
      expect(result.tokens[0].image).toBe("Sheet1!");
      expect(result.tokens[1].tokenType.name).toBe("CellReference");
      expect(result.tokens[1].image).toBe("A1");
    });

    it("should tokenize sheet names with spaces", () => {
      // Sheet names with spaces must be quoted in Google Sheets
      const result = FormulaLexer.tokenize("'My Sheet'!A1");

      expect(result.tokens[0].tokenType.name).toBe("SheetReference");
      expect(result.tokens[0].image).toBe("'My Sheet'!");
    });

    it("should tokenize sheet names with escaped quotes", () => {
      // Sheet names with apostrophes use doubled quotes for escaping
      const result = FormulaLexer.tokenize("'John''s Data'!A1");

      expect(result.tokens[0].tokenType.name).toBe("SheetReference");
      expect(result.tokens[0].image).toBe("'John''s Data'!");
    });
  });

  describe("operators", () => {
    it("should tokenize arithmetic operators", () => {
      const operators = ["+", "-", "*", "/", "^", "%"];
      const expectedTypes = [
        "Plus",
        "Minus",
        "Multiply",
        "Divide",
        "Power",
        "Percent",
      ];

      operators.forEach((op, index) => {
        const result = FormulaLexer.tokenize(op);
        expect(result.tokens[0].tokenType.name).toBe(expectedTypes[index]);
        expect(result.tokens[0].image).toBe(op);
      });
    });

    it("should tokenize comparison operators", () => {
      const operators = ["=", "<>", "<", ">", "<=", ">="];
      const expectedTypes = [
        "Equals",
        "NotEquals",
        "LessThan",
        "GreaterThan",
        "LessThanOrEqual",
        "GreaterThanOrEqual",
      ];

      operators.forEach((op, index) => {
        const result = FormulaLexer.tokenize(op);
        expect(result.tokens[0].tokenType.name).toBe(expectedTypes[index]);
        expect(result.tokens[0].image).toBe(op);
      });
    });

    it("should tokenize string concatenation operator", () => {
      const result = FormulaLexer.tokenize("&");

      expect(result.tokens[0].tokenType.name).toBe("Ampersand");
      expect(result.tokens[0].image).toBe("&");
    });
  });

  describe("functions", () => {
    it("should tokenize function names", () => {
      const result = FormulaLexer.tokenize("SUM");

      expect(result.errors).toHaveLength(0);
      expect(result.tokens).toHaveLength(1);
      expect(result.tokens[0].tokenType.name).toBe("Function");
      expect(result.tokens[0].image).toBe("SUM");
    });

    it("should tokenize case-insensitive function names", () => {
      const result = FormulaLexer.tokenize("sum");

      expect(result.tokens[0].tokenType.name).toBe("Function");
      expect(result.tokens[0].image).toBe("sum");
    });

    it("should tokenize functions with numbers and underscores", () => {
      const result = FormulaLexer.tokenize("FUNC_123");

      expect(result.tokens[0].tokenType.name).toBe("Function");
      expect(result.tokens[0].image).toBe("FUNC_123");
    });

    it("should distinguish TRUE/FALSE from function names", () => {
      const trueResult = FormulaLexer.tokenize("TRUE");
      expect(trueResult.tokens[0].tokenType.name).toBe("True");

      const customResult = FormulaLexer.tokenize("TRUECUSTOM");
      expect(customResult.tokens[0].tokenType.name).toBe("Function");
    });
  });

  describe("punctuation", () => {
    it("should tokenize parentheses", () => {
      const result = FormulaLexer.tokenize("()");

      expect(result.tokens).toHaveLength(2);
      expect(result.tokens[0].tokenType.name).toBe("LeftParen");
      expect(result.tokens[1].tokenType.name).toBe("RightParen");
    });

    it("should tokenize brackets", () => {
      const result = FormulaLexer.tokenize("[]");

      expect(result.tokens[0].tokenType.name).toBe("LeftBracket");
      expect(result.tokens[1].tokenType.name).toBe("RightBracket");
    });

    it("should tokenize braces", () => {
      const result = FormulaLexer.tokenize("{}");

      expect(result.tokens[0].tokenType.name).toBe("LeftBrace");
      expect(result.tokens[1].tokenType.name).toBe("RightBrace");
    });

    it("should tokenize separators", () => {
      const result = FormulaLexer.tokenize(",;:");

      expect(result.tokens).toHaveLength(3);
      expect(result.tokens[0].tokenType.name).toBe("Comma");
      expect(result.tokens[1].tokenType.name).toBe("Semicolon");
      expect(result.tokens[2].tokenType.name).toBe("Colon");
    });
  });

  describe("complex expressions", () => {
    it("should tokenize complete formula", () => {
      const result = FormulaLexer.tokenize("=SUM(A1:A10)+B1*2");

      expect(result.errors).toHaveLength(0);
      expect(result.tokens.map((t) => t.tokenType.name)).toEqual([
        "Equals",
        "Function",
        "LeftParen",
        "RangeReference",
        "RightParen",
        "Plus",
        "CellReference",
        "Multiply",
        "Number",
      ]);
    });

    it("should tokenize nested function calls", () => {
      const result = FormulaLexer.tokenize("=IF(A1>0,SUM(B:B),0)");

      expect(result.errors).toHaveLength(0);
      const tokenTypes = result.tokens.map((t) => t.tokenType.name);
      expect(tokenTypes).toContain("Function"); // IF
      expect(tokenTypes).toContain("Function"); // SUM
      expect(tokenTypes).toContain("GreaterThan");
      expect(tokenTypes).toContain("Comma");
    });

    it("should tokenize array literals", () => {
      const result = FormulaLexer.tokenize("={1,2;3,4}");

      expect(result.errors).toHaveLength(0);
      const tokenTypes = result.tokens.map((t) => t.tokenType.name);
      expect(tokenTypes).toEqual([
        "Equals",
        "LeftBrace",
        "Number",
        "Comma",
        "Number",
        "Semicolon",
        "Number",
        "Comma",
        "Number",
        "RightBrace",
      ]);
    });
  });

  describe("whitespace handling", () => {
    it("should skip whitespace tokens", () => {
      const result = FormulaLexer.tokenize("A1 + B1");

      expect(result.errors).toHaveLength(0);
      expect(result.tokens).toHaveLength(3);
      expect(result.tokens.map((t) => t.tokenType.name)).toEqual([
        "CellReference",
        "Plus",
        "CellReference",
      ]);
    });

    it("should handle tabs and newlines", () => {
      const result = FormulaLexer.tokenize("A1\t+\nB1");

      expect(result.tokens).toHaveLength(3);
      expect(result.tokens.map((t) => t.image)).toEqual(["A1", "+", "B1"]);
    });
  });

  describe("error handling", () => {
    it("should handle invalid characters gracefully", () => {
      const result = FormulaLexer.tokenize("A1 # B1");

      // The lexer should still tokenize valid parts
      expect(result.tokens.length).toBeGreaterThan(0);
      expect(result.tokens[0].image).toBe("A1");
    });

    it("should handle empty input", () => {
      const result = FormulaLexer.tokenize("");

      expect(result.errors).toHaveLength(0);
      expect(result.tokens).toHaveLength(0);
    });
  });
});
