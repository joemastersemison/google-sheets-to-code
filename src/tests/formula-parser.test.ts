import { FormulaParser } from "../parsers/formula-parser.js";

describe("FormulaParser", () => {
  let parser: FormulaParser;

  beforeEach(() => {
    parser = new FormulaParser();
  });

  describe("Literals", () => {
    it("should parse numbers", () => {
      const result = parser.parse("=42");
      expect(result).toEqual({
        type: "literal",
        value: "42",
      });
    });

    it("should parse decimal numbers", () => {
      const result = parser.parse("=3.14");
      expect(result).toEqual({
        type: "literal",
        value: "3.14",
      });
    });

    it("should parse negative numbers", () => {
      const result = parser.parse("=-42");
      expect(result).toEqual({
        type: "operator",
        value: "-",
        children: [
          {
            type: "literal",
            value: "42",
          },
        ],
      });
    });

    it("should parse strings", () => {
      const result = parser.parse('="Hello World"');
      expect(result).toEqual({
        type: "literal",
        value: "Hello World",
      });
    });

    it("should parse boolean literals", () => {
      const trueResult = parser.parse("=TRUE");
      expect(trueResult).toEqual({
        type: "literal",
        value: "TRUE",
      });

      const falseResult = parser.parse("=FALSE");
      expect(falseResult).toEqual({
        type: "literal",
        value: "FALSE",
      });
    });
  });

  describe("Cell References", () => {
    it("should parse simple cell references", () => {
      const result = parser.parse("=A1");
      expect(result).toEqual({
        type: "reference",
        value: "A1",
      });
    });

    it("should parse absolute cell references", () => {
      const result = parser.parse("=$A$1");
      expect(result).toEqual({
        type: "reference",
        value: "$A$1",
      });
    });

    it("should parse sheet references", () => {
      const result = parser.parse("=Sheet1!A1");
      expect(result).toEqual({
        type: "reference",
        value: "Sheet1!A1",
      });
    });

    it("should parse range references", () => {
      const result = parser.parse("=A1:B10");
      expect(result).toEqual({
        type: "reference",
        value: "A1:B10",
      });
    });

    it("should parse sheet range references", () => {
      const result = parser.parse("=Sheet1!A1:B10");
      expect(result).toEqual({
        type: "reference",
        value: "Sheet1!A1:B10",
      });
    });
  });

  describe("Arithmetic Operations", () => {
    it("should parse addition", () => {
      const result = parser.parse("=1+2");
      expect(result).toEqual({
        type: "operator",
        value: "+",
        children: [
          { type: "literal", value: "1" },
          { type: "literal", value: "2" },
        ],
      });
    });

    it("should parse subtraction", () => {
      const result = parser.parse("=5-3");
      expect(result).toEqual({
        type: "operator",
        value: "-",
        children: [
          { type: "literal", value: "5" },
          { type: "literal", value: "3" },
        ],
      });
    });

    it("should parse multiplication", () => {
      const result = parser.parse("=4*3");
      expect(result).toEqual({
        type: "operator",
        value: "*",
        children: [
          { type: "literal", value: "4" },
          { type: "literal", value: "3" },
        ],
      });
    });

    it("should parse division", () => {
      const result = parser.parse("=8/2");
      expect(result).toEqual({
        type: "operator",
        value: "/",
        children: [
          { type: "literal", value: "8" },
          { type: "literal", value: "2" },
        ],
      });
    });

    it("should parse exponentiation", () => {
      const result = parser.parse("=2^3");
      expect(result).toEqual({
        type: "operator",
        value: "^",
        children: [
          { type: "literal", value: "2" },
          { type: "literal", value: "3" },
        ],
      });
    });

    it("should parse percentage", () => {
      const result = parser.parse("=50%");
      expect(result).toEqual({
        type: "operator",
        value: "%",
        children: [{ type: "literal", value: "50" }],
      });
    });
  });

  describe("Comparison Operations", () => {
    it("should parse equals", () => {
      const result = parser.parse("=A1=B1");
      expect(result).toEqual({
        type: "operator",
        value: "=",
        children: [
          { type: "reference", value: "A1" },
          { type: "reference", value: "B1" },
        ],
      });
    });

    it("should parse not equals", () => {
      const result = parser.parse("=A1<>B1");
      expect(result).toEqual({
        type: "operator",
        value: "<>",
        children: [
          { type: "reference", value: "A1" },
          { type: "reference", value: "B1" },
        ],
      });
    });

    it("should parse less than", () => {
      const result = parser.parse("=A1<B1");
      expect(result).toEqual({
        type: "operator",
        value: "<",
        children: [
          { type: "reference", value: "A1" },
          { type: "reference", value: "B1" },
        ],
      });
    });

    it("should parse greater than or equal", () => {
      const result = parser.parse("=A1>=B1");
      expect(result).toEqual({
        type: "operator",
        value: ">=",
        children: [
          { type: "reference", value: "A1" },
          { type: "reference", value: "B1" },
        ],
      });
    });
  });

  describe("String Operations", () => {
    it("should parse concatenation", () => {
      const result = parser.parse('="Hello"&" "&"World"');
      expect(result.type).toBe("operator");
      expect(result.value).toBe("&");
      expect(result.children).toHaveLength(2);
    });
  });

  describe("Functions", () => {
    it("should parse simple functions", () => {
      const result = parser.parse("=SUM(A1:A10)");
      expect(result).toEqual({
        type: "function",
        value: "SUM",
        children: [{ type: "reference", value: "A1:A10" }],
      });
    });

    it("should parse functions with multiple arguments", () => {
      const result = parser.parse('=IF(A1>0,"Positive","Negative")');
      expect(result).toEqual({
        type: "function",
        value: "IF",
        children: [
          {
            type: "operator",
            value: ">",
            children: [
              { type: "reference", value: "A1" },
              { type: "literal", value: "0" },
            ],
          },
          { type: "literal", value: "Positive" },
          { type: "literal", value: "Negative" },
        ],
      });
    });

    it("should parse nested functions", () => {
      const result = parser.parse("=SUM(A1,MAX(B1:B10))");
      expect(result).toEqual({
        type: "function",
        value: "SUM",
        children: [
          { type: "reference", value: "A1" },
          {
            type: "function",
            value: "MAX",
            children: [{ type: "reference", value: "B1:B10" }],
          },
        ],
      });
    });

    it("should parse functions with no arguments", () => {
      const result = parser.parse("=TODAY()");
      expect(result).toEqual({
        type: "function",
        value: "TODAY",
        children: [],
      });
    });

    it("should parse functions with dots in names", () => {
      // T.INV function
      const tInvResult = parser.parse("=T.INV(0.95,10)");
      expect(tInvResult).toEqual({
        type: "function",
        value: "T.INV",
        children: [
          { type: "literal", value: "0.95" },
          { type: "literal", value: "10" },
        ],
      });

      // T.DIST function
      const tDistResult = parser.parse("=T.DIST(2,10,TRUE)");
      expect(tDistResult).toEqual({
        type: "function",
        value: "T.DIST",
        children: [
          { type: "literal", value: "2" },
          { type: "literal", value: "10" },
          { type: "literal", value: "TRUE" },
        ],
      });

      // NORM.S.INV function
      const normSInvResult = parser.parse("=NORM.S.INV(0.95)");
      expect(normSInvResult).toEqual({
        type: "function",
        value: "NORM.S.INV",
        children: [{ type: "literal", value: "0.95" }],
      });

      // Complex formula with T.INV
      const complexResult = parser.parse("=T.INV(A1,B1)*C1");
      expect(complexResult).toEqual({
        type: "operator",
        value: "*",
        children: [
          {
            type: "function",
            value: "T.INV",
            children: [
              { type: "reference", value: "A1" },
              { type: "reference", value: "B1" },
            ],
          },
          { type: "reference", value: "C1" },
        ],
      });
    });
  });

  describe("Complex Expressions", () => {
    it("should parse expressions with parentheses", () => {
      const result = parser.parse("=(A1+B1)*C1");
      expect(result.type).toBe("operator");
      expect(result.value).toBe("*");
      expect(result.children).toHaveLength(2);
    });

    it("should handle operator precedence", () => {
      const result = parser.parse("=A1+B1*C1");
      expect(result.type).toBe("operator");
      expect(result.value).toBe("+");
      expect(result.children?.[1]).toEqual({
        type: "operator",
        value: "*",
        children: [
          { type: "reference", value: "B1" },
          { type: "reference", value: "C1" },
        ],
      });
    });

    it("should parse array literals", () => {
      const result = parser.parse("={1,2,3;4,5,6}");
      expect(result.type).toBe("function");
      expect(result.value).toBe("ARRAY");
      expect(result.children).toHaveLength(2);
    });
  });

  describe("Complex Sheet References", () => {
    it("should parse sheet reference after minus operator", () => {
      const result = parser.parse("=10-SQCdata!J774");
      expect(result).toEqual({
        type: "operator",
        value: "-",
        children: [
          { type: "literal", value: "10" },
          { type: "reference", value: "SQCdata!J774" },
        ],
      });
    });

    it("should parse sheet references in parentheses with operators", () => {
      const result = parser.parse("=(SQCdata!J774-SQCdata!$AH$2)");
      expect(result).toEqual({
        type: "operator",
        value: "-",
        children: [
          { type: "reference", value: "SQCdata!J774" },
          { type: "reference", value: "SQCdata!$AH$2" },
        ],
      });
    });

    it("should parse complex formula with multiple sheet references", () => {
      const formula =
        '=IF(ISBLANK(SQCdata!J774),"",IF(SQCdata!$AJ$2="Yes",(SQCdata!J774-SQCdata!$AH$2)/SQCdata!$AI$2,SQCdata!J774))';
      const result = parser.parse(formula);

      expect(result.type).toBe("function");
      expect(result.value).toBe("IF");
      expect(result.children).toHaveLength(3);

      // First argument: ISBLANK(SQCdata!J774)
      expect(result.children?.[0].type).toBe("function");
      expect(result.children?.[0].value).toBe("ISBLANK");
      expect(result.children?.[0].children?.[0].value).toBe("SQCdata!J774");

      // Second argument: empty string
      expect(result.children?.[1].type).toBe("literal");
      expect(result.children?.[1].value).toBe("");

      // Third argument: nested IF
      expect(result.children?.[2].type).toBe("function");
      expect(result.children?.[2].value).toBe("IF");
    });

    it("should parse sheet reference with spaces in quoted name", () => {
      const result = parser.parse("='My Sheet'!A1");
      expect(result).toEqual({
        type: "reference",
        value: "'My Sheet'!A1",
      });
    });

    it("should parse sheet reference with special characters in quoted name", () => {
      const result = parser.parse("='John''s Data'!B2");
      expect(result).toEqual({
        type: "reference",
        value: "'John''s Data'!B2",
      });
    });

    it("should handle sheet references in arithmetic expressions", () => {
      const result = parser.parse("=(1-SQCdata!AD2)*S781+SQCdata!AD2*Q782");
      expect(result.type).toBe("operator");
      expect(result.value).toBe("+");

      // Left side: (1-SQCdata!AD2)*S781
      const leftSide = result.children?.[0];
      expect(leftSide?.type).toBe("operator");
      expect(leftSide?.value).toBe("*");

      // Right side: SQCdata!AD2*Q782
      const rightSide = result.children?.[1];
      expect(rightSide?.type).toBe("operator");
      expect(rightSide?.value).toBe("*");
      expect(rightSide?.children?.[0].value).toBe("SQCdata!AD2");
      expect(rightSide?.children?.[1].value).toBe("Q782");
    });
  });

  describe("Error Handling", () => {
    it("should throw error for invalid syntax", () => {
      expect(() => parser.parse("=A1+")).toThrow();
    });

    it("should throw error for missing equals sign", () => {
      expect(() => parser.parse("A1+B1")).toThrow();
    });

    it("should throw error for unmatched parentheses", () => {
      expect(() => parser.parse("=(A1+B1")).toThrow();
    });
  });
});
