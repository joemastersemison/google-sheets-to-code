import { CstParser } from "chevrotain";
import type { ParsedFormula } from "../types/index.js";
import {
  createFormulaVisitor,
  type FormulaVisitor,
} from "./formula-visitor.js";
import * as tokens from "./lexer.js";
import { FormulaLexer } from "./lexer.js";

export class FormulaParser extends CstParser {
  private visitor: FormulaVisitor;

  constructor() {
    super(tokens.allTokens);
    this.performSelfAnalysis();
    this.visitor = createFormulaVisitor(this);
  }

  public formula = this.RULE("formula", () => {
    this.CONSUME(tokens.Equals);
    this.SUBRULE(this.expression);
  });

  private expression = this.RULE("expression", () => {
    this.SUBRULE(this.logicalExpression);
  });

  private logicalExpression = this.RULE("logicalExpression", () => {
    this.SUBRULE(this.comparisonExpression);
  });

  private comparisonExpression = this.RULE("comparisonExpression", () => {
    this.SUBRULE(this.concatenationExpression);
    this.MANY(() => {
      this.OR([
        { ALT: () => this.CONSUME(tokens.Equals) },
        { ALT: () => this.CONSUME(tokens.NotEquals) },
        { ALT: () => this.CONSUME(tokens.LessThan) },
        { ALT: () => this.CONSUME(tokens.GreaterThan) },
        { ALT: () => this.CONSUME(tokens.LessThanOrEqual) },
        { ALT: () => this.CONSUME(tokens.GreaterThanOrEqual) },
      ]);
      this.SUBRULE2(this.concatenationExpression);
    });
  });

  private concatenationExpression = this.RULE("concatenationExpression", () => {
    this.SUBRULE(this.additiveExpression);
    this.MANY(() => {
      this.CONSUME(tokens.Ampersand);
      this.SUBRULE2(this.additiveExpression);
    });
  });

  private additiveExpression = this.RULE("additiveExpression", () => {
    this.SUBRULE(this.multiplicativeExpression);
    this.MANY(() => {
      this.OR([
        { ALT: () => this.CONSUME(tokens.Plus) },
        { ALT: () => this.CONSUME(tokens.Minus) },
      ]);
      this.SUBRULE2(this.multiplicativeExpression);
    });
  });

  private multiplicativeExpression = this.RULE(
    "multiplicativeExpression",
    () => {
      this.SUBRULE(this.powerExpression);
      this.MANY(() => {
        this.OR([
          { ALT: () => this.CONSUME(tokens.Multiply) },
          { ALT: () => this.CONSUME(tokens.Divide) },
        ]);
        this.SUBRULE2(this.powerExpression);
      });
    }
  );

  private powerExpression = this.RULE("powerExpression", () => {
    this.SUBRULE(this.percentExpression);
    this.MANY(() => {
      this.CONSUME(tokens.Power);
      this.SUBRULE2(this.percentExpression);
    });
  });

  private percentExpression = this.RULE("percentExpression", () => {
    this.SUBRULE(this.unaryExpression);
    this.OPTION(() => {
      this.CONSUME(tokens.Percent);
    });
  });

  private unaryExpression = this.RULE("unaryExpression", () => {
    this.OR([
      {
        ALT: () => {
          this.OR2([
            { ALT: () => this.CONSUME(tokens.Plus) },
            { ALT: () => this.CONSUME(tokens.Minus) },
          ]);
          this.SUBRULE(this.unaryExpression);
        },
      },
      { ALT: () => this.SUBRULE(this.primaryExpression) },
    ]);
  });

  private primaryExpression = this.RULE("primaryExpression", () => {
    this.OR([
      { ALT: () => this.CONSUME(tokens.NumberToken) },
      { ALT: () => this.CONSUME(tokens.StringToken) },
      { ALT: () => this.CONSUME(tokens.True) },
      { ALT: () => this.CONSUME(tokens.False) },
      { ALT: () => this.SUBRULE(this.cellOrRangeReference) },
      { ALT: () => this.SUBRULE(this.functionCall) },
      { ALT: () => this.SUBRULE(this.arrayLiteral) },
      {
        ALT: () => {
          this.CONSUME(tokens.LeftParen);
          this.SUBRULE(this.expression);
          this.CONSUME(tokens.RightParen);
        },
      },
    ]);
  });

  private cellOrRangeReference = this.RULE("cellOrRangeReference", () => {
    this.OPTION(() => {
      this.CONSUME(tokens.SheetReference);
    });
    this.OR([
      { ALT: () => this.CONSUME(tokens.RangeReference) },
      { ALT: () => this.CONSUME(tokens.CellReference) },
    ]);
  });

  private functionCall = this.RULE("functionCall", () => {
    this.CONSUME(tokens.FunctionToken);
    this.CONSUME(tokens.LeftParen);
    this.OPTION(() => {
      this.SUBRULE(this.argumentList);
    });
    this.CONSUME(tokens.RightParen);
  });

  private argumentList = this.RULE("argumentList", () => {
    this.SUBRULE(this.expression);
    this.MANY(() => {
      this.OR([
        { ALT: () => this.CONSUME(tokens.Comma) },
        { ALT: () => this.CONSUME(tokens.Semicolon) },
      ]);
      this.SUBRULE2(this.expression);
    });
  });

  private arrayLiteral = this.RULE("arrayLiteral", () => {
    this.CONSUME(tokens.LeftBrace);
    this.SUBRULE(this.arrayRow);
    this.MANY(() => {
      this.CONSUME(tokens.Semicolon);
      this.SUBRULE2(this.arrayRow);
    });
    this.CONSUME(tokens.RightBrace);
  });

  private arrayRow = this.RULE("arrayRow", () => {
    this.SUBRULE(this.expression);
    this.MANY(() => {
      this.CONSUME(tokens.Comma);
      this.SUBRULE2(this.expression);
    });
  });

  parse(formula: string): ParsedFormula {
    const lexingResult = FormulaLexer.tokenize(formula);

    if (lexingResult.errors.length > 0) {
      throw new Error(
        `Lexing errors: ${lexingResult.errors.map((e) => e.message).join(", ")}`
      );
    }

    this.input = lexingResult.tokens;
    const cst = this.formula();

    if (this.errors.length > 0) {
      throw new Error(
        `Parsing errors: ${this.errors.map((e) => e.message).join(", ")}`
      );
    }

    return this.visitor.visit(cst);
  }
}
