import type { CstParser } from "chevrotain";
import type { ParsedFormula } from "../types/index.js";
import type {
  AdditiveExpressionContext,
  ArgumentListContext,
  ArrayLiteralContext,
  ArrayRowContext,
  CellOrRangeReferenceContext,
  ComparisonExpressionContext,
  ConcatenationExpressionContext,
  ExpressionContext,
  FormulaContext,
  FunctionCallContext,
  LogicalExpressionContext,
  MultiplicativeExpressionContext,
  PercentExpressionContext,
  PowerExpressionContext,
  PrimaryExpressionContext,
  UnaryExpressionContext,
} from "./visitor-types.js";

export function createFormulaVisitor(parser: CstParser) {
  const BaseVisitor = parser.getBaseCstVisitorConstructor();

  class FormulaVisitorImpl extends BaseVisitor {
    constructor() {
      super();
      this.validateVisitor();
    }

    formula(ctx: FormulaContext): ParsedFormula {
      return this.visit(ctx.expression);
    }

    expression(ctx: ExpressionContext): ParsedFormula {
      return this.visit(ctx.logicalExpression);
    }

    logicalExpression(ctx: LogicalExpressionContext): ParsedFormula {
      return this.visit(ctx.comparisonExpression);
    }

    comparisonExpression(ctx: ComparisonExpressionContext): ParsedFormula {
      let result = this.visit(ctx.concatenationExpression[0]);

      if (ctx.concatenationExpression.length > 1) {
        const operators = [
          ...(ctx.Equals || []),
          ...(ctx.NotEquals || []),
          ...(ctx.LessThan || []),
          ...(ctx.GreaterThan || []),
          ...(ctx.LessThanOrEqual || []),
          ...(ctx.GreaterThanOrEqual || []),
        ].sort((a, b) => a.startOffset - b.startOffset);

        for (let i = 1; i < ctx.concatenationExpression.length; i++) {
          const operator = operators[i - 1];
          const right = this.visit(ctx.concatenationExpression[i]);

          result = {
            type: "operator",
            value: operator.image,
            children: [result, right],
          };
        }
      }

      return result;
    }

    concatenationExpression(
      ctx: ConcatenationExpressionContext
    ): ParsedFormula {
      let result = this.visit(ctx.additiveExpression[0]);

      if (ctx.additiveExpression.length > 1) {
        for (let i = 1; i < ctx.additiveExpression.length; i++) {
          const right = this.visit(ctx.additiveExpression[i]);
          result = {
            type: "operator",
            value: "&",
            children: [result, right],
          };
        }
      }

      return result;
    }

    additiveExpression(ctx: AdditiveExpressionContext): ParsedFormula {
      let result = this.visit(ctx.multiplicativeExpression[0]);

      if (ctx.multiplicativeExpression.length > 1) {
        const operators = [...(ctx.Plus || []), ...(ctx.Minus || [])].sort(
          (a, b) => a.startOffset - b.startOffset
        );

        for (let i = 1; i < ctx.multiplicativeExpression.length; i++) {
          const operator = operators[i - 1];
          const right = this.visit(ctx.multiplicativeExpression[i]);

          result = {
            type: "operator",
            value: operator.image,
            children: [result, right],
          };
        }
      }

      return result;
    }

    multiplicativeExpression(
      ctx: MultiplicativeExpressionContext
    ): ParsedFormula {
      let result = this.visit(ctx.powerExpression[0]);

      if (ctx.powerExpression.length > 1) {
        const operators = [...(ctx.Multiply || []), ...(ctx.Divide || [])].sort(
          (a, b) => a.startOffset - b.startOffset
        );

        for (let i = 1; i < ctx.powerExpression.length; i++) {
          const operator = operators[i - 1];
          const right = this.visit(ctx.powerExpression[i]);

          result = {
            type: "operator",
            value: operator.image,
            children: [result, right],
          };
        }
      }

      return result;
    }

    powerExpression(ctx: PowerExpressionContext): ParsedFormula {
      let result = this.visit(ctx.percentExpression[0]);

      if (ctx.percentExpression.length > 1) {
        for (let i = 1; i < ctx.percentExpression.length; i++) {
          const right = this.visit(ctx.percentExpression[i]);
          result = {
            type: "operator",
            value: "^",
            children: [result, right],
          };
        }
      }

      return result;
    }

    percentExpression(ctx: PercentExpressionContext): ParsedFormula {
      const result = this.visit(ctx.unaryExpression);

      if (ctx.Percent) {
        return {
          type: "operator",
          value: "%",
          children: [result],
        };
      }

      return result;
    }

    unaryExpression(ctx: UnaryExpressionContext): ParsedFormula {
      if (ctx.Plus || ctx.Minus) {
        const operator = ctx.Plus ? "+" : "-";
        if (!ctx.unaryExpression) {
          throw new Error("Expected unary expression");
        }
        const operand = this.visit(ctx.unaryExpression);

        return {
          type: "operator",
          value: operator,
          children: [operand],
        };
      }

      if (!ctx.primaryExpression) {
        throw new Error("Expected primary expression");
      }
      return this.visit(ctx.primaryExpression);
    }

    primaryExpression(ctx: PrimaryExpressionContext): ParsedFormula {
      // Check for consumed tokens (they are stored as arrays in the CST)
      // Token names in the CST may be different from the token class names
      if (ctx.Number && ctx.Number.length > 0) {
        return {
          type: "literal",
          value: ctx.Number[0].image,
        };
      }

      if (ctx.String && ctx.String.length > 0) {
        const stringValue = ctx.String[0].image;
        return {
          type: "literal",
          value: stringValue.slice(1, -1).replace(/\\"/g, '"'),
        };
      }

      if (ctx.True && ctx.True.length > 0) {
        return {
          type: "literal",
          value: "TRUE",
        };
      }

      if (ctx.False && ctx.False.length > 0) {
        return {
          type: "literal",
          value: "FALSE",
        };
      }

      // Check for subrules (they are stored as arrays of CST nodes)
      if (ctx.cellOrRangeReference && ctx.cellOrRangeReference.length > 0) {
        return this.visit(ctx.cellOrRangeReference[0]);
      }

      if (ctx.functionCall && ctx.functionCall.length > 0) {
        return this.visit(ctx.functionCall[0]);
      }

      if (ctx.arrayLiteral && ctx.arrayLiteral.length > 0) {
        return this.visit(ctx.arrayLiteral[0]);
      }

      // For parenthesized expressions
      if (ctx.expression && ctx.expression.length > 0) {
        return this.visit(ctx.expression[0]);
      }

      // Better error message with context info
      const availableKeys = Object.keys(ctx).filter((k) => {
        const value = (ctx as Record<string, unknown>)[k];
        return value && (Array.isArray(value) ? value.length > 0 : true);
      });
      throw new Error(
        `Unknown primary expression. Available context keys: ${availableKeys.join(", ")}`
      );
    }

    cellOrRangeReference(ctx: CellOrRangeReferenceContext): ParsedFormula {
      let reference = "";

      if (ctx.SheetReference) {
        reference = ctx.SheetReference[0].image;
      }

      if (ctx.RangeReference) {
        reference += ctx.RangeReference[0].image;
        return {
          type: "reference",
          value: reference,
        };
      }

      if (ctx.CellReference) {
        reference += ctx.CellReference[0].image;
        return {
          type: "reference",
          value: reference,
        };
      }

      throw new Error("Invalid cell or range reference");
    }

    functionCall(ctx: FunctionCallContext): ParsedFormula {
      const functionName = ctx.Function[0].image;
      const args: ParsedFormula[] =
        ctx.argumentList && ctx.argumentList.length > 0
          ? this.visit(ctx.argumentList[0])
          : [];

      return {
        type: "function",
        value: functionName,
        children: args,
      };
    }

    argumentList(ctx: ArgumentListContext): ParsedFormula[] {
      const args: ParsedFormula[] = [];

      for (const expr of ctx.expression) {
        args.push(this.visit(expr));
      }

      return args;
    }

    arrayLiteral(ctx: ArrayLiteralContext): ParsedFormula {
      const rows: ParsedFormula[] = [];

      for (const row of ctx.arrayRow) {
        rows.push(this.visit(row));
      }

      return {
        type: "function",
        value: "ARRAY",
        children: rows,
      };
    }

    arrayRow(ctx: ArrayRowContext): ParsedFormula {
      const elements: ParsedFormula[] = [];

      for (const expr of ctx.expression) {
        elements.push(this.visit(expr));
      }

      return {
        type: "function",
        value: "ARRAYROW",
        children: elements,
      };
    }
  }

  return new FormulaVisitorImpl();
}

// Export a type for the visitor
export type FormulaVisitor = ReturnType<typeof createFormulaVisitor>;
