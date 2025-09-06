import { ParsedFormula } from '../types/index.js';
import { FormulaParser } from './formula-parser.js';
import { CstNode, IToken } from 'chevrotain';

const parser = new FormulaParser();
const BaseVisitor = parser.getBaseCstVisitorConstructor();

export class FormulaVisitor extends BaseVisitor {
  constructor() {
    super();
    this.validateVisitor();
  }

  formula(ctx: any): ParsedFormula {
    return this.visit(ctx.expression);
  }

  expression(ctx: any): ParsedFormula {
    return this.visit(ctx.logicalExpression);
  }

  logicalExpression(ctx: any): ParsedFormula {
    return this.visit(ctx.comparisonExpression);
  }

  comparisonExpression(ctx: any): ParsedFormula {
    let result = this.visit(ctx.concatenationExpression[0]);
    
    if (ctx.concatenationExpression.length > 1) {
      const operators = [
        ...(ctx.Equals || []),
        ...(ctx.NotEquals || []),
        ...(ctx.LessThan || []),
        ...(ctx.GreaterThan || []),
        ...(ctx.LessThanOrEqual || []),
        ...(ctx.GreaterThanOrEqual || [])
      ].sort((a, b) => a.startOffset - b.startOffset);

      for (let i = 1; i < ctx.concatenationExpression.length; i++) {
        const operator = operators[i - 1];
        const right = this.visit(ctx.concatenationExpression[i]);
        
        result = {
          type: 'operator',
          value: operator.image,
          children: [result, right]
        };
      }
    }
    
    return result;
  }

  concatenationExpression(ctx: any): ParsedFormula {
    let result = this.visit(ctx.additiveExpression[0]);
    
    if (ctx.additiveExpression.length > 1) {
      for (let i = 1; i < ctx.additiveExpression.length; i++) {
        const right = this.visit(ctx.additiveExpression[i]);
        result = {
          type: 'operator',
          value: '&',
          children: [result, right]
        };
      }
    }
    
    return result;
  }

  additiveExpression(ctx: any): ParsedFormula {
    let result = this.visit(ctx.multiplicativeExpression[0]);
    
    if (ctx.multiplicativeExpression.length > 1) {
      const operators = [
        ...(ctx.Plus || []),
        ...(ctx.Minus || [])
      ].sort((a, b) => a.startOffset - b.startOffset);

      for (let i = 1; i < ctx.multiplicativeExpression.length; i++) {
        const operator = operators[i - 1];
        const right = this.visit(ctx.multiplicativeExpression[i]);
        
        result = {
          type: 'operator',
          value: operator.image,
          children: [result, right]
        };
      }
    }
    
    return result;
  }

  multiplicativeExpression(ctx: any): ParsedFormula {
    let result = this.visit(ctx.powerExpression[0]);
    
    if (ctx.powerExpression.length > 1) {
      const operators = [
        ...(ctx.Multiply || []),
        ...(ctx.Divide || [])
      ].sort((a, b) => a.startOffset - b.startOffset);

      for (let i = 1; i < ctx.powerExpression.length; i++) {
        const operator = operators[i - 1];
        const right = this.visit(ctx.powerExpression[i]);
        
        result = {
          type: 'operator',
          value: operator.image,
          children: [result, right]
        };
      }
    }
    
    return result;
  }

  powerExpression(ctx: any): ParsedFormula {
    let result = this.visit(ctx.percentExpression[0]);
    
    if (ctx.percentExpression.length > 1) {
      for (let i = 1; i < ctx.percentExpression.length; i++) {
        const right = this.visit(ctx.percentExpression[i]);
        result = {
          type: 'operator',
          value: '^',
          children: [result, right]
        };
      }
    }
    
    return result;
  }

  percentExpression(ctx: any): ParsedFormula {
    const result = this.visit(ctx.unaryExpression);
    
    if (ctx.Percent) {
      return {
        type: 'operator',
        value: '%',
        children: [result]
      };
    }
    
    return result;
  }

  unaryExpression(ctx: any): ParsedFormula {
    if (ctx.Plus || ctx.Minus) {
      const operator = ctx.Plus ? '+' : '-';
      const operand = this.visit(ctx.unaryExpression);
      
      return {
        type: 'operator',
        value: operator,
        children: [operand]
      };
    }
    
    return this.visit(ctx.primaryExpression);
  }

  primaryExpression(ctx: any): ParsedFormula {
    if (ctx.Number) {
      return {
        type: 'literal',
        value: ctx.Number[0].image
      };
    }
    
    if (ctx.String) {
      const stringValue = ctx.String[0].image;
      return {
        type: 'literal',
        value: stringValue.slice(1, -1).replace(/\\"/g, '"')
      };
    }
    
    if (ctx.True) {
      return {
        type: 'literal',
        value: 'TRUE'
      };
    }
    
    if (ctx.False) {
      return {
        type: 'literal',
        value: 'FALSE'
      };
    }
    
    if (ctx.cellOrRangeReference) {
      return this.visit(ctx.cellOrRangeReference);
    }
    
    if (ctx.functionCall) {
      return this.visit(ctx.functionCall);
    }
    
    if (ctx.arrayLiteral) {
      return this.visit(ctx.arrayLiteral);
    }
    
    if (ctx.expression) {
      return this.visit(ctx.expression);
    }
    
    throw new Error('Unknown primary expression');
  }

  cellOrRangeReference(ctx: any): ParsedFormula {
    let reference = '';
    
    if (ctx.SheetReference) {
      reference = ctx.SheetReference[0].image;
    }
    
    if (ctx.RangeReference) {
      reference += ctx.RangeReference[0].image;
      return {
        type: 'reference',
        value: reference
      };
    }
    
    if (ctx.CellReference) {
      reference += ctx.CellReference[0].image;
      return {
        type: 'reference',
        value: reference
      };
    }
    
    throw new Error('Invalid cell or range reference');
  }

  functionCall(ctx: any): ParsedFormula {
    const functionName = ctx.Function[0].image;
    const args: ParsedFormula[] = ctx.argumentList ? this.visit(ctx.argumentList) : [];
    
    return {
      type: 'function',
      value: functionName,
      children: args
    };
  }

  argumentList(ctx: any): ParsedFormula[] {
    const args: ParsedFormula[] = [];
    
    for (const expr of ctx.expression) {
      args.push(this.visit(expr));
    }
    
    return args;
  }

  arrayLiteral(ctx: any): ParsedFormula {
    const rows: ParsedFormula[] = [];
    
    for (const row of ctx.arrayRow) {
      rows.push(this.visit(row));
    }
    
    return {
      type: 'function',
      value: 'ARRAY',
      children: rows
    };
  }

  arrayRow(ctx: any): ParsedFormula {
    const elements: ParsedFormula[] = [];
    
    for (const expr of ctx.expression) {
      elements.push(this.visit(expr));
    }
    
    return {
      type: 'function',
      value: 'ARRAYROW',
      children: elements
    };
  }
}