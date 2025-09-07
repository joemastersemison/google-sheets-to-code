import type { CstNode, IToken } from "chevrotain";

// Define context types for each rule in the parser
export interface FormulaContext {
  Equals: IToken[];
  expression: CstNode[];
}

export interface ExpressionContext {
  logicalExpression: CstNode[];
}

export interface LogicalExpressionContext {
  comparisonExpression: CstNode[];
}

export interface ComparisonExpressionContext {
  concatenationExpression: CstNode[];
  Equals?: IToken[];
  NotEquals?: IToken[];
  LessThan?: IToken[];
  GreaterThan?: IToken[];
  LessThanOrEqual?: IToken[];
  GreaterThanOrEqual?: IToken[];
}

export interface ConcatenationExpressionContext {
  additiveExpression: CstNode[];
  Ampersand?: IToken[];
}

export interface AdditiveExpressionContext {
  multiplicativeExpression: CstNode[];
  Plus?: IToken[];
  Minus?: IToken[];
}

export interface MultiplicativeExpressionContext {
  powerExpression: CstNode[];
  Multiply?: IToken[];
  Divide?: IToken[];
}

export interface PowerExpressionContext {
  percentExpression: CstNode[];
  Power?: IToken[];
}

export interface PercentExpressionContext {
  unaryExpression: CstNode[];
  Percent?: IToken[];
}

export interface UnaryExpressionContext {
  Plus?: IToken[];
  Minus?: IToken[];
  unaryExpression?: CstNode[];
  primaryExpression?: CstNode[];
}

export interface PrimaryExpressionContext {
  Number?: IToken[];
  String?: IToken[];
  True?: IToken[];
  False?: IToken[];
  cellOrRangeReference?: CstNode[];
  functionCall?: CstNode[];
  arrayLiteral?: CstNode[];
  LeftParen?: IToken[];
  expression?: CstNode[];
  RightParen?: IToken[];
}

export interface CellOrRangeReferenceContext {
  SheetReference?: IToken[];
  RangeReference?: IToken[];
  CellReference?: IToken[];
}

export interface FunctionCallContext {
  Function: IToken[];
  LeftParen: IToken[];
  argumentList?: CstNode[];
  RightParen: IToken[];
}

export interface ArgumentListContext {
  expression: CstNode[];
  Comma?: IToken[];
  Semicolon?: IToken[];
}

export interface ArrayLiteralContext {
  LeftBrace: IToken[];
  arrayRow: CstNode[];
  Semicolon?: IToken[];
  RightBrace: IToken[];
}

export interface ArrayRowContext {
  expression: CstNode[];
  Comma?: IToken[];
}
