import { createToken, Lexer } from "chevrotain";

export const True = createToken({
  name: "True",
  pattern: /TRUE(?![A-Z0-9_])/i,
});
export const False = createToken({
  name: "False",
  pattern: /FALSE(?![A-Z0-9_])/i,
});

export const FunctionToken = createToken({
  name: "Function",
  pattern: /[A-Z][A-Z0-9_]*/i,
  longer_alt: True,
});

export const SheetReference = createToken({
  name: "SheetReference",
  pattern: /[A-Za-z0-9_\- ]+!/,
});

export const CellReference = createToken({
  name: "CellReference",
  pattern: /\$?[A-Z]+\$?[0-9]+/i,
});

export const RangeReference = createToken({
  name: "RangeReference",
  pattern: /\$?[A-Z]+\$?[0-9]+:\$?[A-Z]+\$?[0-9]+|\$?[A-Z]+:\$?[A-Z]+/i,
});

export const NumberToken = createToken({
  name: "Number",
  pattern: /-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?/,
});

export const StringToken = createToken({
  name: "String",
  pattern: /"(?:[^"\\]|\\.)*"/,
});

export const Equals = createToken({ name: "Equals", pattern: /=/ });
export const NotEquals = createToken({ name: "NotEquals", pattern: /<>/ });
export const LessThanOrEqual = createToken({
  name: "LessThanOrEqual",
  pattern: /<=/,
});
export const GreaterThanOrEqual = createToken({
  name: "GreaterThanOrEqual",
  pattern: />=/,
});
export const LessThan = createToken({ name: "LessThan", pattern: /</ });
export const GreaterThan = createToken({ name: "GreaterThan", pattern: />/ });

export const Plus = createToken({ name: "Plus", pattern: /\+/ });
export const Minus = createToken({ name: "Minus", pattern: /-/ });
export const Multiply = createToken({ name: "Multiply", pattern: /\*/ });
export const Divide = createToken({ name: "Divide", pattern: /\// });
export const Power = createToken({ name: "Power", pattern: /\^/ });
export const Percent = createToken({ name: "Percent", pattern: /%/ });

export const Ampersand = createToken({ name: "Ampersand", pattern: /&/ });
export const Comma = createToken({ name: "Comma", pattern: /,/ });
export const Semicolon = createToken({ name: "Semicolon", pattern: /;/ });
export const Colon = createToken({ name: "Colon", pattern: /:/ });

export const LeftParen = createToken({ name: "LeftParen", pattern: /\(/ });
export const RightParen = createToken({ name: "RightParen", pattern: /\)/ });
export const LeftBracket = createToken({ name: "LeftBracket", pattern: /\[/ });
export const RightBracket = createToken({
  name: "RightBracket",
  pattern: /\]/,
});
export const LeftBrace = createToken({ name: "LeftBrace", pattern: /\{/ });
export const RightBrace = createToken({ name: "RightBrace", pattern: /\}/ });

export const WhiteSpace = createToken({
  name: "WhiteSpace",
  pattern: /\s+/,
  group: Lexer.SKIPPED,
});

export const allTokens = [
  WhiteSpace,
  True,
  False,
  NotEquals,
  LessThanOrEqual,
  GreaterThanOrEqual,
  LessThan,
  GreaterThan,
  Equals,
  Plus,
  Minus,
  Multiply,
  Divide,
  Power,
  Percent,
  Ampersand,
  Comma,
  Semicolon,
  Colon,
  LeftParen,
  RightParen,
  LeftBracket,
  RightBracket,
  LeftBrace,
  RightBrace,
  SheetReference,
  RangeReference,
  CellReference,
  FunctionToken,
  NumberToken,
  StringToken,
];

export const FormulaLexer = new Lexer(allTokens);
