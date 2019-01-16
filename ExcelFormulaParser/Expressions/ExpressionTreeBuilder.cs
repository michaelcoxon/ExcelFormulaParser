// https://github.com/psalaets/excel-formula-ast/blob/master/lib/build-tree.js

using ExcelFormulaParser.FormulaTokenizer;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ExcelFormulaParser.Expressions
{
    public class ExpressionTreeBuilder
    {
        private static readonly Dictionary<string, byte> _precendence = new Dictionary<string, byte>
        {
            // cell range union and intersect
            [" "] = 8,
            [","] = 8,

            // raise to power
            ["^"] = 5,

            // multiply, divide
            ["*"] = 4,
            ["/"] = 4,

            // add, subtract
            ["+"] = 3,
            ["-"] = 3,

            // string concat
            ["&"] = 2,

            // comparison
            ["="] = 1,
            ["<>"] = 1,
            ["<="] = 1,
            [">="] = 1,
            [">"] = 1,
            ["<"] = 1
        };

        public static Expression ParseFormula(Token[] tokens)
        {
            var stream = new TokenStream(tokens);
            var shuntingYard = new ShuntingYard.ShuntingYard();

            ParseExpression(stream, shuntingYard);

            var retVal = shuntingYard.operands.Peek();
            if (retVal == null)
            {
                throw new Exception("Syntax error");
            }
            return retVal;
        }

        private static void ParseExpression(TokenStream stream, ShuntingYard.ShuntingYard shuntingYard)
        {
            ParseOperandExpression(stream, shuntingYard);

            var pos = 0;
            while (true)
            {
                if (!stream.NextIsBinaryOperator())
                {
                    break;
                }
                if (pos == stream.Pos())
                {
                    throw new Exception("Invalid syntax!");
                }
                pos = stream.Pos();
                PushOperator(CreateBinaryOperator(stream.GetNext().Value), shuntingYard);
                stream.Consume();
                ParseOperandExpression(stream, shuntingYard);
            }

            while (shuntingYard.operators.Peek() != ShuntingYard.Operator.SENTINEL)
            {
                PopOperator(shuntingYard.operators, shuntingYard.operands);
            }
        }

        private static void ParseOperandExpression(TokenStream stream, ShuntingYard.ShuntingYard shuntingYard)
        {
            if (stream.NextIsTerminal())
            {
                shuntingYard.operands.Push(ParseTerminal(stream));
                // parseTerminal already consumes once so don"t need to consume on line below
                // stream.consume()
            }
            else if (stream.NextIsOpenParen())
            {
                stream.Consume(); // open paren
                WithinSentinel(shuntingYard, () => ParseExpression(stream, shuntingYard));
                stream.Consume(); // close paren
            }
            else if (stream.NextIsPrefixOperator())
            {
                var unaryOperator = CreateUnaryOperator(stream.GetNext().Value);
                PushOperator(unaryOperator, shuntingYard);
                stream.Consume();
                ParseOperandExpression(stream, shuntingYard);
            }
            else if (stream.NextIsFunctionCall())
            {
                ParseFunctionCall(stream, shuntingYard);
            }
        }

        private static void ParseFunctionCall(TokenStream stream, ShuntingYard.ShuntingYard shuntingYard)
        {
            var name = stream.GetNext().Value;
            stream.Consume(); // consume start of function call

            var args = ParseFunctionArgList(stream, shuntingYard);
            shuntingYard.operands.Push(Expression.FunctionCall(name, args));

            stream.Consume(); // consume end of function call
        }

        private static Expression[] ParseFunctionArgList(TokenStream stream, ShuntingYard.ShuntingYard shuntingYard)
        {
            var reverseArgs = new List<Expression>();

            WithinSentinel(shuntingYard, () =>
            {
                var arity = 0;
                var pos = 0;
                while (true)
                {
                    if (stream.NextIsEndOfFunctionCall())
                    {
                        break;
                    }

                    if (pos == stream.Pos())
                    {
                        throw new Exception("Invalid syntax");
                    }
                    pos = stream.Pos();
                    ParseExpression(stream, shuntingYard);
                    arity += 1;

                    if (stream.NextIsFunctionArgumentSeparator())
                    {
                        stream.Consume();
                    }
                }

                for (var i = 0; i < arity; i++)
                {
                    reverseArgs.Add(shuntingYard.operands.Pop());
                }
            });

            reverseArgs.Reverse();

            return reverseArgs.ToArray();
        }

        private static void WithinSentinel(ShuntingYard.ShuntingYard shuntingYard, Action fn)
        {
            shuntingYard.operators.Push(ShuntingYard.Operator.SENTINEL);
            fn();
            shuntingYard.operators.Pop();
        }

        private static void PushOperator(ShuntingYard.Operator @operator, ShuntingYard.ShuntingYard shuntingYard)
        {
            while (shuntingYard.operators.Peek().EvaluatesBefore(@operator))
            {
                PopOperator(shuntingYard.operators, shuntingYard.operands);
            }
            shuntingYard.operators.Push(@operator);
        }

        private static void PopOperator(Stack<ShuntingYard.Operator> operators, Stack<Expression> operands)
        {
            if (operators.Peek().IsBinary())
            {
                var right = operands.Pop();
                var left = operands.Pop();

                var @operator = operators.Pop();
                operands.Push(Expression.BinaryExpression(BinaryOperatorTypeConverter(@operator.symbol), left, right));
            }
            else if (operators.Peek().IsUnary())
            {
                var operand = operands.Pop();

                var @operator = operators.Pop();
                operands.Push(Expression.UnaryExpression(UnaryOperatorTypeConverter(@operator.symbol), operand));
            }
        }

        private static Expression ParseTerminal(TokenStream stream)
        {
            if (stream.NextIsNumber())
            {
                return ParseNumber(stream);
            }

            if (stream.NextIsText())
            {
                return ParseText(stream);
            }

            if (stream.NextIsLogical())
            {
                return ParseLogical(stream);
            }

            if (stream.NextIsRange())
            {
                return ParseRange(stream);
            }

            return null;
        }

        private static Expression ParseRange(TokenStream stream)
        {
            var next = stream.GetNext();
            stream.Consume();
            return CreateCellRange(next.Value);
        }

        private static Expression CreateCellRange(string value)
        {
            var parts = value.Split(":");

            if (parts.Length == 2)
            {
                return Expression.CellRange(
                  Expression.Cell(parts[0], CellRefType(parts[0])),
                  Expression.Cell(parts[1], CellRefType(parts[1]))
                );
            }
            else
            {
                return Expression.Cell(value, CellRefType(value));
            }
        }

        private static CellReferenceType CellRefType(string key)
        {
            if (new Regex(@"^\$[A-Z]+\$\d+$").IsMatch(key))
            {
                return CellReferenceType.Absolute;
            }

            if (new Regex(@"^\$[A-Z]+$").IsMatch(key))
            {
                return CellReferenceType.Absolute;
            }

            if (new Regex(@"^\$\d+$").IsMatch(key))
            {
                return CellReferenceType.Absolute;
            }

            if (new Regex(@"^\$[A-Z]+\d+$").IsMatch(key))
            {
                return CellReferenceType.Mixed;
            }

            if (new Regex(@"^[A-Z]+\$\d+$").IsMatch(key))
            {
                return CellReferenceType.Mixed;
            }

            if (new Regex(@"^[A-Z]+\d+$").IsMatch(key))
            {
                return CellReferenceType.Relative;
            }

            if (new Regex(@"^\d+$").IsMatch(key))
            {
                return CellReferenceType.Relative;
            }

            if (new Regex(@"^[A-Z]+$").IsMatch(key))
            {
                return CellReferenceType.Relative;
            }

            throw new NotSupportedException($"Cell with key '{key}' has an unsupported type");
        }

        private static TextExpression ParseText(TokenStream stream)
        {
            var next = stream.GetNext();
            stream.Consume();
            return Expression.Text(next.Value);
        }

        private static LogicalExpression ParseLogical(TokenStream stream)
        {
            var next = stream.GetNext();
            stream.Consume();
            return Expression.Logical(next.Value == "TRUE");
        }

        private static NumberExpression ParseNumber(TokenStream stream)
        {
            var value = double.Parse(stream.GetNext().Value);
            stream.Consume();

            if (stream.NextIsPostfixOperator())
            {
                value *= 0.01;
                stream.Consume();
            }

            return Expression.Number(value);
        }

        private static ShuntingYard.Operator CreateUnaryOperator(string symbol)
        {
            var precendence = new Dictionary<string, int>
            {
                // negation
                ["-"] = 7
            }
            [symbol];

            return new ShuntingYard.Operator(symbol, precendence, 1, true);
        }

        private static ShuntingYard.Operator CreateBinaryOperator(string symbol)
        {
            var precendence = _precendence[symbol];

            return new ShuntingYard.Operator(symbol, precendence, 2, true);
        }

        private static UnaryOperatorType UnaryOperatorTypeConverter(string symbol)
        {
            switch (symbol) {
                case "-":
                    return UnaryOperatorType.Negate;
                default:
                    throw new NotSupportedException($"Unary operator '{symbol}' is not supported");
            }
        }

        private static BinaryOperatorType BinaryOperatorTypeConverter(string symbol)
        {
            switch (symbol)
            {
                case "+":
                    return BinaryOperatorType.Add;
                case "&":
                    return BinaryOperatorType.Concat;
                case "<>":
                    return BinaryOperatorType.NotEqualTo;
                case ",":
                    return BinaryOperatorType.Comma;
                case " ":
                    return BinaryOperatorType.Whitespace;
                case "/":
                    return BinaryOperatorType.Divide;
                default:
                    throw new NotSupportedException($"Binary operator '{symbol}' is not supported");
            }
        }
    }
}
