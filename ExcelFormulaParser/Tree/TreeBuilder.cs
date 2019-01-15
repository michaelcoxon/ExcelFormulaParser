// https://github.com/psalaets/excel-formula-ast/blob/master/lib/build-tree.js

using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ExcelFormulaParser.Tree
{
    public class TreeBuilder
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

        public Token ParseFormula(Token[] tokens)
        {
            var stream = new TokenStream(tokens);
            var shuntingYard = new ShuntingYard.ShuntingYard();

            this.ParseExpression(stream, shuntingYard);

            var retVal = shuntingYard.operands.Peek();
            if (retVal != null)
            {
                throw new Exception("Syntax error");
            }
            return retVal;
        }

        public void ParseExpression(TokenStream stream, ShuntingYard.ShuntingYard shuntingYard)
        {
            this.ParseOperandExpression(stream, shuntingYard);

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
                this.PushOperator(this.CreateBinaryOperator((string)stream.GetNext().Value), shuntingYard);
                stream.Consume();
                this.ParseOperandExpression(stream, shuntingYard);
            }

            while (shuntingYard.operators.Peek() != ShuntingYard.Operator.SENTINEL)
            {
                this.PopOperator(shuntingYard.operators, shuntingYard.operands);
            }
        }

        public void ParseOperandExpression(TokenStream stream, ShuntingYard.ShuntingYard shuntingYard)
        {
            if (stream.NextIsTerminal())
            {
                shuntingYard.operands.Push(this.ParseTerminal(stream));
                // parseTerminal already consumes once so don"t need to consume on line below
                // stream.consume()
            }
            else if (stream.NextIsOpenParen())
            {
                stream.Consume(); // open paren
                this.WithinSentinel(shuntingYard, () => this.ParseExpression(stream, shuntingYard));
                stream.Consume(); // close paren
            }
            else if (stream.NextIsPrefixOperator())
            {
                var unaryOperator = this.CreateUnaryOperator((string)stream.GetNext().Value);
                this.PushOperator(unaryOperator, shuntingYard);
                stream.Consume();
                this.ParseOperandExpression(stream, shuntingYard);
            }
            else if (stream.NextIsFunctionCall())
            {
                this.ParseFunctionCall(stream, shuntingYard);
            }
        }

        public void ParseFunctionCall(TokenStream stream, ShuntingYard.ShuntingYard shuntingYard)
        {
            var name = (string)stream.GetNext().Value;
            stream.Consume(); // consume start of function call

            var args = this.ParseFunctionArgList(stream, shuntingYard);
            shuntingYard.operands.Push(NodeBuilder.FunctionCall(name, args));

            stream.Consume(); // consume end of function call
        }

        public Token[] ParseFunctionArgList(TokenStream stream, ShuntingYard.ShuntingYard shuntingYard)
        {
            var reverseArgs = new List<Token>();

            this.WithinSentinel(shuntingYard, () =>
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
                    this.ParseExpression(stream, shuntingYard);
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

        public void WithinSentinel(ShuntingYard.ShuntingYard shuntingYard, Action fn)
        {
            shuntingYard.operators.Push(ShuntingYard.Operator.SENTINEL);
            fn();
            shuntingYard.operators.Pop();
        }

        public void PushOperator(ShuntingYard.Operator @operator, ShuntingYard.ShuntingYard shuntingYard)
        {
            while (shuntingYard.operators.Peek().EvaluatesBefore(@operator))
            {
                this.PopOperator(shuntingYard.operators, shuntingYard.operands);
            }
            shuntingYard.operators.Push(@operator);
        }

        public void PopOperator(Stack<ShuntingYard.Operator> operators, Stack<Token> operands)
        {
            if (operators.Peek().IsBinary())
            {
                var right = operands.Pop();
                var left = operands.Pop();

                var @operator = operators.Pop();
                operands.Push(NodeBuilder.BinaryExpression(@operator.symbol, left, right));
            }
            else if (operators.Peek().IsUnary())
            {
                var operand = operands.Pop();

                var @operator = operators.Pop();
                operands.Push(NodeBuilder.UnaryExpression(@operator.symbol, operand));
            }
        }

        public Token ParseTerminal(TokenStream stream)
        {
            if (stream.NextIsNumber())
            {
                return this.ParseNumber(stream);
            }

            if (stream.NextIsText())
            {
                return this.ParseText(stream);
            }

            if (stream.NextIsLogical())
            {
                return this.ParseLogical(stream);
            }

            if (stream.NextIsRange())
            {
                return this.ParseRange(stream);
            }

            return null;
        }

        public Token ParseRange(TokenStream stream)
        {
            var next = stream.GetNext();
            stream.Consume();
            return this.CreateCellRange((string)next.Value);
        }

        public Token CreateCellRange(string value)
        {
            var parts = value.Split(":");

            if (parts.Length == 2)
            {
                return NodeBuilder.CellRange(
                  NodeBuilder.Cell(parts[0], this.CellRefType(parts[0])),
                  NodeBuilder.Cell(parts[1], this.CellRefType(parts[1]))
                );
            }
            else
            {
                return NodeBuilder.Cell(value, this.CellRefType(value));
            }
        }

        public string CellRefType(string key)
        {
            if (new Regex(@"^{\$}[A-Z]+\$\d+$").IsMatch(key))
            {
                return "absolute";
            }

            if (new Regex(@"^\$[A-Z]+$").IsMatch(key))
            {
                return "absolute";
            }

            if (new Regex(@"^\$\d+$").IsMatch(key))
            {
                return "absolute";
            }

            if (new Regex(@"^\$[A-Z]+\d+$").IsMatch(key))
            {
                return "mixed";
            }

            if (new Regex(@"^[A-Z]+\$\d+$").IsMatch(key))
            {
                return "mixed";
            }

            if (new Regex(@"^[A-Z]+\d+$").IsMatch(key))
            {
                return "relative";
            }

            if (new Regex(@"^\d+$").IsMatch(key))
            {
                return "relative";
            }

            if (new Regex(@"^[A-Z]+$").IsMatch(key))
            {
                return "relative";
            }

            return null;
        }

        public Token ParseText(TokenStream stream)
        {
            var next = stream.GetNext();
            stream.Consume();
            return NodeBuilder.Text((string)next.Value);
        }

        public Token ParseLogical(TokenStream stream)
        {
            var next = stream.GetNext();
            stream.Consume();
            return NodeBuilder.Logical((string)next.Value == "TRUE");
        }

        public Token ParseNumber(TokenStream stream)
        {
            var value = (double)stream.GetNext().Value;
            stream.Consume();

            if (stream.NextIsPostfixOperator())
            {
                value *= 0.01;
                stream.Consume();
            }

            return NodeBuilder.Number(value);
        }

        public ShuntingYard.Operator CreateUnaryOperator(string symbol)
        {
            var precendence = new Dictionary<string, int>
            {
                // negation
                ["-"] = 7
            }
            [symbol];

            return new ShuntingYard.Operator(symbol, precendence, 1, true);
        }

        public ShuntingYard.Operator CreateBinaryOperator(string symbol)
        {
            var precendence = _precendence[symbol];

            return new ShuntingYard.Operator(symbol, precendence, 2, true);
        }
    }
}
