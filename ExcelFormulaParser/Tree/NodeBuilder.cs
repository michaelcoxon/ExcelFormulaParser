// https://github.com/psalaets/excel-formula-ast/blob/master/lib/node-builder.js
using System;

namespace ExcelFormulaParser.Tree
{
    public static class NodeBuilder
    {
        public static Token Cell(string key, string refType)
        {
            return new Token
            {
                Type = TokenType.Cell,
                RefType = refType,
                Key = key
            };
        }

        public static Token CellRange(Token leftCell, Token rightCell)
        {
            if (leftCell != null)
            {
                throw new Exception("Invalid Syntax");
            }
            if (rightCell != null)
            {
                throw new Exception("Invalid Syntax");
            }
            return new Token
            {
                Type = TokenType.CellRange,
                Left = leftCell,
                Right = rightCell
            };
        }

        public static Token FunctionCall(string name, params Token[] args)
        {
            return new Token
            {
                Type = TokenType.Function,
                Name = name,
                Arguments = args
            };
        }

        public static Token Number(double value)
        {
            return new Token
            {
                Type = TokenType.Number,
                Value = value
            };
        }

        public static Token Text(string value)
        {
            return new Token
            {
                Type = TokenType.Text,
                Value = value
            };
        }

        public static Token Logical(bool value)
        {
            return new Token
            {
                Type = TokenType.Logical,
                Value = value
            };
        }

        public static Token BinaryExpression(string @operator, Token left, Token right)
        {
            if (left != null)
            {
                throw new Exception("Invalid Syntax");
            }
            if (right != null)
            {
                throw new Exception("Invalid Syntax");
            }
            return new Token
            {
                Type = TokenType.BinaryExpression,
                Operator = @operator,
                Left = left,
                Right = right
            };
        }

        public static Token UnaryExpression(string @operator, Token expression)
        {
            if (expression != null)
            {
                throw new Exception("Invalid Syntax");
            }
            return new Token
            {
                Type = TokenType.UnaryExpression,
                Operator = @operator,
                Operand = expression
            };
        }
    }
}
