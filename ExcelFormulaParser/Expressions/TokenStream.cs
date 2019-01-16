// https://github.com/psalaets/excel-formula-ast/blob/master/lib/token-stream.js

using ExcelFormulaParser.FormulaTokenizer;
using System;
using System.Linq;

namespace ExcelFormulaParser.Expressions
{
    public class TokenStream
    {
        private readonly Token[] _arr;
        private int _index = 0;

        public TokenStream(Token[] tokens)
        {
            this._arr = tokens.Append(null).ToArray();
        }

        public void Consume()
        {
            this._index += 1;
            if (this._index >= this._arr.Length)
            {
                throw new Exception("Invalid Syntax");
            }
        }

        public Token GetNext()
        {
            return this._arr[this._index];
        }

        public bool NextIs(TokenType type, TokenSubType? subType = null)
        {
            var next = this.GetNext();

            if (next is null)
            {
                return false;
            }

            if (this.GetNext().Type != type)
            {
                return false;
            }

            if (subType != null && this.GetNext().SubType != subType)
            {
                return false;
            }

            return true;
        }

        public bool NextIsOpenParen()
        {
            return this.NextIs(TokenType.SubExpression, TokenSubType.Start);
        }

        public bool NextIsTerminal()
        {
            if (this.NextIsNumber())
            {
                return true;
            }

            if (this.NextIsText())
            {
                return true;
            }

            if (this.NextIsLogical())
            {
                return true;
            }

            if (this.NextIsRange())
            {
                return true;
            }

            return false;
        }

        public bool NextIsFunctionCall()
        {
            return this.NextIs(TokenType.Function, TokenSubType.Start);
        }

        public bool NextIsFunctionArgumentSeparator()
        {
            return this.NextIs(TokenType.Argument);
        }

        public bool NextIsEndOfFunctionCall()
        {
            return this.NextIs(TokenType.Function, TokenSubType.Stop);
        }

        public bool NextIsBinaryOperator()
        {
            return this.NextIs(TokenType.OperatorInfix);
        }

        public bool NextIsPrefixOperator()
        {
            return this.NextIs(TokenType.OperatorPrefix);
        }

        public bool NextIsPostfixOperator()
        {
            return this.NextIs(TokenType.OperatorPostfix);
        }

        public bool NextIsRange()
        {
            return this.NextIs(TokenType.Operand, TokenSubType.Range);
        }

        public bool NextIsNumber()
        {
            return this.NextIs(TokenType.Operand, TokenSubType.Number);
        }

        public bool NextIsText()
        {
            return this.NextIs(TokenType.Operand, TokenSubType.Text);
        }

        public bool NextIsLogical()
        {
            return this.NextIs(TokenType.Operand, TokenSubType.Logical);
        }

        public int Pos()
        {
            return this._index;
        }
    }
}
