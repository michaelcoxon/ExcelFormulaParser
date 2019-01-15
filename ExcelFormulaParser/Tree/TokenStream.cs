// https://github.com/psalaets/excel-formula-ast/blob/master/lib/token-stream.js

using System;
using System.Linq;

namespace ExcelFormulaParser.Tree
{
    public class TokenStream
    {
        private readonly Token end;
        private readonly Token[] arr;
        private int index = 0;


        public TokenStream(Token[] tokens)
        {
            this.end = new Token();
            this.arr = tokens.Append(this.end).ToArray();
        }

        public void Consume()
        {
            this.index += 1;
            if (this.index >= this.arr.Length)
            {
                throw new Exception("Invalid Syntax");
            }
        }

        public Token GetNext()
        {
            return this.arr[this.index];
        }

        public bool NextIs(TokenType type, string subtype = null)
        {
            if (this.GetNext().Type != type)
            {
                return false;
            }

            if (subtype != null && this.GetNext().SubType != subtype)
            {
                return false;
            }

            return true;
        }

        public bool NextIsOpenParen()
        {
            return this.NextIs(TokenType.subexpression, "start");
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
            return this.NextIs(TokenType.Function, "start");
        }

        public bool NextIsFunctionArgumentSeparator()
        {
            return this.NextIs(TokenType.argument);
        }

        public bool NextIsEndOfFunctionCall()
        {
            return this.NextIs(TokenType.Function, "stop");
        }

        public bool NextIsBinaryOperator()
        {
            return this.NextIs(TokenType.operatorInfix);
        }

        public bool NextIsPrefixOperator()
        {
            return this.NextIs(TokenType.operatorPrefix);
        }

        public bool NextIsPostfixOperator()
        {
            return this.NextIs(TokenType.operatorPostfix);
        }

        public bool NextIsRange()
        {
            return this.NextIs(TokenType.operand, "range");
        }

        public bool NextIsNumber()
        {
            return this.NextIs(TokenType.operand, "number");
        }

        public bool NextIsText()
        {
            return this.NextIs(TokenType.operand, "text");
        }

        public bool NextIsLogical()
        {
            return this.NextIs(TokenType.operand, "logical");
        }

        public int Pos()
        {
            return this.index;
        }
    }
}
