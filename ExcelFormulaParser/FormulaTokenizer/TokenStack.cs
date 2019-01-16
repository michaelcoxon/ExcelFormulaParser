using System.Collections.Generic;

namespace ExcelFormulaParser.FormulaTokenizer
{
    public class TokenStack
    {
        private readonly Stack<Token> _items;
        public TokenStack()
        {
            this._items = new Stack<Token>();
        }

        public void Push(Token token)
        {
            this._items.Push(token);
        }

        public Token Pop()
        {
            var token = this._items.Pop();
            return new Token("", token.Type, TokenSubType.Stop);
        }

        public Token Token()
        {
            if (this._items.Count > 0)
            {
                return this._items.ToArray()[this._items.Count - 1];
            }
            else
            {
                return null;
            }
        }

        public string Value()
        {
            return this.Token()?.Value ?? "";
        }

        public TokenType? Type()
        {
            return this.Token()?.Type ?? null;
        }

        public TokenSubType? SubType()
        {
            return this.Token()?.SubType ?? null;
        }
    }
}
