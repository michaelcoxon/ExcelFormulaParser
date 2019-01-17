using System.Collections.Generic;

namespace ExcelFormulaParser.FormulaTokenizer
{
    internal class TokenCollection
    {
        private readonly List<Token> _items;
        private int _index;

        public TokenCollection()
        {
            this._items = new List<Token>();
            this._index = -1;
        }

        public Token Add(string value, TokenType type, TokenSubType? subtype = null)
        {
            var token = new Token(value, type, subtype);
            this.AddRef(token);
            return token;
        }

        public void AddRef(Token token)
        {
            this._items.Add(token);
        }

        public void Reset()
        {
            this._index = -1;
        }

        public bool BOF()
        {
            return this._index <= 0;
        }

        public bool EOF()
        {
            return this._index >= this._items.Count - 1;
        }

        public bool MoveNext()
        {
            if (this.EOF())
            {
                return false;
            }

            this._index++;
            return true;
        }

        public Token Current()
        {
            if (this._index == -1)
            {
                return null;
            }

            return this._items[this._index];
        }

        public Token Next()
        {
            if (this.EOF())
            {
                return null;
            }

            return this._items[this._index + 1];
        }

        public Token Previous()
        {
            if (this._index < 1)
            {
                return null;
            }

            return (this._items[this._index - 1]);
        }

        public Token[] ToArray()
        {
            return this._items.ToArray();
        }
    }
}
