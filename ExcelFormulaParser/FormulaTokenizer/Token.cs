namespace ExcelFormulaParser.FormulaTokenizer
{
    public enum TokenType
    {
        Noop,
        Operand,
        Function,
        SubExpression,
        Argument,
        OperatorPrefix,
        OperatorInfix,
        OperatorPostfix,
        WhiteSpace,
        Unknown,
    }

    public enum TokenSubType
    {
        Start,
        Stop,

        Text,
        Number,
        Logical,
        Error,
        Range,

        Math,
        Concatenate,
        Intersect,
        Union,
    }


    public class Token
    {
        public string Value;
        public TokenType Type;
        public TokenSubType? SubType;

        /*
        public string RefType;
        public string Key;
        public string Name;
        public Token[] Arguments;
        public Token Left;
        public Token Right;
        public string Operator;
        public Token Operand;
        */

        public Token(string value, TokenType type, TokenSubType? subType = null)
        {
            this.Value = value;
            this.Type = type;
            this.SubType = subType;
        }
    }
}
