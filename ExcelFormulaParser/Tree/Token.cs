namespace ExcelFormulaParser.Tree
{
    public enum TokenType
    {
        UnaryExpression,
        BinaryExpression,
        Logical,
        Text,
        Number,
        Function,
        CellRange,
        Cell,
        subexpression,
        operand,
        operatorPostfix,
        argument,
        operatorInfix,
        operatorPrefix
    }

    public class Token
    {
        public string Raw;
        public TokenType Type;
        public string SubType;
        public string RefType;
        public string Key;
        public string Name;
        public object Value;
        public Token[] Arguments;
        public Token Left;
        public Token Right;
        public string Operator;
        public Token Operand;
    }
}
