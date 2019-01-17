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

        public Token(string value, TokenType type, TokenSubType? subType = null)
        {
            this.Value = value;
            this.Type = type;
            this.SubType = subType;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is Token item))
            {
                return false;
            }

            return this.Type.Equals(item.Type)
                && this.SubType.Equals(item.SubType)
                && this.Value.Equals(item.Value)
                ;
        }

        public override int GetHashCode()
        {
            return this.Type.GetHashCode()
                ^ this.SubType.GetHashCode()
                ^ this.Value.GetHashCode()
                ;
        }


#if !DEBUG
        public override string ToString()
        {
            return this.Value;
        }
#endif
    }
}
