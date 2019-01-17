using System;

namespace ExcelFormulaParser.Expressions
{
    [Serializable]
    public sealed class InvalidSyntaxException : Exception
    {
        public int Position { get; }
        public InvalidSyntaxException(int pos = -1)
        {
            this.Position = pos;
        }

        public InvalidSyntaxException(string message, int pos = -1) : base(message)
        {
            this.Position = pos;
        }

        public InvalidSyntaxException(string message, Exception inner, int pos = -1) : base(message, inner)
        {
            this.Position = pos;
        }

        protected InvalidSyntaxException(System.Runtime.Serialization.SerializationInfo info, System.Runtime.Serialization.StreamingContext context) : base(info, context)
        {
        }
    }
}
