using System;

namespace ExcelFormulaParser.Expressions.ShuntingYard
{
    public class Operator
    {
        public static readonly Operator SENTINEL = new Operator("S", 0);

        public readonly string symbol;
        public readonly int precendence;
        public readonly int operandCount;
        public readonly bool leftAssociative;

        public Operator(string symbol, int precendence, int operandCount = 2, bool leftAssociative = true)
        {
            if (operandCount < 1 || operandCount > 2)
            {
                throw new Exception($"operandCount cannot be { operandCount }, must be 1 or 2");
            }

            this.symbol = symbol;
            this.precendence = precendence;
            this.operandCount = operandCount;
            this.leftAssociative = leftAssociative;
        }

        public bool IsUnary()
        {
            return this.operandCount == 1;
        }

        public bool IsBinary()
        {
            return this.operandCount == 2;
        }

        public bool EvaluatesBefore(Operator other)
        {
            if (this == Operator.SENTINEL)
            {
                return false;
            }

            if (other == Operator.SENTINEL)
            {
                return true;
            }

            if (other.IsUnary())
            {
                return false;
            }

            if (this.IsUnary())
            {
                return this.precendence >= other.precendence;
            }
            else if (this.IsBinary())
            {
                if (this.precendence == other.precendence)
                {
                    return this.leftAssociative;
                }
                else
                {
                    return this.precendence > other.precendence;
                }
            }

            return false;
        }
    }
}
