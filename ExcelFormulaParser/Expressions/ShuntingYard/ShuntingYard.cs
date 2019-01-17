using System.Collections.Generic;

namespace ExcelFormulaParser.Expressions.ShuntingYard
{
    internal class ShuntingYard
    {
        public Stack<Expression> operands = new Stack<Expression>();
        public Stack<Operator> operators = new Stack<Operator>();

        public ShuntingYard()
        {
            this.operators.Push(Operator.SENTINEL);
        }
    }
}
