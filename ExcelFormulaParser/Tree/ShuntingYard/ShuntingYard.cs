using System.Collections.Generic;

namespace ExcelFormulaParser.Tree.ShuntingYard
{
    public class ShuntingYard
    {
        public Stack<Token> operands = new Stack<Token>();
        public Stack<Operator> operators = new Stack<Operator>();

        public ShuntingYard()
        {
            this.operators.Push(Operator.SENTINEL);
        }
    }
}
