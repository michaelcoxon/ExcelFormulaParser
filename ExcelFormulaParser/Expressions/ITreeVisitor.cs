using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFormulaParser.Expressions
{
    public interface IExpressionVisitor
    {
        void EnterCell(CellExpression node);
        void ExitCell(CellExpression node);
        void EnterCellRange(RangeExpression node);
        void ExitCellRange(RangeExpression node);
        void EnterFunction(FunctionExpression node);
        void ExitFunction(FunctionExpression node);
        void EnterNumber(NumberExpression node);
        void ExitNumber(NumberExpression node);
        void EnterText(TextExpression node);
        void ExitText(TextExpression node);
        void EnterLogical(LogicalExpression node);
        void ExitLogical(LogicalExpression node);
        void EnterBinaryExpression(BinaryExpression node);
        void ExitBinaryExpression(BinaryExpression node);
        void EnterUnaryExpression(UnaryExpression node);
        void ExitUnaryExpression(UnaryExpression node);
    }
}
