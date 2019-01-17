using ExcelFormulaParser.Expressions;
using System.Collections.Generic;
using Xunit;

namespace ExcelFormulaParserTests.Expressions
{
    public class TestHelper
    {
        public static void AssertExpression(Expression actual, Expression expected)
        {
            Assert.Equal(expected, actual);
        }
    }

    public struct Call
    {
        public readonly string MethodName;
        public readonly Expression Node;

        public Call(string methodName, Expression node)
        {
            this.MethodName = methodName;
            this.Node = node;
        }
    }

    public class TestExpressionVisitor : ExpressionVisitor
    {
        public List<Call> Calls { get; }

        public TestExpressionVisitor()
        {
            this.Calls = new List<Call>();
        }

        protected override void EnterBinaryExpression(BinaryExpression node)
        {
            this.Calls.Add(new Call(nameof(EnterBinaryExpression), node));
        }
        protected override void EnterCell(CellExpression node)
        {
            this.Calls.Add(new Call(nameof(EnterCell), node));
        }
        protected override void EnterCellRange(RangeExpression node)
        {
            this.Calls.Add(new Call(nameof(EnterCellRange), node));
        }
        protected override void EnterFunction(FunctionExpression node)
        {
            this.Calls.Add(new Call(nameof(EnterFunction), node));
        }
        protected override void EnterLogical(LogicalExpression node)
        {
            this.Calls.Add(new Call(nameof(EnterLogical), node));
        }
        protected override void EnterNumber(NumberExpression node)
        {
            this.Calls.Add(new Call(nameof(EnterNumber), node));
        }
        protected override void EnterText(TextExpression node)
        {
            this.Calls.Add(new Call(nameof(EnterText), node));
        }
        protected override void EnterUnaryExpression(UnaryExpression node)
        {
            this.Calls.Add(new Call(nameof(EnterUnaryExpression), node));
        }
        protected override void ExitBinaryExpression(BinaryExpression node)
        {
            this.Calls.Add(new Call(nameof(ExitBinaryExpression), node));
        }
        protected override void ExitCell(CellExpression node)
        {
            this.Calls.Add(new Call(nameof(ExitCell), node));
        }
        protected override void ExitCellRange(RangeExpression node)
        {
            this.Calls.Add(new Call(nameof(ExitCellRange), node));
        }
        protected override void ExitFunction(FunctionExpression node)
        {
            this.Calls.Add(new Call(nameof(ExitFunction), node));
        }
        protected override void ExitLogical(LogicalExpression node)
        {
            this.Calls.Add(new Call(nameof(ExitLogical), node));
        }
        protected override void ExitNumber(NumberExpression node)
        {
            this.Calls.Add(new Call(nameof(ExitNumber), node));
        }
        protected override void ExitText(TextExpression node)
        {
            this.Calls.Add(new Call(nameof(ExitText), node));
        }
        protected override void ExitUnaryExpression(UnaryExpression node)
        {
            this.Calls.Add(new Call(nameof(ExitUnaryExpression), node));
        }
    }
}
