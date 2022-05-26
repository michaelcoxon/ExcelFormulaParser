using ExcelFormulaParser.Expressions;
using System.Collections.Generic;
using System.Text;
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

    public class CallsTestExpressionVisitor : ExpressionVisitor
    {
        public List<Call> Calls { get; }

        public CallsTestExpressionVisitor()
        {
            this.Calls = new List<Call>();
        }

        protected override Expression VisitCell(CellExpression node)
        {
            this.Calls.Add(new Call(nameof(VisitCell), node));
            return base.VisitCell(node);
        }

        protected override Expression VisitCellRange(RangeExpression node)
        {
            this.Calls.Add(new Call(nameof(VisitCellRange), node));
            this.Visit(node.Left);
            this.Visit(node.Right);

            return base.VisitCellRange(node);
        }

        protected override Expression VisitFunction(FunctionExpression node)
        {
            this.Calls.Add(new Call(nameof(VisitFunction), node));
            foreach (var arg in node.Arguments)
            {
                this.Visit(arg);
            }
            return base.VisitFunction(node);
        }

        protected override Expression VisitNumber(NumberExpression node)
        {
            this.Calls.Add(new Call(nameof(VisitNumber), node));
            return base.VisitNumber(node);
        }

        protected override Expression VisitText(TextExpression node)
        {
            this.Calls.Add(new Call(nameof(VisitText), node));
            return base.VisitText(node);
        }

        protected override Expression VisitLogical(LogicalExpression node)
        {
            this.Calls.Add(new Call(nameof(VisitLogical), node));
            return base.VisitLogical(node);
        }

        protected override Expression VisitBinaryExpression(BinaryExpression node)
        {
            this.Calls.Add(new Call(nameof(VisitBinaryExpression), node));
            this.Visit(node.Left);
            this.Visit(node.Right);
            return base.VisitBinaryExpression(node);
        }

        protected override Expression VisitUnaryExpression(UnaryExpression node)
        {
            this.Visit(node.Operand);
            this.Calls.Add(new Call(nameof(VisitUnaryExpression), node));
            return base.VisitUnaryExpression(node);
        }
    }

    public class StringBuilderExpressionVisitor : ExpressionVisitor
    {
        private readonly StringBuilder _stringBuilder;

        public StringBuilderExpressionVisitor(StringBuilder stringBuilder)
        {
            this._stringBuilder = stringBuilder;
        }

        protected override Expression VisitCell(CellExpression node)
        {
            //this.Calls.Add(new Call(nameof(VisitCell), node));
            this._stringBuilder.Append(node.Key);
            return base.VisitCell(node);
        }

        protected override Expression VisitCellRange(RangeExpression node)
        {
            //this.Calls.Add(new Call(nameof(VisitCellRange), node));
            this.Visit(node.Left);
            this._stringBuilder.Append(":");
            this.Visit(node.Right);

            return base.VisitCellRange(node);
        }

        protected override Expression VisitFunction(FunctionExpression node)
        {
            //this.Calls.Add(new Call(nameof(VisitFunction), node));
            foreach (var arg in node.Arguments)
            {
                this.Visit(arg);
            }
            return base.VisitFunction(node);
        }

        protected override Expression VisitNumber(NumberExpression node)
        {
            //this.Calls.Add(new Call(nameof(VisitNumber), node));
            return base.VisitNumber(node);
        }

        protected override Expression VisitText(TextExpression node)
        {
            //this.Calls.Add(new Call(nameof(VisitText), node));
            return base.VisitText(node);
        }

        protected override Expression VisitLogical(LogicalExpression node)
        {
            //this.Calls.Add(new Call(nameof(VisitLogical), node));
            return base.VisitLogical(node);
        }

        protected override Expression VisitBinaryExpression(BinaryExpression node)
        {
            //this.Calls.Add(new Call(nameof(VisitBinaryExpression), node));
            this.Visit(node.Left);
            this.Visit(node.Right);
            return base.VisitBinaryExpression(node);
        }

        protected override Expression VisitUnaryExpression(UnaryExpression node)
        {
            this.Visit(node.Operand);
            //this.Calls.Add(new Call(nameof(VisitUnaryExpression), node));
            return base.VisitUnaryExpression(node);
        }
    }

    public class ExecuteExpressionVisitor : ExpressionVisitor
    {
        public ExecuteExpressionVisitor()
        {
        }
    }

}
