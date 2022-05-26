using System;

namespace ExcelFormulaParser.Expressions
{
    public abstract class ExpressionVisitor
    {
        protected virtual Expression VisitCell(CellExpression node) => node;
        protected virtual Expression VisitCellRange(RangeExpression node) => node;
        protected virtual Expression VisitFunction(FunctionExpression node) => node;
        protected virtual Expression VisitNumber(NumberExpression node) => node;
        protected virtual Expression VisitText(TextExpression node) => node;
        protected virtual Expression VisitLogical(LogicalExpression node) => node;
        protected virtual Expression VisitBinaryExpression(BinaryExpression node) => node;
        protected virtual Expression VisitUnaryExpression(UnaryExpression node) => node;

        public Expression Visit(Expression node)
        {
            return this.VisitNode(node);
        }

        public static Expression Visit(Expression node, ExpressionVisitor visitor)
        {
            return visitor.VisitNode(node);
        }

        private Expression VisitNode(Expression node)
        {
            return node switch
            {
                CellExpression expr => this.VisitCell(expr),
                RangeExpression expr => this.VisitCellRange(expr),
                FunctionExpression expr => this.VisitFunction(expr),
                NumberExpression expr => this.VisitNumber(expr),
                TextExpression expr => this.VisitText(expr),
                LogicalExpression expr => this.VisitLogical(expr),
                BinaryExpression expr => this.VisitBinaryExpression(expr),
                UnaryExpression expr => this.VisitUnaryExpression(expr),
                _ => throw new NotSupportedException($"Node of type '{node.Type}' not supported"),
            };
        }
    }
}
