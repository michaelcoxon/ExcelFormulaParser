using System;

namespace ExcelFormulaParser.Expressions
{
    internal class TreeVisitor
    {
        public void Visit(Expression node, IExpressionVisitor visitor)
        {
            this.VisitNode(node, visitor);
        }

        private void VisitNode(Expression node, IExpressionVisitor visitor)
        {
            switch (node)
            {
                case CellExpression expr:
                    this.VisitCell(expr, visitor);
                    break;
                case RangeExpression expr:
                    this.VisitCellRange(expr, visitor);
                    break;
                case FunctionExpression expr:
                    this.VisitFunction(expr, visitor);
                    break;
                case NumberExpression expr:
                    this.VisitNumber(expr, visitor);
                    break;
                case TextExpression expr:
                    this.VisitText(expr, visitor);
                    break;
                case LogicalExpression expr:
                    this.VisitLogical(expr, visitor);
                    break;
                case BinaryExpression expr:
                    this.VisitBinaryExpression(expr, visitor);
                    break;
                case UnaryExpression expr:
                    this.VisitUnaryExpression(expr, visitor);
                    break;
                default:
                    throw new NotSupportedException($"Node of type '{node.Type}' not supported");
            }
        }

        private void VisitCell(CellExpression node, IExpressionVisitor visitor)
        {
            visitor.EnterCell(node);
            visitor.ExitCell(node);
        }

        private void VisitCellRange(RangeExpression node, IExpressionVisitor visitor)
        {
            visitor.EnterCellRange(node);

            this.VisitNode(node.Left, visitor);
            this.VisitNode(node.Right, visitor);

            visitor.ExitCellRange(node);
        }

        private void VisitFunction(FunctionExpression node, IExpressionVisitor visitor)
        {
            visitor.EnterFunction(node);

            foreach (var arg in node.Arguments)
            {
                this.VisitNode(arg, visitor);
            }

            visitor.ExitFunction(node);
        }

        private void VisitNumber(NumberExpression node, IExpressionVisitor visitor)
        {
            visitor.EnterNumber(node);
            visitor.ExitNumber(node);
        }

        private void VisitText(TextExpression node, IExpressionVisitor visitor)
        {
            visitor.EnterText(node);
            visitor.ExitText(node);
        }

        private void VisitLogical(LogicalExpression node, IExpressionVisitor visitor)
        {
            visitor.EnterLogical(node);
            visitor.ExitLogical(node);
        }

        private void VisitBinaryExpression(BinaryExpression node, IExpressionVisitor visitor)
        {
            visitor.EnterBinaryExpression(node);

            this.VisitNode(node.Left, visitor);
            this.VisitNode(node.Right, visitor);

            visitor.ExitBinaryExpression(node);
        }

        private void VisitUnaryExpression(UnaryExpression node, IExpressionVisitor visitor)
        {
            visitor.EnterUnaryExpression(node);

            this.VisitNode(node.Operand, visitor);

            visitor.ExitUnaryExpression(node);
        }
    }
}
