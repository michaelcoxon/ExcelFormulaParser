using System;

namespace ExcelFormulaParser.Expressions
{
    public abstract class ExpressionVisitor
    {
        protected virtual void EnterCell(CellExpression node) { }
        protected virtual void ExitCell(CellExpression node) { }
        protected virtual void EnterCellRange(RangeExpression node) { }
        protected virtual void ExitCellRange(RangeExpression node) { }
        protected virtual void EnterFunction(FunctionExpression node) { }
        protected virtual void ExitFunction(FunctionExpression node) { }
        protected virtual void EnterNumber(NumberExpression node) { }
        protected virtual void ExitNumber(NumberExpression node) { }
        protected virtual void EnterText(TextExpression node) { }
        protected virtual void ExitText(TextExpression node) { }
        protected virtual void EnterLogical(LogicalExpression node) { }
        protected virtual void ExitLogical(LogicalExpression node) { }
        protected virtual void EnterBinaryExpression(BinaryExpression node) { }
        protected virtual void ExitBinaryExpression(BinaryExpression node) { }
        protected virtual void EnterUnaryExpression(UnaryExpression node) { }
        protected virtual void ExitUnaryExpression(UnaryExpression node) { }

        public void Visit(Expression node)
        {
            VisitNode(node, this);
        }

        public static void Visit(Expression node, ExpressionVisitor visitor)
        {
            VisitNode(node, visitor);
        }

        private static void VisitNode(Expression node, ExpressionVisitor visitor)
        {
            switch (node)
            {
                case CellExpression expr:
                    VisitCell(expr, visitor);
                    break;
                case RangeExpression expr:
                    VisitCellRange(expr, visitor);
                    break;
                case FunctionExpression expr:
                    VisitFunction(expr, visitor);
                    break;
                case NumberExpression expr:
                    VisitNumber(expr, visitor);
                    break;
                case TextExpression expr:
                    VisitText(expr, visitor);
                    break;
                case LogicalExpression expr:
                    VisitLogical(expr, visitor);
                    break;
                case BinaryExpression expr:
                    VisitBinaryExpression(expr, visitor);
                    break;
                case UnaryExpression expr:
                    VisitUnaryExpression(expr, visitor);
                    break;
                default:
                    throw new NotSupportedException($"Node of type '{node.Type}' not supported");
            }
        }

        private static void VisitCell(CellExpression node, ExpressionVisitor visitor)
        {
            visitor.EnterCell(node);
            visitor.ExitCell(node);
        }

        private static void VisitCellRange(RangeExpression node, ExpressionVisitor visitor)
        {
            visitor.EnterCellRange(node);

            VisitNode(node.Left, visitor);
            VisitNode(node.Right, visitor);

            visitor.ExitCellRange(node);
        }

        private static void VisitFunction(FunctionExpression node, ExpressionVisitor visitor)
        {
            visitor.EnterFunction(node);

            foreach (var arg in node.Arguments)
            {
                VisitNode(arg, visitor);
            }

            visitor.ExitFunction(node);
        }

        private static void VisitNumber(NumberExpression node, ExpressionVisitor visitor)
        {
            visitor.EnterNumber(node);
            visitor.ExitNumber(node);
        }

        private static void VisitText(TextExpression node, ExpressionVisitor visitor)
        {
            visitor.EnterText(node);
            visitor.ExitText(node);
        }

        private static void VisitLogical(LogicalExpression node, ExpressionVisitor visitor)
        {
            visitor.EnterLogical(node);
            visitor.ExitLogical(node);
        }

        private static void VisitBinaryExpression(BinaryExpression node, ExpressionVisitor visitor)
        {
            visitor.EnterBinaryExpression(node);

            VisitNode(node.Left, visitor);
            VisitNode(node.Right, visitor);

            visitor.ExitBinaryExpression(node);
        }

        private static void VisitUnaryExpression(UnaryExpression node, ExpressionVisitor visitor)
        {
            visitor.EnterUnaryExpression(node);

            VisitNode(node.Operand, visitor);

            visitor.ExitUnaryExpression(node);
        }
    }
}
