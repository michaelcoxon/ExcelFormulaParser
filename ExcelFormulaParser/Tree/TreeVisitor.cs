namespace ExcelFormulaParser.Tree
{
    internal class TreeVisitor
    {
        public void Visit(Token node, ITreeVisitor visitor)
        {
            this.VisitNode(node, visitor);
        }

        private void VisitNode(Token node, ITreeVisitor visitor)
        {
            switch (node.Type)
            {
                case TokenType.Cell:
                    this.VisitCell(node, visitor);
                    break;
                case TokenType.CellRange:
                    this.VisitCellRange(node, visitor);
                    break;
                case TokenType.Function:
                    this.VisitFunction(node, visitor);
                    break;
                case TokenType.Number:
                    this.VisitNumber(node, visitor);
                    break;
                case TokenType.Text:
                    this.VisitText(node, visitor);
                    break;
                case TokenType.Logical:
                    this.VisitLogical(node, visitor);
                    break;
                case TokenType.BinaryExpression:
                    this.VisitBinaryExpression(node, visitor);
                    break;
                case TokenType.UnaryExpression:
                    this.VisitUnaryExpression(node, visitor);
                    break;
            }
        }

        private void VisitCell(Token node, ITreeVisitor visitor)
        {
            visitor.EnterCell(node);
            visitor.ExitCell(node);
        }

        private void VisitCellRange(Token node, ITreeVisitor visitor)
        {
            visitor.EnterCellRange(node);

            this.VisitNode(node.Left, visitor);
            this.VisitNode(node.Right, visitor);

            visitor.ExitCellRange(node);
        }

        private void VisitFunction(Token node, ITreeVisitor visitor)
        {
            visitor.EnterFunction(node);

            foreach (var arg in node.Arguments)
            {
                this.VisitNode(arg, visitor);
            }

            visitor.ExitFunction(node);
        }

        private void VisitNumber(Token node, ITreeVisitor visitor)
        {
            visitor.EnterNumber(node);
            visitor.ExitNumber(node);
        }

        private void VisitText(Token node, ITreeVisitor visitor)
        {
            visitor.EnterText(node);
            visitor.ExitText(node);
        }

        private void VisitLogical(Token node, ITreeVisitor visitor)
        {
            visitor.EnterLogical(node);
            visitor.ExitLogical(node);
        }

        private void VisitBinaryExpression(Token node, ITreeVisitor visitor)
        {
            visitor.EnterBinaryExpression(node);

            this.VisitNode(node.Left, visitor);
            this.VisitNode(node.Right, visitor);

            visitor.ExitBinaryExpression(node);
        }

        private void VisitUnaryExpression(Token node, ITreeVisitor visitor)
        {
            visitor.EnterUnaryExpression(node);

            this.VisitNode(node.Operand, visitor);

            visitor.ExitUnaryExpression(node);
        }
    }
}
