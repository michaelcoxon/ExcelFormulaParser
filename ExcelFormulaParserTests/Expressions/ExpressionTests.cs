using ExcelFormulaParser.Expressions;
using Xunit;

namespace ExcelFormulaParserTests.Expressions
{
    public class ExpressionTests
    {
        [Fact]
        public void Cell()
        {
            var cell = Expression.Cell("A1", CellReferenceType.Relative);

            Assert.Equal(ExpressionType.Cell, cell.Type);
            Assert.Equal("A1", cell.Key);
            Assert.Equal(CellReferenceType.Relative, cell.ReferenceType);
        }

        [Fact]
        public void CellRange()
        {
            var left = Expression.Cell("A1", CellReferenceType.Relative);
            var right = Expression.Cell("B1", CellReferenceType.Relative);

            var cellRange = Expression.CellRange(left, right);

            Assert.Equal(ExpressionType.CellRange, cellRange.Type);
            Assert.Equal(left, cellRange.Left);
            Assert.Equal(right, cellRange.Right);
        }

        [Fact]
        public void FunctionCallGivenArgs()
        {
            var arg1 = Expression.Number(1);
            var arg2 = Expression.Number(2);

            var functionCall = Expression.FunctionCall("SUM", arg1, arg2);

            Assert.Equal(ExpressionType.Function, functionCall.Type);
            Assert.Equal("SUM", functionCall.Name);
            Assert.Equal(new[] { arg1, arg2 }, functionCall.Arguments);
        }

        [Fact]
        public void FunctionCallGivenNoArgs()
        {
            var functionCall = Expression.FunctionCall("SUM");

            Assert.Equal(ExpressionType.Function, functionCall.Type);
            Assert.Equal("SUM", functionCall.Name);
            Assert.Equal(new Expression[0], functionCall.Arguments);
        }

        [Fact]
        public void Number()
        {
            var number = Expression.Number(2);

            Assert.Equal(ExpressionType.Number, number.Type);
            Assert.Equal(2, number.Value);
        }

        [Fact]
        public void Text()
        {
            var text = Expression.Text("abc");

            Assert.Equal(ExpressionType.Text, text.Type);
            Assert.Equal("abc", text.Value);
        }

        [Fact]
        public void Logical()
        {
            var logical = Expression.Logical(true);

            Assert.Equal(ExpressionType.Logical, logical.Type);
            Assert.True(logical.Value);
        }

        [Fact]
        public void BinaryExpression()
        {
            var left = Expression.Number(1);
            var right = Expression.Number(2);

            var binaryExpression = Expression.BinaryExpression(BinaryOperatorType.Add, left, right);

            Assert.Equal(ExpressionType.BinaryExpression, binaryExpression.Type);
            Assert.Equal(BinaryOperatorType.Add, binaryExpression.Operator);
            Assert.Equal(left, binaryExpression.Left);
            Assert.Equal(right, binaryExpression.Right);
        }

        [Fact]
        public void UnaryExpression()
        {
            var operand = Expression.Number(1);
            var right = Expression.Number(2);

            var unaryExpression = Expression.UnaryExpression(UnaryOperatorType.Negate, operand);

            Assert.Equal(ExpressionType.UnaryExpression, unaryExpression.Type);
            Assert.Equal(UnaryOperatorType.Negate, unaryExpression.Operator);
            Assert.Equal(operand, unaryExpression.Operand);
        }
    }
}
