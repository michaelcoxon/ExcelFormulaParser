using ExcelFormulaParser.Expressions;
using Xunit;

namespace ExcelFormulaParserTests.Expressions
{
    public class ExpressionVisitorTests
    {
        [Fact]
        public void CellNode()
        {
            var visitor = new TestExpressionVisitor();
            var A1 = Expression.Cell("A1", CellReferenceType.Relative);

            visitor.Visit(A1);

            Assert.Equal(new[]
            {
                new Call("EnterCell", A1),
                new Call("ExitCell", A1)
            },
            visitor.Calls.ToArray());
        }

        [Fact]
        public void CellRangeNode()
        {
            var visitor = new TestExpressionVisitor();
            var A1 = Expression.Cell("A1", CellReferenceType.Relative);
            var A2 = Expression.Cell("A2", CellReferenceType.Relative);
            var cellRange = Expression.CellRange(A1, A2);

            visitor.Visit(cellRange);

            Assert.Equal(new[]
            {
                new Call("EnterCellRange", cellRange),
                new Call("EnterCell", A1),
                new Call("ExitCell", A1),
                new Call("EnterCell", A2),
                new Call("ExitCell", A2),
                new Call("ExitCellRange", cellRange)
            },
            visitor.Calls.ToArray());
        }
    }
}
