using ExcelFormulaParser.Expressions;
using System;
using System.Text;
using Xunit;

namespace ExcelFormulaParserTests.Expressions
{
    public class ExpressionVisitorTests
    {
        [Fact]
        public void CellNode()
        {
            var visitor = new CallsTestExpressionVisitor();
            var A1 = Expression.Cell("A1", CellReferenceType.Relative);

            visitor.Visit(A1);

            Assert.Equal(new[]
            {
                new Call("VisitCell", A1),
            },
            visitor.Calls.ToArray());
        }

        [Fact]
        public void CellRangeNode()
        {
            var visitor = new CallsTestExpressionVisitor();
            var A1 = Expression.Cell("A1", CellReferenceType.Relative);
            var A2 = Expression.Cell("A2", CellReferenceType.Relative);
            var cellRange = Expression.CellRange(A1, A2);

            visitor.Visit(cellRange);

            Assert.Equal(new[]
            {
                new Call("VisitCellRange", cellRange),
                new Call("VisitCell", A1),
                new Call("VisitCell", A2),
            },
            visitor.Calls.ToArray());
        }

        [Fact]
        public void CellRangeNode_str()
        {
            var sb = new StringBuilder();
            var visitor = new StringBuilderExpressionVisitor(sb);
            var A1 = Expression.Cell("A1", CellReferenceType.Relative);
            var A2 = Expression.Cell("A2", CellReferenceType.Relative);
            var cellRange = Expression.CellRange(A1, A2);

            visitor.Visit(cellRange);

            Console.WriteLine(sb);

        }
    }
}
