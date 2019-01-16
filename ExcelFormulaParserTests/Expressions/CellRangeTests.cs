using ExcelFormulaParser.Expressions;
using ExcelFormulaParser.FormulaTokenizer;
using Xunit;

namespace ExcelFormulaParserTests.Expressions
{
    public class CellRangeTests
    {
        [Fact]
        public void A1()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("A1"));

            TestHelper.AssertExpression(tree, Expression.Cell("A1", CellReferenceType.Relative));
        }

        [Fact]
        public void ADollar1()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("A$1"));

            TestHelper.AssertExpression(tree, Expression.Cell("A$1", CellReferenceType.Mixed));
        }

        [Fact]
        public void DollarA1()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("$A1"));

            TestHelper.AssertExpression(tree, Expression.Cell("$A1", CellReferenceType.Mixed));
        }

        [Fact]
        public void DollarADollar1()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("$A$1"));

            TestHelper.AssertExpression(tree, Expression.Cell("$A$1", CellReferenceType.Absolute));
        }

        [Fact]
        public void A1ToA4()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("A1:A4"));

            TestHelper.AssertExpression(
                tree,
                Expression.CellRange(
                    Expression.Cell("A1", CellReferenceType.Relative),
                    Expression.Cell("A4", CellReferenceType.Relative)
                )
             );
        }

        [Fact]
        public void DollarA1ToADollar4()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("$A1:A$4"));

            TestHelper.AssertExpression(
                tree,
                Expression.CellRange(
                    Expression.Cell("$A1", CellReferenceType.Mixed),
                    Expression.Cell("A$4", CellReferenceType.Mixed)
                )
             );
        }

        [Fact]
        public void DollarADollar1ToDollarADollar4()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("$A$1:$A$4"));

            TestHelper.AssertExpression(
                tree,
                Expression.CellRange(
                    Expression.Cell("$A$1", CellReferenceType.Absolute),
                    Expression.Cell("$A$4", CellReferenceType.Absolute)
                )
             );
        }

        [Fact]
        public void OneToFour()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("1:4"));

            TestHelper.AssertExpression(
                tree,
                Expression.CellRange(
                    Expression.Cell("1", CellReferenceType.Relative),
                    Expression.Cell("4", CellReferenceType.Relative)
                )
             );
        }

        [Fact]
        public void CToG()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("C:G"));

            TestHelper.AssertExpression(
                tree,
                Expression.CellRange(
                    Expression.Cell("C", CellReferenceType.Relative),
                    Expression.Cell("G", CellReferenceType.Relative)
                )
             );
        }

        [Fact]
        public void CToDollarG()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("C:$G"));

            TestHelper.AssertExpression(
                tree,
                Expression.CellRange(
                    Expression.Cell("C", CellReferenceType.Relative),
                    Expression.Cell("$G", CellReferenceType.Absolute)
                )
             );
        }

        [Fact]
        public void CToG5()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("C:G5"));

            TestHelper.AssertExpression(
                tree,
                Expression.CellRange(
                    Expression.Cell("C", CellReferenceType.Relative),
                    Expression.Cell("G5", CellReferenceType.Relative)
                )
             );
        }

        [Fact]
        public void FiveToDFive()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("5:D5"));

            TestHelper.AssertExpression(
                tree,
                Expression.CellRange(
                    Expression.Cell("5", CellReferenceType.Relative),
                    Expression.Cell("D5", CellReferenceType.Relative)
                )
             );
        }

        [Fact]
        public void A1ToB3CommaC1ToD3()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("A1:B3,C1:D3"));

            TestHelper.AssertExpression(
                tree,
                Expression.BinaryExpression(
                    BinaryOperatorType.Comma,
                    Expression.CellRange(
                        Expression.Cell("A1", CellReferenceType.Relative),
                        Expression.Cell("B3", CellReferenceType.Relative)
                    ),
                    Expression.CellRange(
                        Expression.Cell("C1", CellReferenceType.Relative),
                        Expression.Cell("D3", CellReferenceType.Relative)
                    )
                )
             );
        }

        [Fact]
        public void A1ToB3WhitespaceC1ToD3()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("A1:B3 C1:D3"));

            TestHelper.AssertExpression(
                tree,
                Expression.BinaryExpression(
                    BinaryOperatorType.Whitespace,
                    Expression.CellRange(
                        Expression.Cell("A1", CellReferenceType.Relative),
                        Expression.Cell("B3", CellReferenceType.Relative)
                    ),
                    Expression.CellRange(
                        Expression.Cell("C1", CellReferenceType.Relative),
                        Expression.Cell("D3", CellReferenceType.Relative)
                    )
                )
             );
        }
    }
}
