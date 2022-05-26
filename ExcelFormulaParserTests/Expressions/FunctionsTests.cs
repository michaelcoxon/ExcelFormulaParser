using ExcelFormulaParser.Expressions;
using ExcelFormulaParser.FormulaTokenizer;
using Xunit;

namespace ExcelFormulaParserTests.Expressions
{
    public class FunctionsTests
    {
        [Fact]
        public void Sum()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("SUM()"));

            TestHelper.AssertExpression(tree, Expression.FunctionCall("SUM"));
        }

        [Fact]
        public void NegateSum()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("-SUM()"));

            TestHelper.AssertExpression(tree,
                Expression.UnaryExpression(
                    UnaryOperatorType.Negate,
                    Expression.FunctionCall("SUM")
                )
            );
        }

        [Fact]
        public void SumOne()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("SUM(1)"));

            TestHelper.AssertExpression(tree,
                Expression.FunctionCall(
                    "SUM",
                    Expression.Number(1)
                )
            );
        }

        [Fact]
        public void SumOneAndTwo()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("SUM(1, 2)"));

            TestHelper.AssertExpression(tree,
                Expression.FunctionCall(
                    "SUM",
                    Expression.Number(1),
                    Expression.Number(2)
                )
            );
        }

        [Fact]
        public void SumOneAndSumTwoAndThree()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("SUM(1, SUM(2, 3))"));

            TestHelper.AssertExpression(tree,
                Expression.FunctionCall(
                    "SUM",
                    Expression.Number(1),
                    Expression.FunctionCall(
                        "SUM",
                        Expression.Number(2),
                        Expression.Number(3)
                    )
                )
            );
        }

        [Fact]
        public void SumTenDivideFourAndSumTwoAndThree()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("SUM(10 / 4, SUM(2, 3))"));

            TestHelper.AssertExpression(tree,
                Expression.FunctionCall(
                    "SUM",
                    Expression.BinaryExpression(
                        BinaryOperatorType.Divide,
                        Expression.Number(10),
                        Expression.Number(4)
                    ),
                    Expression.FunctionCall(
                        "SUM",
                        Expression.Number(2),
                        Expression.Number(3)
                    )
                )
            );
        }

        [Fact]
        public void TwoAddSumOne()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("2 + SUM(1)"));

            TestHelper.AssertExpression(tree,
                Expression.BinaryExpression(
                    BinaryOperatorType.Add,
                    Expression.Number(2),
                    Expression.FunctionCall(
                        "SUM",
                        Expression.Number(1)
                    )
                )
            );
        }

        [Fact]
        public void TwoAddSumOneAndTwoAndThreeAndFour()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("2 + SUM(1, 2, 3, 4)"));

            TestHelper.AssertExpression(tree,
                Expression.BinaryExpression(
                    BinaryOperatorType.Add,
                    Expression.Number(2),
                    Expression.FunctionCall(
                        "SUM",
                        Expression.Number(1),
                        Expression.Number(2),
                        Expression.Number(3),
                        Expression.Number(4)
                    )
                )
            );
        }

        [Fact]
        public void SumTwoAddSumOne()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("SUM(2) + SUM(1)"));

            TestHelper.AssertExpression(tree,
                Expression.BinaryExpression(
                    BinaryOperatorType.Add,
                    Expression.FunctionCall(
                        "SUM",
                        Expression.Number(2)
                    ),
                    Expression.FunctionCall(
                        "SUM",
                        Expression.Number(1)
                    )
                )
            );
        }

        [Fact]
        public void SumSumOneAndTwoAddThree()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("SUM(SUM(1), 2 + 3)"));

            TestHelper.AssertExpression(tree,
                Expression.FunctionCall(
                    "SUM",
                    Expression.FunctionCall(
                        "SUM",
                        Expression.Number(1)
                    ),
                    Expression.BinaryExpression(
                        BinaryOperatorType.Add,
                        Expression.Number(2),
                        Expression.Number(3)
                    )
                )
            );
        }
    }
}
