using ExcelFormulaParser.Expressions;
using ExcelFormulaParser.FormulaTokenizer;
using Xunit;

namespace ExcelFormulaParserTests.Expressions
{
    public class BasicExpressionTests
    {
        [Fact]
        public void Test_Formula_Is_1()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("1"));

            TestHelper.AssertExpression(tree, Expression.Number(1));
        }

        [Fact]
        public void Test_Formula_Is_1Expo2()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("1E-2"));

            TestHelper.AssertExpression(tree, Expression.Number(0.01));
        }

        [Fact]
        public void Test_Formula_Is_10_Percent()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("10%"));

            TestHelper.AssertExpression(tree, Expression.Number(0.1));
        }

        [Fact]
        public void Test_Formula_Is_Negative_1()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("-1"));

            TestHelper.AssertExpression(
                tree,
                Expression.UnaryExpression(UnaryOperatorType.Negate, Expression.Number(1)));
        }

        [Fact]
        public void Test_Formula_Is_Negative_Negative_Negative_1()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("---1"));

            TestHelper.AssertExpression(tree,
             Expression.UnaryExpression(
               UnaryOperatorType.Negate,
               Expression.UnaryExpression(
                 UnaryOperatorType.Negate,
                 Expression.UnaryExpression(
                   UnaryOperatorType.Negate,
                   Expression.Number(1)
                 )
               )
             )
           );
        }

        [Fact]
        public void Test_Formula_Is_String__abc__()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("\"abc\""));

            TestHelper.AssertExpression(tree, Expression.Text("abc"));
        }

        [Fact]
        public void Test_Formula_Is_TRUE()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("TRUE"));

            TestHelper.AssertExpression(tree, Expression.Logical(true));
        }

        [Fact]
        public void Test_Formula_Is_1_Plus_2()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("1 + 2"));

            TestHelper.AssertExpression(tree,
             Expression.BinaryExpression(
               BinaryOperatorType.Add,
               Expression.Number(1),
               Expression.Number(2)
             )
           );
        }

        [Fact]
        public void Test_Formula_Is_Negative_1_Plus_2()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("-1 + 2"));

            TestHelper.AssertExpression(tree,
             Expression.BinaryExpression(
               BinaryOperatorType.Add,
               Expression.UnaryExpression(
                 UnaryOperatorType.Negate,
                 Expression.Number(1)
               ),
               Expression.Number(2)
             )
           );
        }

        [Fact]
        public void Test_Formula_Is_String__a__And_String__b__()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("\"a\" & \"b\""));

            TestHelper.AssertExpression(tree,
             Expression.BinaryExpression(
               BinaryOperatorType.Concat,
               Expression.Text("a"),
               Expression.Text("b")
             )
           );
        }

        [Fact]
        public void Test_Formula_Is_1_NotEqual_String__b__()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("1 <> \"b\""));

            TestHelper.AssertExpression(tree,
             Expression.BinaryExpression(
               BinaryOperatorType.NotEqualTo,
               Expression.Number(1),
               Expression.Text("b")
             )
           );
        }

    }
}
