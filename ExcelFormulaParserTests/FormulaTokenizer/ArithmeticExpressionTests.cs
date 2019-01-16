using ExcelFormulaParser.FormulaTokenizer;
using Xunit;

namespace ExcelFormulaParserTests.FormulaTokenizer
{
    public class ArithmeticExpressionTests
    {
        [Fact]
        public void One()
        {
            var formula = "1";
            var expected = new[]
            {
                new Token("1", TokenType.Operand, TokenSubType.Number)
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void OnePointFive()
        {
            var formula = "1.5";
            var expected = new[]
            {
                new Token("1.5", TokenType.Operand, TokenSubType.Number)
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void ElevenPointFiveFive()
        {
            var formula = "11.55";
            var expected = new[]
            {
                new Token("11.55", TokenType.Operand, TokenSubType.Number)
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void OneExpoNegativeOne()
        {
            var formula = "1E-1";
            var expected = new[]
            {
                new Token("1E-1", TokenType.Operand, TokenSubType.Number)
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void OnePointFiveExpoNegativeTen()
        {
            var formula = "1.5E-10";
            var expected = new[]
            {
                new Token("1.5E-10", TokenType.Operand, TokenSubType.Number)
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void ElevenPointFiveFiveExpoOneHundred()
        {
            var formula = "11.55E+100";
            var expected = new[]
            {
                new Token("11.55E+100", TokenType.Operand, TokenSubType.Number)
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void OnePlusTwo()
        {
            var formula = "1 + 2";
            var expected = new[]
            {
                new Token("1", TokenType.Operand, TokenSubType.Number),
                new Token("+", TokenType.OperatorInfix, TokenSubType.Math),
                new Token("2", TokenType.Operand, TokenSubType.Number),
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void OnePlusTwoThenMinusThree()
        {
            var formula = "(1 + 2) - 3";
            var expected = new[]
            {
                new Token("", TokenType.SubExpression, TokenSubType.Start),
                new Token("1", TokenType.Operand, TokenSubType.Number),
                new Token("+", TokenType.OperatorInfix, TokenSubType.Math),
                new Token("2", TokenType.Operand, TokenSubType.Number),
                new Token("", TokenType.SubExpression, TokenSubType.Stop),
                new Token("-", TokenType.OperatorInfix, TokenSubType.Math),
                new Token("3", TokenType.Operand, TokenSubType.Number),
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void NegativeOne()
        {
            var formula = "-1";
            var expected = new[]
            {
                new Token("-", TokenType.OperatorPrefix),
                new Token("1", TokenType.Operand, TokenSubType.Number),
            };

            TestHelper.AssertFormula(formula, expected);
        }


        [Fact]
        public void PositiveOne()
        {
            var formula = "+1";
            var expected = new[]
            {
                new Token("1", TokenType.Operand, TokenSubType.Number),
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void OnePercent()
        {
            var formula = "1%";
            var expected = new[]
            {
                new Token("1", TokenType.Operand, TokenSubType.Number),
                new Token("%", TokenType.OperatorPostfix),
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void OnePlusTwoThenNegateIt()
        {
            var formula = "-(1 + 2)";
            var expected = new[]
            {
                new Token("-", TokenType.OperatorPrefix),
                new Token("", TokenType.SubExpression, TokenSubType.Start),
                new Token("1", TokenType.Operand, TokenSubType.Number),
                new Token("+", TokenType.OperatorInfix, TokenSubType.Math),
                new Token("2", TokenType.Operand, TokenSubType.Number),
                new Token("", TokenType.SubExpression, TokenSubType.Stop),
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void OnePlusTwoThenPercent()
        {
            var formula = "(1 + 2)%";
            var expected = new[]
            {
                new Token("", TokenType.SubExpression, TokenSubType.Start),
                new Token("1", TokenType.Operand, TokenSubType.Number),
                new Token("+", TokenType.OperatorInfix, TokenSubType.Math),
                new Token("2", TokenType.Operand, TokenSubType.Number),
                new Token("", TokenType.SubExpression, TokenSubType.Stop),
                new Token("%", TokenType.OperatorPostfix),
            };

            TestHelper.AssertFormula(formula, expected);
        }
    }
}
