using ExcelFormulaParser.FormulaTokenizer;
using Xunit;

namespace ExcelFormulaParserTests.FormulaTokenizer
{
    public class StringExpressionTests
    {
        [Fact]
        public void Cat()
        {
            var formula = "\"cat\"";
            var expected = new[]
            {
                new Token("cat", TokenType.Operand, TokenSubType.Text)
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void Con_And_Cat()
        {
            var formula = "\"con\"&\"cat\"";
            var expected = new[]
            {
                new Token("con", TokenType.Operand, TokenSubType.Text),
                new Token("&", TokenType.OperatorInfix, TokenSubType.Concatenate),
                new Token("cat", TokenType.Operand, TokenSubType.Text),
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void Con_And_Cat_And_Enate()
        {
            var formula = "\"con\"&\"cat\"&\"enate\"";
            var expected = new[]
            {
                new Token("con", TokenType.Operand, TokenSubType.Text),
                new Token("&", TokenType.OperatorInfix, TokenSubType.Concatenate),
                new Token("cat", TokenType.Operand, TokenSubType.Text),
                new Token("&", TokenType.OperatorInfix, TokenSubType.Concatenate),
                new Token("enate", TokenType.Operand, TokenSubType.Text),
            };

            TestHelper.AssertFormula(formula, expected);
        }
    }
}
