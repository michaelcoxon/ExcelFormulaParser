using ExcelFormulaParser.FormulaTokenizer;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace ExcelFormulaParserTests.FormulaTokenizer
{
    public class ComparisonExpressionTests
    {
        [Fact]
        public void OneGreaterThanTwo()
        {
            var formula = "1 > 2";
            var expected = new[]
            {
                new Token("1", TokenType.Operand, TokenSubType.Number),
                new Token(">", TokenType.OperatorInfix, TokenSubType.Logical),
                new Token("2", TokenType.Operand, TokenSubType.Number),
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void OneGreaterThanEqualToTwo()
        {
            var formula = "1 >= 2";
            var expected = new[]
            {
                new Token("1", TokenType.Operand, TokenSubType.Number),
                new Token(">=", TokenType.OperatorInfix, TokenSubType.Logical),
                new Token("2", TokenType.Operand, TokenSubType.Number),
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void OneEqualToTwo()
        {
            var formula = "1 = 2";
            var expected = new[]
            {
                new Token("1", TokenType.Operand, TokenSubType.Number),
                new Token("=", TokenType.OperatorInfix, TokenSubType.Logical),
                new Token("2", TokenType.Operand, TokenSubType.Number),
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void OneNotEqualToTwo()
        {
            var formula = "1 <> 2";
            var expected = new[]
            {
                new Token("1", TokenType.Operand, TokenSubType.Number),
                new Token("<>", TokenType.OperatorInfix, TokenSubType.Logical),
                new Token("2", TokenType.Operand, TokenSubType.Number),
            };

            TestHelper.AssertFormula(formula, expected);
        }
    }
}
