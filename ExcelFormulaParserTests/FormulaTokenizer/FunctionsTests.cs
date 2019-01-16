using ExcelFormulaParser.FormulaTokenizer;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace ExcelFormulaParserTests.FormulaTokenizer
{
    public class FunctionsTests
    {
        [Fact]
        public void Today()
        {
            var formula = "TODAY()";
            var expected = new[]
            {
                new Token("TODAY", TokenType.Function, TokenSubType.Start),
                new Token("", TokenType.Function, TokenSubType.Stop),
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void SumOne()
        {
            var formula = "SUM(1)";
            var expected = new[]
            {
                new Token("SUM", TokenType.Function, TokenSubType.Start),
                new Token("1", TokenType.Operand, TokenSubType.Number),
                new Token("", TokenType.Function, TokenSubType.Stop),
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void AtSumOne()
        {
            var formula = "@SUM(1)";
            var expected = new[]
            {
                new Token("SUM", TokenType.Function, TokenSubType.Start),
                new Token("1", TokenType.Operand, TokenSubType.Number),
                new Token("", TokenType.Function, TokenSubType.Stop),
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void SumOneAndTwo()
        {
            var formula = "SUM(1, 2)";
            var expected = new[]
            {
                new Token("SUM", TokenType.Function, TokenSubType.Start),
                new Token("1", TokenType.Operand, TokenSubType.Number),
                new Token(",", TokenType.Argument),
                new Token("2", TokenType.Operand, TokenSubType.Number),
                new Token("", TokenType.Function, TokenSubType.Stop),
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void SumOneAndSumTwoAndThree()
        {
            var formula = "SUM(1, SUM(2, 3))";
            var expected = new[]
            {
                new Token("SUM", TokenType.Function, TokenSubType.Start),
                new Token("1", TokenType.Operand, TokenSubType.Number),
                new Token(",", TokenType.Argument),
                new Token("SUM", TokenType.Function, TokenSubType.Start),
                new Token("2", TokenType.Operand, TokenSubType.Number),
                new Token(",", TokenType.Argument),
                new Token("3", TokenType.Operand, TokenSubType.Number),
                new Token("", TokenType.Function, TokenSubType.Stop),
                new Token("", TokenType.Function, TokenSubType.Stop),
            };

            TestHelper.AssertFormula(formula, expected);
        }
    }
}
