using ExcelFormulaParser.FormulaTokenizer;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace ExcelFormulaParserTests.FormulaTokenizer
{
    public class ArrayTests
    {
        [Fact]
        public void VectorArray()
        {
            var formula = "{1,2,3}";
            var expected = new[]
            {
                new Token("ARRAY", TokenType.Function, TokenSubType.Start),
                new Token("ARRAYROW", TokenType.Function, TokenSubType.Start),
                new Token("1", TokenType.Operand, TokenSubType.Number),
                new Token(",", TokenType.Argument),
                new Token("2", TokenType.Operand, TokenSubType.Number),
                new Token(",", TokenType.Argument),
                new Token("3", TokenType.Operand, TokenSubType.Number),
                new Token("", TokenType.Function, TokenSubType.Stop),
                new Token("", TokenType.Function, TokenSubType.Stop),
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void MatrixArray()
        {
            var formula = "{1;2;3}";
            var expected = new[]
            {
                new Token("ARRAY", TokenType.Function, TokenSubType.Start),
                new Token("ARRAYROW", TokenType.Function, TokenSubType.Start),
                new Token("1", TokenType.Operand, TokenSubType.Number),
                new Token("", TokenType.Function, TokenSubType.Stop),
                new Token(",", TokenType.Argument),
                new Token("ARRAYROW", TokenType.Function, TokenSubType.Start),
                new Token("2", TokenType.Operand, TokenSubType.Number),
                new Token("", TokenType.Function, TokenSubType.Stop),
                new Token(",", TokenType.Argument),
                new Token("ARRAYROW", TokenType.Function, TokenSubType.Start),
                new Token("3", TokenType.Operand, TokenSubType.Number),
                new Token("", TokenType.Function, TokenSubType.Stop),
                new Token("", TokenType.Function, TokenSubType.Stop),
            };

            TestHelper.AssertFormula(formula, expected);
        }
    }
}
