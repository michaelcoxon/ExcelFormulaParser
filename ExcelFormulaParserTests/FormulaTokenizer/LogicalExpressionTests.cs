using ExcelFormulaParser.FormulaTokenizer;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace ExcelFormulaParserTests.FormulaTokenizer
{
    public class LogicalExpressionTests
    {
        [Fact]
        public void True()
        {
            var formula = "TRUE";
            var expected = new[]
            {
                new Token("TRUE", TokenType.Operand, TokenSubType.Logical)
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void False()
        {
            var formula = "FALSE";
            var expected = new[]
            {
                new Token("FALSE", TokenType.Operand, TokenSubType.Logical)
            };

            TestHelper.AssertFormula(formula, expected);
        }
    }
}
