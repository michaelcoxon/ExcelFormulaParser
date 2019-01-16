using ExcelFormulaParser.FormulaTokenizer;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace ExcelFormulaParserTests.FormulaTokenizer
{
    public class RangeOperatorTests
    {
        [Fact]
        public void A1ToC1UnionA2ToC2()
        {
            var formula = "A1:C1,A2:C2";
            var expected = new[]
            {
                new Token("A1:C1", TokenType.Operand, TokenSubType.Range),
                new Token(",", TokenType.OperatorInfix, TokenSubType.Union),
                new Token("A2:C2", TokenType.Operand, TokenSubType.Range),
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void A1ToC1IntersectA2ToC2()
        {
            var formula = "A1:C1 A2:C2";
            var expected = new[]
            {
                new Token("A1:C1", TokenType.Operand, TokenSubType.Range),
                new Token(" ", TokenType.OperatorInfix, TokenSubType.Intersect),
                new Token("A2:C2", TokenType.Operand, TokenSubType.Range),
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void A1ToC1IntersectA2ToC2_MultipleSpaces()
        {
            var formula = "A1:C1   A2:C2";
            var expected = new[]
            {
                new Token("A1:C1", TokenType.Operand, TokenSubType.Range),
                new Token(" ", TokenType.OperatorInfix, TokenSubType.Intersect),
                new Token("A2:C2", TokenType.Operand, TokenSubType.Range),
            };

            TestHelper.AssertFormula(formula, expected);
        }
    }
}
