using ExcelFormulaParser.FormulaTokenizer;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace ExcelFormulaParserTests.FormulaTokenizer
{
    public class CellRangesTests
    {
        [Fact]
        public void A1()
        {
            var formula = "A1";
            var expected = new[]
            {
                new Token("A1", TokenType.Operand, TokenSubType.Range)
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void DollarADollar1()
        {
            var formula = "$A$1";
            var expected = new[]
            {
                new Token("$A$1", TokenType.Operand, TokenSubType.Range)
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void ADollar1()
        {
            var formula = "A$1";
            var expected = new[]
            {
                new Token("A$1", TokenType.Operand, TokenSubType.Range)
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void DollarA1()
        {
            var formula = "$A1";
            var expected = new[]
            {
                new Token("$A1", TokenType.Operand, TokenSubType.Range)
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void A10ToA20()
        {
            var formula = "A10:A20";
            var expected = new[]
            {
                new Token("A10:A20", TokenType.Operand, TokenSubType.Range)
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void A1ToC1()
        {
            var formula = "A1:C1";
            var expected = new[]
            {
                new Token("A1:C1", TokenType.Operand, TokenSubType.Range)
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void FiveToFive()
        {
            var formula = "5:5";
            var expected = new[]
            {
                new Token("5:5", TokenType.Operand, TokenSubType.Range)
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void FiveToTen()
        {
            var formula = "5:10";
            var expected = new[]
            {
                new Token("5:10", TokenType.Operand, TokenSubType.Range)
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void HToH()
        {
            var formula = "H:H";
            var expected = new[]
            {
                new Token("H:H", TokenType.Operand, TokenSubType.Range)
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void HToJ()
        {
            var formula = "H:J";
            var expected = new[]
            {
                new Token("H:J", TokenType.Operand, TokenSubType.Range)
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void A10ToE20()
        {
            var formula = "A10:E20";
            var expected = new[]
            {
                new Token("A10:E20", TokenType.Operand, TokenSubType.Range)
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void R1C1()
        {
            var formula = "R1C1";
            var expected = new[]
            {
                new Token("R1C1", TokenType.Operand, TokenSubType.Range)
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void RNegative2C()
        {
            var formula = "R[-2]C";
            var expected = new[]
            {
                new Token("R[-2]C", TokenType.Operand, TokenSubType.Range)
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void RC3()
        {
            var formula = "RC[3]";
            var expected = new[]
            {
                new Token("RC[3]", TokenType.Operand, TokenSubType.Range)
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void R2C2()
        {
            var formula = "R[2]C[2]";
            var expected = new[]
            {
                new Token("R[2]C[2]", TokenType.Operand, TokenSubType.Range)
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void RNegativeOne()
        {
            var formula = "R[-1]";
            var expected = new[]
            {
                new Token("R[-1]", TokenType.Operand, TokenSubType.Range)
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void CNegativeOne()
        {
            var formula = "C[-1]";
            var expected = new[]
            {
                new Token("C[-1]", TokenType.Operand, TokenSubType.Range)
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void R()
        {
            var formula = "R";
            var expected = new[]
            {
                new Token("R", TokenType.Operand, TokenSubType.Range)
            };

            TestHelper.AssertFormula(formula, expected);
        }

        [Fact]
        public void C()
        {
            var formula = "C";
            var expected = new[]
            {
                new Token("C", TokenType.Operand, TokenSubType.Range)
            };

            TestHelper.AssertFormula(formula, expected);
        }
    }
}
