using ExcelFormulaParser.FormulaTokenizer;
using System.Linq;
using Xunit;

namespace ExcelFormulaParserTests.FormulaTokenizer
{
    public class TestHelper
    {
        public static void AssertFormula(string formula, Token[] expected, TokenizerOptions options = null)
        {
            var result = Tokenizer.Tokenize(formula, options);
            AssertTokens(result, expected);
        }

        private static void AssertTokens(Token[] actual, Token[] expected)
        {
            Assert.Equal(expected.Length, actual.Length);

            for (var i = 0; i < actual.Length; i++)
            {
                Assert.Equal(expected[i].Value, actual[i].Value);
                Assert.Equal(expected[i].Type, actual[i].Type);
                Assert.Equal(expected[i].SubType, actual[i].SubType);
            }
        }
    }
}
