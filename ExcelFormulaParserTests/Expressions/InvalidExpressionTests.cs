using ExcelFormulaParser.Expressions;
using ExcelFormulaParser.FormulaTokenizer;
using Xunit;

namespace ExcelFormulaParserTests.Expressions
{
    public class InvalidExpressionTests
    {
        [Fact]
        public void Sum()
        {
            var ex = Assert.Throws<InvalidSyntaxException>(() =>
            {
                var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("SUM("));
            });

            Assert.Equal(4, ex.Position);
        }

        [Fact]
        public void Add()
        {
            var ex = Assert.Throws<InvalidSyntaxException>(() =>
            {
                var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("+"));
            });

            Assert.Equal(0, ex.Position);
        }

        [Fact]
        public void SumCommaComma()
        {
            var ex = Assert.Throws<InvalidSyntaxException>(() =>
            {
                var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("SUM(,,"));
            });

            Assert.Equal(6, ex.Position);
        }

        [Fact]
        public void GreaterThan()
        {
            var ex = Assert.Throws<InvalidSyntaxException>(() =>
            {
                var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize(">"));
            });

            Assert.Equal(1, ex.Position);
        }

        [Fact]
        public void AGreaterThan()
        {
            var ex = Assert.Throws<InvalidSyntaxException>(() =>
            {
                var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("A >"));
            });

            Assert.Equal(-1, ex.Position);
        }

        [Fact]
        public void GreaterThanB()
        {
            var ex = Assert.Throws<InvalidSyntaxException>(() =>
            {
                var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("> B"));
            });

            Assert.Equal(1, ex.Position);
        }
    }
}
