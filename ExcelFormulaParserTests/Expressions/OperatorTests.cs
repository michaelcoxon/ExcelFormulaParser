using ExcelFormulaParser.Expressions;
using ExcelFormulaParser.FormulaTokenizer;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace ExcelFormulaParserTests.Expressions
{
    public class OperatorTests
    {
        [Fact]
        public void Precendence1()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("1 + 2 >= 3 - 4"));

            TestHelper.AssertExpression(
                tree,
                Expression.BinaryExpression(
                    BinaryOperatorType.GreaterThanOrEqualTo,
                    Expression.BinaryExpression(
                        BinaryOperatorType.Add,
                        Expression.Number(1),
                        Expression.Number(2)
                    ),
                    Expression.BinaryExpression(
                        BinaryOperatorType.Subtract,
                        Expression.Number(3),
                        Expression.Number(4)
                    )
                )
            );
        }

        [Fact]
        public void Precendence2()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("1 + 2 & \"a\""));

            TestHelper.AssertExpression(
                tree,
                Expression.BinaryExpression(
                    BinaryOperatorType.Concat,
                    Expression.BinaryExpression(
                        BinaryOperatorType.Add,
                        Expression.Number(1),
                        Expression.Number(2)
                    ),
                    Expression.Text("a")                    
                )
            );
        }

        [Fact]
        public void Precendence3()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("1 + 2 * 3"));

            TestHelper.AssertExpression(
                tree,
                Expression.BinaryExpression(
                    BinaryOperatorType.Add,
                    Expression.Number(1),
                    Expression.BinaryExpression(
                        BinaryOperatorType.Multiply,
                        Expression.Number(2),
                        Expression.Number(3)
                    )
                )
            );
        }

        [Fact]
        public void Precendence4()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("1 * 2 ^ 3"));

            TestHelper.AssertExpression(
                tree,
                Expression.BinaryExpression(
                    BinaryOperatorType.Multiply,
                    Expression.Number(1),
                    Expression.BinaryExpression(
                        BinaryOperatorType.Pow,
                        Expression.Number(2),
                        Expression.Number(3)
                    )
                )
            );
        }

        [Fact]
        public void Precendence5()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("(1 * 2) ^ 3"));

            TestHelper.AssertExpression(
                tree,
                Expression.BinaryExpression(
                    BinaryOperatorType.Pow,                    
                    Expression.BinaryExpression(
                        BinaryOperatorType.Multiply,
                        Expression.Number(1),
                        Expression.Number(2)
                    ),
                    Expression.Number(3)
                )
            );
        }

        [Fact]
        public void Associativity1()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("1 + 2 + 3"));

            TestHelper.AssertExpression(
                tree,
                Expression.BinaryExpression(
                    BinaryOperatorType.Add,
                    Expression.BinaryExpression(
                        BinaryOperatorType.Add,
                        Expression.Number(1),
                        Expression.Number(2)
                    ),
                    Expression.Number(3)                    
                )
            );
        }

        [Fact]
        public void Associativity2()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("1 + (2 + 3)"));

            TestHelper.AssertExpression(
                tree,
                Expression.BinaryExpression(
                    BinaryOperatorType.Add,
                    Expression.Number(1),
                    Expression.BinaryExpression(
                        BinaryOperatorType.Add,
                        Expression.Number(2),
                        Expression.Number(3)
                    )
                )
            );
        }

        [Fact]
        public void Associativity3()
        {
            var tree = ExpressionTreeBuilder.ParseFormula(Tokenizer.Tokenize("1 / 2 / 3"));

            TestHelper.AssertExpression(
                tree,
                Expression.BinaryExpression(
                    BinaryOperatorType.Divide,
                    Expression.BinaryExpression(
                        BinaryOperatorType.Divide,
                        Expression.Number(1),
                        Expression.Number(2)
                    ),
                    Expression.Number(3)
                )
            );
        }
    }
}
