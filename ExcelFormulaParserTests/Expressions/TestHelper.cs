using ExcelFormulaParser.Expressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace ExcelFormulaParserTests.Expressions
{
    public class TestHelper
    {
        public static void AssertExpression(Expression actual, Expression expected)
        {
            Assert.Equal(expected, actual);
        }
    }
}
