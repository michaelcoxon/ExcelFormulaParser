using ExcelFormulaParser.Tree;

namespace ExcelFormulaParserTests
{
    public class BasicExpressions
    {
        public void Test_Formula_Is_1()
        {
            var tree = new TreeBuilder(tokenize('1'));

            deepStrictEqual(tree, NodeBuilder.number(1));
        }

        public void Test_Formula_Is_1Expo2()
        {
            convarst tree = new TreeBuilder(tokenize('1E-2'));

            deepStrictEqual(tree, NodeBuilder.number(0.01));
        }

        public void Test_Formula_Is_10_Percent()
        {
            var tree = new TreeBuilder(tokenize('10%'));

            deepStrictEqual(tree, NodeBuilder.number(0.1));
        }

        public void Test_Formula_Is_Negative_1()
        {
            var tree = new TreeBuilder(tokenize('-1'));

            deepStrictEqual(tree,
              NodeBuilder.unaryExpression(
                '-',
                NodeBuilder.number(1)
              )
            );
        }

        public void Test_Formula_Is_Negative_Negative_Negative_1()
        {
            var tree = new TreeBuilder(tokenize('---1'));

            deepStrictEqual(tree,
              NodeBuilder.unaryExpression(
                '-',
                NodeBuilder.unaryExpression(
                  '-',
                  NodeBuilder.unaryExpression(
                    '-',
                    NodeBuilder.number(1)
                  )
                )
              )
            );
        }

        public void Test_Formula_Is_String__abc__()
        {
            var tree = new TreeBuilder(tokenize('"abc"'));

            deepStrictEqual(tree, NodeBuilder.text('abc'));
        }

        public void Test_Formula_Is_TRUE()
        {
            var tree = new TreeBuilder(tokenize('TRUE'));

            deepStrictEqual(tree, NodeBuilder.logical(true));
        }

        public void Test_Formula_Is_1_Plus_2()
        {
            var tree = new TreeBuilder(tokenize('1 + 2'));

            deepStrictEqual(tree,
              NodeBuilder.binaryExpression(
                '+',
                NodeBuilder.number(1),
                NodeBuilder.number(2)
              )
            );
        }

        public void Test_Formula_Is_Negative_1_Plus_2()
        {
            var tree = new TreeBuilder(tokenize('-1 + 2'));

            deepStrictEqual(tree,
              NodeBuilder.binaryExpression(
                '+',
                NodeBuilder.unaryExpression(
                  '-',
                  NodeBuilder.number(1)
                ),
                NodeBuilder.number(2)
              )
            );
        }

        public void Test_Formula_Is_String__a__And_String__b__()
        {
            var tree = new TreeBuilder(tokenize('"a" & "b"'));

            deepStrictEqual(tree,
              NodeBuilder.binaryExpression(
                '&',
                NodeBuilder.text('a'),
                NodeBuilder.text('b')
              )
            );
        }

        public void Test_Formula_Is_1_NotEqual_String__b__()
        {
            var tree = new TreeBuilder(tokenize('1 <> "b"'));

            deepStrictEqual(tree,
              NodeBuilder.binaryExpression(
                '<>',
                NodeBuilder.number(1),
                NodeBuilder.text('b')
              )
            );
        }
    }
}
