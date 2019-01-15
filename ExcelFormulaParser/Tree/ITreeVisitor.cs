using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFormulaParser.Tree
{
    public interface ITreeVisitor
    {
        void EnterCell(Token node);
        void ExitCell(Token node);
        void EnterCellRange(Token node);
        void ExitCellRange(Token node);
        void EnterFunction(Token node);
        void ExitFunction(Token node);
        void EnterNumber(Token node);
        void ExitNumber(Token node);
        void EnterText(Token node);
        void ExitText(Token node);
        void EnterLogical(Token node);
        void ExitLogical(Token node);
        void EnterBinaryExpression(Token node);
        void ExitBinaryExpression(Token node);
        void EnterUnaryExpression(Token node);
        void ExitUnaryExpression(Token node);
    }
}
