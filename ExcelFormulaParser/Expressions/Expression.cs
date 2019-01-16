using EnsureFramework;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelFormulaParser.Expressions
{
    public enum ExpressionType
    {
        BinaryExpression,
        UnaryExpression,
        Logical,
        Text,
        Number,
        Function,
        CellRange,
        Cell
    }

    public enum CellReferenceType
    {
        Absolute,
        Mixed,
        Relative
    }

    public enum UnaryOperatorType
    {
        Negate
    }

    public enum BinaryOperatorType
    {
        Add,
        Concat,
        NotEqualTo,
        Comma,
        Whitespace,
        Divide,
        Subtract,
        Pow,
        Multiply,
        EqualTo,
        LessThanOrEqualTo,
        GreaterThanOrEqualTo,
        GreaterThan,
        LessThan
    }

    public abstract class Expression
    {
        public ExpressionType Type { get; }

        protected Expression(ExpressionType type)
        {
            this.Type = type;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is Expression item))
            {
                return false;
            }

            return this.Type.Equals(item.Type)
                ;
        }

        public override int GetHashCode()
        {
            return this.Type.GetHashCode()
                ;
        }

        public static CellExpression Cell(string key, CellReferenceType refType)
        {
            return new CellExpression(key, refType);
        }

        public static RangeExpression CellRange(Expression leftCell, Expression rightCell)
        {
            if (leftCell == null)
            {
                throw new Exception("Invalid Syntax");
            }
            if (rightCell == null)
            {
                throw new Exception("Invalid Syntax");
            }
            return new RangeExpression(leftCell, rightCell);
        }

        public static FunctionExpression FunctionCall(string name, params Expression[] args)
        {
            return new FunctionExpression(name, args);
        }

        public static NumberExpression Number(double value)
        {
            return new NumberExpression(value);
        }

        public static TextExpression Text(string value)
        {
            return new TextExpression(value);
        }

        public static LogicalExpression Logical(bool value)
        {
            return new LogicalExpression(value);
        }

        public static BinaryExpression BinaryExpression(BinaryOperatorType @operator, Expression left, Expression right)
        {
            if (left == null)
            {
                throw new Exception("Invalid Syntax");
            }
            if (right == null)
            {
                throw new Exception("Invalid Syntax");
            }
            return new BinaryExpression(@operator, left, right);
        }

        public static Expression UnaryExpression(UnaryOperatorType @operator, Expression expression)
        {
            if (expression == null)
            {
                throw new Exception("Invalid Syntax");
            }
            return new UnaryExpression(@operator, expression);
        }
    }

    public sealed class FunctionExpression : Expression
    {
        public string Name { get; }
        public IEnumerable<Expression> Arguments { get; }

        internal FunctionExpression(string name, IEnumerable<Expression> arguments) : base(ExpressionType.Function)
        {
            Ensure.Arg(name, nameof(name)).IsNotNullOrEmpty();
            Ensure.Arg(arguments, nameof(arguments)).IsNotNull();
            this.Name = name;
            this.Arguments = arguments;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is FunctionExpression item))
            {
                return false;
            }

            return base.Equals(item)
                && this.Name.Equals(item.Name)
                && this.Arguments.SequenceEqual(item.Arguments);
        }

        public override int GetHashCode()
        {
            return base.GetHashCode()
                ^ this.Name.GetHashCode()
                ^ this.Arguments.Aggregate(0, (a, c) => a ^ c.GetHashCode());
        }

#if !DEBUG
        public override string ToString()
        {
            return $"{this.Name}({string.Join(", ", this.Arguments)})";
        }
#endif
    }

    public sealed class CellExpression : Expression
    {
        public CellReferenceType ReferenceType { get; }
        public string Key { get; }

        internal CellExpression(string key, CellReferenceType referenceType) : base(ExpressionType.Cell)
        {
            this.Key = key;
            this.ReferenceType = referenceType;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is CellExpression item))
            {
                return false;
            }

            return base.Equals(item)
                && this.ReferenceType.Equals(item.ReferenceType)
                && this.Key.Equals(item.Key);
        }

        public override int GetHashCode()
        {
            return base.GetHashCode()
                ^ this.ReferenceType.GetHashCode()
                ^ this.Key.GetHashCode();
        }

#if !DEBUG

        public override string ToString()
        {
            return this.Key.ToString();
        }
#endif
    }

    public sealed class NumberExpression : Expression
    {
        public double Value { get; }

        internal NumberExpression(double value) : base(ExpressionType.Number)
        {
            this.Value = value;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is NumberExpression item))
            {
                return false;
            }

            return base.Equals(item)
                && this.Value.Equals(item.Value)
                ;
        }

        public override int GetHashCode()
        {
            return base.GetHashCode()
                ^ this.Value.GetHashCode()
                ;
        }

#if !DEBUG
        public override string ToString()
        {
            return this.Value.ToString();
        }
#endif
    }

    public sealed class TextExpression : Expression
    {
        public string Value { get; }

        internal TextExpression(string value) : base(ExpressionType.Text)
        {
            this.Value = value;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is TextExpression item))
            {
                return false;
            }

            return base.Equals(item)
                && this.Value.Equals(item.Value)
                ;
        }

        public override int GetHashCode()
        {
            return base.GetHashCode()
                ^ this.Value.GetHashCode()
                ;
        }

#if !DEBUG
        public override string ToString()
        {
            return $"\"{this.Value}\"";
        }
#endif
    }

    public sealed class LogicalExpression : Expression
    {
        public bool Value { get; }

        internal LogicalExpression(bool value) : base(ExpressionType.Logical)
        {
            this.Value = value;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is LogicalExpression item))
            {
                return false;
            }

            return base.Equals(item)
                && this.Value.Equals(item.Value)
                ;
        }

        public override int GetHashCode()
        {
            return base.GetHashCode()
                ^ this.Value.GetHashCode()
                ;
        }

#if !DEBUG
        public override string ToString()
        {
            return this.Value ? "TRUE" : "FALSE";
        }
#endif
    }

    public sealed class BinaryExpression : Expression
    {
        public Expression Left;
        public Expression Right;
        public BinaryOperatorType Operator;

        internal BinaryExpression(BinaryOperatorType @operator, Expression left, Expression right) : base(ExpressionType.BinaryExpression)
        {
            this.Left = left;
            this.Right = right;
            this.Operator = @operator;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is BinaryExpression item))
            {
                return false;
            }

            return base.Equals(item)
                && this.Left.Equals(item.Left)
                && this.Right.Equals(item.Right)
                && this.Operator.Equals(item.Operator)
                ;
        }

        public override int GetHashCode()
        {
            return base.GetHashCode()
                ^ this.Left.GetHashCode()
                ^ this.Right.GetHashCode()
                ^ this.Operator.GetHashCode()
                ;
        }

#if !DEBUG
        public override string ToString()
        {
            return $"{this.Left} {this.Operator} {this.Right}";
        }
#endif
    }

    public sealed class UnaryExpression : Expression
    {
        public Expression Operand;
        public UnaryOperatorType Operator;

        internal UnaryExpression(UnaryOperatorType @operator, Expression operand) : base(ExpressionType.UnaryExpression)
        {
            this.Operator = @operator;
            this.Operand = operand;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is UnaryExpression item))
            {
                return false;
            }

            return base.Equals(item)
                && this.Operand.Equals(item.Operand)
                && this.Operator.Equals(item.Operator)
                ;
        }

        public override int GetHashCode()
        {
            return base.GetHashCode()
                ^ this.Operand.GetHashCode()
                ^ this.Operator.GetHashCode()
                ;
        }
    }

    public sealed class RangeExpression : Expression
    {
        public Expression Left;
        public Expression Right;

        internal RangeExpression(Expression left, Expression right) : base(ExpressionType.CellRange)
        {
            this.Left = left;
            this.Right = right;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is RangeExpression item))
            {
                return false;
            }

            return base.Equals(item)
                && this.Left.Equals(item.Left)
                && this.Right.Equals(item.Right)
                ;
        }

        public override int GetHashCode()
        {
            return base.GetHashCode()
                ^ this.Left.GetHashCode()
                ^ this.Right.GetHashCode()
                ;
        }

#if !DEBUG
        public override string ToString()
        {
            return $"{this.Left}:{this.Right}";
        }
#endif
    }
}
