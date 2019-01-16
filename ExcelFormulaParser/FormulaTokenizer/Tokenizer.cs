// https://github.com/psalaets/excel-formula-tokenizer/blob/master/index.js

using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelFormulaParser.FormulaTokenizer
{
    public class TokenizerOptions
    {
    }

    public class Tokenizer
    {
        public static Token[] Tokenize(string formula, TokenizerOptions options = null)
        {
            options = options ?? new TokenizerOptions();

            var language = new Language();

            var tokens = new TokenCollection();
            var tokenStack = new TokenStack();

            var offset = 0;

            string currentChar()
            {
                if (offset < formula.Length)
                {
                    return formula.Substring(offset, 1);
                }
                else
                {
                    return string.Empty;
                }
            }

            string doubleChar()
            {
                if (offset < formula.Length)
                {
                    if (offset + 1 < formula.Length)
                    {
                        return formula.Substring(offset, 2);
                    }
                    return currentChar();
                }

                return string.Empty;
            }

            string nextChar()
            {
                if (offset + 1 < formula.Length)
                {
                    return formula.Substring(offset + 1, 1);
                }
                else
                {
                    return string.Empty;
                }
            }

            bool EOF()
            {
                return offset >= formula.Length;
            }

            bool isPreviousNonDigitBlank()
            {
                var offsetCopy = offset;
                if (offsetCopy == 0)
                {
                    return true;
                }

                while (offsetCopy > 0)
                {
                    if (!new Regex(@"\d").IsMatch(formula[offsetCopy].ToString()))
                    {
                        return new Regex(@"\s").IsMatch(formula[offsetCopy].ToString());
                    }

                    offsetCopy -= 1;
                }
                return false;
            };

            bool isNextNonDigitTheRangeOperator()
            {
                var offsetCopy = offset;

                while (offsetCopy < formula.Length)
                {
                    if (!new Regex(@"\d").IsMatch(formula[offsetCopy].ToString()))
                    {
                        return new Regex(@":").IsMatch(formula[offsetCopy].ToString());
                    }

                    offsetCopy += 1;
                }
                return false;
            };

            var strToken = string.Empty;

            var inString = false;
            var inPath = false;
            var inRange = false;
            var inError = false;
            var inNumeric = false;

            while (formula.Length > 0)
            {
                if (formula.Substring(0, 1) == " ")
                {
                    formula = formula.Substring(1);
                }
                else
                {
                    if (formula.Substring(0, 1) == "=")
                    {
                        formula = formula.Substring(1);
                    }
                    break;
                }
            }

            var regexSN = new Regex(@"^[1-9]{1}(\.[0-9]+)?E{1}$");

            while (!EOF())
            {

                // state-dependent character evaluation (order is important)

                // double-quoted strings
                // embeds are doubled
                // end marks token

                if (inString)
                {
                    if (currentChar() == "\"")
                    {
                        if (nextChar() == "\"")
                        {
                            strToken += "\"";
                            offset += 1;
                        }
                        else
                        {
                            inString = false;
                            tokens.Add(strToken, TokenType.Operand, TokenSubType.Text);
                            strToken = string.Empty;
                        }
                    }
                    else
                    {
                        strToken += currentChar();
                    }
                    offset += 1;
                    continue;
                }

                // single-quoted strings (links)
                // embeds are double
                // end does not mark a token

                if (inPath)
                {
                    if (currentChar() == "'")
                    {
                        if (nextChar() == "'")
                        {
                            strToken += "'";
                            offset += 1;
                        }
                        else
                        {
                            inPath = false;
                        }
                    }
                    else
                    {
                        strToken += currentChar();
                    }
                    offset += 1;
                    continue;
                }

                // bracked strings (range offset or linked workbook name)
                // no embeds (changed to "()" by Excel)
                // end does not mark a token

                if (inRange)
                {
                    if (currentChar() == "]")
                    {
                        inRange = false;
                    }
                    strToken += currentChar();
                    offset += 1;
                    continue;
                }

                // error values
                // end marks a token, determined from absolute list of values

                if (inError)
                {
                    strToken += currentChar();
                    offset += 1;
                    if (new[] { "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A" }.Contains(strToken))
                    {
                        inError = false;
                        tokens.Add(strToken, TokenType.Operand, TokenSubType.Error);
                        strToken = string.Empty;
                    }
                    continue;
                }

                if (inNumeric)
                {
                    if (new[] { language.DecimalSeparator, "E", "+", "-" }.Contains(currentChar()) || new Regex(@"\d").IsMatch(currentChar()))
                    {
                        inNumeric = true;
                        strToken += currentChar();

                        offset += 1;
                        continue;
                    }
                    else
                    {
                        inNumeric = false;
                        tokens.Add(strToken, TokenType.Operand, TokenSubType.Number);
                        strToken = string.Empty;
                    }
                }

                // scientific notation check

                if ("+-".Contains(currentChar()))
                {
                    if (strToken.Length > 1)
                    {
                        if (regexSN.IsMatch(strToken))
                        {
                            strToken += currentChar();
                            offset += 1;
                            continue;
                        }
                    }
                }

                // independent character evaulation (order not important)

                // function, subexpression, array parameters

                if (currentChar() == language.ArgumentSeparator)
                {
                    if (strToken.Length > 0)
                    {
                        tokens.Add(strToken, TokenType.Operand);
                        strToken = string.Empty;
                    }

                    if (tokenStack.Type() == TokenType.Function)
                    {
                        tokens.Add(",", TokenType.Argument);

                        offset += 1;
                        continue;
                    }
                }

                if (currentChar() == ",")
                {
                    if (strToken.Length > 0)
                    {
                        tokens.Add(strToken, TokenType.Operand);
                        strToken = string.Empty;
                    }

                    tokens.Add(currentChar(), TokenType.OperatorInfix, TokenSubType.Union);

                    offset += 1;
                    continue;
                }

                // establish state-dependent character evaluations

                if (new Regex(@"\d").IsMatch(currentChar()) && isPreviousNonDigitBlank() && !isNextNonDigitTheRangeOperator())
                {
                    inNumeric = true;
                    strToken += currentChar();
                    offset += 1;
                    continue;
                }

                if (currentChar() == "\"")
                {
                    if (strToken.Length > 0)
                    {
                        // not expected
                        tokens.Add(strToken, TokenType.Unknown);
                        strToken = string.Empty;
                    }
                    inString = true;
                    offset += 1;
                    continue;
                }

                if (currentChar() == "'")
                {
                    if (strToken.Length > 0)
                    {
                        // not expected
                        tokens.Add(strToken, TokenType.Unknown);
                        strToken = string.Empty;
                    }
                    inPath = true;
                    offset += 1;
                    continue;
                }

                if (currentChar() == "[")
                {
                    inRange = true;
                    strToken += currentChar();
                    offset += 1;
                    continue;
                }

                if (currentChar() == "#")
                {
                    if (strToken.Length > 0)
                    {
                        // not expected
                        tokens.Add(strToken, TokenType.Unknown);
                        strToken = string.Empty;
                    }
                    inError = true;
                    strToken += currentChar();
                    offset += 1;
                    continue;
                }

                // mark start and end of arrays and array rows

                if (currentChar() == "{")
                {
                    if (strToken.Length > 0)
                    {
                        // not expected
                        tokens.Add(strToken, TokenType.Unknown);
                        strToken = string.Empty;
                    }
                    tokenStack.Push(tokens.Add("ARRAY", TokenType.Function, TokenSubType.Start));
                    tokenStack.Push(tokens.Add("ARRAYROW", TokenType.Function, TokenSubType.Start));
                    offset += 1;
                    continue;
                }

                if (currentChar() == ";")
                {
                    if (strToken.Length > 0)
                    {
                        tokens.Add(strToken, TokenType.Operand);
                        strToken = string.Empty;
                    }
                    tokens.AddRef(tokenStack.Pop());
                    tokens.Add(",", TokenType.Argument);
                    tokenStack.Push(tokens.Add("ARRAYROW", TokenType.Function, TokenSubType.Start));
                    offset += 1;
                    continue;
                }

                if (currentChar() == "}")
                {
                    if (strToken.Length > 0)
                    {
                        tokens.Add(strToken, TokenType.Operand);
                        strToken = string.Empty;
                    }
                    tokens.AddRef(tokenStack.Pop());
                    tokens.AddRef(tokenStack.Pop());
                    offset += 1;
                    continue;
                }

                // trim white-space

                if (currentChar() == " ")
                {
                    if (strToken.Length > 0)
                    {
                        tokens.Add(strToken, TokenType.Operand);
                        strToken = string.Empty;
                    }
                    tokens.Add(currentChar(), TokenType.WhiteSpace);
                    offset += 1;
                    while ((currentChar() == " ") && (!EOF()))
                    {
                        offset += 1;
                    }
                    continue;
                }

                // multi-character comparators

                if (new[] { ">=", "<=", "<>" }.Contains(doubleChar()))
                {
                    if (strToken.Length > 0)
                    {
                        tokens.Add(strToken, TokenType.Operand);
                        strToken = string.Empty;
                    }
                    tokens.Add(doubleChar(), TokenType.OperatorInfix, TokenSubType.Logical);
                    offset += 2;
                    continue;
                }

                // standard infix operators

                if ("+-*/^&=><".Contains(currentChar()))
                {
                    if (strToken.Length > 0)
                    {
                        tokens.Add(strToken, TokenType.Operand);
                        strToken = string.Empty;
                    }
                    tokens.Add(currentChar(), TokenType.OperatorInfix);
                    offset += 1;
                    continue;
                }

                // standard postfix operators

                if ("%".Contains(currentChar()))
                {
                    if (strToken.Length > 0)
                    {
                        tokens.Add(strToken, TokenType.Operand);
                        strToken = string.Empty;
                    }
                    tokens.Add(currentChar(), TokenType.OperatorPostfix);
                    offset += 1;
                    continue;
                }

                // start subexpression or function

                if (currentChar() == "(")
                {
                    if (strToken.Length > 0)
                    {
                        tokenStack.Push(tokens.Add(strToken, TokenType.Function, TokenSubType.Start));
                        strToken = string.Empty;
                    }
                    else
                    {
                        tokenStack.Push(tokens.Add("", TokenType.SubExpression, TokenSubType.Start));
                    }
                    offset += 1;
                    continue;
                }

                // stop subexpression

                if (currentChar() == ")")
                {
                    if (strToken.Length > 0)
                    {
                        tokens.Add(strToken, TokenType.Operand);
                        strToken = string.Empty;
                    }
                    tokens.AddRef(tokenStack.Pop());
                    offset += 1;
                    continue;
                }

                // token accumulation

                strToken += currentChar();
                offset += 1;

            }

            // dump remaining accumulation

            if (strToken.Length > 0)
            {
                tokens.Add(strToken, TokenType.Operand);
            }

            // move all tokens to a new collection, excluding all unnecessary white-space tokens

            var tokens2 = new TokenCollection();

            while (tokens.MoveNext())
            {
                var token = tokens.Current();

                if (token.Type == TokenType.WhiteSpace)
                {
                    if ((tokens.BOF()) || (tokens.EOF()))
                    {
                        // no-op
                    }
                    else if (!(
                             ((tokens.Previous().Type == TokenType.Function) && (tokens.Previous().SubType == TokenSubType.Stop)) ||
                             ((tokens.Previous().Type == TokenType.SubExpression) && (tokens.Previous().SubType == TokenSubType.Stop)) ||
                             (tokens.Previous().Type == TokenType.Operand)
                            )
                          )
                    {
                        // no-op
                    }
                    else if (!(
                               ((tokens.Next().Type == TokenType.Function) && (tokens.Next().SubType == TokenSubType.Start)) ||
                               ((tokens.Next().Type == TokenType.SubExpression) && (tokens.Next().SubType == TokenSubType.Start)) ||
                               (tokens.Next().Type == TokenType.Operand)
                               )
                             )
                    {
                        // no-op
                    }
                    else
                    {
                        tokens2.Add(token.Value, TokenType.OperatorInfix, TokenSubType.Intersect);
                    }
                    continue;
                }

                tokens2.AddRef(token);

            }

            // switch infix "-" operator to prefix when appropriate, switch infix "+" operator to noop when appropriate, identify operand
            // and infix-operator subtypes, pull "@" from in front of function names

            while (tokens2.MoveNext())
            {
                var token = tokens2.Current();

                if ((token.Type == TokenType.OperatorInfix) && (token.Value == "-"))
                {
                    if (tokens2.BOF())
                    {
                        token.Type = TokenType.OperatorPrefix;
                    }
                    else if (
                           ((tokens2.Previous().Type == TokenType.Function) && (tokens2.Previous().SubType == TokenSubType.Stop)) ||
                           ((tokens2.Previous().Type == TokenType.SubExpression) && (tokens2.Previous().SubType == TokenSubType.Stop)) ||
                           (tokens2.Previous().Type == TokenType.OperatorPostfix) ||
                           (tokens2.Previous().Type == TokenType.Operand)
                         )
                    {
                        token.SubType = TokenSubType.Math;
                    }
                    else
                    {
                        token.Type = TokenType.OperatorPrefix;
                    }
                    continue;
                }

                if ((token.Type == TokenType.OperatorInfix) && (token.Value == "+"))
                {
                    if (tokens2.BOF())
                    {
                        token.Type = TokenType.Noop;
                    }
                    else if (
                           ((tokens2.Previous().Type == TokenType.Function) && (tokens2.Previous().SubType == TokenSubType.Stop)) ||
                           ((tokens2.Previous().Type == TokenType.SubExpression) && (tokens2.Previous().SubType == TokenSubType.Stop)) ||
                           (tokens2.Previous().Type == TokenType.OperatorPostfix) ||
                           (tokens2.Previous().Type == TokenType.Operand)
                         )
                    {
                        token.SubType = TokenSubType.Math;
                    }
                    else
                    {
                        token.Type = TokenType.Noop;
                    }
                    continue;
                }

                if ((token.Type == TokenType.OperatorInfix) && (token.SubType == null))
                {
                    if ("<>=".Contains(token.Value.Substring(0, 1)))
                    {
                        token.SubType = TokenSubType.Logical;
                    }
                    else if (token.Value == "&")
                    {
                        token.SubType = TokenSubType.Concatenate;
                    }
                    else
                    {
                        token.SubType = TokenSubType.Math;
                    }
                    continue;
                }

                if ((token.Type == TokenType.Operand) && (token.SubType == null))
                {
                    if (!double.TryParse(language.ReformatNumberForJsParsing(token.Value), out var numberValue))
                    {
                        if (token.Value == language.True)
                        {
                            token.SubType = TokenSubType.Logical;
                            token.Value = "TRUE";
                        }
                        else if (token.Value == language.False)
                        {
                            token.SubType = TokenSubType.Logical;
                            token.Value = "FALSE";
                        }
                        else
                        {
                            token.SubType = TokenSubType.Range;
                        }
                    }
                    else
                    {
                        token.SubType = TokenSubType.Number;
                        token.Value = language.ReformatNumberForJsParsing(token.Value);
                    }
                    continue;
                }

                if (token.Type == TokenType.Function)
                {
                    if (token.Value.Length > 0 && token.Value.Substring(0, 1) == "@")
                    {
                        token.Value = token.Value.Substring(1);
                    }
                    continue;
                }

            }

            tokens2.Reset();

            // move all tokens to a new collection, excluding all noops

            tokens = new TokenCollection();

            while (tokens2.MoveNext())
            {
                if (tokens2.Current().Type != TokenType.Noop)
                {
                    tokens.AddRef(tokens2.Current());
                }
            }

            tokens.Reset();

            return tokens.ToArray();
        }
    }
}
