// https://github.com/psalaets/excel-formula-tokenizer/blob/master/index.js

using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelFormulaParser.Tree
{
    public class TokenizerOptions
    {
    }

    public class Tokenizer
    {
        public const string TOK_TYPE_NOOP = "noop";
        public const string TOK_TYPE_OPERAND = "operand";
        public const string TOK_TYPE_FUNCTION = "function";
        public const string TOK_TYPE_SUBEXPR = "subexpression";
        public const string TOK_TYPE_ARGUMENT = "argument";
        public const string TOK_TYPE_OP_PRE = "operator-prefix";
        public const string TOK_TYPE_OP_IN = "operator-infix";
        public const string TOK_TYPE_OP_POST = "operator-postfix";
        public const string TOK_TYPE_WSPACE = "white-space";
        public const string TOK_TYPE_UNKNOWN = "unknown";
        public const string TOK_SUBTYPE_START = "start";
        public const string TOK_SUBTYPE_STOP = "stop";
        public const string TOK_SUBTYPE_TEXT = "text";
        public const string TOK_SUBTYPE_NUMBER = "number";
        public const string TOK_SUBTYPE_LOGICAL = "logical";
        public const string TOK_SUBTYPE_ERROR = "error";
        public const string TOK_SUBTYPE_RANGE = "range";
        public const string TOK_SUBTYPE_MATH = "math";
        public const string TOK_SUBTYPE_CONCAT = "concatenate";
        public const string TOK_SUBTYPE_INTERSECT = "intersect";
        public const string TOK_SUBTYPE_UNION = "union";

        public void tokenize(string formula, TokenizerOptions options = null)
        {
            options = options ?? new TokenizerOptions();

            var language = new Language();

            var tokens = new Tokens();
            var tokenStack = new TokenStack();

            var offset = 0;

            string currentChar  () => formula.Substring(offset, 1);
            string doubleChar  ()=>  formula.Substring(offset, 2);
            string nextChar  () => formula.Substring(offset + 1, 1);
            bool EOF  ()=>  offset >= formula.Length;            

            bool isPreviousNonDigitBlank  ()
            {
                var offsetCopy = offset;
                if (offsetCopy == 0)
                    return true;

                while (offsetCopy > 0)
                {
                    if (!new Regex(@"\d").IsMatch(formula[offsetCopy].ToString())) {
                    return new Regex(@"\s").IsMatch(formula[offsetCopy].ToString());
                    }

                    offsetCopy -= 1;
                }
                return false;
            };

            bool isNextNonDigitTheRangeOperator  ()
            {
                var offsetCopy = offset;

                while (offsetCopy < formula.Length)
                {
                    if (!new Regex(@"\d").IsMatch(formula[offsetCopy].ToString())) {
                        return new Regex(@":").IsMatch(formula[offsetCopy].ToString());
                    }

                    offsetCopy += 1;
                }
                return false;
            };

            Token token = new Token();

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

  while (!EOF()) {

    // state-dependent character evaluation (order is important)

    // double-quoted strings
    // embeds are doubled
    // end marks token

    if (inString) {
      if (currentChar() == "\"") {
        if (nextChar() == "\"") {
          token.Raw += "\"";
          offset += 1;
        } else {
          inString = false;
          tokens.add(token, TOK_TYPE_OPERAND, TOK_SUBTYPE_TEXT);
          token = "";
        }
      } else {
        token.Raw += currentChar();
      }
      offset += 1;
      continue;
    }

    // single-quoted strings (links)
    // embeds are double
    // end does not mark a token

    if (inPath) {
      if (currentChar() == "'") {
        if (nextChar() == "'") {
          token.Raw += "'";
          offset += 1;
        } else {
          inPath = false;
        }
      } else {
        token.Raw += currentChar();
      }
      offset += 1;
      continue;
    }

    // bracked strings (range offset or linked workbook name)
    // no embeds (changed to "()" by Excel)
    // end does not mark a token

    if (inRange) {
      if (currentChar() == "]") {
        inRange = false;
      }
      token.Raw += currentChar();
offset += 1;
      continue;
    }

    // error values
    // end marks a token, determined from absolute list of values

    if (inError) {
      token.Raw += currentChar();
offset += 1;
      if (new[] { "#NULL!","#DIV/0!","#VALUE!","#REF!","#NAME?","#NUM!","#N/A" }.Contains( token )) {
        inError = false;
        tokens.add(token, TOK_TYPE_OPERAND, TOK_SUBTYPE_ERROR);
        token = "";
      }
      continue;
    }

    if (inNumeric) {
      if (new [] { language.DecimalSeparator, "E", "+", "-" }.Contains(currentChar()) || new Regex(@"\d").IsMatch(currentChar())) {
        inNumeric = true;
        token.Raw += currentChar();

offset += 1;
        continue;
      } else {
        inNumeric = false;
        tokens.add(token.Raw, TOK_TYPE_OPERAND, TOK_SUBTYPE_NUMBER);
        token.Raw = "";
      }
    }

    // scientific notation check

    if ("+-".Contains(currentChar())) {
      if (token.Raw.Length > 1) {
        if (regexSN.IsMatch(token.Raw)) {
          token.Raw += currentChar();
offset += 1;
          continue;
        }
      }
    }

    // independent character evaulation (order not important)

    // function, subexpression, array parameters

    if (currentChar() == language.ArgumentSeparator) {
      if (token.Raw.Length > 0) {
        tokens.add(token.Raw, TOK_TYPE_OPERAND);
        token = "";
      }

      if (tokenStack.type() == Tokenizer.TOK_TYPE_FUNCTION) {
        tokens.add(",", TOK_TYPE_ARGUMENT);

        offset += 1;
        continue;
      }
    }

    if (currentChar() == ",") {
      if (token.Raw.Length > 0) {
        tokens.add(token.Raw, TOK_TYPE_OPERAND);
        token = "";
      }

      tokens.add(currentChar(), TOK_TYPE_OP_IN, TOK_SUBTYPE_UNION);

      offset += 1;
      continue;
    }

    // establish state-dependent character evaluations

    if (new Regex(@"\d").IsMatch(currentChar()) && isPreviousNonDigitBlank() && !isNextNonDigitTheRangeOperator()) {
      inNumeric = true;
      token += currentChar();
offset += 1;
      continue;
    }

    if (currentChar() == "\"") {
      if (token.Length > 0) {
        // not expected
        tokens.add(token, TOK_TYPE_UNKNOWN);
        token = "";
      }
      inString = true;
      offset += 1;
      continue;
    }

    if (currentChar() == "'") {
      if (token.Length > 0) {
        // not expected
        tokens.add(token, TOK_TYPE_UNKNOWN);
        token = "";
      }
      inPath = true;
      offset += 1;
      continue;
    }

    if (currentChar() == "[") {
      inRange = true;
      token += currentChar();
offset += 1;
      continue;
    }

    if (currentChar() == "#") {
      if (token.Length > 0) {
        // not expected
        tokens.add(token, TOK_TYPE_UNKNOWN);
        token = "";
      }
      inError = true;
      token += currentChar();
offset += 1;
      continue;
    }

    // mark start and end of arrays and array rows

    if (currentChar() == "{") {
      if (token.Length > 0) {
        // not expected
        tokens.add(token, TOK_TYPE_UNKNOWN);
        token = "";
      }
      tokenStack.push(tokens.add("ARRAY", TOK_TYPE_FUNCTION, TOK_SUBTYPE_START));
      tokenStack.push(tokens.add("ARRAYROW", TOK_TYPE_FUNCTION, TOK_SUBTYPE_START));
      offset += 1;
      continue;
    }

    if (currentChar() == ";") {
      if (token.Length > 0) {
        tokens.add(token, TOK_TYPE_OPERAND);
        token = "";
      }
      tokens.addRef(tokenStack.pop());
      tokens.add(",", TOK_TYPE_ARGUMENT);
      tokenStack.push(tokens.add("ARRAYROW", TOK_TYPE_FUNCTION, TOK_SUBTYPE_START));
      offset += 1;
      continue;
    }

    if (currentChar() == "}") {
      if (token.Length > 0) {
        tokens.add(token, TOK_TYPE_OPERAND);
        token = "";
      }
      tokens.addRef(tokenStack.pop());
      tokens.addRef(tokenStack.pop());
      offset += 1;
      continue;
    }

    // trim white-space

    if (currentChar() == " ") {
      if (token.Length > 0) {
        tokens.add(token, TOK_TYPE_OPERAND);
        token = "";
      }
      tokens.add(currentChar(), TOK_TYPE_WSPACE);
      offset += 1;
      while ((currentChar() == " ") && (!EOF())) {
        offset += 1;
      }
      continue;
    }

    // multi-character comparators

    if (new[] { ">=","<=","<>" }.Contains(doubleChar())) {
      if (token.Length > 0) {
        tokens.add(token, TOK_TYPE_OPERAND);
        token = "";
      }
      tokens.add(doubleChar(), TOK_TYPE_OP_IN, TOK_SUBTYPE_LOGICAL);
      offset += 2;
      continue;
    }

    // standard infix operators

    if ("+-*/^&=><".Contains(currentChar())) {
      if (token.Length > 0) {
        tokens.add(token, TOK_TYPE_OPERAND);
        token = "";
      }
      tokens.add(currentChar(), TOK_TYPE_OP_IN);
      offset += 1;
      continue;
    }

    // standard postfix operators

    if ("%".Contains(currentChar())) {
      if (token.Length > 0) {
        tokens.add(token, TOK_TYPE_OPERAND);
        token = "";
      }
      tokens.add(currentChar(), TOK_TYPE_OP_POST);
      offset += 1;
      continue;
    }

    // start subexpression or function

    if (currentChar() == "(") {
      if (token.Length > 0) {
        tokenStack.push(tokens.add(token, TOK_TYPE_FUNCTION, TOK_SUBTYPE_START));
        token = "";
      } else {
        tokenStack.push(tokens.add("", TOK_TYPE_SUBEXPR, TOK_SUBTYPE_START));
      }
      offset += 1;
      continue;
    }

    // stop subexpression

    if (currentChar() == ")") {
      if (token.Length > 0) {
        tokens.add(token, TOK_TYPE_OPERAND);
        token = "";
      }
      tokens.addRef(tokenStack.pop());
      offset += 1;
      continue;
    }

    // token accumulation

    token += currentChar();
offset += 1;

  }

  // dump remaining accumulation

  if (token.Length > 0) tokens.add(token, TOK_TYPE_OPERAND);

  // move all tokens to a new collection, excluding all unnecessary white-space tokens

  var tokens2 = new Tokens();

  while (tokens.moveNext()) {

    token = tokens.current();

    if (token.type == TOK_TYPE_WSPACE) {
      if ((tokens.BOF()) || (tokens.EOF())) {
        // no-op
      } else if (!(
                 ((tokens.previous().type == TOK_TYPE_FUNCTION) && (tokens.previous().subtype == TOK_SUBTYPE_STOP)) ||
                 ((tokens.previous().type == TOK_TYPE_SUBEXPR) && (tokens.previous().subtype == TOK_SUBTYPE_STOP)) ||
                 (tokens.previous().type == TOK_TYPE_OPERAND)
                )
              ) {
                // no-op
              }
      else if (!(
                 ((tokens.next().type == TOK_TYPE_FUNCTION) && (tokens.next().subtype == TOK_SUBTYPE_START)) ||
                 ((tokens.next().type == TOK_TYPE_SUBEXPR) && (tokens.next().subtype == TOK_SUBTYPE_START)) ||
                 (tokens.next().type == TOK_TYPE_OPERAND)
                 )
               ) {
                 // no-op
               }
      else {
        tokens2.add(token.value, TOK_TYPE_OP_IN, TOK_SUBTYPE_INTERSECT);
      }
      continue;
    }

    tokens2.addRef(token);

  }

  // switch infix "-" operator to prefix when appropriate, switch infix "+" operator to noop when appropriate, identify operand
  // and infix-operator subtypes, pull "@" from in front of function names

  while (tokens2.moveNext()) {

    token = tokens2.current();

    if ((token.type == TOK_TYPE_OP_IN) && (token.value == "-")) {
      if (tokens2.BOF()) {
        token.type = TOK_TYPE_OP_PRE;
      } else if (
               ((tokens2.previous().type == TOK_TYPE_FUNCTION) && (tokens2.previous().subtype == TOK_SUBTYPE_STOP)) ||
               ((tokens2.previous().type == TOK_TYPE_SUBEXPR) && (tokens2.previous().subtype == TOK_SUBTYPE_STOP)) ||
               (tokens2.previous().type == TOK_TYPE_OP_POST) ||
               (tokens2.previous().type == TOK_TYPE_OPERAND)
             ) {
        token.subtype = TOK_SUBTYPE_MATH;
      } else {
        token.type = TOK_TYPE_OP_PRE;
      }
      continue;
    }

    if ((token.type == TOK_TYPE_OP_IN) && (token.value == "+")) {
      if (tokens2.BOF()) {
        token.type = TOK_TYPE_NOOP;
      } else if (
               ((tokens2.previous().type == TOK_TYPE_FUNCTION) && (tokens2.previous().subtype == TOK_SUBTYPE_STOP)) ||
               ((tokens2.previous().type == TOK_TYPE_SUBEXPR) && (tokens2.previous().subtype == TOK_SUBTYPE_STOP)) ||
               (tokens2.previous().type == TOK_TYPE_OP_POST) ||
               (tokens2.previous().type == TOK_TYPE_OPERAND)
             ) {
        token.subtype = TOK_SUBTYPE_MATH;
      } else {
        token.type = TOK_TYPE_NOOP;
      }
      continue;
    }

    if ((token.type == TOK_TYPE_OP_IN) && (token.subtype.length == 0)) {
      if (("<>=").indexOf(token.value.substr(0, 1)) != -1) {
        token.subtype = TOK_SUBTYPE_LOGICAL;
      } else if (token.value == "&") {
        token.subtype = TOK_SUBTYPE_CONCAT;
      } else {
        token.subtype = TOK_SUBTYPE_MATH;
      }
      continue;
    }

    if ((token.type == TOK_TYPE_OPERAND) && (token.subtype.length == 0)) {
      if (isNaN(Number(language.reformatNumberForJsParsing(token.value)))) {
        if (token.value == language.true) {
          token.subtype = TOK_SUBTYPE_LOGICAL;
          token.value = 'TRUE';
        } else if (token.value == language.false) {
          token.subtype = TOK_SUBTYPE_LOGICAL;
          token.value = 'FALSE';
        } else {
          token.subtype = TOK_SUBTYPE_RANGE;
        }
      } else {
        token.subtype = TOK_SUBTYPE_NUMBER;
        token.value = language.reformatNumberForJsParsing(token.value);
      }
      continue;
    }

    if (token.type == TOK_TYPE_FUNCTION) {
      if (token.value.substr(0, 1) == "@") {
        token.value = token.value.substr(1);
      }
      continue;
    }

  }

  tokens2.reset();

  // move all tokens to a new collection, excluding all noops

  tokens = new Tokens();

  while (tokens2.moveNext()) {
    if (tokens2.current().type != TOK_TYPE_NOOP) {
      tokens.addRef(tokens2.current());
    }
  }

  tokens.reset();

  return tokens.toArray();
}
    }

    public class Tokens
    {
        public Tokens()
        {
            this.items = [];
            this.index = -1;
        }

    public Token add(string value, string type, string subtype=null)
        {
            const token = createToken(value, type, subtype);
            this.addRef(token);
            return token;
        }

    public void addRef(Token token)
        {
            this.items.push(token);
        }

    public void reset()
        {
            this.index = -1;
        }

    public void BOF()
        {
            return this.index <= 0;
        }

    public void EOF()
        {
            return this.index >= this.items.length - 1;
        }

    public bool moveNext()
        {
            if (this.EOF())
                return false;
            this.index++;
            return true;
        }

    public Token current()
        {
            if (this.index == -1)
                return null;
            return this.items[this.index];
        }

    public Token next()
        {
            if (this.EOF())
                return null;
            return this.items[this.index + 1];
        }

    public Token previous()
        {
            if (this.index < 1)
                return null;
            return (this.items[this.index - 1]);
        }

    public Token[] toArray()
        {
            return this.items;
        }
    }

    public class TokenStack
    {
        public TokenStack()
        {
            this.items = [];
        }

    public void push(Token token)
        {
            this.items.push(token);
        }

    public Token pop()
        {
            const token = this.items.pop();
            return createToken("", token.type, TOK_SUBTYPE_STOP);
        }

    public void token()
        {
            if (this.items.length > 0)
            {
                return this.items[this.items.length - 1];
            }
            else
            {
                return null;
            }
        }

    public string value()
        {
            return this.token() ? this.token().value : '';
        }

    public string type()
        {
            return this.token() ? this.token().type : '';
        }

    public string subtype()
        {
            return this.token() ? this.token().subtype : '';
        }
    }
}
