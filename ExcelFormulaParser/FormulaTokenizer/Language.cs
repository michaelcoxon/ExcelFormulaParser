// https://github.com/psalaets/excel-formula-tokenizer/blob/master/languages.js

namespace ExcelFormulaParser.FormulaTokenizer
{
    public class Language
    {
        // value for true
        public readonly string True = "TRUE";
        // value for false
        public readonly string False = "FALSE";
        // separates function arguments
        public readonly string ArgumentSeparator = ",";
        // decimal point in numbers
        public readonly string DecimalSeparator = ".";
        // returns number string that can be parsed by Number()
        public string ReformatNumberForJsParsing(string n)
        {
            return n;
        }
    }
}
