// https://github.com/psalaets/excel-formula-tokenizer/blob/master/languages.js

namespace ExcelFormulaParser.FormulaTokenizer
{
    internal static class BuiltinLanguages
    {
        private class LangEnglish : ILanguage
        {
            public string True { get; } = "TRUE";
            public string False { get; } = "FALSE";
            public string ArgumentSeparator { get; } = ",";
            public string DecimalSeparator { get; } = ".";
            public string ReformatNumberForJsParsing(string n)
            {
                return n;
            }
        }

        public static ILanguage English = new LangEnglish();
    }

    public interface ILanguage
    {
        /// <summary>
        /// Gets the value for true
        /// </summary>
        string True { get; }

        /// <summary>
        /// Gets the value for false  
        /// </summary>
        string False { get; }

        /// <summary>
        /// Gets the argument separator. Separates function arguments 
        /// </summary>
        string ArgumentSeparator { get; }

        /// <summary>
        /// Gets the decimal separator for numbers.
        /// </summary>
        string DecimalSeparator { get; }

        /// <summary>
        /// Reformats the number for js parsing. returns number string that can be parsed by <see cref="double.Parse(string)"/>
        /// </summary>
        /// <param name="n">The number.</param>
        /// <returns></returns>
        string ReformatNumberForJsParsing(string n);
    }
}
