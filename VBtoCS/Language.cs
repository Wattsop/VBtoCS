using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VBtoCS
{
    internal static class Language
    {
        private static Dictionary<string, string> _accessModifiers = new Dictionary<string, string>();
        private static Dictionary<string, string> _dataTypes = new Dictionary<string, string>();
        private static Dictionary<string, string> _excludedKeywords = new Dictionary<string, string>();
        private static Dictionary<string, string> _flowKeywords = new Dictionary<string, string>();

        private static Dictionary<string, string> _identifierTypeChars = new Dictionary<string, string>();
        private static Dictionary<string, string[]> _pairedKeywords = new Dictionary<string, string[]>();
        private static Dictionary<string, string> _typeKeywords = new Dictionary<string, string>();

        static Language()
        {
            AddKeywords();
        }

        private static void AddKeywords()
        {
            // The first few items are technically not
            // access modifiers but are included here
            // for convenience.
            _accessModifiers.Add("const", "const");
            _accessModifiers.Add("readonly", "readonly");
            _accessModifiers.Add("static", "static");

            _accessModifiers.Add("friend", "internal");
            _accessModifiers.Add("private", "private");
            _accessModifiers.Add("protected", "protected");
            _accessModifiers.Add("public", "public");

            _typeKeywords.Add("as", "as");

            _dataTypes.Add("double", "double");
            _dataTypes.Add("integer", "int");
            _dataTypes.Add("int", "int");
            _dataTypes.Add("list", "List");
            _dataTypes.Add("dictionary", "Dictionary");
            _dataTypes.Add("string", "string");

            _excludedKeywords.Add("byval", "");
            _excludedKeywords.Add("dim", "");

            // TO-DO: ByRef/ref needs to be processed
            _flowKeywords.Add("byref", "ref");

            _flowKeywords.Add("and", "&&");
            _flowKeywords.Add("class", "class");
            _flowKeywords.Add("do", "do");

            _flowKeywords.Add("end class", "}");
            _flowKeywords.Add("end enum", "}");
            _flowKeywords.Add("end function", "}");
            _flowKeywords.Add("end if", "}");
            _flowKeywords.Add("end sub", "}");
            _flowKeywords.Add("end while", "}");

            _flowKeywords.Add("enum", "enum");
            _flowKeywords.Add("for", "for");
            _flowKeywords.Add("function", "");
            _flowKeywords.Add("if", "if");

            _flowKeywords.Add("loop", "}");
            _flowKeywords.Add("loop until", "}");
            _flowKeywords.Add("loop while", "}");
            _flowKeywords.Add("new", "new");
            _flowKeywords.Add("next", "}");
            _flowKeywords.Add("of", "of");

            _flowKeywords.Add("sub", "void");
            _flowKeywords.Add("sub new", "sub new");
            _flowKeywords.Add("then", "then");

            _flowKeywords.Add("or", "||");
            _flowKeywords.Add("rem", "//");
            _flowKeywords.Add("to", "to");

            _identifierTypeChars.Add("%", "int");
            _identifierTypeChars.Add("&", "long");
            _identifierTypeChars.Add("@", "decimal");
            _identifierTypeChars.Add("!", "float");
            _identifierTypeChars.Add("#", "double");
            _identifierTypeChars.Add("$", "string");

            _pairedKeywords.Add("if", new string[] { "then" });
            _pairedKeywords.Add("loop", new string[] { "until", "while" });
            _pairedKeywords.Add("end", new string[] { "class", "enum", "function", "if", "sub", "while" });
            _pairedKeywords.Add("sub", new string[] { "new" });


            //keywords.Add("", "");
        }

        internal static bool ConvertAccessModifier(string modifier, out string? newModifier)
        {
            return _accessModifiers.TryGetValue(modifier.ToLower(), out newModifier);
        }

        internal static bool ConvertDataType(string dataType, out string? newDataType)
        {
            return _dataTypes.TryGetValue(dataType.ToLower(), out newDataType);
        }

        internal static bool ConvertIdTypeChar(string idTypeChar, out string? newDataType)
        {
            return _identifierTypeChars.TryGetValue(idTypeChar, out newDataType);
        }

        internal static bool IsAccessModifier(string str)
        {
            return _accessModifiers.ContainsKey(str);
        }

        internal static bool IsDataType(string str)
        {
            return _dataTypes.ContainsKey(str);
        }

        internal static bool IsExcludedKeyword(string str)
        {
            return _excludedKeywords.ContainsKey(str);
        }

        internal static bool IsIdTypeChar(char c)
        {
            return IsIdTypeChar(c.ToString());
        }

        internal static bool IsIdTypeChar(string str)
        {
            return _identifierTypeChars.ContainsKey(str);
        }

        internal static bool IsUnregistered(string word)
        {
            return (!VBConverter.IsNumeric(word) &&
                (!VBConverter.IsOperator(word)) &&
                (!VBConverter.IsGroupingSymbol(word)) &&
                (!_accessModifiers.ContainsKey(word)) &&
                (!_dataTypes.ContainsKey(word)) &&
                (!_excludedKeywords.ContainsKey(word)) &&
                (!_flowKeywords.ContainsKey(word)) &&
                (!_pairedKeywords.ContainsKey(word)) &&
                (!_typeKeywords.ContainsKey(word)));
        }
    }
}
