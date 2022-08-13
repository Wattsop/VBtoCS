using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VBtoCS
{
    internal static class VBConverter
    {
        private static Dictionary<string, EnumInfo> _enums = new Dictionary<string, EnumInfo>();
        private static Dictionary<string, ClassInfo> _classes = new Dictionary<string, ClassInfo>();
        private static Dictionary<string, FunctionInfo> _functions = new Dictionary<string, FunctionInfo>();
        private static Dictionary<string, VariableInfo> _variables = new Dictionary<string, VariableInfo>();

        // Line end comments using ' are not supported and
        // will be skipped in the output.
        //private static Dictionary<int, string> _lineEndComments = new Dictionary<int, string>();

        // inputLines are the strings read from a file
        private static List<string>? _inputLines;

        // outputLines are the processed strings that will be
        // written to the output file.
        private static List<string> _outputLines = new List<string>();

        // The "_insideClasses" list is used to keep track of which class
        // the code is in. Every time the loop passes a line
        // where a class starts, it adds the class to _insideClasses.
        //
        // Every time the loop passes a line that ends a class,
        // it removes the last class added to the list.
        private static List<string> _insideClasses = new List<string>();
        private static FunctionInfo? _insideFunction = null;

        private static List<string> _insideEnums = new List<string>();

        // _insideEnum will be true when the loop passes the
        // first line of an enum definition, and false when
        // the enum ends.
        private static bool _insideEnum = false;

        private static bool _insideModule = false;
        private static bool _insideClass = false;
        //private static bool _insideFunction = false;
        private static bool _insideDoLoop = false;

        // The _funcInputLineIndex variable keeps track of the
        // first line of a function definition (the signature).
        private static int _funcInputLineIndex = -1;

        // When true, static keyword will be listed after
        // the first access modifier using SwapStaticKeyword()
        // after the main conversion loop has run.
        private static bool _swapStaticKeyword = true;

        // Reset has to be called before or after
        // each file is loaded.
        private static void Reset()
        {
            _classes = new Dictionary<string, ClassInfo>();
            _functions = new Dictionary<string, FunctionInfo>();
            //_lineEndComments = new Dictionary<int, string>();
            _outputLines = new List<string>();
            _insideClasses = new List<string>();
            _insideEnum = false;

            _insideModule = false;
            _insideClass = false;
            _insideFunction = null;

            _funcInputLineIndex = -1;
        }

        // MAIN()
        static void Main(string[] args)
        {
            // Get the path to the folder that contains the converter.
            string baseDir = AppContext.BaseDirectory;

            // Store the input/output folder paths.
            string inputDir = baseDir + "input";
            string outputDir = baseDir + "output";

            // Create input/output folders if they
            // don't exist in the converter folder.
            Directory.CreateDirectory(inputDir);
            Directory.CreateDirectory(outputDir);

            // NOTE: Sub-folders in the input directory are
            // not supported.

            // Get a collection of file paths from the input folder.
            var inputFilePaths = Directory.EnumerateFiles(inputDir);

            // Print a message if there were no files.
            if (inputFilePaths.Count() < 1)
            {
                Console.WriteLine("The input folder does not contain any files.");
            }

            // Loop through the vb file paths.
            foreach (string inputPath in inputFilePaths)
            {
                // Get the file name from the file path.
                string fileName = Path.GetFileName(inputPath);

                // Change .vb to .cs extension.
                fileName = Path.ChangeExtension(fileName, ".cs");

                // Combine the outputDir path and the file name.
                string outputPath = Path.Combine(outputDir, fileName);

                // Load the file using the inputPath string.
                Console.WriteLine($"Loading {inputPath}");
                _inputLines = Load(inputPath);

                // Check if the file was empty.
                if (_inputLines == null)
                {
                    Console.WriteLine("The input file did not contain data.");
                }
                else
                {
                    // Enter primary conversion loop.
                    Convert();

                    // Save the result.
                    Save(outputPath, _outputLines);

                    // Reset VBConverter variables for next file.
                    Reset();
                }
            }

            Console.WriteLine();
            Console.ReadLine();
        }
        // END OF MAIN() METHOD

        // The Convert() method contains the primary conversion loop.
        private static void Convert()
        {
            if (_inputLines != null)
            {
                // This removes line end comments by default.
                //CommentLineEndings(_inputLines);

                // Convert a comma-separated Dim statement into
                // single variable Dim statements on separate lines.
                SeparateDimStatements(_inputLines);

                Console.WriteLine("");

                // Main conversion loop (which goes through each
                // line of input from the VB file and uses an
                // if statement to check whether certain keywords
                // are in that line)
                for (int lineIndex = 0; lineIndex < _inputLines.Count; lineIndex++)
                {
                    // Trim the space before and after the line.
                    _inputLines[lineIndex] = _inputLines[lineIndex].Trim();

                    // Remove comments from the end of the line.
                    _inputLines[lineIndex] = CommentLineEndUndo(_inputLines[lineIndex]);

                    // Separate a line string into "ops" (operand/operator) that are
                    // short strings that can be either words, numbers, or operators such
                    // as * + , ( ) .
                    List<string> lineOps = SeparateExpressionOps(_inputLines[lineIndex]);

                    // After separating ops, remove spaces from the input string.
                    lineOps = RemoveSpaces(lineOps);

                    // Change concat '&' symbols to '+'
                    AmpersandConcatToPlus(lineOps);


                    // ========================
                    // PRIMARY CONVERSION CHECK
                    // ========================


                    // This set of if-else statements calls
                    // converter methods that translate
                    // vb statements into cs.

                    // The first check is for comments.
                    // Note: Comments such as this one
                    // will be converted, but comments
                    // at the end of a line will
                    // be removed from cs output.
                    if (_inputLines[lineIndex].StartsWith("'"))
                    {
                        Comment(lineIndex);
                    }
                    // Class statement
                    else if (IsOutsideQuotes(lineOps, "Class") && (lineOps[0] != "End"))
                    {
                        _insideClass = true;
                        ClassSignature(lineOps);
                    }
                    // Dim statement (vb local variable)
                    else if ((lineOps.Count > 0) && (lineOps[0] == "Dim"))
                    {
                        VariableStatement(lineOps);
                    }
                    // Do statement
                    else if ((lineOps.Count > 0) && (lineOps[0] == "Do"))
                    {
                        _insideDoLoop = true;
                        DoLoop();
                    }
                    // End statement
                    else if ((lineOps.Count > 0) && (lineOps[0] == "End"))
                    {
                        EndStatement(lineOps);
                    }
                    // Enum statement
                    else if ((lineOps.Count > 0) && (lineOps[0] == "Enum"))
                    {
                        EnumStatement(lineOps);
                    }
                    // For statement
                    else if ((lineOps.Count > 0) && (lineOps[0] == "For"))
                    {
                        ForLoop(lineOps);
                    }
                    // Function statement
                    else if (IsOutsideQuotes(lineOps, "Function") && (lineOps[0] != "End"))
                    {
                        //_insideFunction = true;
                        FuncOrSub(lineOps, false);
                    }
                    // Get statement (The "Get" accessor from vb Property)
                    else if (IsOutsideQuotes(lineOps, "Get") && (lineOps[0] != "End"))
                    {
                        GetOrSet("Get", lineOps);
                    }
                    // If statement
                    else if ((lineOps.Count > 0) && (lineOps[0] == "If"))
                    {
                        IfStatement(lineOps);
                    }
                    // Loop keyword (from Do ... Loop While, Do ... Loop Until)
                    else if (IsOutsideQuotes(lineOps, "Loop"))
                    {
                        LoopStatement(lineOps);
                    }
                    // Module statement
                    else if (IsOutsideQuotes(lineOps, "Module") && (lineOps[0] != "End"))
                    {
                        _insideModule = true;
                        ClassSignature(lineOps);
                    }
                    else if (lineOps.Count > 0 && lineOps[0] == "Namespace")
                    {
                        Namespace(lineOps);
                    }
                    // Next keyword (part of For loop)
                    else if (lineOps.Count > 0 && lineOps[0] == "Next")
                    {
                        _outputLines.Add("}");
                    }
                    // Property definition
                    else if (IsOutsideQuotes(lineOps, "Property") && (lineOps[0] != "End"))
                    {
                        //_insideFunction = true;
                        Property(lineIndex);
                    }
                    // Return keyword (in Function definition)
                    else if (lineOps.Count > 0 && lineOps[0] == "Return")
                    {
                        ReturnStatement(lineOps);
                    }
                    // Set statement (The "Set" accessor from vb Property)
                    else if (IsOutsideQuotes(lineOps, "Set") && (lineOps[0] != "End"))
                    {
                        GetOrSet("Set", lineOps);
                    }
                    // Sub definition
                    else if (IsOutsideQuotes(lineOps, "Sub") && (lineOps[0] != "End"))
                    {
                        Sub(lineOps);
                    }
                    // While keyword from While ... End While.
                    // See "Loop" check for Do .. Loop While.
                    else if (lineOps.Contains("While"))
                    {
                        WhileLoop(lineOps);
                    }
                    // If no recognized keywords were found,
                    // try processing as a variable declaration, or
                    // let the line pass through to the cs output
                    // for human review.
                    else
                    {
                        bool isVariableDeclaration = false;

                        for (int i = 0; i < lineOps.Count; i++)
                        {
                            // If an access modifier is encountered
                            // this far into conversion, the input
                            // line should be a variable declaration.
                            if (Language.IsAccessModifier(lineOps[i].ToLower()))
                            {
                                isVariableDeclaration = true;
                            }
                        }

                        if (isVariableDeclaration)
                        {
                            VariableStatement(lineOps);
                        }
                        else
                        {
                            // If the converter was unable to
                            // translate the input vb line,
                            // let the input line pass through
                            // to the cs output for human review.
                            Passthrough(lineOps);
                        }

                    }
                }
                // END OF PRIMARY CONVERSION LOOP

                // Swap static keyword with the first
                // access modifier in the output lines.
                if (_swapStaticKeyword)
                {
                    SwapStaticKeyword();
                }

                // Add spacing to converted text
                int blankLineCount = 0;
                int tabCount = 0;
                string tabString = "";

                for (int i = 0; i < _outputLines.Count; i++)
                {
                    if (_outputLines[i] == "")
                    {
                        blankLineCount++;
                    }

                    if (blankLineCount > 1)
                    {
                        _outputLines.RemoveAt(i);
                        blankLineCount = 0;
                    }

                    if (_outputLines[i].StartsWith("{"))
                    {
                        _outputLines[i] = tabString + _outputLines[i];
                        tabCount++;
                        tabString = new String('\t', tabCount);
                    }
                    else if (_outputLines[i].StartsWith("}"))
                    {
                        tabCount--;
                        tabString = new String('\t', tabCount);
                        _outputLines[i] = tabString + _outputLines[i];
                    }
                    else
                    {
                        _outputLines[i] = tabString + _outputLines[i];
                    }
                }

                for (int i = 0; i < _outputLines.Count; i++)
                {
                    Console.WriteLine($"CS [{i}]: {_outputLines[i]}");
                }

                Console.WriteLine();
            }
        }
        // END OF CONVERT() METHOD

        private static void Sub(List<string> lineOps)
        {
            // "Sub New" will become
            // a cs class constructor.
            //
            // "Sub" by itself will be converted to a
            // function and then to a cs class method.
            //_insideFunction = true;

            int subIndex = lineOps.IndexOf("Sub");

            if (subIndex + 1 < lineOps.Count)
            {
                if (lineOps[subIndex + 1] == "Main")
                {
                    SubMain(lineOps);
                }
                else if (lineOps[subIndex + 1] == "New")
                {
                    // Convert "Sub New" to something like
                    // "public ClassName()
                    // {"
                    SubNew(lineOps);
                }
                else
                {
                    // If not Sub New, then change Sub to Function.
                    //lineOps[subIndex] = "Function";

                    // Process Sub as Function.
                    FuncOrSub(lineOps, true);
                }
            }
        }

        private static void SubMain(List<string> lineOps)
        {
            FunctionInfo subMainInfo = new FunctionInfo();
            subMainInfo.Name = "Main";

            _insideFunction = subMainInfo;

            _outputLines.Add("");

            // Sub Main() will become static void Main().
            // A static constructor can't have access modifiers.
            List<int> modIndices = new List<int>();

            for (int i = 0; i < lineOps.Count; i++)
            {
                if (Language.IsAccessModifier(lineOps[i].ToLower()))
                {
                    modIndices.Add(i);
                }
            }

            foreach (int index in modIndices)
            {
                lineOps.RemoveAt(index);
            }

            // Remove Sub keyword
            lineOps.RemoveAt(lineOps.IndexOf("Sub"));

            // OpsInfo is a struct that holds info about
            // how to combine ops into a line and a
            // method to combine them.
            OpsInfo combineOpsInfo = new OpsInfo(lineOps, true);
            combineOpsInfo.PrependString = "static void";
            string newLine = combineOpsInfo.CombineOps();

            _outputLines.Add(newLine);
            _outputLines.Add("{");
        }

        private static bool IsOutsideQuotes(List<string> lineOps, string searchString)
        {
            bool insideQuotes = false;

            for (int i = 0; i < lineOps.Count; i++)
            {
                if (lineOps[i] == "\"")
                {
                    insideQuotes = !insideQuotes;
                }

                if ((lineOps[i] == searchString) && !insideQuotes)
                {
                    return true;
                }
            }

            return false;
        }

        private static void Namespace(List<string> lineOps)
        {
            string firstWord = lineOps[0].ToLower();

            if (firstWord == "namespace")
            {
                lineOps[0] = firstWord;
                string newLine = "";

                foreach (string op in lineOps)
                {
                    newLine += op + " ";
                }

                newLine = newLine.Trim();

                _outputLines.Add(newLine);
                _outputLines.Add("{");
            }
        }

        // SwapStaticKeyword() swaps the static keyword
        // with the first access modifier in the lines.
        private static void SwapStaticKeyword()
        {
            for (int i = 0; i < _outputLines.Count; i++)
            {
                if (_outputLines[i].StartsWith("static"))
                {
                    int firstSpaceIndex = _outputLines[i].IndexOf(' ');
                    int secondSpaceIndex = _outputLines[i].IndexOf(' ', firstSpaceIndex + 1);
                    int secondWordLength = (secondSpaceIndex - 1) - firstSpaceIndex;
                    string secondWord = "";

                    if (firstSpaceIndex + secondWordLength < _outputLines[i].Length)
                    {
                        secondWord = _outputLines[i].Substring(firstSpaceIndex + 1, secondWordLength);
                    }

                    if (Language.IsAccessModifier(secondWord))
                    {
                        _outputLines[i] = $"{secondWord} static {_outputLines[i].Substring(secondSpaceIndex + 1)}";
                    }
                }
            }
        }

        // This method handles anything not handled by other methods
        // by slightly formatting it and letting it pass through.
        //
        // In some cases, this may result in errors in c# that
        // will need to be fixed manually.
        private static void Passthrough(List<string> lineOps)
        {
            string output = FormatGenericString(lineOps);

            // Place a semicolon at the end of a line if certain characters are not there
            if (output != "" && output[output.Length - 1] != '{' && output[output.Length - 1] != '}' && output[output.Length - 1] != ',')
            {
                output += ";";
            }

            // Add the processed line to the output list
            _outputLines.Add(output);
        }

        private static string FormatGenericString(List<string> lineOps)
        {
            string output = "";
            string previousOp = "";

            // The insideDoubleQuotes bool is used to tell when
            // the a portion of text is inside quotation marks, which
            // would mean it should be copied exactly as it appears.
            bool insideDoubleQuotes = false;

            // Loops through the parts of the input line
            for (int j = 0; j < lineOps.Count; j++)
            {
                // Escape characters, such as \", are used in this if statement.
                // == "\"" will check if the op is a " by using the \" escape character
                // and != "\\" will check if the two ops combined make
                // the \" escape character, which is used to print a "
                if (lineOps[j] == "\"" && previousOp != "\\")
                {
                    // Toggle the insideDoubleQuotes bool to true/false
                    insideDoubleQuotes = !insideDoubleQuotes;
                }

                // If the op is a dot or closing parenthesis
                if (lineOps[j] == "." || lineOps[j] == "," || lineOps[j] == ")")
                {
                    // Trim the whitespace off the ends of the output string
                    output = output.Trim();
                }

                // Add the op to the output string
                output += lineOps[j];

                // If the op is not a dot or '(' and not inside parentheses
                if (lineOps[j] != "." && lineOps[j] != "(" && !insideDoubleQuotes)
                {
                    // Add a space after it
                    output += " ";
                }

                // Store the op in order to check for escape
                // characters, such as \", at the top of the loop
                previousOp = lineOps[j];
            }

            // Make sure there are no spaces at the beginning or end of output
            output = output.Trim();

            // Place a comma after each value in an enum
            if (_insideEnum)
            {
                output += ",";
            }

            // Place a semicolon at the end of a line if certain characters are not there
            //if (output != "" && output[output.Length - 1] != '{' && output[output.Length - 1] != '}' && output[output.Length - 1] != ',')
            //{
            //output += ";";
            //}

            return output;
        }

        // Changes the VB concat symbol '&' to a '+' sign
        private static void AmpersandConcatToPlus(List<string> lineOps)
        {
            string previousOp = "";

            // Loop through the parts of a line
            for (int j = 0; j < lineOps.Count; j++)
            {
                string nextOp = "";

                // If the next index is inside the bounds of the list
                if (j + 1 < lineOps.Count)
                {
                    // Look ahead to the next op and store it
                    nextOp = lineOps[j + 1];
                }

                // Check if ops combine to form the && or ++ operators
                if ((lineOps[j] == "&") && (nextOp != "&") && (previousOp != "+"))
                {
                    // Change & to + if the next op is not one of these
                    if (nextOp != "=" && nextOp != "(" && nextOp != "")
                    {
                        lineOps[j] = "+";
                    }
                }

                previousOp = lineOps[j];
            }
        }

        // Writes the first two lines of a do loop
        // 
        private static void DoLoop()
        {
            _outputLines.Add("do");
            _outputLines.Add("{");
        }

        // Rewrites a VB "Sub New" as a C# constructor
        private static void SubNew(List<string> lineOps)
        {
            FunctionInfo subNewInfo = new FunctionInfo();
            _insideFunction = subNewInfo;

            // VB Modules are converted to static C# classes.
            // This constructor (Sub New()) was inside a VB module,
            // so this needs to be a static constructor.
            if (_insideModule && !_insideClass)
            {
                // Static constructor can't have access modifiers
                List<int> modIndices = new List<int>();

                for (int i = 0; i < lineOps.Count; i++)
                {
                    if (Language.IsAccessModifier(lineOps[i].ToLower()))
                    {
                        modIndices.Add(i);
                    }
                }

                foreach (int index in modIndices)
                {
                    lineOps.RemoveAt(index);
                }

                lineOps.Insert(0, "static");
            }
            else
            {
                // If the Sub New is not directly inside
                // a module, then add an access modifier.
                lineOps.Insert(0, "public");
            }

            // Set subNewInfo.Name to the class name.
            subNewInfo.Name = _insideClasses.Last();

            // Replace the New keyword with the name of the class.
            lineOps[lineOps.IndexOf("New")] = subNewInfo.Name;

            // Remove Sub keyword
            lineOps.RemoveAt(lineOps.IndexOf("Sub"));

            string str = "";

            for (int j = 0; j < lineOps.Count; j++)
            {
                str += lineOps[j];

                if (lineOps[j] != "(")
                {
                    str += " ";
                }
            }

            str = str.Trim();

            _outputLines.Add(str);
            _outputLines.Add("{");
        }

        private static void WhileLoop(List<string> lineOps)
        {
            string str = "while (";

            for (int j = 1; j < lineOps.Count; j++)
            {
                str += lineOps[j];

                if (lineOps[j] != "(")
                {
                    str += " ";
                }
            }

            str = str.Trim();
            str += ")";

            _outputLines.Add(str);
            _outputLines.Add("{");
        }

        private static void ReturnStatement(List<string> lineOps)
        {
            string str = "return";

            if (lineOps.Count > 1)
            {
                str += " " + lineOps[1];
            }

            str += ";";

            _outputLines.Add(str);
        }

        private static void LoopStatement(List<string> lineOps)
        {
            string loopCondition = "true";

            // Loop...Until is re-written as a while loop with
            // the condition reversed, so Loop...Until (a != b)
            // becomes while (a == b).
            if (lineOps.Contains("Until") && lineOps.Count > 2)
            {
                loopCondition = "";

                for (int j = 2; j < lineOps.Count; j++)
                {
                    string previousOp = "";
                    string nextOp = "";

                    if (j - 1 >= 0)
                    {
                        previousOp = lineOps[j - 1];
                    }

                    if (j + 1 < lineOps.Count)
                    {
                        nextOp = lineOps[j + 1];
                    }

                    if (lineOps[j] == "<")
                    {
                        if (nextOp == "=")
                        {
                            lineOps[j] = ">";
                        }
                        else
                        {
                            lineOps[j] = ">=";
                        }
                    }
                    else if (lineOps[j] == ">")
                    {
                        if (nextOp == "=")
                        {
                            lineOps[j] = "<";
                        }
                        else
                        {
                            lineOps[j] = "<=";
                        }
                    }
                    else if (lineOps[j] == "=")
                    {
                        if (previousOp != "<" && previousOp != ">")
                        {
                            if (nextOp != "<" && nextOp != ">")
                            {
                                lineOps[j] = "!=";
                            }
                        }
                    }
                    else if (lineOps[j] == "!")
                    {
                        if (nextOp == "=")
                        {
                            lineOps[j] = "=";
                        }
                        else
                        {
                            lineOps[j] = "";
                        }
                    }

                    if (lineOps[j] != "")
                    {
                        loopCondition += lineOps[j] + " ";
                    }
                }

                loopCondition = loopCondition.Trim();
            }
            else if (lineOps.Contains("While") && lineOps.Count > 2)
            {
                loopCondition = "";

                for (int j = 2; j < lineOps.Count; j++)
                {
                    loopCondition += lineOps[j] + " ";
                }

                loopCondition = loopCondition.Trim();
            }

            _outputLines.Add($"}} while ({loopCondition});");
        }

        private static void IfStatement(List<string> lineOps)
        {
            int thenIndex = lineOps.Count;
            bool singleLineIf = false;

            for (int opIndex = 0; opIndex < lineOps.Count; opIndex++)
            {
                if (lineOps[opIndex] == "If")
                {
                    lineOps[opIndex] = "if";
                }
                else if (lineOps[opIndex] == "=")
                {
                    if (opIndex < thenIndex)
                    {
                        lineOps[opIndex] = "==";
                    }
                }
                else if (lineOps[opIndex] == "And")
                {
                    lineOps[opIndex] = "&&";
                }
                else if (lineOps[opIndex] == "Or")
                {
                    lineOps[opIndex] = "||";
                }
                else if (lineOps[opIndex] == "Then")
                {
                    lineOps[opIndex] = ")";
                    thenIndex = opIndex;

                    if (opIndex + 1 < lineOps.Count)
                    {
                        singleLineIf = true;
                    }
                }
            }

            string str = "if (";

            for (int j = 1; j < lineOps.Count; j++)
            {
                if (lineOps[j] == ")")
                {
                    str = str.Trim();
                }

                str += lineOps[j] + " ";
            }

            str = str.Trim();

            if (singleLineIf)
            {
                _outputLines.Add(str + ";");
            }
            else
            {
                _outputLines.Add(str);
                _outputLines.Add("{");
            }
        }

        private static void ForLoop(List<string> lineOps)
        {
            string varName = lineOps[1];
            string item1 = "";
            string item2 = "";

            int toIndex = lineOps.IndexOf("To");

            for (int j = 3; j < toIndex; j++)
            {
                item1 += lineOps[j];
            }

            for (int j = toIndex + 1; j < lineOps.Count; j++)
            {
                item2 += lineOps[j];
            }

            string str;

            if (int.TryParse(item1, out int firstInt) && int.TryParse(item2, out int secondInt))
            {
                str = $"for (int {varName} = {item1}; {varName} != {item2}; {varName} += ({item1} < {item2}) ? 1 : -1)";
            }
            else if (decimal.TryParse(item1, out decimal firstDecimal) && decimal.TryParse(item2, out decimal secondDecimal))
            {
                str = $"for (decimal {varName} = {item1}; {varName} != {item2}; {varName} += ({item1} < {item2}) ? 1M : -1M)";
            }
            else
            {
                str = $"for (int {varName} = {item1}; {varName} != {item2}; {varName} += ({item1} < {item2}) ? 1 : -1)";
            }

            _outputLines.Add(str);
            _outputLines.Add("{");
        }

        private static void EndStatement(List<string> lineOps)
        {
            if (lineOps.Contains("Class"))
            {
                _insideClass = false;
                _insideClasses.Remove(_insideClasses.Last());
            }
            else if (lineOps.Contains("Enum"))
            {
                _insideEnum = false;
                _insideEnums.Remove(_insideEnums.Last());
            }
            else if (lineOps.Contains("Function"))
            {
                //_insideFunction = false;
                _insideFunction = null;
                //if (inferFuncReturnType)
                //{
                //    
                //}
            }
            else if (lineOps.Contains("Module"))
            {
                _insideModule = false;
                _insideClasses.Remove(_insideClasses.Last());
            }
            else if (lineOps.Contains("Property"))
            {
                //_insideFunction = false;
                _insideFunction = null;
                return;
            }
            else if (lineOps.Contains("Sub"))
            {
                //_insideFunction = false;
                _insideFunction = null;
            }

            _outputLines.Add("}");
        }

        private static void Comment(int inputLineIndex)
        {
            if (_inputLines != null)
            {
                string str = "//";

                for (int j = 1; j < _inputLines[inputLineIndex].Length; j++)
                {
                    str += _inputLines[inputLineIndex][j];
                }

                _outputLines.Add(str);
            }
        }

        // CommentLineEndings() removes line ending comments
        // because they are frowned upon and due to potential
        // confusion for a programmer reading the generated code,
        // which may contain multiple lines generated for one
        // line of VB or have VB lines removed entirely.
        private static void CommentLineEndings(List<string> inputLines)
        {
            for (int i = 0; i < inputLines.Count; i++)
            {
                int commentIndex = -1;
                bool inQuotes = false;

                for (int j = inputLines[i].Length - 1; j >= 0; j--)
                {
                    if (inputLines[i][j] == '\"')
                    {
                        inQuotes = !inQuotes;
                    }

                    if (inputLines[i][j] == '\'' && !inQuotes)
                    {
                        commentIndex = j;
                    }
                }

                if (commentIndex > 1)
                {
                    //_lineEndComments.Add(i, inputLines[i].Substring(commentIndex));
                    //inputLines[i] = inputLines[i].Remove(commentIndex);
                }

                inputLines[i] = inputLines[i].Trim();
            }
        }

        private static string CommentLineEndUndo(string line)
        {
            int commentIndex = -1;
            bool inQuotes = false;

            for (int i = line.Length - 1; i >= 0; i--)
            {
                if (line[i] == '\"')
                {
                    inQuotes = !inQuotes;
                }

                if (line[i] == '\'' && !inQuotes)
                {
                    commentIndex = i;
                }
            }

            if (commentIndex > 1)
            {
                //_lineEndComments.Add(i, inputLines[i].Substring(commentIndex));
                line = line.Remove(commentIndex);
            }

            return line;
        }

        private static void EnumStatementOld(List<string> lineOps)
        {
            string str = "enum ";

            for (int j = 1; j < lineOps.Count; j++)
            {
                str += lineOps[j] + " ";
            }

            str = str.Trim();

            _outputLines.Add(str);
            _outputLines.Add("{");

            _insideEnum = true;
        }

        private static void EnumStatement(List<string> lineOps)
        {
            EnumInfo enumInfo = new EnumInfo();
            string outputString = "";

            for (int i = 0; i < lineOps.Count; i++)
            {
                if (Language.IsAccessModifier(lineOps[i]))
                {
                    enumInfo.AddAccessModifier(lineOps[i]);
                }
                else if ((lineOps[i] == "Enum") && (i + 1 < lineOps.Count))
                {
                    string enumName = lineOps[i + 1];
                    enumInfo.AddName(enumName);
                    _enums.Add(enumName, enumInfo);
                    _insideEnums.Add(enumName);
                    outputString += "enum ";
                }
                else
                {
                    outputString += lineOps[i] + " ";
                }
            }

            outputString = outputString.Trim();

            _outputLines.Add(outputString);
            _outputLines.Add("{");

            _insideEnum = true;
        }

        private static void VariableStatement(List<string> lineOps)
        {
            VariableInfo variableInfo = new VariableInfo();
            string previousOp = "";
            string nextOp = "";

            if (_insideModule && !_insideClass && (_insideFunction == null))
            {
                lineOps.Insert(0, "static");
            }

            for (int i = 0; i < lineOps.Count; i++)
            {
                if (i + 1 < lineOps.Count)
                {
                    nextOp = lineOps[i + 1];
                }
                else
                {
                    nextOp = "";
                }

                if (Language.IsAccessModifier(lineOps[i].ToLower()))
                {
                    variableInfo.AccessModifiers.Add(lineOps[i].ToLower());
                }

                if (previousOp == "Dim" || Language.IsAccessModifier(previousOp.ToLower()))
                {
                    if (nextOp == "(")
                    {
                        variableInfo.IsArray = true;
                    }

                    if (nextOp == "(" || nextOp == "" || nextOp == "As" || nextOp == "=" ||
                        Language.IsIdTypeChar(nextOp) || Language.IsDataType(nextOp))
                    {
                        variableInfo.Name = lineOps[i];

                        if (_insideFunction.HasValue)
                        {
                            _insideFunction.Value.AddVariable(variableInfo);
                        }
                        else
                        {
                            _classes[_insideClasses.Last()].AddVariable(variableInfo);
                        }
                    }

                    if (Language.ConvertIdTypeChar(nextOp, out string? variableType) && variableType != null)
                    {
                        variableInfo.VariableType = variableType;
                    }
                }

                if (Language.IsIdTypeChar(lineOps[i]))
                {
                    if (previousOp == variableInfo.Name && nextOp == "(")
                    {
                        variableInfo.IsArray = true;
                    }
                }

                if (previousOp == "As" || previousOp == "=")
                {
                    if (lineOps[i] == "New")
                    {
                        variableInfo.IsNew = true;
                    }
                    else
                    {
                        nextOp = lineOps[i];
                    }

                    if (_classes.ContainsKey(nextOp) || _enums.ContainsKey(nextOp))
                    {
                        variableInfo.IsNew = true;
                        variableInfo.VariableType = nextOp;
                    }
                    else if (Language.ConvertDataType(nextOp.ToLower(), out string? asDataType))
                    {
                        if (asDataType != null)
                        {
                            variableInfo.VariableType = asDataType;
                        }
                    }
                    else
                    {
                        List<string> assignment = lineOps.GetRange(i, lineOps.Count - i);
                        AmpersandConcatToPlus(assignment);
                        variableInfo.Assignment = FormatGenericString(assignment);
                    }
                }

                if (Language.IsDataType(previousOp.ToLower()) && lineOps[i] == "(")
                {
                    if (nextOp == "Of")
                    {
                        string str = "";

                        for (int j = i + 2; j < lineOps.Count; j++)
                        {
                            str += lineOps[j];
                        }

                        str = str.Trim(')');

                        string[] strArray = str.Split(',');

                        for (int k = 0; k < strArray.Length; k++)
                        {
                            if (Language.ConvertDataType(strArray[k].ToLower(), out string? ofDataType))
                            {
                                variableInfo.VariableType += "<" + ofDataType + ">";
                            }
                            else if (_classes.ContainsKey(strArray[k]))
                            {
                                variableInfo.VariableType += "<" + strArray[k] + ">";
                            }
                        }

                        variableInfo.VariableType = variableInfo.VariableType.Replace("><", ", ");
                    }
                    else if (!variableInfo.IsNew)
                    {
                        variableInfo.IsArray = true;
                    }
                }

                previousOp = lineOps[i];
            }

            if (Language.ConvertDataType(variableInfo.VariableType.ToLower(), out string? convertedDataType))
            {
                if (convertedDataType != null)
                {
                    variableInfo.VariableType = convertedDataType;
                }
            }

            if (variableInfo.IsArray)
            {
                variableInfo.VariableType += "[]";
            }

            if (variableInfo.VariableType != "")
            {
                if (variableInfo.IsNew && !variableInfo.IsArray)
                {
                    variableInfo.Assignment += $"new {variableInfo.VariableType}()";
                }

                string str = "";

                foreach (string mod in variableInfo.AccessModifiers)
                {
                    str += mod + " ";
                }

                if (variableInfo.Assignment == "")
                {
                    str += $"{variableInfo.VariableType} {variableInfo.Name};";
                }
                else
                {
                    str += $"{variableInfo.VariableType} {variableInfo.Name} = {variableInfo.Assignment};";
                }

                _outputLines.Add(str);
            }
        }

        private static void ClassSignature(List<string> lineOps)
        {
            string vbElementType = "";

            for (int j = 0; j < lineOps.Count; j++)
            {
                if (lineOps[j] == "Class" || lineOps[j] == "Module")
                {
                    vbElementType = lineOps[j];

                    lineOps[j] = "class";

                    if (j > 0)
                    {
                        for (int modIndex = 0; modIndex < j; modIndex++)
                        {
                            lineOps[modIndex] = lineOps[modIndex].ToLower();
                        }
                    }

                    if (j + 1 < lineOps.Count)
                    {
                        string name = lineOps[j + 1];
                        _classes.Add(name, new ClassInfo(name));
                        _insideClasses.Add(name);
                    }
                }
            }

            if (vbElementType == "Module")
            {
                lineOps.Insert(0, "static");
            }

            string str = "";

            for (int j = 0; j < lineOps.Count; j++)
            {
                str += lineOps[j] + " ";
            }

            _outputLines.Add(str);
            _outputLines.Add("{");
        }

        private static List<string> RemoveSpaces(List<string> ops)
        {
            List<string> vbOps = new List<string>(ops);

            string previousOp = "";
            bool insideDoubleQuotes = false;

            for (int i = 0; i < vbOps.Count; i++)
            {
                if (vbOps[i] == "\"" && previousOp != "\\")
                {
                    insideDoubleQuotes = !insideDoubleQuotes;
                }

                if (!insideDoubleQuotes)
                {
                    vbOps[i] = vbOps[i].Replace("  ", " ");
                    vbOps[i] = vbOps[i].Replace(" ", "");
                }

                previousOp = vbOps[i];
            }

            List<string> newOps = new List<string>();

            for (int i = 0; i < vbOps.Count; i++)
            {
                if (vbOps[i] != "")
                {
                    newOps.Add(vbOps[i]);
                }
            }

            return newOps;
        }

        // If a Dim statement has two or more parts,
        // separated by comma, this method writes them
        // as separate Dim statements.
        private static void SeparateDimStatements(List<string> lines)
        {
            // Loop through each line of input
            for (int i = 0; i < lines.Count; i++)
            {
                // If a local variable is declared
                if (lines[i].ToLower().StartsWith("dim"))
                {
                    // Split the input line by commas
                    string[] strArray = lines[i].Split(',');

                    // prependStr will hold strings that
                    // should not be separated by commas
                    // (things that appear in grouping symbols).
                    //
                    // Example: {1, 2, 3} will be added
                    // back together.
                    //
                    // prependStr will:
                    // hold {1 on the first pass,
                    // hold {1, 2 on the second pass, and
                    // hold {1, 2, 3} on the third pass.
                    string prependStr = "";

                    // Loop through strings from the input line
                    for (int j = 0; j < strArray.Length; j++)
                    {
                        // openGroupSymbol and closeGroupSymbol
                        // will be compared later to determine
                        // when the loop is outside grouping symbols
                        int openGroupSymbol = 0;
                        int closeGroupSymbol = 0;

                        // If prependStr is not empty, then there was
                        // an open grouping symbol without a closing one.
                        if (prependStr != "")
                        {
                            // Previous string is prepended to the current one
                            strArray[j] = prependStr + "," + strArray[j];
                            prependStr = "";
                        }

                        // Searching for grouping symbols
                        foreach (char c in strArray[j])
                        {
                            if (c == '(' || c == '[' || c == '{')
                            {
                                openGroupSymbol++;
                            }
                            else if (c == ')' || c == ']' || c == '}')
                            {
                                closeGroupSymbol++;
                            }
                        }

                        // If there was an open grouping symbol but no closing one,
                        if (openGroupSymbol != closeGroupSymbol)
                        {
                            // then save the current string in prependStr
                            prependStr = strArray[j];

                            // and erase it from the array.
                            strArray[j] = "";
                        }

                        // prependStr will be used at the top
                        // of the loop to rebuild the string that
                        // should not have been split by comma.
                    }

                    // Loops through the string array
                    // and inserts "Dim" in front of any
                    // string that does not already have one.
                    for (int j = 0; j < strArray.Length; j++)
                    {
                        if ((strArray[j] != "") && !strArray[j].ToLower().StartsWith("dim"))
                        {
                            strArray[j] = "Dim" + strArray[j];
                        }
                    }

                    // Removes the original, multi-variable "Dim" statement from the input line list.
                    lines.RemoveAt(i);

                    // Loops through the string array
                    for (int j = strArray.Length - 1; j >= 0; j--)
                    {
                        // If the string is not empty
                        if (strArray[j] != "")
                        {
                            // this inserts each new, single-variable
                            // "Dim" statement into the input line list.
                            lines.Insert(i, strArray[j]);
                        }
                    }
                }
            }
        }

        // Gets the function name from its signature, which has
        // been divided into a list of "ops" (operator/operand)
        // that can be words, numbers, or symbols.
        //
        // Example line divided into ops: SomeFunction, (, a, +, b, )
        // Each of the parts separated by commas is an op.
        private static string GetFuncName(List<string> lineOps)
        {
            int nameIndex = lineOps.IndexOf("(") - 1;

            if (nameIndex > -1)
            {
                return lineOps[nameIndex];
            }

            return "";
        }

        // Gets the access modifiers (public, private, etc.)
        // listed before the function name in the function signature.
        //
        // The function signature has been divided into a list
        // of "ops" (operator/operand) that can be words,
        // numbers, or symbols.
        //
        // Example line divided into ops: SomeFunction, (, a, +, b, )
        // Each of the parts separated by commas is an op.
        private static List<string> GetFuncAccessMods(List<string> lineOps)
        {
            List<string> accessModifiers = new List<string>();

            for (int j = 0; j < lineOps.Count; j++)
            {
                if (Language.IsAccessModifier(lineOps[j].ToLower()))
                {
                    accessModifiers.Add(lineOps[j].ToLower());
                }
            }

            return accessModifiers;
        }

        // Gets parameters from the function signature
        private static List<List<string>> GetFuncParams(List<string> lineOps)
        {
            List<List<string>> parameters = new List<List<string>>();

            int leftParenthIndex = lineOps.IndexOf("(");
            int rightParenthIndex = lineOps.LastIndexOf(")");
            List<string> parameterOps = new List<string>();

            for (int j = leftParenthIndex + 1; j < rightParenthIndex; j++)
            {
                if ((lineOps[j] != ",") && (lineOps[j] != ")"))
                {
                    if (lineOps[j] != "ByVal")
                    {
                        parameterOps.Add(lineOps[j].ToLower());
                    }
                }

                if ((lineOps[j] == ",") || (j == rightParenthIndex - 1))
                {
                    parameters.Add(parameterOps);
                    parameterOps = new List<string>();
                }
            }

            return parameters;
        }

        // Gets the VB variable type listed after the
        // variable name, converts that type to a c# type,
        // and inserts the c# type before the variable name
        private static void SetFuncParamTypes(FunctionInfo funcInfo)
        {
            string previousOp = "";

            foreach (var parOps in funcInfo.Parameters)
            {
                for (int j = 0; j < parOps.Count; j++)
                {
                    if (j + 1 < parOps.Count)
                    {
                        if (parOps[j].ToLower() == "as")
                        {
                            if (Language.ConvertDataType(parOps[j + 1].ToLower(), out string? newDataType) && newDataType != null)
                            {
                                parOps.RemoveAt(j + 1);
                                parOps.RemoveAt(j);
                                parOps.Insert(0, newDataType);

                                if (!funcInfo.ContainsVariable(previousOp))
                                {
                                    funcInfo.AddVariable(new VariableInfo(previousOp, newDataType));
                                }
                            }
                        }
                    }

                    previousOp = parOps[j];
                }
            }
        }

        // Gets the return type from the function signature.
        // If no return type is listed, "object" is used by default.
        private static string GetFuncReturnType(List<string> lineOps)
        {
            string returnType = "object";

            if (lineOps.Count > 1)
            {
                if (lineOps[lineOps.Count - 2].ToLower() == "as")
                {
                    Language.ConvertDataType(lineOps.Last().ToLower(), out string? newDataType);

                    if (newDataType != null)
                    {
                        returnType = newDataType;
                    }
                }
            }

            return returnType;
        }

        // Uses the info contained in funcInfo to write
        // a C# version of the VB function signature.
        private static string RewriteFuncSignature(FunctionInfo funcInfo)
        {
            // Begins with an empty output string
            string str = "";

            // Adds each access modifier, separated by spaces
            foreach (string mod in funcInfo.AccessModifiers)
            {
                str += mod + " ";
            }

            // Completes the part before the '(' in the function signature
            str += funcInfo.ReturnType + " " + funcInfo.Name + "(";

            // Each parameter has been divided into a list of strings.
            // The "parameters" list contains all of those lists.
            // Example: parameters<parameter<part>>
            List<List<string>> parameters = funcInfo.Parameters;

            // Loop through parameters
            for (int parIndex = 0; parIndex < parameters.Count; parIndex++)
            {
                // Get each list of ops (a list of parts of a parameter)
                List<string> parOps = parameters[parIndex];

                // Loop through the ops (parameter parts)
                for (int opIndex = 0; opIndex < parOps.Count; opIndex++)
                {
                    // Add each op to the output string
                    // Example output: int
                    str += parOps[opIndex];

                    // Add a space between each op in the parameter
                    // Example output: int x
                    if (opIndex < parOps.Count - 1)
                    {
                        str += " ";
                    }
                }

                // Add a comma after each parameter in the string
                // Example output: int x, int y
                if (parIndex < parameters.Count - 1)
                {
                    str += ", ";
                }
            }

            // Ends the method signature with a ')'
            str += ")";

            return str;
        }

        // The Property() method stores the index for
        // the first line of the VB property. The actual
        // processing of the VB property is done in GetOrSet().
        private static void Property(int inputLineIndex)
        {
            _funcInputLineIndex = inputLineIndex;
        }

        // GetOrSet() creates two C# methods to replace a VB property.
        // The getOrSet string should be either "Get" or "Set".
        private static void GetOrSet(string getOrSet, List<string> lineOps)
        {
            if (_inputLines != null)
            {
                // Uses the _funcInputLineIndex to get the
                // property signature from a previous line
                string propSignature = _inputLines[_funcInputLineIndex];

                // Separates parts of the property signature
                List<string> propSigOps = SeparateExpressionOps(propSignature);

                // Removes spaces so that every piece
                // of following code doesn't have to
                propSigOps = RemoveSpaces(propSigOps);

                // Fills a struct with info from property
                FunctionInfo propInfo = new FunctionInfo();
                propInfo.AccessModifiers = GetFuncAccessMods(propSigOps);
                propInfo.ReturnType = GetFuncReturnType(propSigOps);
                propInfo.Name = GetFuncName(propSigOps);
                propInfo.Parameters = GetFuncParams(propSigOps);

                // Fills a struct with info from Get or Set accessor
                FunctionInfo accessorInfo = new FunctionInfo();
                accessorInfo.AccessModifiers = GetFuncAccessMods(lineOps);
                accessorInfo.ReturnType = propInfo.ReturnType;
                accessorInfo.Parameters = GetFuncParams(lineOps);
                accessorInfo.Parameters.AddRange(propInfo.Parameters);

                _insideFunction = accessorInfo;

                // If there are no access mods in the accessor,
                // then the property mods are copied to it.
                if (accessorInfo.AccessModifiers.Count < 1)
                {
                    accessorInfo.AccessModifiers = propInfo.AccessModifiers;
                }

                if (_insideModule && !_insideClass)
                {
                    accessorInfo.AccessModifiers.Insert(0, "static");
                }

                if (getOrSet.ToLower() == "set")
                {
                    accessorInfo.ReturnType = "void";
                }

                // Sets the type for each parameter in the accessor's parentheses
                SetFuncParamTypes(accessorInfo);

                // Attempt to give the created C# method a name.
                // If Get/Set + Name already exists, "Prop_" is prepended.
                if (!_functions.ContainsKey(getOrSet + propInfo.Name))
                {
                    accessorInfo.Name = getOrSet + propInfo.Name;
                }
                else
                {
                    accessorInfo.Name = "Prop_" + getOrSet + propInfo.Name;
                }

                // Writes the new C# method signature representing VB Get or Set
                propSignature = RewriteFuncSignature(accessorInfo);

                // Adds method signature to the list of output lines
                _outputLines.Add("");
                _outputLines.Add(propSignature);
                _outputLines.Add("{");
            }
        }

        private static void FuncOrSub(List<string> lineOps, bool isSub)
        {
            _outputLines.Add("");

            bool isMain = IsOutsideQuotes(lineOps, "Main");

            if (isMain || (_insideModule && !_insideClass))
            {
                lineOps.Insert(0, "static");
            }

            // Fills a struct with info from function or sub
            FunctionInfo funcInfo = new FunctionInfo();
            funcInfo.AccessModifiers = GetFuncAccessMods(lineOps);
            funcInfo.ReturnType = GetFuncReturnType(lineOps);
            funcInfo.Name = GetFuncName(lineOps);
            funcInfo.Parameters = GetFuncParams(lineOps);

            _insideFunction = funcInfo;

            if (isSub)
            {
                funcInfo.ReturnType = "void";
            }

            // Sets the type for each parameter in the function's parentheses
            SetFuncParamTypes(funcInfo);

            // Writes the C# method signature
            string funcSignature = RewriteFuncSignature(funcInfo);

            // Adds C# method signature to the output line list
            if (funcInfo.Name != "")
            {
                _outputLines.Add(funcSignature);
                _outputLines.Add("{");
            }
        }

        internal static bool IsGroupingSymbol(char c)
        {
            return IsGroupingSymbol(c.ToString());
        }

        internal static bool IsGroupingSymbol(string str)
        {
            string[] groupingSymbols = { " ", "(", ")", "[", "]", "{", "}", "," };

            for (int i = 0; i < groupingSymbols.Length; i++)
            {
                if (str == groupingSymbols[i])
                {
                    return true;
                }
            }

            return false;
        }

        internal static bool IsOperator(char c)
        {
            return IsOperator(c.ToString());
        }

        internal static bool IsOperator(string str)
        {
            string[] operators = {
                ";", "'", "\"", ".", "&", "|", "@", "<", ">", "=", "*", "/", "+", "-"
            };

            for (int i = 0; i < operators.Length; i++)
            {
                if (str == operators[i])
                {
                    return true;
                }
            }

            return false;
        }

        internal static bool IsNumeric(char c)
        {
            return (c >= 48 && c <= 57);
        }

        internal static bool IsNumeric(string str)
        {
            for (int i = 0; i < str.Length; i++)
            {
                if (!IsNumeric(str[i]))
                {
                    return false;
                }
            }
            return true;
        }

        internal static bool IsAlpha(char c)
        {
            return (c >= 65 && c <= 90) || (c >= 97 && c <= 122);
        }

        internal static bool IsValidVarNameChar(char c)
        {
            return (IsAlpha(c) || IsNumeric(c) || c == 95);
        }

        internal static List<string> SeparateExpressionOps(string expression)
        {
            expression = expression.Trim();
            List<string> ops = new List<string>();
            string operatorStr = "";
            string operandStr = "";

            for (int i = 0; i < expression.Length; i++)
            {
                if (IsOperator(expression[i]))
                {
                    if (operandStr != "")
                    {
                        //Console.WriteLine("operandStr: " + operandStr);
                        ops.Add(operandStr);
                        operandStr = "";
                    }

                    operatorStr += expression[i];
                }

                if (IsValidVarNameChar(expression[i]))
                {
                    if (operatorStr != "")
                    {
                        //Console.WriteLine("operatorStr: " + operatorStr);
                        ops.Add(operatorStr);
                        operatorStr = "";
                    }

                    operandStr += expression[i];
                }

                bool isGroupingSymbol = IsGroupingSymbol(expression[i]);

                if (isGroupingSymbol || (i == expression.Length - 1))
                {
                    if (operandStr != "")
                    {
                        //Console.WriteLine("operandStr: " + operandStr);
                        ops.Add(operandStr);
                        operandStr = "";
                    }
                    if (operatorStr != "")
                    {
                        //Console.WriteLine("operatorStr: " + operatorStr);
                        ops.Add(operatorStr);
                        operatorStr = "";
                    }

                    if (isGroupingSymbol)
                    {
                        // Add grouping symbol
                        ops.Add(expression[i].ToString());
                    }
                }
            }

            return ops;
        }

        private static List<string> Load(string filePath)
        {
            //string docPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            List<string> inputLines = new List<string>();
            int lineNumber = 1;

            Console.WriteLine($"Converting {filePath}");
            Console.WriteLine();

            using (StreamReader inputFile = new StreamReader(Path.Combine(filePath)))
            {
                while (!inputFile.EndOfStream)
                {
                    string? line = inputFile.ReadLine();

                    if (line != null)
                    {
                        inputLines.Add(line);
                        Console.WriteLine($"VB [{lineNumber}]: {line}");
                        lineNumber++;
                    }
                }
            }

            Console.WriteLine();

            return inputLines;
        }

        private static void Save(string outputPath, List<string> outputLines)
        {
            //string docPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            using (StreamWriter outputFile = new StreamWriter(Path.Combine(outputPath), false))
            {
                foreach (string line in outputLines)
                {
                    outputFile.WriteLine(line);
                }
            }
        }
    }
}