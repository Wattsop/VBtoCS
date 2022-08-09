using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VBtoCS
{
    internal struct FunctionInfo
    {
        private List<string> _accessModifiers = new List<string>();
        private string _returnType = "object";
        private string _name = "";
        private List<List<string>> _parameters = new List<List<string>>();
        private Dictionary<string, VariableInfo> _variables = new Dictionary<string, VariableInfo>();

        internal List<string> AccessModifiers
        {
            get { return _accessModifiers; }
            set { _accessModifiers = value; }
        }

        internal string ReturnType
        {
            get { return _returnType; }
            set { _returnType = value; }
        }

        internal string Name
        {
            get { return _name; }
            set { _name = value; }
        }

        internal List<List<string>> Parameters
        {
            get { return _parameters; }
            set { _parameters = value; }
        }

        public FunctionInfo()
        {
        }

        internal FunctionInfo(int outputLineIndex, List<string> accessModifiers, string returnType, string name, List<List<string>> parameters)
        {
            _accessModifiers = accessModifiers;
            _returnType = returnType;
            _name = name;
            _parameters = parameters;
        }

        internal void AddParameter(List<string> parameterOps)
        {
            _parameters.Add(parameterOps);
        }

        internal void AddVariable(VariableInfo variable)
        {
            if (!_variables.ContainsKey(variable.Name))
            {
                _variables.Add(variable.Name, variable);
            }
        }

        internal bool ContainsVariable(string vName)
        {
            return _variables.ContainsKey(vName);
        }

        internal void SetVariableType(string vName, string vType)
        {
            if (!_variables.ContainsKey(vName))
            {
                _variables.Add(vName, new VariableInfo(vName, vType));
            }
            else
            {
                _variables[vName] = new VariableInfo(vName, vType);
            }
        }
    }
}
