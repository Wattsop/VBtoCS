using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VBtoCS
{
    internal struct VariableInfo
    {
        private List<string> _accessModifiers = new List<string>();
        private string _variableType = "object";
        private bool _isArray = false;
        private bool _isNew = false;
        private string _variableName = "";
        private string _assignment = "";

        internal List<string> AccessModifiers
        {
            get { return _accessModifiers; }
            set { _accessModifiers = value; }
        }

        internal string VariableType
        {
            get { return _variableType; }
            set { _variableType = value; }
        }

        internal bool IsArray
        {
            get { return _isArray; }
            set { _isArray = value; }
        }

        internal bool IsNew
        {
            get { return _isNew; }
            set { _isNew = value; }
        }

        internal string Name
        {
            get { return _variableName; }
            set { _variableName = value; }
        }

        internal string Assignment
        {
            get { return _assignment; }
            set { _assignment = value; }
        }

        public VariableInfo()
        {
        }

        internal VariableInfo(string variableName, string variableType)
        {
            _variableName = variableName;
            _variableType = variableType;
        }
    }
}
