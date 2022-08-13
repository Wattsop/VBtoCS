using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VBtoCS
{
    internal struct ClassInfo
    {
        private string _name = "";
        private Dictionary<string, VariableInfo> _variables = new Dictionary<string, VariableInfo>();

        public ClassInfo()
        {
        }

        internal ClassInfo(string name)
        {
            _name = name;
        }

        internal void AddVariable(VariableInfo vInfo)
        {
            _variables.Add(vInfo.Name, vInfo);
        }
    }
}
