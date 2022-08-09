using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VBtoCS
{
    internal struct ClassInfo
    {
        string _name;
        private Dictionary<string, VariableInfo> _variables = new Dictionary<string, VariableInfo>();

        internal ClassInfo(string name)
        {
            _name = name;
        }
    }
}
