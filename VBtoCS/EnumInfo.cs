using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VBtoCS
{
    internal struct EnumInfo
    {
        private List<string> _accessModifiers = new List<string>();
        private string _name = "";
        private Dictionary<string, int> _values = new Dictionary<string, int>();

        public EnumInfo()
        {
        }

        internal void AddAccessModifier(string mod)
        {
            if (!_accessModifiers.Contains(mod))
            {
                _accessModifiers.Add(mod);
            }
        }

        internal void AddValue(string valueName, int value)
        {
            _values.TryAdd(valueName, value);
        }

        internal void AddName(string name)
        {
            if (_name == "")
            {
                _name = name;
            }
        }
    }
}
