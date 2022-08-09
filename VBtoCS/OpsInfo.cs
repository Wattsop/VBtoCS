using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VBtoCS
{
    internal struct OpsInfo
    {
        internal List<string> LineOps = new List<string>();
        internal List<string> SkipPadLeft = new List<string>();
        internal List<string> SkipPadRight = new List<string>();
        internal string PadString = " ";
        internal string PrependString = "";
        internal string AppendString = "";

        public OpsInfo()
        {
        }

        internal OpsInfo(List<string> lineOps, bool defaultPadding = false)
        {
            LineOps = lineOps;

            if (defaultPadding)
            {
                SkipPadLeft.Add("(");
                SkipPadRight.Add("(");

                SkipPadLeft.Add(")");
                SkipPadRight.Add(")");
            }
        }

        internal string CombineOps()
        {
            if (PrependString != "")
            {
                LineOps.Insert(0, PrependString);
            }

            string line = "";

            for (int i = 0; i < LineOps.Count; i++)
            {
                string nextOp = "";

                if (i + 1 < LineOps.Count)
                {
                    nextOp = LineOps[i + 1];
                }

                line += LineOps[i];

                if (!SkipPadRight.Contains(LineOps[i]) && !SkipPadLeft.Contains(nextOp))
                {
                    line += " ";
                }
            }

            line = line.Trim();
            line += AppendString;

            return line;
        }
    }
}
