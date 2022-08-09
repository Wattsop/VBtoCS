using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VBtoCS
{
    internal struct CombineOpsInfo
    {
        internal List<string> LineOps = new List<string>();
        internal List<string> SkipPadLeft = new List<string>();
        internal List<string> SkipPadRight = new List<string>();
        internal string PadString = " ";
        internal string PrependString = "";
        internal string AppendString = "";

        public CombineOpsInfo()
        {
        }

        internal CombineOpsInfo(List<string> lineOps)
        {
            LineOps = lineOps;
        }

        internal string CombineOps()
        {
            string line = PrependString;

            for (int i = 1; i < LineOps.Count; i++)
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
