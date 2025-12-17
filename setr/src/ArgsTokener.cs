using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Tmm
{
    public class ArgsTokener
    {
        private string Buf { get; set; }
        private int Seq { get; set; }
        public int StartPos { get; set; }
        public int Pos { get; set; }

        public ArgsTokener()
        {
            Reset();
        }

        public void Reset()
        {
            Buf = "";
            StartPos = 0;
            Pos = 0;
        }

        //////////////////////////////////////////////////////////////////////////
        /// 

        public IEnumerable<string> Tokens(string s)
        {
            var sb = new StringBuilder();
            var seq = 0;
            var buf = s + " ";
            for (var i = 0; i < buf.Length; i++)
            {
                var c = buf[i];
                switch (seq)
                {
                    case 0:         //idle
                        if (Char.IsWhiteSpace(c)) continue;
                        else if (c == '^') seq = 2;
                        else { sb.Append(c); seq = 1; }
                        if (c == '"') seq = 3;
                        StartPos = i;
                        break;
                    case 1:         //normal
                        if (Char.IsWhiteSpace(c))
                        {
                            Pos = i;
                            yield return sb.ToString();
                            sb.Clear();
                            seq = 0;
                        }
                        else if (c == '^') seq = 2;
                        else sb.Append(c);
                        if (c == '"') seq = 3;
                        break;
                    case 2:         //escape
                        sb.Append(c);
                        seq = 1;
                        break;
                    case 3:         //string
                        if (c == '^') seq = 4;
                        else sb.Append(c);
                        if (c == '"') seq = 1;
                        break;
                    case 4:         //string-escape
                        sb.Append(c);
                        seq = 3;
                        break;
                    default:
                        break;
                }
            }
            if (sb.Length > 0) yield return sb.ToString();
        }
    }
}
