using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SHMicroMWordLib.WordText
{
    public class WTextPartProperties
    {
        private bool prp_Bold = false;
        private bool prp_Italic = false;
        private bool prp_Underline = false;

        public bool Bold { get => prp_Bold; set => prp_Bold = value; }
        public bool Italic { get => prp_Italic; set => prp_Italic = value; }
        public bool Underline { get => prp_Underline; set => prp_Underline = value; }
    }
}
