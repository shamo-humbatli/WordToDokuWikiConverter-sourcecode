using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MicroMWordLib.WordOperations
{
    public abstract class WParameters
    {
        private static System.Reflection.Missing ref_Missing;

        public static Missing Missing { get => ref_Missing; }
    }
}
