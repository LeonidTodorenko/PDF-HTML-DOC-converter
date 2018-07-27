using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommonHelper
{
    public  class MCFParams
    {
        public MCFParams(String m, String c, String f)
        {
            M = m;
            C = c;
            F = f;
        }

        public MCFParams()
        {
        }

        public String M { get; set; }

        public String C { get; set; }

        public String F { get; set; }

    }
}
