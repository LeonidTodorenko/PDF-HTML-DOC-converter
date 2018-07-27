using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DocConvert.Helpers
{
    public class ConvertDocumentAttribute : Attribute
    {
        public String MediaType { get; set; }

        public String Extension { get; set; }

        public String ConvertType { get; set; }
    }
}