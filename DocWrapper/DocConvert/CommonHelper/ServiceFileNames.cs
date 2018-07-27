using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommonHelper
{
    public class ServiceFileNames
    {
        public ServiceFileNames(String outputName, String inputName, String outputExtension)
        {
            OutputName = outputName;
            InputName = inputName;
            OutputExtension = outputExtension;
        }

        public ServiceFileNames(String outputName, String inputName)
        {
            OutputName = outputName;
            InputName = inputName;
        }

        public ServiceFileNames()
        {
        }

        public String OutputName { get; set; }
        public String InputName { get; set; }
        public String OutputExtension { get; set; }
    }
}
