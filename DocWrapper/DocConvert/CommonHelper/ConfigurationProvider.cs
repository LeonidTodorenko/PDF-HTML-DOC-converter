using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace CommonHelper
{
    public class ConfigurationProvider
    {


        public Configuration Configuration
        {
            get
            {
                String conf = System.Reflection.Assembly.GetEntryAssembly().Location + "\\App.config";
                ExeConfigurationFileMap map = new ExeConfigurationFileMap { ExeConfigFilename = conf };
                Configuration configuration = ConfigurationManager.OpenMappedExeConfiguration(map, ConfigurationUserLevel.None);
                return configuration;
            }
        }



        public String ConvertDocRoute
        {
            get
            {
                String key = "convertDocRoute";
                return Configuration.AppSettings.Settings[key].Value;
            }
        }
    }
}
