using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;


namespace CommonHelper
{
    public class NameConstructor
    {
        /// <summary>
        /// Gets the global routes.
        /// </summary>
        /// <returns></returns>
        public GlobalRoutes GetGlobalRoutes()
        {
            GlobalRoutes globalRoutes = new GlobalRoutes();

            Configuration configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            globalRoutes.ConvertDocRoute = configuration.AppSettings.Settings["convertDocRoute"].Value;
            globalRoutes.DiffDocRoute = configuration.AppSettings.Settings["diffDocRoute"].Value;

            globalRoutes.InputConvertRoute = configuration.AppSettings.Settings["inputConvertRoute"].Value;
            globalRoutes.OutputConvertRoute = configuration.AppSettings.Settings["OutputConvertRoute"].Value;
            globalRoutes.DiffStartRoute = configuration.AppSettings.Settings["DiffStartRoute"].Value;
            globalRoutes.DiffEndRoute = configuration.AppSettings.Settings["DiffEndRoute"].Value;
            globalRoutes.DiffresultRoute = configuration.AppSettings.Settings["DiffresultRoute"].Value;
            globalRoutes.DiffTempRoute = configuration.AppSettings.Settings["DiffTempRoute"].Value;
            globalRoutes.LogRoute = configuration.AppSettings.Settings["LogRoute"].Value;
            globalRoutes.TimeDelete = Int32.Parse(configuration.AppSettings.Settings["TimeDelete"].Value);

            return globalRoutes;
        }

        /// <summary>
        /// Gets the global routes service.
        /// </summary>
        /// <returns></returns>
        public GlobalRoutes GetGlobalRoutesService()
        {
            GlobalRoutes globalRoutes = new GlobalRoutes();

            var configuration = ConfigurationManager.AppSettings;
            globalRoutes.ConvertDocRoute = configuration["convertDocRoute"];
            globalRoutes.DiffDocRoute = configuration["diffDocRoute"];

            globalRoutes.InputConvertRoute = configuration["inputConvertRoute"];
            globalRoutes.OutputConvertRoute = configuration["OutputConvertRoute"];
            globalRoutes.DiffStartRoute = configuration["DiffStartRoute"];
            globalRoutes.DiffEndRoute = configuration["DiffEndRoute"];
            globalRoutes.DiffresultRoute = configuration["DiffresultRoute"];
            globalRoutes.DiffTempRoute = configuration["DiffTempRoute"];
            globalRoutes.LogRoute = configuration["LogRoute"];

            return globalRoutes;
        }






        /// <summary>
        /// /generate names and M C F params
        /// </summary>
        /// <param name="inputExtension">The input extension.</param>
        /// <param name="outputExtension">The output extension.</param>
        /// <returns></returns>
        public ConvertParametersForService GenerateName(String inputExtension, String outputExtension)
        {




            ConvertParametersForService convertParametersForService = new ConvertParametersForService();

            convertParametersForService.McfParams.M = "1";
            convertParametersForService.McfParams.F = "";

            switch (outputExtension)
            {
                case "html":
                    convertParametersForService.McfParams.C = "8";
                    break;
                case "docx":
                    convertParametersForService.McfParams.C = "12";
                    break;
                case "pdf":
                    convertParametersForService.McfParams.C = "17";
                    break;
            }

            convertParametersForService.ServiceFileNames.InputName = inputExtension + "_to_" + outputExtension + "_params_" + "cpr" + convertParametersForService.McfParams.C + "fpr" + convertParametersForService.McfParams.F + "mpr" + convertParametersForService.McfParams.M + "_" + Path.GetRandomFileName();
            convertParametersForService.ServiceFileNames.OutputName = convertParametersForService.ServiceFileNames.InputName + "_converted";

            return convertParametersForService;
        }

     
        /// <summary>
        /// generate output name and M C F params 
        /// </summary>
        /// <param name="inputName">Name of the input.</param>
        /// <returns></returns>
        public ConvertParametersForService ParseName(String inputName)
        {
            ConvertParametersForService convertParametersForService = new ConvertParametersForService();
            convertParametersForService.ServiceFileNames.OutputName = GetOutputName(inputName);
            convertParametersForService.ServiceFileNames.OutputExtension = GetExt(inputName);
            convertParametersForService.McfParams.C = GetCparam(inputName);
            convertParametersForService.McfParams.M = GetMparam(inputName);
            convertParametersForService.McfParams.F = GetFparam(inputName);

            return convertParametersForService;
        }

        /// <summary>
        /// Cenereates the name of the diff doc.
        /// </summary>
        /// <param name="originalName">Name of the original.</param>
        /// <param name="modifiedName">Name of the modified.</param>
        /// <param name="originalExtension">The original extension.</param>
        /// <param name="modifiedExtension">The modified extension.</param>
        /// <returns></returns>
        public ConvertParametersForDiff CenereateDiffDocName(String originalName, String modifiedName, String originalExtension, String modifiedExtension)
        {
            ConvertParametersForDiff convertParametersForDiff = new ConvertParametersForDiff();
            String genName = Path.GetRandomFileName();
            convertParametersForDiff.OriginalDiffName = "original" + "_" + genName + "." + GetExtDiff(originalName);
            convertParametersForDiff.ModifiedDiffName = "modified" + "_" + genName + "." + GetExtDiff(modifiedName);
            convertParametersForDiff.DiffResult = "diffResult" + "_" + genName + ".doc";
            convertParametersForDiff.DiffResultEnd = "diffResult" + "_" + genName + ".html";

            return convertParametersForDiff;
        }

        /// <summary>
        /// Parses the name of the diff.
        /// </summary>
        /// <param name="inputName">Name of the input.</param>
        /// <returns></returns>
        public ConvertParametersForDiff ParseDiffName(String inputName)
        {
            ConvertParametersForDiff convertParametersForDiff = new ConvertParametersForDiff();

            if (CheckIfFileOriginal(inputName))
            {
                convertParametersForDiff.ModifiedDiffName = convertParametersForDiff.PartnerDiffName = GenerateModifiedName(inputName);
                convertParametersForDiff.OriginalDiffName = inputName;
            }
            else
            {
                convertParametersForDiff.OriginalDiffName = convertParametersForDiff.PartnerDiffName = GenerateOriginalName(inputName);
                convertParametersForDiff.ModifiedDiffName = inputName;
            }

            String genName = GetGenName(inputName);
            convertParametersForDiff.DiffResult = "diffResult" + "_" + genName + ".doc";
            convertParametersForDiff.DiffResultEnd = "diffResult" + "_" + genName + ".html";


            return convertParametersForDiff;
        }

        #region Private

        private String GetGenName(String inputName)
        {
            Int32 ind1 = inputName.IndexOf('_');
            Int32 ind2 = inputName.LastIndexOf('.');
            return inputName.Substring(ind1 + 1, ind2 - ind1 - 1);
        }

        private String GenerateModifiedName(String input)
        {
            return input.Replace("original", "modified");
        }

        private String GenerateOriginalName(String input)
        {
            return input.Replace("modified", "original");
        }

        private Boolean CheckIfFileOriginal(String input)
        {
            if (input.Substring(0, 8) == "original")
            {
                return true;
            }

            return false;
        }

        private String GetOutputName(String name)
        {
            Int32 ind1 = name.LastIndexOf('.');
            name = name.Substring(0, ind1);
            return name;
        }

        private String GetCparam(String name)
        {
            Int32 ind1 = name.IndexOf("cpr");
            Int32 ind2 = name.IndexOf("fpr");
            name = name.Substring(ind1 + 3, ind2 - ind1 - 3);
            return name;
        }

        private String GetMparam(String name)
        {

            Int32 ind1 = name.IndexOf("mpr");
            name = name.Substring(ind1 + 3, 1);

            return name;
        }

        private String GetFparam(String name)
        {
            Int32 ind1 = name.IndexOf("fpr");
            Int32 ind2 = name.IndexOf("mpr");
            name = name.Substring(ind1 + 3, ind2 - ind1 - 3);
            if (name == "0")
            {
                name = "";
            }
            return name;
        }


        private String GetExt(String name)
        {
            Int32 ind1 = name.IndexOf("_to_");
            Int32 ind2 = name.IndexOf("_params_");
            name = name.Substring(ind1 + 4, ind2 - ind1 - 4);

            return name;
        }

        private static String GetExtDiff(String name)
        {
            Int32 ind1 = name.LastIndexOf('.');
            name = name.Substring(ind1 + 1, name.Length - ind1 - 1);
            return name;
        }

        #endregion Private
    }
}
