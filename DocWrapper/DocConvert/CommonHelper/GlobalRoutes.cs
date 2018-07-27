using System;

namespace CommonHelper
{
   public class GlobalRoutes
    {
       public GlobalRoutes()
       {
       }

       public GlobalRoutes(String diffresultRoute, String diffEndRoute, String diffStartRoute, String outputConvertRoute, String inputConvertRoute, String diffDocRoute, String convertDocRoute, String diffTempRoute, Int32 timeDelete, String logRoute)
       {
           DiffresultRoute = diffresultRoute;
           DiffEndRoute = diffEndRoute;
           DiffStartRoute = diffStartRoute;
           OutputConvertRoute = outputConvertRoute;
           InputConvertRoute = inputConvertRoute;
           DiffDocRoute = diffDocRoute;
           ConvertDocRoute = convertDocRoute;
           DiffTempRoute = diffTempRoute;
           TimeDelete = timeDelete;
           LogRoute = logRoute;
       }

       public String ConvertDocRoute { get; set; }
       public String DiffDocRoute { get; set; }

       public String InputConvertRoute { get; set; }
       public String OutputConvertRoute { get; set; }
       public String DiffStartRoute { get; set; }
       public String DiffEndRoute { get; set; }
       public String DiffresultRoute { get; set; }
       public String DiffTempRoute { get; set; }
       public String LogRoute { get; set; }

       public Int32 TimeDelete { get; set; }


    }
}
