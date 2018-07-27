using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;

namespace DocConvert
{
    public static class WebApiConfig
    {
        /// <summary>
        /// Registers the specified config.
        /// </summary>
        /// <param name="config">The config.</param>
        public static void Register(HttpConfiguration config)
        {
            // Web API routes
            config.MapHttpAttributeRoutes();

            config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "{controller}/{id}",
                defaults: new { id = RouteParameter.Optional }
            );
        }
    }
}
