using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http;

namespace DocConvert.Controllers
{
    public class DefaultController : ApiController
    {
        [Route("")]
        public HttpResponseMessage Get()
        {
            String json = @"{ 'name': 'legalthings\/doctools' , 'version':'0.1.0',   'description':'Document conversion and comparison tool'  }";

            HttpResponseMessage result = new HttpResponseMessage();

            result = new HttpResponseMessage(HttpStatusCode.OK);
            result.Content = new StringContent(json);
            result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/json");

            return result;
        }
    }
}
