﻿using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using System.Web;
using CommonHelper;
using System.Windows.Markup;
using DocConvert.Models;
using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http;
using DocConvert.Helpers;
using HtmlAgilityPack;
using HttpMultipartParser;
using Novacode;
using Formatting = Novacode.Formatting;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;

namespace DocConvert.Controllers
{
    public class ConvertController : ApiController
    {
        public String Get()
        {
            return "test";
        }

        public Boolean FlagCreated { get; set; }
        public FileStream CreateddFile { get; set; }
        public String OutputDirecroy { get; set; }
        public MediaTypeEnum AcceptTypePubl { get; set; }
        public Boolean ConvertingImage { get; set; }
        public String GlobalInputExt { get; set; }

        static readonly NameConstructor nameConstructorRoutes = new NameConstructor();
        readonly GlobalRoutes globalRoutes = nameConstructorRoutes.GetGlobalRoutesService();

        public HttpResponseMessage Post()
        {
            HttpResponseMessage resultError = new HttpResponseMessage();

            MediaTypeHeaderValue content = Request.Content.Headers.ContentType;
            if (content == null)
            {
                resultError = new HttpResponseMessage(HttpStatusCode.NotAcceptable); // 406
                resultError.Content = new StringContent("Unable to convert to specified format");
                resultError.Content.Headers.ContentType = new MediaTypeHeaderValue("text/plain");

                return resultError;
            }

            MediaTypeEnum contentType = ResolveMediaType(content.MediaType);
            if (contentType == MediaTypeEnum.Empty)
            {
                resultError = new HttpResponseMessage(HttpStatusCode.NotAcceptable); // 406
                resultError.Content = new StringContent("Unable to convert to specified format");
                resultError.Content.Headers.ContentType = new MediaTypeHeaderValue("text/plain");

                return resultError;
            }

            MediaTypeWithQualityHeaderValue accept = Request.Headers.Accept.FirstOrDefault();
            if (accept == null)
            {
                resultError = new HttpResponseMessage(HttpStatusCode.NotAcceptable); // 406
                resultError.Content = new StringContent("Unable to convert to specified format");
                resultError.Content.Headers.ContentType = new MediaTypeHeaderValue("text/plain");

                return resultError;
            }

            MediaTypeEnum acceptType = ResolveMediaType(accept.MediaType);
            if (acceptType == MediaTypeEnum.Empty)
            {
                resultError = new HttpResponseMessage(HttpStatusCode.NotAcceptable); // 406
                resultError.Content = new StringContent("Unable to convert to specified format");
                resultError.Content.Headers.ContentType = new MediaTypeHeaderValue("text/plain");

                return resultError;
            }

            AcceptTypePubl = acceptType;

            Byte[] fileStream = null;

            if (Request.Content != null)
            {
                fileStream = GetFileStream(Request.Content);
            }
            else
            {
                resultError = new HttpResponseMessage(HttpStatusCode.UnsupportedMediaType); // 415
                resultError.Content = new StringContent("Unable to convert the supplied document");
                resultError.Content.Headers.ContentType = new MediaTypeHeaderValue("text/plain");

                return resultError;
            }

            NameConstructor nameConstructor = new NameConstructor();



            String inputExtension = "";
            MediaTypeEnum multipartType = new MediaTypeEnum();
            String outputExtension = acceptType.Attribute<ConvertDocumentAttribute>().Extension;

            if (contentType == MediaTypeEnum.MultiPart)
            {
                multipartType = ResolveMediaType(GlobalInputExt);
                inputExtension = multipartType.Attribute<ConvertDocumentAttribute>().Extension;
            }
            else
            {
                inputExtension = contentType.Attribute<ConvertDocumentAttribute>().Extension;
            }

            ConvertParametersForService convertParametersForService = nameConstructor.GenerateName(inputExtension,
                                                                                                   outputExtension);

            //  save input
            String inputName = convertParametersForService.ServiceFileNames.InputName;
            String input = String.Format("{2}{0}.{1}", inputName, inputExtension, globalRoutes.InputConvertRoute);

            //  save output
            String outputName = convertParametersForService.ServiceFileNames.OutputName;
            String output = String.Format("{2}{0}.{1}", outputName, outputExtension, globalRoutes.OutputConvertRoute);
            OutputDirecroy = output;

            //before save parse html and save images
            if (contentType == MediaTypeEnum.Html || multipartType == MediaTypeEnum.Html)
            {
                ParseHtmlAndCreateInages(input, fileStream, inputName, inputExtension);
            }
            else
            {
                //save file
                FileStream wFile = new FileStream(input, FileMode.Create);
                wFile.Write(fileStream, 0, fileStream.Length);
                wFile.Close();
            }


            Stream resultStream = null;
            HttpResponseMessage result = new HttpResponseMessage();
            try
            {
                resultStream = Work(outputName + "." + outputExtension);
                result = new HttpResponseMessage(HttpStatusCode.OK);
                result.Content = new StreamContent(resultStream);
                result.Content.Headers.ContentType = new MediaTypeHeaderValue(accept.MediaType);
            }
            catch (Exception ex)
            {
                result = new HttpResponseMessage(HttpStatusCode.UnsupportedMediaType); // 415
                result.Content = new StringContent("Unable to convert the supplied document");
                result.Content.Headers.ContentType = new MediaTypeHeaderValue("text/plain");

                LogError(ex);
            }



            return result;
        }

        #region Private

        /// <summary>
        /// Parses the HTML and create inages.
        /// </summary>
        /// <param name="input">The input.</param>
        private void ParseHtmlAndCreateInages(String input, Byte[] fileStream, String inputName, String inputExtension)
        {
            String inputTemp = String.Format("{2}{0}.{1}", inputName, inputExtension, globalRoutes.DiffTempRoute);
            FileStream wFile = new FileStream(inputTemp, FileMode.Create);
            wFile.Write(fileStream, 0, fileStream.Length);
            wFile.Close();

            WhaitFileFree(inputTemp);

            HtmlWeb hw = new HtmlWeb();
            HtmlAgilityPack.HtmlDocument htmlDoc = hw.Load(inputTemp);
            if (htmlDoc.DocumentNode != null)
            {
                htmlDoc.DocumentNode.InnerHtml = ClearFromGarbage(htmlDoc.DocumentNode.InnerHtml);

                if (htmlDoc.DocumentNode.SelectNodes("//img") != null)
                {

                    foreach (HtmlNode link in htmlDoc.DocumentNode.SelectNodes("//img"))
                    {
                        String currSrc = link.Attributes["src"].Value;
                        if (currSrc.Contains("base64"))
                        {
                            String imgName = Path.GetRandomFileName();
                            String newImageName = globalRoutes.InputConvertRoute + imgName + ".bmp";
                            SaveByteArrayAsImage(newImageName, currSrc.Substring(22));
                            link.Attributes["src"].Value = imgName + ".bmp";
                        }
                    }
                }
            }

            if (htmlDoc.DocumentNode != null)
            {
                var t = new HtmlDocument();
                t.LoadHtml(htmlDoc.DocumentNode.InnerHtml);
                t.Save(input);
            }
            else
            {
                htmlDoc.Save(input);
            }

        }

        /// <summary>
        /// Creates valid lists.
        /// </summary>
        /// <param name="htmlDocument">The HTML document.</param>
        /// <returns></returns>
        private HtmlDocument CreateLists(HtmlDocument htmlDocument)
        {
            // todo: for letters a,b,c and I,II,III 
            // todo: sublists with levels

            Boolean startUl = false;
            Boolean startOl = false;

            if (htmlDocument.DocumentNode.SelectNodes("//p") != null)
            {
              
                Int32 previousNumber = -1;
                Boolean wasDeleteNum = false;
                HtmlNodeCollection pNodes = htmlDocument.DocumentNode.SelectNodes("//p");
                for (int i = 0; i < pNodes.Count; i++)   //foreach (HtmlNode pNode in htmlDocument.DocumentNode.SelectNodes("//p"))
                {

                    if (pNodes[i].OuterHtml.Contains("Symbol"))
                    {
                        if (pNodes[i].InnerText.StartsWith("&#61623;"))
                        {
                            String newNodeStr = "<ul><li>";
                            newNodeStr = newNodeStr + pNodes[i].OuterHtml.Replace("&#61623;", "") + "</li></ul>";
                            var newNode = HtmlNode.CreateNode(newNodeStr);
                            pNodes[i].ParentNode.ReplaceChild(newNode, pNodes[i]);
                        }

                        if (pNodes[i].InnerText.StartsWith("·"))
                        {
                            String newNodeStr = "<ul><li>";
                            newNodeStr = newNodeStr + pNodes[i].OuterHtml.Replace(">·<", "><") + "</li></ul>";
                            var newNode = HtmlNode.CreateNode(newNodeStr);
                            pNodes[i].ParentNode.ReplaceChild(newNode, pNodes[i]);
                        }

                        if (pNodes[i].InnerText.StartsWith("."))
                        {
                            String newNodeStr = "<ul><li>";
                            newNodeStr = newNodeStr + pNodes[i].OuterHtml.Replace(">.<", "><") + "</li></ul>";
                            var newNode = HtmlNode.CreateNode(newNodeStr);
                            pNodes[i].ParentNode.ReplaceChild(newNode, pNodes[i]);
                        }



                    }

                    Boolean numberOperated = false;
                    Int32 currNumber = -1;
                    if (pNodes[i].InnerText.Length > 2)
                    {
                        currNumber = GetFirstNum(pNodes[i].InnerText);
                    }
                    if (currNumber > 0)
                    {

                        Int32 numLen = currNumber.ToString().Length;

                        String checkPoint = pNodes[i].InnerText.Substring(numLen, 1);
                        if (checkPoint == ".")
                        {
                            numberOperated = true;
                            if (previousNumber >= 0)
                            {
                                if (currNumber - previousNumber == 1)
                                {
                                    if (!wasDeleteNum)
                                    {
                                        String oldNodeStr = "<ol><li>";
                                        Int32 oldNumberIndex = pNodes[i - 1].OuterHtml.IndexOf(">" + previousNumber.ToString() + ".");
                                        oldNodeStr = oldNodeStr + pNodes[i - 1].OuterHtml.Remove(oldNumberIndex + 1, previousNumber.ToString().Length + 1) + "</li></ol>";
                                        var oldNode = HtmlNode.CreateNode(oldNodeStr);
                                        pNodes[i - 1].ParentNode.ReplaceChild(oldNode, pNodes[i - 1]);
                                    }

                                    String newNodeStr = "<ol><li>";
                                    Int32 newNumberIndex = pNodes[i].OuterHtml.IndexOf(">" + currNumber.ToString() + ".");
                                    newNodeStr = newNodeStr + pNodes[i].OuterHtml.Remove(newNumberIndex + 1, currNumber.ToString().Length + 1) + "</li></ol>";
                                    var newNode = HtmlNode.CreateNode(newNodeStr);
                                    pNodes[i].ParentNode.ReplaceChild(newNode, pNodes[i]);

                                    wasDeleteNum = true;
                                    previousNumber = currNumber;
                                }
                            }
                            else
                            {
                                previousNumber = currNumber;
                                wasDeleteNum = false;
                            }
                        }
                    }
                }
            }

            if (htmlDocument.DocumentNode.InnerHtml.Contains("</ul>\r\n\r\n<ul>"))
            {
                htmlDocument.DocumentNode.InnerHtml = htmlDocument.DocumentNode.InnerHtml.Replace("</ul>\r\n\r\n<ul>", "");
            }
            if (htmlDocument.DocumentNode.InnerHtml.Contains("</ol>\r\n\r\n<ol>"))
            {
                htmlDocument.DocumentNode.InnerHtml = htmlDocument.DocumentNode.InnerHtml.Replace("</ol>\r\n\r\n<ol>", "");
            }
            return htmlDocument;
        }

        /// <summary>
        /// Gets the first num.
        /// </summary>
        /// <param name="input">The input.</param>
        /// <returns></returns>
        private static Int32 GetFirstNum(String input)
        {
            String final = "0"; //if there's nothing, it'll return -1
            foreach (Char c in input) //loop the string
            {
                try
                {
                    Convert.ToInt32(c.ToString()); //if it can convert
                    final += c.ToString(); //add to final string
                }
                catch (FormatException) //if NaN
                {
                    break; //break out of loop
                }
            }

            return Convert.ToInt32(final); //return the int
        }

        /// <summary>
        /// Clears from bullets.
        /// </summary>
        /// <param name="html">The HTML.</param>
        /// <returns></returns>
        private String ClearFromBullets(String html)
        {
            String startBullet = "<![if !supportLists]>";
            String endBullet = "<![endif]>";

            Int32 startBulletIndex = html.IndexOf(startBullet);
            Int32 endBulletIndex = html.IndexOf(endBullet);

            if (startBulletIndex > 0 || endBulletIndex > 0)
            {
                html = html.Replace(startBullet, "");
                html = html.Replace(endBullet, "");
            }

            // fix styles for symbols
            if (html.IndexOf("Segoe UI Symbol") > 0)
            {
                html = html.Replace("Segoe UI Symbol", "Symbol");
            }

            return html;
        }

        /// <summary>
        /// Clears from garbage.
        /// </summary>
        /// <param name="html">The HTML.</param>
        /// <returns></returns>
        private String ClearFromGarbage(String html)
        {
            if (html.IndexOf("v:shapes") > 0)
            {
                html = html.Replace("v:shapes", "datgar");
            }

            if (html.IndexOf("v:shape") > 0)
            {
                html = html.Replace("v:shape", "datgar");
            }


            //mso-list make ul corrupted
            if (html.IndexOf("mso-list") > 0)
            {
                html = html.Replace("mso-list", "morbo-list");
            }

            return html;
        }

        /// <summary>
        /// Saves the byte array as image.
        /// </summary>
        /// <param name="fullOutputPath">The full output path.</param>
        /// <param name="base64String">The base64 string.</param>
        private void SaveByteArrayAsImage(string fullOutputPath, string base64String)
        {
            byte[] bytes = Convert.FromBase64String(base64String);

            FileStream wFile = new FileStream(fullOutputPath, FileMode.Create);
            wFile.Write(bytes, 0, bytes.Length);
            wFile.Close();
        }

        /// <summary>
        /// Checks and convert image.
        /// </summary>
        /// <returns>converted FileStream</returns>
        private void CheckAndConvertImage()
        {
            HtmlWeb hw = new HtmlWeb();
            HtmlAgilityPack.HtmlDocument htmlDoc = hw.Load(OutputDirecroy);
            if (htmlDoc.DocumentNode != null)
            {
                // clearing bullets
                htmlDoc.DocumentNode.InnerHtml = ClearFromBullets(htmlDoc.DocumentNode.InnerHtml);

                // create normal lists
                htmlDoc = CreateLists(htmlDoc);


                // images to base64
                if (htmlDoc.DocumentNode.SelectNodes("//img") != null)
                {
                    foreach (HtmlNode link in htmlDoc.DocumentNode.SelectNodes("//img"))
                    {

                        String currSrc = link.Attributes["src"].Value;
                        currSrc = currSrc.Replace("%20", " ");
                        if (currSrc.Contains("file:"))
                        {
                            currSrc = currSrc.Substring(8); // for full patch
                        }
                        else
                        {
                            currSrc = String.Format("{1}{0}", currSrc, globalRoutes.OutputConvertRoute); // for part of patch
                        }

                        link.Attributes["src"].Value = MakeImageSrcData(currSrc);
                    }
                }
            }

            if (htmlDoc.DocumentNode != null)
            {
                var t = new HtmlDocument();
                t.LoadHtml(htmlDoc.DocumentNode.InnerHtml);
                t.Save(OutputDirecroy);
            }
            else
            {
                htmlDoc.Save(OutputDirecroy);
            }


            WhaitFileFree(OutputDirecroy);

            CreateddFile = new FileStream(OutputDirecroy, FileMode.Open);


        }


        /// <summary>
        /// Logs the error.
        /// </summary>
        /// <param name="ex">The ex.</param>
        private void LogError(Exception ex)
        {
            String filename = globalRoutes.LogRoute + "log.txt"; 
            if (!File.Exists(filename))
            {
                var fs = File.Create(filename);
                fs.Close();
            }
            if (File.ReadAllBytes(filename).Length >= 100 * 1024 * 1024) // (100mB) File to big? Create new
            {
                var fs = File.Create(filename);
                fs.Close();
            }

            String errorText = "Some error occured - " + DateTime.Now + "-" + "Message:" + ex.Message + "InnerException:" + ex.InnerException + "StackTrace:" + ex.StackTrace;
            StreamWriter log = File.AppendText(filename);
            log.WriteLine(errorText);
            log.WriteLine();
            log.Close();
        }

        
        /// <summary>
        /// convert src to base64 -    <img src="data:image/jpeg;base64,[data]">
        /// </summary>
        /// <param name="filename">The filename.</param>
        /// <returns></returns>
        private string MakeImageSrcData(string filename)
        {
            FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read);
            byte[] filebytes = new byte[fs.Length];
            fs.Read(filebytes, 0, Convert.ToInt32(fs.Length));
            fs.Close();
            return "data:image/png;base64," + Convert.ToBase64String(filebytes, Base64FormattingOptions.None);
        }

        private Byte[] GetFileStream(HttpContent httpContent)
        {
            Byte[] FileStream = null;
            if (httpContent.Headers.ContentType.MediaType != "multipart/form-data")
            {
                FileStream = httpContent.ReadAsByteArrayAsync().Result;
            }
            else
            {
                MultipartFormDataParser parser = new MultipartFormDataParser(httpContent.ReadAsStreamAsync().Result);
                foreach (var currentFile in parser.Files)
                {
                    if (currentFile.Name == "file")
                    {
                        GlobalInputExt = currentFile.ContentType;
                        FileStream = ToByteArray(currentFile.Data);
                    }
                }
            }
            return FileStream;
        }

        private static byte[] ToByteArray(Stream stream)
        {
            byte[] buffer = new byte[32768];
            using (MemoryStream ms = new MemoryStream())
            {
                while (true)
                {
                    int read = stream.Read(buffer, 0, buffer.Length);
                    if (read <= 0)
                        return ms.ToArray();
                    ms.Write(buffer, 0, read);
                }
            }
        }

        private MediaTypeEnum ResolveMediaType(String value)
        {
            return Enum.GetValues(typeof(MediaTypeEnum))
                       .OfType<MediaTypeEnum>()
                       .Where(e => e.Attribute<ConvertDocumentAttribute>().MediaType == value)
                       .ToList()
                       .FirstOrDefault();
        }

        private Stream Work(String wholeName)
        {

            Int32 countDown = 1;

            while (!File.Exists(globalRoutes.OutputConvertRoute + wholeName))
            {
                if (countDown == 3000) //  timeout 5 min  
                {
                    throw new TimeoutException();  
                }
                Thread.Sleep(100);
                countDown++;

            }

            OnChanged();

            return CreateddFile;
        }

        private void OnChanged()
        {
            if (WhaitFileFree(OutputDirecroy)) 
            {
                if (AcceptTypePubl == MediaTypeEnum.Html)
                {
                    CheckAndConvertImage();
                }
                else // PDF  Docx
                {
                    CreateddFile = new FileStream(OutputDirecroy, FileMode.Open);
                }
            }
        }


        /// <summary>
        /// Whaits the file free.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        private static Boolean WhaitFileFree(String filePath)
        {
            Int32 countDown = 1;
            Boolean checkLock = true;
            while (checkLock)
            {
                checkLock = IsFileLocked(filePath);
                Thread.Sleep(100);
                if (countDown == 1800) // todo: обсудить timeout 3 min
                {
                    throw new TimeoutException();
                }
                countDown++;
            }
            return true;
        }

        /// <summary>
        /// Determines whether [is file locked] [the specified file path].
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <returns>
        ///   <c>true</c> if [is file locked] [the specified file path]; otherwise, <c>false</c>.
        /// </returns>
        private static bool IsFileLocked(String filePath)
        {
            FileStream stream = null;
            try
            {
                stream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException e)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }
            //file is not locked
            return false;
        }



        #endregion Private
    }
}
