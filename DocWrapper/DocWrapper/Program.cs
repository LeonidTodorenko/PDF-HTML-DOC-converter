using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Linq;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using CommonHelper;
using System.Configuration;
using Microsoft.Office.Interop.Word;


namespace DocWrapper
{
    class Program
    {
        static readonly NameConstructor NameConstructorRoutes = new NameConstructor();
        static readonly GlobalRoutes globalRoutes = NameConstructorRoutes.GetGlobalRoutes();

        static void Main(string[] args)
        {
            try
            {
                Thread convertListen = new Thread(ConvertListen);
                convertListen.Start();

                Thread diffStart = new Thread(DiffStart);
                diffStart.Start();

                Thread diffEnd = new Thread(DiffEnd);
                diffEnd.Start();

                Thread diffEndResult = new Thread(DiffEndResult);
                diffEndResult.Start();

                Thread threadClean = new Thread(CleanAll);
                threadClean.Start();
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
        }

        /// <summary>
        /// Cleans all.
        /// </summary>
        private static void CleanAll()
        {
            try
            {
                while (true)
                {

                    //// Temp
                    System.IO.DirectoryInfo directoryTemp = new DirectoryInfo(globalRoutes.DiffTempRoute);
                    var filesTemp = directoryTemp.GetFiles();

                    foreach (var file in filesTemp)
                    {
                        var createTime = DateTime.Now.Subtract(file.CreationTime).TotalHours;
                        if (createTime >= globalRoutes.TimeDelete)
                        { file.Delete(); }
                    }

                    //////////////// Input

                    System.IO.DirectoryInfo directoryInput = new DirectoryInfo(globalRoutes.InputConvertRoute);
                    var filesInput = directoryInput.GetFiles();

                    foreach (var file in filesInput)
                    {
                        var createTime = DateTime.Now.Subtract(file.CreationTime).TotalHours;
                        if (createTime >= globalRoutes.TimeDelete)
                        { file.Delete(); }
                    }

                    //////////////// Output

                    System.IO.DirectoryInfo directoryOutput = new DirectoryInfo(globalRoutes.OutputConvertRoute);
                    var filesOutput = directoryOutput.GetFiles();

                    foreach (var file in filesOutput)
                    {
                        var createTime = DateTime.Now.Subtract(file.CreationTime).TotalHours;
                        if (createTime >= globalRoutes.TimeDelete)
                        { file.Delete(); }
                    }


                    //////////////// DiffStart

                    System.IO.DirectoryInfo directoryDiffStart = new DirectoryInfo(globalRoutes.DiffStartRoute);
                    var filesDiffStart = directoryDiffStart.GetFiles();

                    foreach (var file in filesDiffStart)
                    {
                        var createTime = DateTime.Now.Subtract(file.CreationTime).TotalHours;
                        if (createTime >= globalRoutes.TimeDelete)
                        { file.Delete(); }
                    }


                    //////////////// DiffEnd

                    System.IO.DirectoryInfo directoryDiffEnd = new DirectoryInfo(globalRoutes.DiffEndRoute);
                    var filesDiffEnd = directoryDiffEnd.GetFiles();

                    foreach (var file in filesDiffEnd)
                    {
                        var createTime = DateTime.Now.Subtract(file.CreationTime).TotalHours;
                        if (createTime >= globalRoutes.TimeDelete)
                        { file.Delete(); }
                    }


                    //////////////// DiffResult

                    System.IO.DirectoryInfo directoryDiffResult = new DirectoryInfo(globalRoutes.DiffresultRoute);
                    var filesDiffResult = directoryDiffResult.GetFiles();

                    foreach (var file in filesDiffResult)
                    {
                        var createTime = DateTime.Now.Subtract(file.CreationTime).TotalHours;
                        if (createTime >= globalRoutes.TimeDelete)
                        { file.Delete(); }
                    }

                    Thread.Sleep(600000); // check every 10 min
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
        }

        /// <summary>
        /// Diffs the end result.
        /// </summary>
        private static void DiffEndResult()
        {
            try
            {
                FileSystemWatcher endDiffWatcher = new FileSystemWatcher();
                endDiffWatcher.Path = globalRoutes.DiffresultRoute;
                endDiffWatcher.Created += OnChangedEndResultDiff;
                endDiffWatcher.EnableRaisingEvents = true;
                while (true)
                {
                    endDiffWatcher.WaitForChanged(WatcherChangeTypes.Created);
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
        }

        /// <summary>
        ///   Listen files for convert.
        /// </summary>
        private static void ConvertListen()
        {
            try
            {
                FileSystemWatcher convertWatcher = new FileSystemWatcher();
                convertWatcher.Path = globalRoutes.InputConvertRoute;
                convertWatcher.Created += OnChangedConvert;
                convertWatcher.EnableRaisingEvents = true;
                while (true)
                {
                    convertWatcher.WaitForChanged(WatcherChangeTypes.Created);
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
        }



        /// <summary>
        /// Diffs the start.
        /// </summary>
        private static void DiffStart()
        {
            try
            {
                FileSystemWatcher startDiffWatcher = new FileSystemWatcher();
                startDiffWatcher.Path = globalRoutes.DiffStartRoute;
                startDiffWatcher.Created += OnChangedStartDiff;
                startDiffWatcher.EnableRaisingEvents = true;
                while (true)
                {
                    startDiffWatcher.WaitForChanged(WatcherChangeTypes.Created);
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
        }

        /// <summary>
        /// Diffs the end.
        /// </summary>
        private static void DiffEnd()
        {
            try
            {
                FileSystemWatcher endDiffWatcher = new FileSystemWatcher();
                endDiffWatcher.Path = globalRoutes.DiffEndRoute;
                endDiffWatcher.Created += OnChangedEndDiff;
                endDiffWatcher.EnableRaisingEvents = true;
                while (true)
                {
                    endDiffWatcher.WaitForChanged(WatcherChangeTypes.Created);
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
        }


        /// <summary>
        /// Called when [changed start diff].
        /// </summary>
        /// <param name="source">The source.</param>
        /// <param name="e">The <see cref="FileSystemEventArgs"/> instance containing the event data.</param>
        private static void OnChangedStartDiff(object source, FileSystemEventArgs e)
        {
            try
            {
                // prevent of garbage
                if (GetExt(e.Name) == "html" || GetExt(e.Name) == "docx" || GetExt(e.Name) == "htm" || GetExt(e.Name) == "doc" || GetExt(e.Name) == "pdf")
                {
                    String currentExt = GetExt(e.Name);
                    String directoryEnd = globalRoutes.DiffEndRoute;

                    if (currentExt == "htm" || currentExt == "html" || currentExt == "pdf")
                    {
                        WorkConsoleDiffConvertToDoc(e.FullPath, directoryEnd + ChangeToDoc(e.Name), currentExt, e.Name);
                    }

                    else
                    {
                        if (WhaitFileFree(e.FullPath))
                        {
                            // todo:  bullets ro default style
                            // ChangeBulletsToDefault(e.FullPath, directoryEnd + e.Name, e.Name);
                            File.Copy(e.FullPath, directoryEnd + e.Name, true);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
        }

        //// todo:  bullets ro default style
        //private static void ChangeBulletsToDefault(String input, String output, String name)
        //{
        //    String directoryTemp = globalRoutes.DiffTempRoute + name;
        //    BulletsToDefault(input, directoryTemp);
        //    if (WhaitFileFree(directoryTemp))
        //    {
        //        File.Copy(directoryTemp, output, true);
        //    }
        //}

        /// <summary>
        /// Called when [changed end result diff].
        /// </summary>
        /// <param name="source">The source.</param>
        /// <param name="e">The <see cref="FileSystemEventArgs"/> instance containing the event data.</param>
        private static void OnChangedEndResultDiff(object source, FileSystemEventArgs e)
        {
            try
            {
                String directory = globalRoutes.DiffresultRoute;

                if (GetExt(e.Name) == "doc")
                {
                    WorkConsoleDiffConvertToHtml(directory + e.Name, directory + ChangeToHtml(e.Name));
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
        }

        /// <summary>
        /// Called when [changed end diff].
        /// </summary>
        /// <param name="source">The source.</param>
        /// <param name="e">The <see cref="FileSystemEventArgs"/> instance containing the event data.</param>
        private static void OnChangedEndDiff(object source, FileSystemEventArgs e)
        {
            try
            {
                String directoryDiffResult = globalRoutes.DiffresultRoute;
                String directoryEnd = globalRoutes.DiffEndRoute;

                NameConstructor nameConstructor = new NameConstructor();
                ConvertParametersForDiff convertParametersForDiff = nameConstructor.ParseDiffName(e.Name);
                if (File.Exists(directoryEnd + convertParametersForDiff.PartnerDiffName))
                {
                    WorkConsoleDiffDocs(directoryEnd + convertParametersForDiff.OriginalDiffName, directoryEnd + convertParametersForDiff.ModifiedDiffName, directoryDiffResult + convertParametersForDiff.DiffResult);
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
        }

        /// <summary>
        /// Called when [changed convert].
        /// </summary>
        /// <param name="source">The source.</param>
        /// <param name="e">The <see cref="FileSystemEventArgs"/> instance containing the event data.</param>
        private static void OnChangedConvert(object source, FileSystemEventArgs e)
        {
            try
            {
                // prevent of garbage
                if (GetExt(e.Name) == "html" || GetExt(e.Name) == "docx" || GetExt(e.Name) == "htm" || GetExt(e.Name) == "doc" || GetExt(e.Name) == "pdf")
                {
                    NameConstructor nameConstructor = new NameConstructor();
                    ConvertParametersForService convertParametersForService = nameConstructor.ParseName(e.Name);
                    if (GetExt(e.Name) == "html" && convertParametersForService.McfParams.C == "12")
                    {
                        // html - pdf - docx start
                        String input = String.Format("{1}{0}", e.Name, globalRoutes.InputConvertRoute);
                        WorkConsoleConvertHtmlDocStart(input, convertParametersForService);
                    }
                    else if (GetExt(e.Name) == "pdf" && convertParametersForService.McfParams.C == "12")
                    {
                        //   pdf - docx start
                        String input = String.Format("{1}{0}", e.Name, globalRoutes.InputConvertRoute);
                        String output = String.Format("{2}{0}.{1}", convertParametersForService.ServiceFileNames.OutputName + "_converted", convertParametersForService.ServiceFileNames.OutputExtension, globalRoutes.OutputConvertRoute);
                        WorkConsoleConvertDocPdfStart(input, output, convertParametersForService);
                    }
                    else
                    {
                        String input = String.Format("{1}{0}", e.Name, globalRoutes.InputConvertRoute);
                        String output = String.Format("{2}{0}.{1}", convertParametersForService.ServiceFileNames.OutputName + "_converted", convertParametersForService.ServiceFileNames.OutputExtension, globalRoutes.OutputConvertRoute);
                        WorkConsoleConvert(input, output, convertParametersForService.McfParams.C, convertParametersForService.McfParams.F, convertParametersForService.McfParams.M);
                    }
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
        }

       
        /// <summary>
        ///pdf - docx.
        /// </summary>
        /// <param name="input">The input.</param>
        /// <param name="output">The output.</param>
        /// <param name="convertParametersForService">The convert parameters for service.</param>
        private static void WorkConsoleConvertDocPdfStart(String input, String output, ConvertParametersForService convertParametersForService)
        {
            try
            {
                String path = globalRoutes.ConvertDocRoute;
                String command = String.Format("{0}ConvertDoc.exe", path);
                String outputTemp = String.Format("{2}{0}.{1}", convertParametersForService.ServiceFileNames.OutputName, "docx", globalRoutes.DiffTempRoute);
                String argument = String.Format("/S {0}  /T {1} /M{3} /C{2}", input, outputTemp, "12", "1");  // M1

                if (WhaitFileFree(input))
                {
                    ProcessStartInfo processStartInfo = new ProcessStartInfo(command, argument)
                    {
                        UseShellExecute = true,
                        WorkingDirectory = path
                    };

                    Process process = Process.Start(processStartInfo);
                    process.WaitForExit();
                    if (WhaitFileFree(outputTemp))
                    {
                        CheckCleanBullets(outputTemp);
                    }
                    if (WhaitFileFree(outputTemp))
                    {
                        File.Copy(outputTemp, output, true);  
                    }
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
        }


        /// <summary>
        ///  html - pdf - docx start
        /// </summary>
        /// <param name="input">The input.</param>
        /// <param name="convertParametersForService">The convert parameters for service.</param>
        private static void WorkConsoleConvertHtmlDocStart(String input, ConvertParametersForService convertParametersForService)
        {
            try
            {
                String path = globalRoutes.ConvertDocRoute;
                String command = String.Format("{0}ConvertDoc.exe", path);
                String output = String.Format("{2}{0}.{1}", convertParametersForService.ServiceFileNames.OutputName, "pdf", globalRoutes.DiffTempRoute);
                String argument = String.Format("/S {0}  /T {1} /M{3} /C{2}", input, output, "17", "1");  // M1

                if (WhaitFileFree(input))
                {
                    ProcessStartInfo processStartInfo = new ProcessStartInfo(command, argument)
                    {
                        UseShellExecute = true,
                        WorkingDirectory = path
                    };

                    Process process = Process.Start(processStartInfo);
                    process.WaitForExit();

                    // html - PDF - DOCX END
                    WorkConsoleConvertHtmlDocEnd(output, convertParametersForService);
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
        }


        /// <summary>
        ///html - PDF - DOCX END
        /// </summary>
        /// <param name="input">The input.</param>
        /// <param name="convertParametersForService">The convert parameters for service.</param>
        private static void WorkConsoleConvertHtmlDocEnd(String input, ConvertParametersForService convertParametersForService)
        {
            try
            {
                String path = globalRoutes.ConvertDocRoute;
                String command = String.Format("{0}ConvertDoc.exe", path);
                String outputTemp = String.Format("{2}{0}.{1}", convertParametersForService.ServiceFileNames.OutputName + "_converted", convertParametersForService.ServiceFileNames.OutputExtension, globalRoutes.DiffTempRoute);
                String output = String.Format("{2}{0}.{1}", convertParametersForService.ServiceFileNames.OutputName + "_converted", convertParametersForService.ServiceFileNames.OutputExtension, globalRoutes.OutputConvertRoute);
                String argument = String.Format("/S {0}  /T {1} /M{3} /C{2}", input, outputTemp, "12", "1");  // M1

                if (WhaitFileFree(input))
                {
                    ProcessStartInfo processStartInfo = new ProcessStartInfo(command, argument)
                    {
                        UseShellExecute = true,
                        WorkingDirectory = path
                    };

                    Process process = Process.Start(processStartInfo);
                    process.WaitForExit();
                }

                if (WhaitFileFree(outputTemp))
                {
                    CheckCleanBullets(outputTemp);
                }
                if (WhaitFileFree(outputTemp))
                {
                    File.Copy(outputTemp, output, true);  
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
        }



        //// todo: bullets  to default style
        //private static void BulletsToDefault(String patch, String output)
        //{

        //    Microsoft.Office.Interop.Word.Document docs = null;
        //    Microsoft.Office.Interop.Word.Application word = null;
        //    try
        //    {
        //        word = new Microsoft.Office.Interop.Word.Application();

        //        Object miss = System.Reflection.Missing.Value;
        //        Object path = patch;
        //        Object readOnly = false;

        //        docs = word.Documents.Open(ref path, ref miss, ref readOnly,
        //                                                                          ref miss, ref miss, ref miss, ref miss,
        //                                                                          ref miss, ref miss, ref miss, ref miss,
        //                                                                          ref miss, ref miss, ref miss, ref miss,
        //                                                                          ref miss);
        //        Int32 previousNumber = -1;
        //        Boolean wasDeleteNum = false;
        //        // Boolean startNumber = true;
        //        if (docs.Lists != null)
        //        {
        //            WdListType bullet = WdListType.wdListBullet;
        //            WdListType numbering = WdListType.wdListSimpleNumbering;
        //            for (int i = 0; i < docs.Lists.Count; i++)
        //            {
        //                WdListType currentListType = docs.Lists[i + 1].Range.ListFormat.ListType;

        //                if (currentListType == bullet)
        //                {
        //                    docs.Lists[i + 1].Range.ListFormat.ApplyBulletDefault();
        //                    docs.Lists[i + 1].Range.ListFormat.ApplyBulletDefault();
        //                }

        //                if (currentListType == numbering)
        //                {
        //                    docs.Lists[i + 1].Range.ListFormat.ApplyNumberDefault();
        //                    docs.Lists[i + 1].Range.ListFormat.ApplyNumberDefault();
        //                }
        //            }

        //        }
        //        docs.SaveAs(@"C:\_test\test.docx");
        //        if (docs != null) docs.Close();
        //        if (word != null) word.Quit();

        //    }
        //    catch (Exception ex)
        //    {
        //        LogError(ex);
        //    }
        //    finally
        //    {
        //        if (docs != null) docs.Close();
        //        if (word != null) word.Quit();
        //    }
        //}



        /// <summary>
        /// Checks and clean bullets.
        /// </summary>
        private static void CheckCleanBullets(String patch)
        {
            Microsoft.Office.Interop.Word.Document docs = null;
            Microsoft.Office.Interop.Word.Application word = null;
            try
            {
                word = new Microsoft.Office.Interop.Word.Application();

                Object miss = System.Reflection.Missing.Value;
                Object path = patch;
                Object readOnly = false;

                docs = word.Documents.Open(ref path, ref miss, ref readOnly,
                                                                                  ref miss, ref miss, ref miss, ref miss,
                                                                                  ref miss, ref miss, ref miss, ref miss,
                                                                                  ref miss, ref miss, ref miss, ref miss,
                                                                                  ref miss);
                Int32 previousNumber = -1;
                Boolean wasDeleteNum = false;
                for (int i = 0; i < docs.Paragraphs.Count; i++)
                {
                    if (docs.Paragraphs[i + 1].Range.Text.Contains(""))
                    {
                        var paragraphStyle = docs.Paragraphs[i + 1].Range.get_Style();
                        docs.Paragraphs[i + 1].Range.Text = docs.Paragraphs[i + 1].Range.Text.Replace("", "");
                        docs.Paragraphs[i + 1].Range.set_Style(ref paragraphStyle);
                        docs.Paragraphs[i + 1].Range.ListFormat.ApplyBulletDefault();
                        docs.Paragraphs[i + 1].Outdent();
                    }

                    Boolean numberOperated = false;
                    Int32 currNumber = -1;

                    if (docs.Paragraphs[i + 1].Range.Text.Length > 2)
                    {
                        currNumber = GetFirstNum(docs.Paragraphs[i + 1].Range.Text);
                    }
                    if (currNumber > 0)
                    {

                        Int32 numLen = currNumber.ToString().Length;

                        String checkPoint = docs.Paragraphs[i + 1].Range.Text.Substring(numLen, 1);
                        if (checkPoint == ".")
                        {
                            numberOperated = true;
                            if (previousNumber >= 0)
                            {
                                if (currNumber - previousNumber == 1)
                                {
                                    //delete numbers at start
                                    if (!wasDeleteNum)
                                    {
                                        var paragraphStyleOld = docs.Paragraphs[i].Range.get_Style();

                                        docs.Paragraphs[i].Range.Text = docs.Paragraphs[i].Range.Text.Remove(0, numLen + 1);

                                        docs.Paragraphs[i].Range.set_Style(ref paragraphStyleOld);

                                        docs.Paragraphs[i].Range.ListFormat.ApplyNumberDefault();
                                        docs.Paragraphs[i].Outdent();
                                    }

                                    var paragraphStyleNew = docs.Paragraphs[i + 1].Range.get_Style();

                                    docs.Paragraphs[i + 1].Range.Text = docs.Paragraphs[i + 1].Range.Text.Remove(0, numLen + 1);

                                    docs.Paragraphs[i + 1].Range.set_Style(ref paragraphStyleNew);

                                    docs.Paragraphs[i + 1].Range.ListFormat.ApplyNumberDefault();
                                    docs.Paragraphs[i + 1].Outdent();

                                    wasDeleteNum = true;
                                    // startNumber = false;
                                    previousNumber = currNumber;
                                }
                            }
                            else
                            {
                                previousNumber = currNumber;
                                wasDeleteNum = false;
                                //   startNumber = true;
                            }
                        }
                    }

                    if (!numberOperated)
                    {
                        previousNumber = -1;
                        wasDeleteNum = false;
                    }

                }

                docs.SaveAs(patch);
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
            finally
            {
                if (docs != null) docs.Close();
                if (word != null) word.Quit();
            }
        }

        /// <summary>
        /// Parses the int.
        /// </summary>
        /// <param name="str">The STR.</param>
        /// <returns></returns>
        public Int32 ParseInt(String str)
        {
            Int32 val = -1;
            Regex reg = new Regex(@"^([\d]+).*$");
            Match match = reg.Match(str);
            if (match != null) Int32.TryParse(match.Groups[1].Value, out val); // ??
            return val;
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
        ///  diff docs.
        /// </summary>
        /// <param name="originalDiff">The original diff.</param>
        /// <param name="modifiedDiff">The modified diff.</param>
        /// <param name="reportDiff">The report diff.</param>
        private static void WorkConsoleDiffDocs(String originalDiff, String modifiedDiff, String reportDiff)
        {
            try
            {
                if (WhaitFileFree(originalDiff))
                {
                    CheckCleanBullets(originalDiff);
                }
                if (WhaitFileFree(modifiedDiff))
                {
                    CheckCleanBullets(modifiedDiff);
                }
                String path = globalRoutes.DiffDocRoute;
                String command = String.Format("{0}DiffDoc.exe", path);
                String argument = String.Format("/M {0}  /S {1}  /F1 /R4 /T {2} /X /E /P /W", originalDiff, modifiedDiff, reportDiff);

                WhaitFileFree(originalDiff);
                WhaitFileFree(modifiedDiff);

                ProcessStartInfo processStartInfo = new ProcessStartInfo(command, argument)
                    {
                        UseShellExecute = true,
                        WorkingDirectory = path
                    };

                Process process = Process.Start(processStartInfo);
                process.WaitForExit();
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
        }


        /// <summary>
        /// Diff Convert To Html
        /// </summary>
        /// <param name="input">The input.</param>
        /// <param name="output">The output.</param>
        private static void WorkConsoleDiffConvertToHtml(String input, String output)
        {
            try
            {
                String path = globalRoutes.ConvertDocRoute;
                String command = String.Format("{0}ConvertDoc.exe", path);
                String argument = String.Format("/S {0}  /T {1} /M{3} /C{2}", input, output, "8", "1");  // M1

                if (WhaitFileFree(input))
                {
                    ProcessStartInfo processStartInfo = new ProcessStartInfo(command, argument)
                        {
                            UseShellExecute = true,
                            WorkingDirectory = path
                        };

                    Process process = Process.Start(processStartInfo);
                    process.WaitForExit();
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
        }


        /// <summary>
        ///  Diff Convert HTML or PDF To Docs
        /// </summary>
        /// <param name="input">The input.</param>
        /// <param name="output">The output.</param>
        /// <param name="currentExt">The current ext.</param>
        /// <param name="name">The name.</param>
        private static void WorkConsoleDiffConvertToDoc(String input, String output, String currentExt, String name)
        {
            try
            {
                if (currentExt == "html" || currentExt == "htm")
                {
                    //html - pdf - doc diff start
                    WorkConsoleDiffConvertHtmlToDocStart(input, output, name);

                }
                else
                {
                    String path = globalRoutes.ConvertDocRoute;
                    String directoryTemp = globalRoutes.DiffTempRoute + ChangeToDoc(name);  
                    String command = String.Format("{0}ConvertDoc.exe", path);
                    String argument = String.Format("/S {0}  /T {1} /M{3} /C{2}", input, directoryTemp, "12", "1");      // M1

                    if (WhaitFileFree(input))
                    {
                        ProcessStartInfo processStartInfo = new ProcessStartInfo(command, argument)
                        {
                            UseShellExecute = true,
                            WorkingDirectory = path
                        };

                        Process process = Process.Start(processStartInfo);
                        process.WaitForExit();
                    }

                    if (WhaitFileFree(directoryTemp))
                    {
                        File.Copy(directoryTemp, output, true);  
                    }
                }

            }
            catch (Exception ex)
            {
                LogError(ex);
            }

        }


        /// <summary>
        /// html - pdf - doc diff start
        /// </summary>
        /// <param name="input">The input.</param>
        /// <param name="output">The output.</param>
        /// <param name="name">The name.</param>
        private static void WorkConsoleDiffConvertHtmlToDocStart(String input, String output, String name)
        {
            try
            {
                String path = globalRoutes.ConvertDocRoute; //globalRoutes.ConvertDocRoute; 
                String command = String.Format("{0}ConvertDoc.exe", path);
                String outputEnd = globalRoutes.DiffTempRoute + ChangeToPdf(name);
                String argument = String.Format("/S {0}  /T {1} /M{3} /C{2}", input, outputEnd, "17", "1");      // M1

                if (WhaitFileFree(input))
                {
                    ProcessStartInfo processStartInfo = new ProcessStartInfo(command, argument)
                    {
                        UseShellExecute = true,
                        WorkingDirectory = path
                    };

                    Process process = Process.Start(processStartInfo);
                    process.WaitForExit();

                    //html - pdf - doc diff end
                    WorkConsoleDiffConvertHtmlToDocEnd(outputEnd, output, name);
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
        }


        /// <summary>
        /// html - pdf - doc diff end
        /// </summary>
        /// <param name="input">The input.</param>
        /// <param name="output">The output.</param>
        /// <param name="name">The name.</param>
        private static void WorkConsoleDiffConvertHtmlToDocEnd(String input, String output, String name)
        {
            try
            {


                String path = globalRoutes.ConvertDocRoute;
                String command = String.Format("{0}ConvertDoc.exe", path);
                String outputEnd = globalRoutes.DiffTempRoute + ChangeToDoc(name);
                String argument = String.Format("/S {0}  /T {1} /M{3} /C{2}", input, outputEnd, "12", "1");      // M1

                if (WhaitFileFree(input))
                {
                    ProcessStartInfo processStartInfo = new ProcessStartInfo(command, argument)
                    {
                        UseShellExecute = true,
                        WorkingDirectory = path
                    };

                    Process process = Process.Start(processStartInfo);
                    process.WaitForExit();

                    if (WhaitFileFree(outputEnd))
                    {
                        //copy from temp
                        File.Copy(outputEnd, output, true);
                    }

                }
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
        }


        /// <summary>
        /// simple comvert
        /// </summary>
        /// <param name="input">The input.</param>
        /// <param name="output">The output.</param>
        /// <param name="outputFormat">The output format.</param>
        /// <param name="inputformat">The inputformat.</param>
        /// <param name="convertFormat">The convert format.</param>
        private static void WorkConsoleConvert(String input, String output, String outputFormat, String inputformat, String convertFormat)
        {
            try
            {
                String path = globalRoutes.ConvertDocRoute; //globalRoutes.ConvertDocRoute; 

                String command = String.Format("{0}ConvertDoc.exe", path);
                String argument = "";
                if (inputformat == "")
                {
                    argument = String.Format("/S {0} /T {1} /M{4} /C{2}", input, output, outputFormat, inputformat, convertFormat);
                }


                if (WhaitFileFree(input))
                {
                    ProcessStartInfo processStartInfo = new ProcessStartInfo(command, argument)
                    {
                        UseShellExecute = true,
                        WorkingDirectory = path
                    };

                    Process process = Process.Start(processStartInfo);
                    process.WaitForExit();
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
        }

        /// <summary>
        /// Logs the error.
        /// </summary>
        /// <param name="ex">The ex.</param>
        private static void LogError(Exception ex)
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
                if (countDown == 1800) //  timeout 3 min
                {
                    return false;
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

        private static String ChangeToDoc(String name)
        {
            try
            {
                Int32 ind1 = name.LastIndexOf('.');
                name = name.Substring(0, ind1);
                name = name + ".docx";
                return name;
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
            return String.Empty;
        }

        private static String ChangeToPdf(String name)
        {
            try
            {
                Int32 ind1 = name.LastIndexOf('.');
                name = name.Substring(0, ind1);
                name = name + ".pdf";
                return name;
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
            return String.Empty;
        }

        private static String ChangeToHtml(String name)
        {
            try
            {
                Int32 ind1 = name.LastIndexOf('.');
                name = name.Substring(0, ind1);
                name = name + ".html";
                return name;
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
            return String.Empty;
        }

        private static String GetExt(String name)
        {
            try
            {
                Int32 ind1 = name.LastIndexOf('.');
                name = name.Substring(ind1 + 1, name.Length - ind1 - 1);
                return name;
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
            return String.Empty;
        }
    }
}
