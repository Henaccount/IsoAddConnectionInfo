using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.ProcessPower.DataObjects;
using Autodesk.ProcessPower.PartsRepository;
using Autodesk.ProcessPower.PartsRepository.Specification;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.Runtime;
using System.Text;
using System.Text.RegularExpressions;

[assembly: CommandClass(typeof(IsoAddConnectionInfo.Program))]

namespace IsoAddConnectionInfo
{
    class Program
    {


        public static bool theErrorFlag = false;
        public static bool theChangeFlag = false;
        public static bool isImperial = false;
        public static StringBuilder logtext = new StringBuilder();
        public static string continuation = ""; //
        public static int targetmaxlength = 100; //EndConnectionsPipe, has to be smaller then the inner piping continuation (EndConnectionsTo, EndConnectionsFrom)
        public static string freetext = "";
        public static string thepath = "";
        public static List<string[]> hvdims = new List<string[]>();

        public static PromptSelectionResult findMtextContains()
        {
            TypedValue[] filterlist = new TypedValue[1];
            filterlist[0] = new TypedValue(0, "MTEXT");
            SelectionFilter thefilter = new SelectionFilter(filterlist);
            PromptSelectionResult res = Helper.oEditor.SelectAll(thefilter);
            return res;
        }

        [CommandMethod("IsoAddConnectionInfo", CommandFlags.Session)]
        public static void IsoAddConnectionInfo()
        {
            hvdims = new List<string[]>();
            loopOverDrawings("collect");
            loopOverDrawings("modify");
            Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("script execution ended");
        }


        public static void loopOverDrawings(string collectOrmodify)
        {
            Helper.Initialize();
            string configstr = "";



            logtext = new StringBuilder();


            try
            {
                if (thepath.Equals(""))
                {

                    PromptResult pr = Helper.oEditor.GetString("Provide a string: ");

                    if (pr.Status != PromptStatus.OK)
                    {

                        logtext.Append("\r\nNo configuration string was provided\r\n");

                        return;

                    }
                    else
                        configstr = pr.StringResult;


                    if (!configstr.Equals(""))
                    {
                        string[] configArr = configstr.Split(new char[] { ',' });
                        string tmpstr = "";
                        tmpstr = configArr[0].Split(new char[] { '=' })[1];
                        if (!tmpstr.Trim().Equals("")) thepath = tmpstr.Trim();
                        tmpstr = configArr[1].Split(new char[] { '=' })[1];
                        if (!tmpstr.Trim().Equals("")) continuation = tmpstr.Trim();
                        tmpstr = configArr[2].Split(new char[] { '=' })[1];
                        if (!tmpstr.Trim().Equals("")) targetmaxlength = Convert.ToInt32(tmpstr.Trim());
                    }

                }

                foreach (string drawing in Directory.EnumerateFiles(thepath, "*.dwg"))
                {
                    isImperial = false;
                    theChangeFlag = false;
                    Document docToWorkOn = Application.DocumentManager.Open(drawing, false);

                    string[] justfilename = docToWorkOn.Name.Split(new Char[] { '\\' });
                    System.IO.Directory.CreateDirectory(thepath + "/changedfiles/");
                    string strDWGName = thepath + "/changedfiles/" + justfilename[justfilename.Length - 1];
                    logtext.Append("\r\n\r\n" + justfilename[justfilename.Length - 1] + "\r\n");


                    Application.DocumentManager.MdiActiveDocument = docToWorkOn;
                    Helper.Terminate();
                    using (docToWorkOn.LockDocument())
                    {

                        theMain(collectOrmodify);
                    }

                    if (theChangeFlag || theErrorFlag)
                    {
                        docToWorkOn.Database.SaveAs(strDWGName, true, DwgVersion.Current, docToWorkOn.Database.SecurityParameters);
                        docToWorkOn.CloseAndDiscard();

                    }
                    else
                    {
                        docToWorkOn.CloseAndDiscard();
                    }

                }


            }

            catch (System.Exception e)
            {

                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(e, true);
                logtext.Append(trace.ToString());
                logtext.Append("Line: " + trace.GetFrame(0).GetFileLineNumber());
                logtext.Append("message: " + e.Message);

            }
            finally
            {

                File.WriteAllText(thepath + @"\log.txt", logtext.ToString());
            }


        }

        public static void theMain(string collectOrmodify)
        {
            Helper.Initialize();

            PromptSelectionResult res = findMtextContains();

            try
            {


                using (Transaction t = Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.StartTransaction())
                {
                    SelectionSet selSetalldims = res.Value;
                    ObjectId[] alldims = null;
                    if (selSetalldims != null)
                    {
                        alldims = selSetalldims.GetObjectIds();
                    }


                    foreach (ObjectId oid in alldims)
                    {

                        MText adim = (MText)t.GetObject(oid, OpenMode.ForWrite);

                        string[] oidtext = adim.Contents.Split(new string[] { "\r\n" }, System.StringSplitOptions.RemoveEmptyEntries);
                        if (oidtext.Length > 5) logtext.Append("\r\nlabel contains already more than five lines, skipping: " + adim.Contents);
                        string[] filepathparts = Application.DocumentManager.MdiActiveDocument.Name.Split(new char[] { '\\' });
                        string filenamepart = filepathparts[filepathparts.Length - 1];

                        if (oidtext[0].IndexOf(continuation) != -1 && oidtext[1].Length <= targetmaxlength)
                        {
                            if (collectOrmodify.Equals("collect"))
                            {

                                Array.Resize(ref oidtext, oidtext.Length + 1);
                                oidtext[5] = filenamepart;
                                hvdims.Add(oidtext);
                                logtext.Append("\r\ncollected: " + adim.Contents);
                            }
                            else
                            {
                                foreach (string[] oitext in hvdims)
                                {
                                    if ((oitext[2].Equals(oidtext[2]) && oitext[3].Equals(oidtext[3]) && oitext[4].Equals(oidtext[4])) && !oitext[5].Equals(filenamepart))
                                    {
                                        string newlabel = oidtext[0] + "\r\n" + oidtext[1] + "\r\n(" + oitext[5] + ")\r\n" + oidtext[2] + "\r\n" + oidtext[3] + "\r\n" + oidtext[4];
                                        logtext.Append("\r\nnewlabel: " + newlabel);
                                        adim.Contents = newlabel;
                                        theChangeFlag = true;
                                        break;
                                    }
                                }
                            }
                        }


                    }


                    t.Commit();
                }
            }
            catch (System.Exception ex)
            {

                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex, true);
                logtext.Append("\r\ntrace: " + trace.ToString());
                logtext.Append("\r\nLine: " + trace.GetFrame(0).GetFileLineNumber());
                logtext.Append("\r\nmessage: " + ex.Message);

            }
            finally
            {

                Helper.Terminate();

            }
        }

    }
}
