using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

[assembly: CommandClass(typeof(IsoMetric2imperial.Program))]

namespace IsoMetric2imperial
{
    class Program
    {

        /* [CommandMethod("thetest123", CommandFlags.Modal)]
         public static void thetest() // This method can have any name
         {
             Helper.Initialize();
             string configstr = "";

             TypedValue[] filterlist = new TypedValue[1];
             filterlist[0] = new TypedValue(0, "TEXT,MTEXT");
             SelectionFilter thefilter = new SelectionFilter(filterlist);
             PromptSelectionResult result = Helper.oEditor.SelectAll(thefilter);

             using (Transaction t = Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.StartTransaction())
             {

                 try
                 {
                     string text = "";
                     foreach (var id in result.Value.GetObjectIds())
                     {
                         string type = id.ObjectClass.Name;
                         if (type.Equals("AcDbMText"))
                         {
                             MText ent = (MText)t.GetObject(id, OpenMode.ForWrite);
                             text = ent.Text;
                         }
                         else
                         {
                             DBText ent = (DBText)t.GetObject(id, OpenMode.ForWrite);
                             text = ent.TextString;
                         }

                         //String pattern = @"(\d+')?-?(\d+[^/'\d])?(\d+/\d+)?[^/]?";
                         string output = "";

                         foreach (Match m in Regex.Matches(text, configstr))
                         {
                             int u = 0;
                             foreach (var group in m.Groups)
                             {
                                 ++u;
                                 foreach (Capture capture in m.Captures)
                                 {
                                     output += "found in " + type + " of handle: " + id.Handle + " at pos: " + capture.Index + ": " + capture.ToString() + " (in group: " + u + ") search term: " + configstr + " ,full text: " + text + "\n";
                                 }
                             }

                         }
                         Helper.oEditor.WriteMessage(output);
                     }

                 }
                 catch (System.Exception e)
                 {

                     System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(e, true);
                     Helper.oEditor.WriteMessage(trace.ToString());
                     Helper.oEditor.WriteMessage("Line: " + trace.GetFrame(0).GetFileLineNumber());
                     Helper.oEditor.WriteMessage("message: " + e.Message);

                 }
             }
             Helper.Terminate();
         }*/

        public static string thenorth = "UpperLeft"; //UpperRight
        public static string zdecline = "";
        public static string thepath = "";
        public static int precision = 4;
        public static double AngleTolerance = 0.1;
        public static bool theErrorFlag = false;
        public static bool theChangeFlag = false;
        public static bool isImperial = false;
        public static StringBuilder logtext = new StringBuilder();
        public static IDictionary<string, string> sizemap;

        [CommandMethod("IsoMetric2imperial", CommandFlags.Session)]
        public static void loopOverDrawings()
        {
            sizemap = Helper.LoadCsvAsDict(@"D:\Program Files\Autodesk\AutoCAD 2024\NominalDiameterMap.csv");
            logtext = new StringBuilder();

            try
            {

                OpenFileDialog thedialog = new OpenFileDialog("Select one file from the folder", "", "dwg", "dialogName", OpenFileDialog.OpenFileDialogFlags.SearchPath);
                thedialog.ShowDialog();
                thepath = (new FileInfo(thedialog.Filename)).Directory.FullName;
                if (thepath.Length < 3) return;

                if (!Directory.Exists(thepath + Path.DirectorySeparatorChar + "results")) Directory.CreateDirectory(thepath + Path.DirectorySeparatorChar + "results");

                foreach (string drawing in Directory.EnumerateFiles(thepath, "*.dwg"))
                {
                    isImperial = false;
                    theChangeFlag = false;
                    using (Database db = new Database(false, true))
                    {
                        string justfilename = new FileInfo(drawing).Name;
                        string strDWGpath = thepath + Path.DirectorySeparatorChar + "results" + Path.DirectorySeparatorChar + justfilename;
                        logtext.Append("\n\n" + justfilename + "\n");
                        db.ReadDwgFile(drawing, FileOpenMode.OpenForReadAndAllShare, false, null);

                        theMain(db);


                        //docToWorkOn.Database.SaveAs(strDWGName, true, DwgVersion.Current, docToWorkOn.Database.SecurityParameters);
                        db.SaveAs(strDWGpath, DwgVersion.Current);

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
                File.WriteAllText(thepath + "/results" + @"\log.txt", logtext.ToString());
                Helper.CmdLineMessage("script execution finished");
            }


        }

        public static void theMain(Database db)
        {

            try
            {


                using (Transaction t = db.TransactionManager.StartTransaction())
                {

                    BlockTableRecord btr = (BlockTableRecord)t.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                    foreach (ObjectId id in btr)
                    {
                        if (id.ObjectClass.Name.Equals("AcDbRotatedDimension"))
                        { //todo: other dimensions like: AcDbAlignedDimension etc?
                            Dimension adim = (Dimension)t.GetObject(id, OpenMode.ForWrite);

                            string oidtext = adim.DimensionText;

                            adim.DimensionText = MetricToImperial(oidtext);
                        }
                        else if (id.ObjectClass.Name.Equals("AcDbMText"))
                        {
                            //pattern endlabels
                            string label = @"OPEN END
W 4723
N 6866
EL -2269";
                            MText ent = (MText)t.GetObject(id, OpenMode.ForWrite);
                            //string text = ent.Contents.Replace("{\\Q0;","").Replace("}", "");
                            string text = ent.Contents;
                            string output = "";

                            String pattern = @"^(.*?\r?\n[WE] )(-?\d+)(\r?\n[NS] )(-?\d+)(\r?\nEL )(-?\d+)(\r?\n?}?)$";
                            foreach (Match m in Regex.Matches(text, pattern))
                            {
                                output = m.Groups[1].Value + MetricToImperial(m.Groups[2].Value) +
                                       m.Groups[3].Value + MetricToImperial(m.Groups[4].Value) +
                                       m.Groups[5].Value + MetricToImperial(m.Groups[6].Value) +
                                       m.Groups[7].Value;

                                break;
                                
                            }
                            //Helper.CmdLineMessage("\ntext: " + text);
                            if (!output.Equals("")) ent.Contents = output;

                            //pattern linenumber?
                        }
                        else if (id.ObjectClass.Name.Equals("AcDbTable"))
                        {
                            //repace metric with imperial in certain columns, use "D:\Program Files\Autodesk\AutoCAD 2024\NominalDiameterMap.csv"


                            Table table = (Table)t.GetObject(id, OpenMode.ForWrite);

                            // Find the column with the name "size"
                            int sizeColumnIndex = -1;


                            for (int i = 0; i < table.Columns.Count; i++)
                            {

                                if (table.Cells[1, i].Value.ToString().Equals("ND", StringComparison.OrdinalIgnoreCase))
                                {
                                    sizeColumnIndex = i;

                                    break;
                                }

                            }

                            //Helper.CmdLineMessage("\nsizeColumnIndex " + sizeColumnIndex);

                            if (sizeColumnIndex == -1)
                            {
                                Helper.CmdLineMessage("\nTable does not contain a column named 'ND'.");
                            }


                            for (int i = 2; i < table.Rows.Count; i++)
                            {
                                Cell cell = table.Cells[i, sizeColumnIndex];
                                string existingText = cell.GetTextString(FormatOption.FormatOptionNone);
                                string newText = "";

                                string[] theparts = existingText.Split(new Char[] { 'X' });
                                for (int j = 0; j < theparts.Length; j++)
                                {
                                    if (j != 0) newText += "X";
                                    if (sizemap.ContainsKey(existingText))
                                        newText = sizemap[existingText] + "\"";
                                    else
                                        newText += MetricToImperial(theparts[j].Trim());
                                }

                                cell.Value = newText;

                            }
                        }
                    }
                    t.Commit();

                }
            }
            catch (System.Exception ex)
            {

                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex, true);
                logtext.Append(trace.ToString());
                logtext.Append("Line: " + trace.GetFrame(0).GetFileLineNumber());
                logtext.Append("message: " + ex.Message);

            }
            finally
            {
                //File.WriteAllText(thepath + "/results" + @"\log.txt", logtext.ToString());
            }
        }


        private static double DegreeToRadian(double angle)
        {
            return Math.PI * angle / 180.0;
        }


        public static void writeArrayToLog(Dimension[] dimArr)
        {
            //logtext.Append("\n");

            for (int i = 0; i < dimArr.Length; i++)
            {
                if (dimArr[i] != null)
                    logtext.Append(dimArr[i].DimensionText + ", ");
                else
                    logtext.Append("null, ");
            }

            logtext.Append("\n");
        }




        public static string MetricToImperial(string metric)
        {
            string imperial = Converter.DistanceToString(Convert.ToDouble(metric) / 25.4, DistanceUnitFormat.Architectural, precision);
            return imperial;
        }
    }



}

