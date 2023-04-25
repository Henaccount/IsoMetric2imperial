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
        public static string thepath = "";
        public static int precision = 4;
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
                        //Helper.CmdLineMessage("\nclass: " + id.ObjectClass.Name);

                        if (id.ObjectClass.Name.Equals("AcDbRotatedDimension"))
                        { //todo: other dimensions like: AcDbAlignedDimension etc?
                            Dimension adim = (Dimension)t.GetObject(id, OpenMode.ForWrite);

                            string oidtext = adim.DimensionText;

                            adim.DimensionText = MetricToImperial(oidtext);
                        }
                        else if (id.ObjectClass.Name.Equals("AcDbMText"))
                        {
                            MText ent = (MText)t.GetObject(id, OpenMode.ForWrite);

                            string text = ent.Contents;
                            string output = "";
                            try
                            {
                                String pattern = @"\A([\w\W\s]*)(West |East )(-?\d+)(\s+|\\P)(North |South )(-?\d+)(\s+|\\P)(El \+?)(-?\d+)([\w\W\s]*)\z";
                                foreach (Match m in Regex.Matches(text, pattern))
                                {
                                    /*for (int j = 1; j < 10; j++)
                                        Helper.CmdLineMessage("\nm.Groups[j].Value: " + m.Groups[j].Value);*/
                                    if (IsNumeric(m.Groups[3].Value) && IsNumeric(m.Groups[6].Value) && IsNumeric(m.Groups[9].Value))
                                    {
                                        output = m.Groups[1].Value + m.Groups[2].Value + MetricToImperial(m.Groups[3].Value) +
                                               m.Groups[4].Value + m.Groups[5].Value + MetricToImperial(m.Groups[6].Value) +
                                               m.Groups[7].Value + m.Groups[8].Value + MetricToImperial(m.Groups[9].Value) +
                                               m.Groups[10].Value;
                                    }
                                    break;

                                }
                            }
                            catch { }
                            if (!output.Equals("")) { 

                                ent.Contents = output;                            
                            }

                            //pattern linenumber?
                        }
                        else if (id.ObjectClass.Name.Equals("AcDbTable"))
                        {
                            //repace metric with imperial in certain columns, use "D:\Program Files\Autodesk\AutoCAD 2024\NominalDiameterMap.csv"


                            Table table = (Table)t.GetObject(id, OpenMode.ForWrite);

                            // Find the column with the name "size"
                            int sizeColumnIndex = -1;
                            int cutpiecelengthIndex = -1;
                            int partcodeIndex = -1;
                            int quantityIndex = -1;
                            int headerrow = -1;

                            for (int j = 0; j < table.Rows.Count; j++)
                            {
                                if (!table.Rows[j].IsSingleCell)
                                {
                                    for (int i = 0; i < table.Columns.Count; i++)
                                    {

                                        if (table.Cells[j, i].Value.ToString().Equals("Size", StringComparison.OrdinalIgnoreCase))
                                        {
                                            sizeColumnIndex = i;
                                            headerrow = j;
                                            //Helper.CmdLineMessage("\nsizeColumnIndex found");
                                        }
                                        else if (table.Cells[j, i].Value.ToString().Equals("Qty", StringComparison.OrdinalIgnoreCase))
                                        {
                                            quantityIndex = i;
                                            //Helper.CmdLineMessage("\nquantityIndex found");
                                        }
                                        else if (table.Cells[j, i].Value.ToString().Equals("Cut length", StringComparison.OrdinalIgnoreCase))
                                        {
                                            cutpiecelengthIndex = i;
                                            //Helper.CmdLineMessage("\ncutpiecelengthIndex found");
                                        }
                                        else if (table.Cells[j, i].Value.ToString().Equals("Part Code", StringComparison.OrdinalIgnoreCase))
                                        {
                                            partcodeIndex = i;
                                        }
                                    }
                                }
                            }

                            int[] colsToWorkon = { sizeColumnIndex, cutpiecelengthIndex, quantityIndex };

                            for (int i = 0; i < table.Rows.Count; i++)
                            {
                                //Helper.CmdLineMessage("\nstyle: " + table.Rows[i].Style);
                                if (!table.Rows[i].IsSingleCell && i != headerrow)
                                {
                                    for (int k=0; k< colsToWorkon.Length; k++)
                                    {
                                        if (colsToWorkon[k] == -1) continue;
                                        Cell cell = table.Cells[i, colsToWorkon[k]];

                                        string existingText = cell.GetTextString(FormatOption.FormatOptionNone);

                                        //Helper.CmdLineMessage("\ntablecell: " + existingText);

                                        string newText = "";
                                        bool exceptionflag = false;
                                        if (k == 2) exceptionflag = true;

                                        if (partcodeIndex != -1)
                                        {
                                            string partcode = table.Cells[i, partcodeIndex].GetTextString(FormatOption.FormatOptionNone);
                                            if (partcode.StartsWith("BH-") || partcode.StartsWith("BW-") || partcode.StartsWith("BN-"))
                                            {
                                                exceptionflag = true;
                                            }
                                            if (partcode.StartsWith("PP-"))
                                            {
                                                exceptionflag = false;
                                            }
                                        }

                                        if (!exceptionflag)
                                        {
                                            string[] theparts = existingText.Split(new Char[] { 'X', '/' });
                                            for (int j = 0; j < theparts.Length; j++)
                                            {
                                                if (j != 0)
                                                {
                                                    if (existingText.Contains("X"))
                                                        newText += "X";
                                                    else
                                                        newText += "/";
                                                }
                                                if (sizemap.ContainsKey(theparts[j]))
                                                    newText += sizemap[theparts[j]] + "\"";
                                                else
                                                    newText += MetricToImperial(theparts[j].Trim());
                                            }

                                            cell.Value = newText;
                                        }

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
                logtext.Append(trace.ToString());
                logtext.Append("Line: " + trace.GetFrame(0).GetFileLineNumber());
                logtext.Append("message: " + ex.Message);

            }
            finally
            {
                //File.WriteAllText(thepath + "/results" + @"\log.txt", logtext.ToString());
            }
        }

        public static bool IsNumeric(string text) { double _out; return double.TryParse(text, out _out); }
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
            string imperial = "";
            if (!metric.Equals(""))
            {
                double dmetric = 0.0;
                metric = metric.Replace("mm", "");
                if (metric.Contains("M"))
                {
                    metric = metric.Replace("M", "");
                    if (IsNumeric(metric))
                        dmetric = Convert.ToDouble(metric) * 1000;
                }
                else
                {
                    if (IsNumeric(metric))
                        dmetric = Convert.ToDouble(metric);
                }
                imperial = Converter.DistanceToString(dmetric / 25.4, DistanceUnitFormat.Architectural, precision);
            }
            return imperial;
        }
    }



}

