//
//////////////////////////////////////////////////////////////////////////////
//
//  Copyright 2015 Autodesk, Inc.  All rights reserved.
//
//  Use of this software is subject to the terms of the Autodesk license 
//  agreement provided at the time of installation or download, or which 
//  otherwise accompanies this software in either electronic or hard copy form.   
//
//////////////////////////////////////////////////////////////////////////////
// if just one type of hose exists, shortdescription should be "HOSE"


using Autodesk.AutoCAD.ApplicationServices;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace IsoMetric2imperial
{
    /// <summary>
    /// Helper class including some static helper functions.
    /// </summary>
    /// v1: imperial compatible, autoapprove with zdecline=0.98


    public class Helper
    {


        public static IDictionary<string, string> LoadCsvAsDict(string thepath)
        {
            IDictionary<string, string> dict = new Dictionary<string, string>();
            try
            {
                string whole_file = System.IO.File.ReadAllText(thepath);



                // Split into lines.
                whole_file = whole_file.Replace('\n', '\r');
                string[] lines = whole_file.Split(new char[] { '\r' }, StringSplitOptions.RemoveEmptyEntries);

                // See how many rows and columns there are.
                int num_rows = lines.Length;
                int num_cols = lines[0].Split(',').Length;


                // Load the dictionary.
                for (int r = 0; r < num_rows; r++)
                {
                    string[] line_r = lines[r].Split(',');
                    if (line_r[2].Equals("mm"))
                    {
                        dict.Add(line_r[1], line_r[4]);

                    }

                }


            }
            catch (System.Exception ex)
            {

                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex, true);
                Program.logtext.Append(trace.ToString());
                Program.logtext.Append("Line: " + trace.GetFrame(0).GetFileLineNumber());
                Program.logtext.Append("message: " + ex.Message);

            }                // Return the values.
            return dict;
        }

        public static void Message(System.Exception ex)
        {
            try
            {
                System.Windows.Forms.MessageBox.Show(
                    ex.ToString(),
                    "Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
            catch
            {
            }
        }

        public static void InfoMessageBox(string str)
        {
            try
            {
                System.Windows.Forms.MessageBox.Show(
                    str,
                    "Events Watcher",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
            }
            catch (System.Exception ex)
            {
                Helper.Message(ex);
            }
        }

        // Please check the valibility of Editor object before calling this.
        public static void CmdLineMessage(string str)
        {
            try
            {
                if (Application.DocumentManager.MdiActiveDocument != null
                    && Application.DocumentManager.Count != 0
                    && Application.DocumentManager.MdiActiveDocument.Editor != null)
                {
                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(str);
                }
                else
                    return;
            }
            catch (System.Exception ex)
            {
                Helper.Message(ex);
            }
        }






    }


}

