//
// (C) Copyright 2013 by Autodesk, Inc.
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted,
// provided that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC.
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
//
// Use, duplication, or disclosure by the U.S. Government is subject to
// restrictions set forth in FAR 52.227-19 (Commercial Computer
// Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
// (Rights in Technical Data and Computer Software), as applicable.
//

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.ProcessPower.ProjectManager;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using PlantApp = Autodesk.ProcessPower.PlantInstance.PlantApplication;
using Autodesk.AutoCAD.Windows;
using Autodesk.AutoCAD.EditorInput;
using System;

namespace BatchExports
{
    class Program
    {



        public static async void BatchExportToAutoCAD(string projectpart, bool head)
        {
            string errorstack = "";

            try
            {
                string sourcefolder = "";
                string destfolder = "";


                if (!head)
                {
                    PromptResult pr = AcadApp.DocumentManager.MdiActiveDocument.Editor.GetString("\nsourcefolder: ");
                    if (pr.Status != PromptStatus.OK)
                    {
                        //Helper.oEditor.WriteMessage("No configuration string was provided, using defaults\n");
                    }
                    else
                        sourcefolder = pr.StringResult;

                    pr = AcadApp.DocumentManager.MdiActiveDocument.Editor.GetString("\ndestinationfolder: ");
                    if (pr.Status != PromptStatus.OK)
                    {
                        //Helper.oEditor.WriteMessage("No configuration string was provided, using defaults\n");
                    }
                    else
                        destfolder = pr.StringResult;

                }

                if (destfolder.Equals(""))
                {
                    OpenFileDialog thedialog = new OpenFileDialog("Select source folder", "defaultName", "extension", "dialogName", OpenFileDialog.OpenFileDialogFlags.AllowFoldersOnly);
                    thedialog.ShowDialog();
                    sourcefolder = thedialog.Filename;

                    OpenFileDialog thedialog1 = new OpenFileDialog("Select destination folder", "defaultName", "extension", "dialogName", OpenFileDialog.OpenFileDialogFlags.AllowFoldersOnly);
                    thedialog1.ShowDialog();
                    destfolder = thedialog1.Filename;
                }

                if (destfolder.Equals(""))
                {
                    //ed.WriteMessage("no folder specified, command stopped...\n");
                    return;
                }



                /*
                string[] files = System.IO.Directory.GetFiles(destfolder);
                if (files.Length > 0)
                {
                    System.Windows.Forms.MessageBox.Show("folder has to be empty!");
                    return;
                }
                */

                Project prj = PlantApp.CurrentProject.ProjectParts[projectpart];

                System.Collections.Generic.List<PnPProjectDrawing> dwgList = prj.GetPnPDrawingFiles();



                System.IO.DirectoryInfo projectdwgfolderobj = System.IO.Directory.CreateDirectory(prj.ProjectDwgDirectory);

                System.IO.DirectoryInfo destfolderobj = System.IO.Directory.CreateDirectory(destfolder);

                if (projectdwgfolderobj.FullName.TrimEnd('\\').Equals(destfolderobj.FullName.TrimEnd('\\')))
                {
                    System.Windows.Forms.MessageBox.Show("You are trying to replace your design files, stopping!");
                    return;
                }


                //destfolder += "\\";

                foreach (PnPProjectDrawing dwg in dwgList)
                {
                    Document docToWorkOn = null;
                    try
                    {
                        errorstack += dwg.ResolvedFilePath + "\n";

                        if (!System.IO.File.Exists(dwg.ResolvedFilePath) || !dwg.ResolvedFilePath.Contains(sourcefolder))
                            continue;

                        string pathstub = dwg.ResolvedFilePath.Substring(projectdwgfolderobj.FullName.TrimEnd('\\').Length);

                        string pathstubexclfile = pathstub.Substring(0, pathstub.LastIndexOf("\\"));

                        System.IO.Directory.CreateDirectory(destfolder + pathstubexclfile);

                        if (System.IO.File.Exists(destfolder + pathstub))
                        {
                            System.IO.File.Delete(destfolder + pathstub);
                        }

                        docToWorkOn = AcadApp.DocumentManager.Open(dwg.ResolvedFilePath, true);


                        AcadApp.DocumentManager.MdiActiveDocument = docToWorkOn;


                        using (docToWorkOn.LockDocument())
                        {
                            await AcadApp.DocumentManager.ExecuteInCommandContextAsync(async (obj) => { await AcadApp.DocumentManager.MdiActiveDocument.Editor.CommandAsync("_.-EXPORTTOAUTOCAD", "2013", destfolder + pathstub); }, null);
                            // this works: await AcadApp.DocumentManager.ExecuteInCommandContextAsync(async (obj) => { await AcadApp.DocumentManager.MdiActiveDocument.Editor.CommandAsync("filedia", "0", "NWCOUT", destfolder + pathstub + ".nwc", "filedia", "1"); }, null);
                        }


                    }
                    catch (System.Exception ex)
                    {

                        System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex, true);
                        errorstack += trace.ToString() + "\n";
                        errorstack += "Line: " + trace.GetFrame(0).GetFileLineNumber() + "\n";
                        errorstack += "message: " + ex.Message + "\n";

                    }
                    finally
                    {
                        if (docToWorkOn != null) { 
                            docToWorkOn.CloseAndDiscard();
                            }
                    }

                }



                if (head)
                    System.Windows.Forms.MessageBox.Show("Batch Export to AutoCAD finished. Selected folder: " + destfolder);
            }
            catch (System.Exception e)
            {

                errorstack += "message: " + e.Message + "\n";

            }
            finally
            {
                AcadApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage(errorstack);
            }
            //here
        }


    }
}

