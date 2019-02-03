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

//version 1: headless commands, pnid commands, overwrite existent (no request for empty folder), check if file exists, check if destination folder is design folder -> stop if so
//version 2: list dwgs not by dwg.AbsoluteFileName (original path) but by dwg.ResolvedFilePath, create folder system in destination folder, complete error handling, removed final message from HL command

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using System.Windows.Forms;
using Autodesk.AutoCAD.EditorInput;


[assembly: CommandClass(typeof(BatchExports.Commands))]

namespace BatchExports
{
    public class Commands
    {
        #region Commands
        [CommandMethod("PnIdBatchExportToAutoCADHL", CommandFlags.Session)]
        public static void PnIdBatchExportToAutoCADHL()
        {
            Program.BatchExportToAutoCAD("PnId", false);
        }

        [CommandMethod("BatchExportToAutoCADHL", CommandFlags.Session)]
        public static void BatchExportToAutoCADHL()
        {
            Program.BatchExportToAutoCAD("Piping", false);
        }

        [CommandMethod("BatchExportToAutoCAD", CommandFlags.Session)]
        public static void BatchExportToAutoCAD()
        {
            Program.BatchExportToAutoCAD("Piping", true);
        }


        [CommandMethod("PnIdBatchExportToAutoCAD", CommandFlags.Session)]
        public static void PnIdBatchExportToAutoCAD()
        {
            Program.BatchExportToAutoCAD("PnId", true);
        }
        #endregion
    }
}
