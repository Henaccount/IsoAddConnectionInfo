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


using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;


using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.ProjectManager;
using Autodesk.ProcessPower.PlantInstance;
using Autodesk.AutoCAD.EditorInput;

using System;
using System.Runtime.InteropServices;
using PlantApp = Autodesk.ProcessPower.PlantInstance.PlantApplication;

using System.Collections.Generic;
using System.Reflection;
using Autodesk.ProcessPower.PnP3dObjects;
using Autodesk.ProcessPower.DataObjects;
using Autodesk.ProcessPower.PartsRepository.Specification;
using Autodesk.ProcessPower.P3dUI;
using Autodesk.ProcessPower.PartsRepository;
using System.IO;

namespace IsoAddConnectionInfo
{
    /// <summary>
    /// Helper class including some static helper functions.
    /// </summary>
    /// v1: imperial compatible, autoapprove with zdecline=0.98


    public class Helper
    {
        public static Project currentProject { get; set; }
        public static Document ActiveDocument { get; set; }
        public static DataLinksManager ActiveDataLinksManager { get; set; }
        public static Database oDatabase { get; set; }
        public static Editor oEditor { get; set; }

        public static bool Initialize()
        {
            if (PlantApp.CurrentProject == null)
                return false;

            currentProject = PlantApp.CurrentProject.ProjectParts["Piping"];
            ActiveDataLinksManager = currentProject.DataLinksManager;
            ActiveDocument = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            oDatabase = ActiveDocument.Database;
            oEditor = ActiveDocument.Editor;
            return true;
        }

        public static void Terminate()
        {
            currentProject = null;
            ActiveDataLinksManager = null;
            ActiveDocument = null;
            oDatabase = null;
            oEditor = null;
        }


    }


}

