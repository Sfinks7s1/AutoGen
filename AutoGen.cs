using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Windows;
//using GetSharedLibs.Properties;
//using Autodesk.AutoCAD.Interop;
using System.Linq.Expressions;
using AcExportLayout = Autodesk.AutoCAD.ExportLayout;



using System;

using System.IO;
using System.Drawing;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Reflection;
using System.Text.RegularExpressions;
using TDMS;
using TDMS.Interop;

using Microsoft.Win32;
using Microsoft.VisualBasic;

using acadApp = Autodesk.AutoCAD.ApplicationServices.Application;

using AcApp = Autodesk.AutoCAD.ApplicationServices;
using AcDb = Autodesk.AutoCAD.DatabaseServices;
using AcGm = Autodesk.AutoCAD.Geometry;
using AcRtm = Autodesk.AutoCAD.Runtime;

namespace AutoCADLibs
{
    public sealed class XREF
    {
        private void Initialize()
        {
        }
        private void Terminate()
        {
        }

        //[CommandMethod("updateImage")]

        //public static void updateImage()
        //{
        //    Document doc = Application.DocumentManager.MdiActiveDocument;
        //    Database db = doc.Database;
        //    Editor ed = doc.Editor;

        //    PromptEntityOptions options = new PromptEntityOptions("\nSelect Raster image to change");

        //    options.SetRejectMessage("\nSelect only Raster image");

        //    options.AddAllowedClass(typeof(RasterImage), false);

        //    PromptEntityResult acSSPrompt = ed.GetEntity(options);

        //    if (acSSPrompt.Status != PromptStatus.OK)
        //        return;

        //    using (Transaction Tx = db.TransactionManager.StartTransaction())
        //    {
        //        //get the mleader
        //        RasterImage image = Tx.GetObject(acSSPrompt.ObjectId,
        //                                   OpenMode.ForRead) as RasterImage;

        //        RasterImageDef ImageDef = Tx.GetObject(image.ImageDefId,
        //                               OpenMode.ForWrite) as RasterImageDef;

        //        ImageDef.SourceFileName = "c:\\temp\\new.jpeg";
        //        ImageDef.ActiveFileName = "c:\\temp\\new.jpeg";

        //        ImageDef.Load();
        //        Tx.Commit();
        //    }
        //}
        //[CommandMethod("ReloadXRefs")]

        //static public void ReloadXRefs()
        //{
        //    Document doc = Application.DocumentManager.MdiActiveDocument;
        //    Database db = doc.Database;
        //    ObjectIdCollection ids = new ObjectIdCollection();
        //    using (Transaction tr = db.TransactionManager.StartTransaction())
        //    {
        //        BlockTable table = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        //        foreach (ObjectId id in table)
        //        {
        //            BlockTableRecord record = tr.GetObject(id, OpenMode.ForRead) as BlockTableRecord;
        //            if (record.IsFromExternalReference)
        //            {
        //                ids.Add(id);
        //            }
        //        }
        //        tr.Commit();
        //    }
        //    if (ids.Count != 0)
        //    {
        //        db.ReloadXrefs(ids);
        //    }
        //}

       

        //Сохранение главного файла и объектов внешних ссылок (dwg и картинок) в TDMS (внешние ссылки в каталоге XREF в TDMS)
        [CommandMethod("REFRESHMAP")]
        public void RefreshMap()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                Condition check = new Condition();
                if (check.CheckTDMSProcess() == true)
                {
                    Document acDoc = Application.DocumentManager.MdiActiveDocument;
                    //путь к файлу, включая название файла, который является текущим в Autocad.
                    string filePath = acDoc.Name; /*Directory.GetFiles(res.StringResult, "*.dwg", SearchOption.AllDirectories)*/;

                    editor.WriteMessage("\n File Name: " + filePath);

                    Database db = acDoc.Database;

                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        editor.WriteMessage("\n----------XRefs_Details----------");
                        db.ResolveXrefs(true, false);
                        XrefGraph xg = db.GetHostDwgXrefGraph(true);
                        GraphNode root = xg.RootNode;
                        ObjectIdCollection objcoll = new ObjectIdCollection();

                        TDMSApplication tdmsApp = new TDMSApplication();
                        Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Minimized;
                        tdmsApp.Visible = false;
                        tdmsApp.Visible = true;

                        string moduleName = "CMD_SYSLIB";
                        string SelectPSDFolder = "SelectPsdFolder";
                        string LoadFileToPSD = "LoadFileToPSD";
                        bool fl_main_file = false;

                        string PathName = tdmsApp.ExecuteScript(moduleName, SelectPSDFolder);
                        //string PathName = @"c:\Temp\TDMS\{0070876A-6C3C-4B79-8D96-DB7CDD198AC0} - SYSADMIN\{42FF0CA7-5910-47C9-B9CA-32128E0A1E6F}\";
                        editor.WriteMessage("\n Path output from TDMS " + PathName);
                        
                        string parseGuidFromFile = null;
                        int xrefcount = xg.NumNodes - 1;
                        
                        if (xrefcount == 0)
                        {
                            Condition Check = new Condition();
                            parseGuidFromFile = Check.StartParseGUID(PathName);
                            editor.WriteMessage("\n Parse GUID: " + parseGuidFromFile);

                            fl_main_file = true;

                            editor.WriteMessage("\n No xref found in drawing");
                            acDoc.Database.SaveAs(filePath, true, DwgVersion.Current, acDoc.Database.SecurityParameters);
                            string mainFile = tdmsApp.ExecuteScript(moduleName, LoadFileToPSD, filePath, parseGuidFromFile, fl_main_file); 
                        }
                        else
                        {
                            Condition Check = new Condition();
                            parseGuidFromFile = Check.StartParseGUID(PathName);
                            editor.WriteMessage("\n Parse GUID: " + parseGuidFromFile);
                            for (int r = 1; r < (xrefcount + 1); r++)
                            {
                                XrefGraphNode child = xg.GetXrefNode(r);
                                if (child.XrefStatus == XrefStatus.Resolved)
                                {
                                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(child.BlockTableRecordId, OpenMode.ForWrite);

                                    db.XrefEditEnabled = true;
                                    try
                                    {
                                        string originalPath = btr.PathName;
                                        string childname = Path.GetFileName(originalPath);

                                        editor.WriteMessage("\n childname:         {0}", childname);
                                        editor.WriteMessage("\n moduleName:        {0}", moduleName);
                                        editor.WriteMessage("\n LoadFileToPSD:     {0}", LoadFileToPSD);
                                        editor.WriteMessage("\n originalPath:      {0}", originalPath);
                                        editor.WriteMessage("\n parseGuidFromFile: {0}", parseGuidFromFile);
                                        editor.WriteMessage("\n fl_main_file:      {0}", fl_main_file.ToString());

                                        string newpath = tdmsApp.ExecuteScript(moduleName, LoadFileToPSD, originalPath, parseGuidFromFile, fl_main_file) + @"\" + childname;
                                        //string newpath = @"d:\good\";
                                        
                                        btr.PathName = newpath;

                                        //изменение путей внешних ссылок для внешних ссылок
                                        refremap(originalPath, newpath);
                                        
                                        editor.WriteMessage("\n xref old path: " + originalPath);
                                        editor.WriteMessage("\n xref new path: " + newpath + " xref fixed !!!");
                                    }
                                    catch (System.Exception ex)
                                    {
                                        editor.WriteMessage("\n Path Xref Not found: " + ex);
                                    } 
                                }
                            }
                            fl_main_file = true;

                            editor.WriteMessage("\n moduleName:         {0}", moduleName);
                            editor.WriteMessage("\n LoadFileToPSD:      {0}", LoadFileToPSD);
                            editor.WriteMessage("\n filePath:           {0}", filePath);
                            editor.WriteMessage("\n parseGuidFromFile:  {0}", parseGuidFromFile);
                            editor.WriteMessage("\n fl_main_file:       {0}", fl_main_file.ToString());
                            string mainFile = tdmsApp.ExecuteScript(moduleName, LoadFileToPSD, filePath, parseGuidFromFile, fl_main_file);  
                        }
                        tr.Commit();
                        acDoc.Database.SaveAs(filePath, true, DwgVersion.Current, acDoc.Database.SecurityParameters);
                    }
                    RefreshXREF();
                    Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Maximized;
                }
                else
                {
                    Application.ShowAlertDialog("Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                }
            }
            catch (System.Exception ex)
            {
               editor.WriteMessage("\n Exception caught " + ex);
            }
        }

        private void refremap(string pathXref, string newPath)
        {
            Document Doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = Doc.Editor;
            Database db = new Database(false, true);
            using (db)
            {
                try
                {
                    db.ReadDwgFile(pathXref, System.IO.FileShare.ReadWrite, false, "");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage("\n Unable to read the drawingfile: " + ex);
                    return;
                }
                try
                {
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        ed.WriteMessage("\n --------Xrefs Details--------");
                        db.ResolveXrefs(true, false);

                        XrefGraph xg = db.GetHostDwgXrefGraph(true);
                        GraphNode root = xg.RootNode;

                        ObjectIdCollection objcoll = new ObjectIdCollection();

                        int xrefcount = xg.NumNodes - 1;

                        if (xrefcount == 0)
                        {
                            ed.WriteMessage("\n No xrefs found in the drawing");
                            return;
                        }
                        else
                        {
                            for (int r = 1; r < (xrefcount + 1); r++)
                            {
                                XrefGraphNode child = xg.GetXrefNode(r);
                                if (child.XrefStatus == XrefStatus.Resolved)
                                {
                                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(child.BlockTableRecordId, OpenMode.ForWrite);

                                    db.XrefEditEnabled = true;

                                    string originalpath = btr.PathName;
                                    string childname = Path.GetFileName(originalpath);
                                    string newpath = newPath + childname;
                                    
                                    refremap(originalpath, newpath);
                                    btr.PathName = newpath;

                                    ed.WriteMessage("\n xref old path: " + originalpath);
                                    ed.WriteMessage("\n xref new path: " + newpath + " xref fixed !!");
                                }
                            }
                            db.SaveAs(pathXref, true, DwgVersion.Current, Doc.Database.SecurityParameters);
                        }
                        tr.Commit();
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage("Work with base drawing Error: " + ex);
                }
            }
        }
        // переподгрузка всех внешних ссылок
        [CommandMethod("RefreshXREF")]
        public void RefreshXREF()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            try
            {
                //RegistryKey myKey = Registry.CurrentUser.OpenSubKey("Software\\Autodesk\\AutoCAD\\R19.0\\ACAD-B001:409\\LanguagePack"); //для AutoCAD 2013
                RegistryKey myKey = Registry.CurrentUser.OpenSubKey("Software\\Autodesk\\AutoCAD\\R19.0\\ACAD-B004:409\\LanguagePack"); //для AutoCAD Architecture 2013
                String path = (String)(myKey.GetValue("LocalRootFolder"));
                editor.WriteMessage("\n " + path);

                string local = path.ToLower();
                string parseGuid = null;
                Regex reg = new Regex("rus", RegexOptions.IgnoreCase);
                MatchCollection mc = reg.Matches(local);
                foreach (Match mat in mc)
                {
                    parseGuid = mat.Value.ToString();
                }
                if (parseGuid == "rus")
                {
                    //Для русской версии AutoCAD
                    doc.SendStringToExecute(".-ССЫЛКА О * " + "\n", true, false, false);
                    doc.SendStringToExecute(".-ИЗОБ О * " + "\n", true, false, false);
                    editor.WriteMessage("\n " + parseGuid);
                }
                else
                {
                    //Для английской версии AutoCAD
                    doc.SendStringToExecute(".-XREF R * " + "\n", true, false, false);
                    doc.SendStringToExecute(".-IMAGE R * " + "\n", true, false, false);
                    editor.WriteMessage("\n " + parseGuid);
                }
                editor.WriteMessage("\n Внешние ссылки обновлены.");
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught", ex);
            }
        }
        // Поиск и выгрузка из ТДМС изображений
        public static void FindImages()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = doc.Database;
            Editor ed = doc.Editor;
            XREF externalReferences = new XREF();
            try
            {
                string parseGuid = null;
                TDMSApplication tdmsApp = new TDMSApplication();
                TDMSObject tdmsObj = null;
                Regex reg = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    TypedValue[] filterlist = new TypedValue[1];
                    filterlist[0] = new TypedValue(0, "IMAGE");
                    SelectionFilter filter = new SelectionFilter(filterlist);
                    PromptSelectionResult selRes = ed.SelectAll(filter);
                    if (selRes.Status != PromptStatus.OK)
                    {
                        ed.WriteMessage("\n No Images selected");
                        return;
                    }
                    SelectionSet oSS = selRes.Value;
                   
                    //ed.WriteMessage("\nNumber of raster images in dwg "+ oSS.Count.ToString());
                    for (int i = 0; i < oSS.Count; i++)
                    {
                        RasterImage oRaster;
                        oRaster = (RasterImage)acTrans.GetObject(oSS[i].ObjectId,OpenMode.ForRead);
                        BlockTableRecord oBlkTblRec =(BlockTableRecord)acTrans.GetObject(oRaster.BlockId,OpenMode.ForRead);
                        if (oBlkTblRec.IsLayout)
                        {
                            Layout lyt = (Layout)acTrans.GetObject(oBlkTblRec.LayoutId, OpenMode.ForRead);
                            RasterImageDef oRasterIDef = (RasterImageDef)acTrans.GetObject(oRaster.ImageDefId,OpenMode.ForRead);
                            ed.WriteMessage("\n Raster image" + oRasterIDef.SourceFileName.ToString()); //получаем сохранённый путь растров

                            Document acDoc = Application.DocumentManager.MdiActiveDocument;
                            string ImagePath = oRasterIDef.SourceFileName.ToString();
                            ImagePath = ImagePath.Remove(7);
                            if (ImagePath != "c:\\Temp")
                            {
                                ed.WriteMessage("\n parsePath: {0}", ImagePath);
                                ed.WriteMessage("\n Объект внешней ссылки не принадлежит TDMS!");
                                externalReferences.RefreshXREF();
                            }
                            else
                            {
                                MatchCollection mc = reg.Matches(oRasterIDef.SourceFileName.ToString());
                                foreach (Match mat in mc)
                                {
                                    parseGuid += mat.Value.ToString();
                                }
                                parseGuid = parseGuid.Remove(0, 38);
                                ed.WriteMessage("\n Parse Path: {0}", parseGuid);
                                tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);
                                //выгрузка чертежа из базы тдмс
                                //tdmsObj.CheckOut();
                                var mainFile = tdmsObj.Files.Main;
                                mainFile.CheckOut(oRasterIDef.SourceFileName.ToString());
                                parseGuid = null;
                            }
                        }
                    }
                    acTrans.Commit();
                }
               externalReferences.RefreshXREF();
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage(ex.ToString());
            }
        }
        //Регистрация реакции на событие "Создание документа"
        [CommandMethod("AddEvent")]
        public void AddEvent()
        {
           //Application.DocumentManager.DocumentActivated += new DocumentCollectionEventHandler(docRefresh);
           Application.DocumentManager.DocumentToBeActivated += new DocumentCollectionEventHandler(docRefresh);
        }
        //Вызов реакции на событие
        //public void docActions(object senderObj, DocumentCollectionEventArgs docColDocActEvtArgs)
        //{
           
        //}
        public void docRefresh(object senderObj, DocumentCollectionEventArgs docColDocActEvtArgs)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            doc.SendStringToExecute("_TDMSRIBBON " + "\n", true, false, false);
            Technical rib = new Technical();
            rib.TDMSRIBBON();
            doc.SendStringToExecute("_REFRESHXREF " + "\n", true, false, false);
            RefreshXREF();
            //следует вызывать только при активации документов, 
            //т.к. если использовать при создании, то получатся двойные надписи в заголовках.
            Technical.TitleDoc();
            
        }
///////////////////////////////////////////////////////////////////////////////////////
        [CommandMethod("XREFUpdate")]
        public void XREFUpdate()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            try
            {
                Condition Check = new Condition();
                if (Check.CheckPath() == true)
                {
                    if (Check.CheckTDMSProcess() == true)
                    {
                        string parseP;
                        Database acCurDb = Application.DocumentManager.MdiActiveDocument.Database;
                        Database db = doc.Database;
                        string parseGuid = null;
                        TDMSApplication tdmsApp = new TDMSApplication();
                        TDMSObject tdmsObj = null;
                        Regex reg = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);

                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
                            foreach (ObjectId id in bt)
                            {
                                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForWrite);
                                if (btr.IsFromExternalReference)
                                {
                                    ed.WriteMessage("\n PathName: {0}", btr.PathName);
                                    ed.WriteMessage("\n Name: {0}", btr.Name);

                                    parseP = btr.PathName;
                                    parseP = parseP.Remove(7);

                                    if (parseP != "c:\\Temp")
                                    {
                                        ed.WriteMessage("\n parsePath: {0}", parseP);
                                        ed.WriteMessage("\n Объект внешней ссылки не принадлежит TDMS!");
                                        RefreshXREF();
                                    }
                                    else
                                    {
                                        //парсинг пути, получение актуального GUID объекта
                                        MatchCollection mc = reg.Matches(btr.PathName);
                                        foreach (Match mat in mc)
                                        {
                                            parseGuid += mat.Value.ToString();
                                        }
                                        parseGuid = parseGuid.Remove(0, 38);
                                        ed.WriteMessage("\n Parse Path: {0}", parseGuid);
                                        tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);
                                        //выгрузка чертежа из базы тдмс
                                        tdmsObj.CheckOut();
                                        var mainFile = tdmsObj.Files.Main;
                                        mainFile.CheckOut(btr.PathName);
                                        
                                        parseGuid = null;

                                        RefreshXREF();
                                    }
                                }
                            }
                        }
                        FindImages();
                    }
                    else
                    {
                        RefreshXREF();
                        //Application.ShowAlertDialog("\n Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                        ed.WriteMessage("\n Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                    }
                }
                else
                {
                    RefreshXREF();
                    //Application.ShowAlertDialog("\n Документ не принадлежит TDMS!");
                    ed.WriteMessage("\n Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\n Exception caught" + ex); 
            }
        }
///////////////////////////////////////////////////////////////////////////////////////
        public void XREFUpdateWithoutMsg()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            try
            {
                Condition Check = new Condition();
                if (Check.CheckPath() == true)
                {
                    if (Check.CheckTDMSProcess() == true)
                    {
                        string pPath;
                        Database acCurDb = Application.DocumentManager.MdiActiveDocument.Database;
                        Database db = doc.Database;
                        string parseGuid = null;
                        TDMSApplication tdmsApp = new TDMSApplication();
                        TDMSObject tdmsObj = null;
                        Regex reg = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);

                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
                            foreach (ObjectId id in bt)
                            {
                                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForWrite);
                                if (btr.IsFromExternalReference)
                                {
                                    pPath = btr.PathName;
                                    pPath = pPath.Remove(7);
                                    if (pPath != "c:\\Temp")
                                    {
                                        ed.WriteMessage("\n parsePath: {0}", pPath);
                                        ed.WriteMessage("\n Документ не принадлежит TDMS!");
                                    }
                                    else
                                    {
                                        //парсинг пути, получение актуального GUID объекта
                                        MatchCollection mc = reg.Matches(btr.PathName);
                                        foreach (Match mat in mc)
                                        {
                                            parseGuid += mat.Value.ToString();
                                        }
                                        parseGuid = parseGuid.Remove(0, 38);
                                        ed.WriteMessage("\n Parse Path: {0}", parseGuid);
                                        tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);
                                        //выгрузка чертежа из базы тдмс
                                        //tdmsObj.CheckOut();
                                        var mainFile = tdmsObj.Files.Main;
                                        mainFile.CheckOut(btr.PathName);
                                        parseGuid = null;
                                        RefreshXREF();
                                    }
                                }
                            }
                        }
                        FindImages();
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\n Exception caught" + ex);
            }
        }
///////////////////////////////////////////////////////////////////////////////////////

        //[CommandMethod("TDMSXREFDWG")]
        //public void TDMSXREFDWG()
        //{
        //    Document doc = Application.DocumentManager.MdiActiveDocument;
        //    Editor ed = doc.Editor;

        //    Technical checkP = new Technical();
        //    if (checkP.CheckPath() == true)
        //    {
        //        if (Check.CheckTDMSProcess() == true)
        //        {
        //            TDMSApplication tdmsAppDlg = new TDMSApplication();
        //            string path;
        //            Database acCurDb;
        //            PromptPointOptions pointOptions = new PromptPointOptions("УКАЖИТЕ ТОЧКУ: ");
        //            acCurDb = Application.DocumentManager.MdiActiveDocument.Database;
        //            string moduleName = "CMD_SYSLIB";
        //            string functionName = "CheckOutSelObj";
        //            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
        //            {
        //                string PathName = tdmsAppDlg.ExecuteScript(moduleName, functionName);
        //                PromptPointResult pointResult = ed.GetPoint(pointOptions);
        //                Point3d PointFrame = pointResult.Value;
        //                path = PathName;

        //                ObjectId acXrefId = acCurDb.AttachXref(path, "Exterior Elevations");
        //                if (!acXrefId.IsNull)
        //                {
        //                    Point3d insPt = new Point3d(PointFrame.X, PointFrame.Y, 0);//Point3d(PointFrame.X, PointFrame.Y, 0);
        //                    using (BlockReference acBlkRef = new BlockReference(insPt, acXrefId))
        //                    {
        //                        BlockTableRecord acBlkTblRec;
        //                        acBlkTblRec = acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

        //                        acBlkTblRec.AppendEntity(acBlkRef);
        //                        acTrans.AddNewlyCreatedDBObject(acBlkRef, true);
        //                    }
        //                }
        //                ed.WriteMessage(PathName);
        //                acTrans.Commit();
        //            }
        //        }
        //        else
        //        {
        //            ed.WriteMessage("Клиент ТДМС не запущен!");
        //        }
        //    }
        //    else
        //    {
        //        ed.WriteMessage("Чертёж не сохранён в ТДМС");
        //    }
        //}
        //Вставка изображения как внешней ссылки

///////////////////////////////////////////////////////////////////////////////////////
        [CommandMethod("TDMSXREFIMG")]
        public void TDMSXREFIMG()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            try
            {
                Condition Check = new Condition();
                if (Check.CheckPath() == true)
                {
                    if (Check.CheckTDMSProcess() == true)
                    {
                        TDMSApplication tdmsApp = new TDMSApplication();
                        Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Minimized;
                        tdmsApp.Visible = false;
                        tdmsApp.Visible = true;
                        string path;

                        Database acCurDb;
                        PromptPointOptions pointOptions = new PromptPointOptions("\n УКАЖИТЕ ТОЧКУ: ");
                        acCurDb = Application.DocumentManager.MdiActiveDocument.Database;
                        string moduleName = "CMD_SYSLIB";
                        string functionName = "CheckOutSelObj";

                        string PathName = tdmsApp.ExecuteScript(moduleName, functionName);

                        if (PathName != "")
                        {
                            TDMSObject tdmsObj;
                            //Берём GUID из пути файла выгруженного на диск
                            string docName = doc.Name;
                            string guidFromFile = docName;
                            string parseGuidFromFile = null;
                            Regex regFF = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                            MatchCollection mcFF = regFF.Matches(guidFromFile);
                            foreach (Match mat in mcFF)
                            {
                                parseGuidFromFile += mat.Value.ToString();
                            }
                            parseGuidFromFile = parseGuidFromFile.Remove(0, 38);
                            //Берём GUID из выбранного в TDMS файла - внешней ссылки
                            string guidFromTDMS = PathName;
                            string parseGuidFromTDMS = null;
                            Regex regFT = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                            MatchCollection mcFT = regFT.Matches(guidFromTDMS);
                            foreach (Match mat in mcFT)
                            {
                                parseGuidFromTDMS += mat.Value.ToString();
                            }
                            parseGuidFromTDMS = parseGuidFromTDMS.Remove(0, 38);
                            tdmsObj = tdmsApp.GetObjectByGUID(parseGuidFromTDMS);
                            string rastrName = tdmsObj.Files.Main.FileName;

                            if (parseGuidFromFile != parseGuidFromTDMS)
                            {
                                string parseGuidJPEG = null;
                                string parseGuidTIFF = null;
                                string parseGuidPNG = null;
                                string parseGuidBMP = null;

                                string parseGuidJPEGUpper = null;
                                string parseGuidTIFFUpper = null;
                                string parseGuidPNGUpper = null;
                                string parseGuidBMPUpper = null;

                                Regex regJPEG = new Regex(".jpg", RegexOptions.IgnoreCase);
                                MatchCollection mcJPEG = regJPEG.Matches(PathName);
                                foreach (Match mat in mcJPEG)
                                {
                                    parseGuidJPEG += mat.Value.ToString();
                                }

                                Regex regTIFF = new Regex(".tif", RegexOptions.IgnoreCase);
                                MatchCollection mcTIFF = regTIFF.Matches(PathName);
                                foreach (Match mat in mcTIFF)
                                {
                                    parseGuidTIFF += mat.Value.ToString();
                                }

                                Regex regPNG = new Regex(".png", RegexOptions.IgnoreCase);
                                MatchCollection mcPNG = regPNG.Matches(PathName);
                                foreach (Match mat in mcPNG)
                                {
                                    parseGuidPNG += mat.Value.ToString();
                                }

                                Regex regBMP = new Regex(".bmp", RegexOptions.IgnoreCase);
                                MatchCollection mcBMP = regBMP.Matches(PathName);
                                foreach (Match mat in mcBMP)
                                {
                                    parseGuidBMP += mat.Value.ToString();
                                }

                                if (parseGuidJPEG == ".jpg" | parseGuidTIFF == ".tif" | parseGuidTIFF == ".TIF" | parseGuidJPEG == ".JPG" | parseGuidPNG == ".png" | parseGuidPNGUpper == ".PNG" | parseGuidBMP == ".bmp" | parseGuidBMPUpper == ".BMP")
                                {
                                    ed.WriteMessage("\n jpg: " + parseGuidJPEG);
                                    ed.WriteMessage("\n tif: " + parseGuidTIFF);
                                    ed.WriteMessage("\n png: " + parseGuidPNG);
                                    ed.WriteMessage("\n bmp: " + parseGuidBMP);

                                    ed.WriteMessage("\n JPG: " + parseGuidJPEGUpper);
                                    ed.WriteMessage("\n TIF: " + parseGuidTIFFUpper);
                                    ed.WriteMessage("\n png: " + parseGuidPNGUpper);
                                    ed.WriteMessage("\n bmp: " + parseGuidBMPUpper);

                                    //tdmsObj.CheckOut();
                                    var mainFile = tdmsObj.Files.Main;
                                    mainFile.CheckOut(PathName);

                                    parseGuidFromTDMS = null;

                                    Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Maximized;

                                    using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                                    {
                                        PromptPointResult pointResult = ed.GetPoint(pointOptions);
                                        Point3d PointFrame = pointResult.Value;
                                        path = PathName;
                                        string strImgName = rastrName;
                                        RasterImageDef acRasterDef;
                                        bool bRasterDefCreated = false;
                                        ObjectId acImgDefId;
                                        ObjectId acImgDctID = RasterImageDef.GetImageDictionary(acCurDb);
                                        if (acImgDctID.IsNull)
                                        {
                                            acImgDctID = RasterImageDef.CreateImageDictionary(acCurDb);
                                        }
                                        DBDictionary acImgDict = acTrans.GetObject(acImgDctID, OpenMode.ForRead) as DBDictionary;
                                        if (acImgDict.Contains(strImgName))
                                        {
                                            acImgDefId = acImgDict.GetAt(strImgName);
                                            acRasterDef = acTrans.GetObject(acImgDefId, OpenMode.ForWrite) as RasterImageDef;
                                        }
                                        else
                                        {
                                            RasterImageDef acRasterDefNew = new RasterImageDef();
                                            acRasterDefNew.SourceFileName = path;
                                            acRasterDefNew.Load();
                                            acImgDict.UpgradeOpen();
                                            acImgDefId = acImgDict.SetAt(strImgName, acRasterDefNew);
                                            acTrans.AddNewlyCreatedDBObject(acRasterDefNew, true);
                                            acRasterDef = acRasterDefNew;
                                            bRasterDefCreated = true;
                                        }
                                        BlockTable acBlkTbl;
                                        acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                                        BlockTableRecord acBlkTblRec;
                                        acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                                        OpenMode.ForWrite) as BlockTableRecord;
                                        using (RasterImage acRaster = new RasterImage())
                                        {
                                            acRaster.ImageDefId = acImgDefId;
                                            Vector3d width;
                                            Vector3d height;
                                            if (acCurDb.Measurement == MeasurementValue.English)
                                            {
                                                width = new Vector3d((acRasterDef.ResolutionMMPerPixel.X * acRaster.ImageWidth) / 25.4, 0, 0);
                                                height = new Vector3d(0, (acRasterDef.ResolutionMMPerPixel.Y * acRaster.ImageHeight) / 25.4, 0);
                                            }
                                            else
                                            {
                                                width = new Vector3d(acRasterDef.ResolutionMMPerPixel.X * acRaster.ImageWidth, 0, 0);
                                                height = new Vector3d(0, acRasterDef.ResolutionMMPerPixel.Y * acRaster.ImageHeight, 0);
                                            }
                                            Point3d insPt = new Point3d(PointFrame.X, PointFrame.Y, 0.0);
                                            CoordinateSystem3d coordinateSystem = new CoordinateSystem3d(insPt, width * 2, height * 2);
                                            acRaster.Orientation = coordinateSystem;
                                            acRaster.Rotation = 0;

                                            acBlkTblRec.AppendEntity(acRaster);
                                            acTrans.AddNewlyCreatedDBObject(acRaster, true);

                                            RasterImage.EnableReactors(true);
                                            acRaster.AssociateRasterDef(acRasterDef);

                                            if (bRasterDefCreated)
                                            {
                                                acRasterDef.Dispose();
                                            }
                                        }
                                        acTrans.Commit();
                                    }
                                    //XREFUpdate();
                                }
                                else
                                {
                                    Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Maximized;
                                    Application.ShowAlertDialog("\n Несоответствие формата выбранного файла!");
                                }
                            }
                            else
                            {
                                Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Maximized;
                                ed.WriteMessage("\n PathNameFromTDMS: " + parseGuidFromTDMS);
                                ed.WriteMessage("\n docName: " + parseGuidFromFile);
                                Application.ShowAlertDialog("\n Нельзя использовать исходный чертёж как ссылку на самого себя!");
                            }
                        }
                        else
                        {
                            Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Maximized;
                            ed.WriteMessage("\n Вставка внешней ссылки отменена!");
                        }
                    }
                    else
                    {
                        Application.ShowAlertDialog("\n Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                    }
                }
                else
                {
                    Application.ShowAlertDialog("\n Документ не принадлежит TDMS!");
                }
            }
            catch(System.Exception ex)
            {
                ed.WriteMessage("\n Exception caught" + ex);
            }
        }

        //Вставка DWG как внешней ссылки
        [AcRtm.CommandMethod("TDMSXREFDWG", AcRtm.CommandFlags.Modal)]
        public void TDMSXREFDWG()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            string docName = doc.Name;
            Editor ed = doc.Editor;
            try
            {
                Condition Check = new Condition();
                if (Check.CheckPath() == true)
                {
                    if (Check.CheckPath() == true)
                    {
                        TDMSApplication tdmsApp = new TDMSApplication();
                        string path;
                        PromptPointOptions pointOptions = new PromptPointOptions("\n УКАЖИТЕ ТОЧКУ: ");
                        string moduleName = "CMD_SYSLIB";
                        string functionName = "CheckOutSelObj";
                        AcDb.Database db = doc.Database;

                        Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Minimized;
                        tdmsApp.Visible = false;
                        tdmsApp.Visible = true;

                        string PathName = tdmsApp.ExecuteScript(moduleName, functionName);
                        ed.WriteMessage("\n Return GUID from ExecuteScript " + PathName);
                        if (PathName != "")
                        {
                            TDMSObject tdmsObj;
                            //Берём GUID из пути файла выгруженного на диск
                            string guidFromFile = docName;
                            string parseGuidFromFile = null;
                            Regex regFF = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                            MatchCollection mcFF = regFF.Matches(guidFromFile);
                            foreach (Match mat in mcFF)
                            {
                                parseGuidFromFile += mat.Value.ToString();
                            }
                            parseGuidFromFile = parseGuidFromFile.Remove(0, 38);
                            //Берём GUID из выбранного в TDMS файла - внешней ссылки
                            string guidFromTDMS = PathName;
                            string parseGuidFromTDMS = null;
                            Regex regFT = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                            MatchCollection mcFT = regFT.Matches(guidFromTDMS);
                            foreach (Match mat in mcFT)
                            {
                                parseGuidFromTDMS += mat.Value.ToString();
                            }
                            parseGuidFromTDMS = parseGuidFromTDMS.Remove(0, 38);
                            tdmsObj = tdmsApp.GetObjectByGUID(parseGuidFromTDMS);
                            if (parseGuidFromFile != parseGuidFromTDMS)
                            {
                                string guid = PathName;
                                string parseGuid = null;
                                Regex reg = new Regex(".dwg", RegexOptions.IgnoreCase);
                                MatchCollection mc = reg.Matches(guid);
                                foreach (Match mat in mc)
                                {
                                    parseGuid += mat.Value.ToString();
                                }
                                if (parseGuid == ".dwg" | parseGuid == ".DWG")
                                {
                                    ed.WriteMessage("\n PathNameFromTDMS: " + parseGuidFromTDMS);
                                    ed.WriteMessage("\n docName: " + parseGuidFromFile);

                                    //tdmsObj.CheckOut();
                                    var mainFile = tdmsObj.Files.Main;
                                    mainFile.CheckOut(PathName);
                                    
                                    Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Maximized;
                                    PromptPointResult pointResult = ed.GetPoint(pointOptions);
                                    Point3d PointFrame = pointResult.Value;
                                    path = PathName;
                                    AcDb.ObjectId xRefId = db.AttachXref(Path.Combine(Environment.GetFolderPath(
                                        Environment.SpecialFolder.MyDocuments), path), path);
                                    using (AcDb.Transaction tr = db.TransactionManager.StartTransaction())
                                    {
                                        //Создаём экземпляр вхождения блока для подключенной ссылки
                                        AcDb.BlockReference br = new AcDb.BlockReference(new AcGm.Point3d(PointFrame.X, PointFrame.Y, 0), xRefId);
                                        AcDb.BlockTable blockTable = (AcDb.BlockTable)tr.GetObject(db.BlockTableId, AcDb.OpenMode.ForRead);
                                        //Вставлять вхождение блока будем в пространстве Model
                                        AcDb.BlockTableRecord modalSpace = (AcDb.BlockTableRecord)
                                            tr.GetObject(blockTable[AcDb.BlockTableRecord.ModelSpace], AcDb.OpenMode.ForWrite);
                                        modalSpace.AppendEntity(br);
                                        tr.AddNewlyCreatedDBObject(br, true);
                                        tr.Commit();
                                    }
                                    XREFUpdate();
                                }
                                else
                                {
                                    Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Maximized;
                                    Application.ShowAlertDialog("\n Несоответствие формата выбранного файла!");
                                }
                            }
                            else
                            {
                                Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Maximized;
                                ed.WriteMessage("\n PathNameFromTDMS: " + parseGuidFromTDMS);
                                ed.WriteMessage("\n docName: " + parseGuidFromFile);
                                Application.ShowAlertDialog("Нельзя использовать исходный чертёж как ссылку на самого себя!");
                            }
                        }
                        else
                        {
                            Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Maximized;
                            ed.WriteMessage("\n Вставка внешней ссылки отменена!");
                        }
                        ////Внедряем содержимое ссылки в чертёж
                        ////ВНИМАНИЕ!
                        ////Следующий фрагмент кода должен находиться ПОСЛЕ блока "using", т.к. должен быть выполнен 
                        ////после строки кода tr.Commit(); (т.е. когда изменения будут зафиксированы в базе данных)
                        ////AcDb.ObjectIdCollection ids = new AcDb.ObjectIdCollection();            
                        ////ids.Add(xRefId); 
                        ////db.BindXrefs(ids, true);  
                    }
                    else
                    {
                        Application.ShowAlertDialog("\n Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                    }
                }
                else
                {
                    Application.ShowAlertDialog("Документ не принадлежит TDMS!");
                    ed.WriteMessage("\n Документ не принадлежит TDMS!");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\n Exception caught" + ex);
            }
        }

        //синхронизация внешних ссылок с внешними ссылками в ТДМС.
        [CommandMethod("XREFSYNCHRO")]
        public void XREFSynchro()
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                Condition ChangePath = new Condition();
                string strDwgName = acDoc.Name;
                //{
                //    if (checkP.CheckPath() == true)
                //    {
                //TDMSApplication tdmsApp = new TDMSApplication();
                //TDMSObject tdmsObj = null;
                //object obj = Application.GetSystemVariable("DWGTITLED");
                List<string> masPath = new List<string>();
                Database db = acDoc.Database;
                //Возвращение пути к файлу чертежа
                string guid = strDwgName;
                string parseGuid = null;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    editor.WriteMessage("\n PATH XREFERENCES: ");
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
                    foreach (ObjectId id in bt)
                    {
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForWrite);
                        //if (btr.IsFromExternalReference)
                        //{
                        if (btr.XrefStatus.ToString() == "Resolved")
                        {
                            if (ChangePath.StartCheckPath(btr.PathName) == false)
                            {
                                editor.WriteMessage("\n PathName: {0}", btr.PathName);
                                editor.WriteMessage("\n Name: {0}", btr.Name);
                                editor.WriteMessage("\n Status: {0}", btr.XrefStatus.ToString());
                                masPath.Add(btr.PathName);
                            }
                        }
                        //}
                    }
                }

                for (int j = 0; j < masPath.Count; j++)
                {
                    //вызываю функцию Лили, передаю ей путь каждой внешней ссылки или массив ссылок, получаю время изменения документа в TDMS, если время одинаковое,
                    //то ничего не делаю, если разное, то объект должен выгрузиться и должно пройти обновление внешних ссылок.
                    editor.WriteMessage("\n ArrayPath: {0}", masPath[j]);
                }

                //Получение объекта по GUID

                //tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);
                //bool lockOwner = tdmsObj.Permissions.LockOwner;
                //Заблокирован ли чертёж текущим пользователем?
                //if (lockOwner)
                //{
                //Сохранение изменений в текущем чертеже
                //загрузить в базу ТДМС, сохранить изменения
                //    tdmsObj.CheckIn();
                //    tdmsObj.Update();
                //    Technical.titleDoc();
                //}
                //else
                //{
                //    editor.WriteMessage("Документ открыт на просмотр; Изменения не будут сохранены в TDMS!");
                //    acDoc.Database.SaveAs(strDwgName, true, DwgVersion.Current, acDoc.Database.SecurityParameters);
                //}
                //}
                //else
                //{
                //   //здесь main file должен тоже сохраняться в ТДМС.
                //}
                //}
                //else
                //{
                //    Application.ShowAlertDialog("Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                //    acDoc.Database.SaveAs(strDwgName, true, DwgVersion.Current, acDoc.Database.SecurityParameters);
                //}
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("Exception caught" + ex);
            }
        }
    }

    public class Commands
    {
        
        [CommandMethod("InsertingBlockWithAnAttribute")]
        public void InsertingBlockWithAnAttribute()
        {
            Database acCurDb = Application.DocumentManager.MdiActiveDocument.Database;
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                BlockTable acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                ObjectId blkRecId = ObjectId.Null;

                if (!acBlkTbl.Has("CircleBlockWithAttributes"))
                {
                    using (BlockTableRecord acBlkTblRec = new BlockTableRecord())
                    {
                        acBlkTblRec.Name = "CircleBlockWithAttributes";

                        acBlkTblRec.Origin = new Point3d(0, 0, 0);

                        // Add a circle to the block
                        using (Circle acCirc = new Circle())
                        {
                            acCirc.Center = new Point3d(0, 0, 0);
                            acCirc.Radius = 2;
                            acBlkTblRec.AppendEntity(acCirc);
                            
                            // Add an attribute definition to the block
                            using (AttributeDefinition acAttDef = new AttributeDefinition())
                            {
                                acAttDef.Position = new Point3d(0, 0, 0);
                                acAttDef.Prompt = "Door #: ";
                                acAttDef.Tag = "Door#";
                                
                                MText acMText = new MText();
                                acMText.SetDatabaseDefaults();
                                acMText.Location = new Point3d(12, 23, 0);

                                acMText.Width = 10;
                                acMText.TextHeight = 2.25;

                                acMText.Contents = "\\pxqc;" + "DXX1231231 1231 123 123 12";

                                //acAttDef.TextString = "DXX";
                                
                                acAttDef.Height = 1;
                                acAttDef.IsMTextAttributeDefinition = true;
                                
                                acAttDef.MTextAttributeDefinition = acMText;
                                acAttDef.Justify = AttachmentPoint.MiddleCenter;
                                acBlkTblRec.AppendEntity(acAttDef);

                                acBlkTbl.UpgradeOpen();
                                acBlkTbl.Add(acBlkTblRec);
                                acTrans.AddNewlyCreatedDBObject(acBlkTblRec, true);
                            }
                        }

                        blkRecId = acBlkTblRec.Id;
                    }
                }
                else
                {
                    blkRecId = acBlkTbl["CircleBlockWithAttributes"];
                }
                if (blkRecId != ObjectId.Null)
                {
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = acTrans.GetObject(blkRecId, OpenMode.ForRead) as BlockTableRecord;

                    // Create and insert the new block reference
                    using (BlockReference acBlkRef = new BlockReference(new Point3d(2, 2, 0), blkRecId))
                    {
                        BlockTableRecord acCurSpaceBlkTblRec;
                        acCurSpaceBlkTblRec = acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
                        acTrans.AddNewlyCreatedDBObject(acBlkRef, true);

                        if (acBlkTblRec.HasAttributeDefinitions)
                        {
                            // Add attributes from the block table record
                            foreach (ObjectId objID in acBlkTblRec)
                            {
                                DBObject dbObj = acTrans.GetObject(objID, OpenMode.ForRead) as DBObject;

                                if (dbObj is AttributeDefinition)
                                {
                                    AttributeDefinition acAtt = dbObj as AttributeDefinition;

                                    if (!acAtt.Constant)
                                    {
                                        using (AttributeReference acAttRef = new AttributeReference())
                                        {
                                            acAttRef.SetAttributeFromBlock(acAtt, acBlkRef.BlockTransform);
                                            acAttRef.Position = acAtt.Position.TransformBy(acBlkRef.BlockTransform);
                                                                                     
                                            acAttRef.TextString = acAtt.TextString;
                                            
                                            acBlkRef.AttributeCollection.AppendAttribute(acAttRef);

                                            acTrans.AddNewlyCreatedDBObject(acAttRef, true);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // Save the new object to the database
                acTrans.Commit();

                // Dispose of the transaction
            }
        }
        //[CommandMethod("UAIF")]
        //public void UpdateAttributeInFiles()
        //{
        //    Document doc = Application.DocumentManager.MdiActiveDocument;
        //    Editor ed = doc.Editor;
        //    PromptResult pr = ed.GetString("\n Enter folder containing DWGs to process: ");
        //    if (pr.Status != PromptStatus.OK)
        //        return;
        //    string pathName = pr.StringResult;

        //    pr = ed.GetString("\nEnter name of block to search for: ");

        //    if (pr.Status != PromptStatus.OK)
        //        return;
        //    string blockName = pr.StringResult.ToUpper();

        //    pr = ed.GetString("\nEnter tag of attribute to update: ");

        //    if (pr.Status != PromptStatus.OK)
        //        return;
        //    string attbName = pr.StringResult.ToUpper();

        //    pr = ed.GetString("\nEnter new value for attribute: ");

        //    if (pr.Status != PromptStatus.OK)
        //        return;
        //    string attbValue = pr.StringResult;

        //    string[] fileNames = Directory.GetFiles(pathName, "*.dwg");

        //    int processed = 0, saved = 0, problem = 0;

        //    foreach (string fileName in fileNames)
        //    {
        //        if (fileName.EndsWith(".dwg", StringComparison.CurrentCultureIgnoreCase))
        //        {
        //            string outputName = fileName.Substring(0, fileName.Length - 4) + "_updated.dwg";
        //            Database db = new Database(false, false);
        //            using (db)
        //            {
        //                try
        //                {
        //                    ed.WriteMessage("\n\nProcessing file: " + fileName);
        //                    db.ReadDwgFile(fileName, FileShare.ReadWrite, false, "");
        //                    int attributesChanged = UpdateAttributesInDatabase(db, blockName, attbName, attbValue);
        //                    // Display the results
        //                    ed.WriteMessage("\nUpdated {0} instance{1} of " + "attribute {2}.", attributesChanged, attributesChanged == 1 ? "" : "s", attbName);
        //                    // Only save if we changed something
        //                    if (attributesChanged > 0)
        //                    {
        //                        ed.WriteMessage("\nSaving to file: {0}", outputName);
        //                        db.SaveAs(outputName, DwgVersion.Current);
        //                        saved++;
        //                    }
        //                    processed++;
        //                }
        //                catch (System.Exception ex)
        //                {
        //                    ed.WriteMessage("\nProblem processing file: {0} - \"{1}\"", fileName, ex.Message);
        //                    problem++;
        //                }
        //            }
        //        }
        //    }
        //    ed.WriteMessage("\n\nSuccessfully processed {0} files, of which {1} had " + "attributes to update and an additional {2} had errors " + "during reading/processing.", processed, saved, problem);
        //}
        
       
        
       
       
     

        [CommandMethod("UA")]
        public void UpdateAttribute()
        {
            Condition Check = new Condition();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor Editor = doc.Editor;
            try
            {
                if (Check.CheckPath() == true)
                {
                    if (Check.CheckTDMSProcess() == true)
                    {
                        Technical Refresh = new Technical();
                        string moduleName = "CMD_SYSLIB";
                        string _GDFOFA = "GETDATAFROMOBJFORACAD";
                        string _GDFPOFA = "GETDATAFROMPAROBJFORACAD";
                        string _GDFARTFA = "GETDATAFROMACTIVEROUTETABLEFORACAD";
                        string _GDFSPFA = "GETDATAFROMSYSPROPSFORACAD";
                        string _GOA = "GETOPADDRESS";
                        string _GON = "GETOPNAME";

                        //Наименования атрибутов
                        string nameObj = "A_DESIGN_OBJ_REF";
                        string oboznach = "A_OBOZN_DOC";
                        string inventNum = "A_ARCH_SIGN";
                        string pageNum = "A_PAGE_NUM";
                        string pageCount = "A_PAGE_COUNT";
                        string tomPage = "A_TOM_PAGE_NUM";

                        //Обновление обычных атрибутов
                        Refresh.RefreshAttribute(nameObj);
                        Refresh.RefreshAttribute(oboznach);
                        Refresh.RefreshAttribute(inventNum);
                        Refresh.RefreshAttribute(pageNum);
                        Refresh.RefreshAttribute(pageCount);
                        Refresh.RefreshAttribute(tomPage);

                        //Обновление атрибутов функций с разным количеством параметров
                        Refresh.RefreshAttribute(moduleName, _GDFOFA, "A_STAGE_CLSF"); // Стадия
                        Refresh.RefreshAttribute(moduleName, _GDFPOFA, "A_INSTEAD_OF_NUM"); // взамен инв №
                        Refresh.RefreshAttribute(moduleName, _GDFSPFA, "GKAB_"); //Гл.констр.АБ
                        Refresh.RefreshAttribute(moduleName, _GOA);// Адрес
                        Refresh.RefreshAttribute(moduleName, _GON);// Наименование объекта

                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "DEVELOP", "A_User"); // Разработал
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "CHECK", "A_User"); //Проверил
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "NORMKL", "A_User"); //Нормоконтроль
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GR_HEAD", "A_User"); //Руководитель группы
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GIP_", "A_User"); //ГИП
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GAP_", "A_User"); //ГAП
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GKP_", "A_User"); //Главный конструктор проекта

                        //Обновление дат
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "DEVELOP", "A_DATE");
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "CHECK", "A_DATE");
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "NORMKL", "A_DATE");
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GR_HEAD", "A_DATE");
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GIP_", "A_DATE");
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GAP_", "A_DATE");
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GKP_", "A_DATE");
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GKAB_", "A_DATE");

                        Editor.Regen();
                    }
                    else
                    {
                        Application.ShowAlertDialog("\n Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                    }
                }
                else
                {
                    Application.ShowAlertDialog("Документ не принадлежит TDMS!");
                }
            }
            catch (System.Exception ex)
            {
                Editor.WriteMessage("Exception caught " + ex);
            }
        }

        public void UpdateAttributeWithoutMsg()
        {
            Condition Check = new Condition();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor Editor = doc.Editor;
            try
            {
                if (Check.CheckPath() == true)
                {
                    if (Check.CheckTDMSProcess() == true)
                    {

                        Technical Refresh = new Technical();
                        string moduleName = "CMD_SYSLIB";
                        string _GDFOFA = "GETDATAFROMOBJFORACAD";
                        string _GDFPOFA = "GETDATAFROMPAROBJFORACAD";
                        string _GDFARTFA = "GETDATAFROMACTIVEROUTETABLEFORACAD";
                        string _GDFSPFA = "GETDATAFROMSYSPROPSFORACAD";
                        string _GOA = "GETOPADDRESS";
                        string _GON = "GETOPNAME";

                        //Наименования атрибутов
                        string nameObj = "A_DESIGN_OBJ_REF";
                        string oboznach = "A_OBOZN_DOC";
                        string inventNum = "A_ARCH_SIGN";
                        string pageNum = "A_PAGE_NUM";
                        string pageCount = "A_PAGE_COUNT";
                        string tomPage = "A_TOM_PAGE_NUM";
                        //Обновление обычных атрибутов
                        Refresh.RefreshAttribute(nameObj);
                        Refresh.RefreshAttribute(oboznach);
                        Refresh.RefreshAttribute(inventNum);
                        Refresh.RefreshAttribute(pageNum);
                        Refresh.RefreshAttribute(pageCount);
                        Refresh.RefreshAttribute(tomPage);

                        //Обновление атрибутов функций с разным количеством параметров
                        Refresh.RefreshAttribute(moduleName, _GDFOFA, "A_STAGE_CLSF"); // Стадия
                        Refresh.RefreshAttribute(moduleName, _GDFPOFA, "A_INSTEAD_OF_NUM"); // взамен инв №
                        Refresh.RefreshAttribute(moduleName, _GDFSPFA, "GKAB_"); //Гл.констр.АБ
                        Refresh.RefreshAttribute(moduleName, _GOA);// Адрес
                        Refresh.RefreshAttribute(moduleName, _GON);// Наименование объекта

                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "DEVELOP", "A_User"); // Разработал
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "CHECK", "A_User"); //Проверил
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "NORMKL", "A_User"); //Нормоконтроль
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GR_HEAD", "A_User"); //Руководитель группы
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GIP_", "A_User"); //ГИП
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GAP_", "A_User"); //ГAП
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GKP_", "A_User"); //Главный конструктор проекта

                        //Обновление дат
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "DEVELOP", "A_DATE");
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "CHECK", "A_DATE");
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "NORMKL", "A_DATE");
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GR_HEAD", "A_DATE");
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GIP_", "A_DATE");
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GAP_", "A_DATE");
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GKP_", "A_DATE");
                        Refresh.RefreshAttribute(moduleName, _GDFARTFA, "GKAB_", "A_DATE");

                        Editor.Regen();
                    }
                    else
                    {
                        Editor.WriteMessage("\n Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                    }
                }
                else
                {
                    Editor.WriteMessage("Документ не принадлежит TDMS!");
                }
            }
            catch (System.Exception ex)
            {
                Editor.WriteMessage("Exception caught" + ex);
            }
        }


        public int UpdateAttributesInDatabase(Database db, string blockName, string attbName, string attbValue)
        {
            //ed.WriteMessage("\n имя блока:         " + blockName);
            //ed.WriteMessage("\n имя атрибута:      " + attbName);
            //ed.WriteMessage("\n значение атрибута: " + attbValue);

            ObjectId msId, psId;
            Transaction tr = db.TransactionManager.StartTransaction();
            using (tr)
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                //msId = bt[BlockTableRecord.ModelSpace];
                psId = bt[BlockTableRecord.PaperSpace];
                tr.Commit();
            }
            //int msCount = UpdateAttributesInBlock(msId, blockName, attbName, attbValue);
            int psCount = UpdateAttributesInBlock(psId, blockName, attbName, attbValue);
            return /*msCount +*/  psCount;
        }
        public int UpdateAttributesInBlock(ObjectId btrId, string blockName, string attbName, string attbValue)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor Editor = doc.Editor;
            int changedCount = 0;
            Database db = doc.Database;

            Transaction tr = doc.TransactionManager.StartTransaction();
            using (tr)
            {
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                foreach (ObjectId entId in btr)
                {
                    Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                    if (ent != null)
                    {
                        BlockReference br = ent as BlockReference;
                        if (br != null)
                        {
                            BlockTableRecord bd = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                            foreach (ObjectId arId in br.AttributeCollection)
                            {
                                DBObject obj = tr.GetObject(arId, OpenMode.ForRead);
                                AttributeReference ar = obj as AttributeReference;
                                if (ar != null)
                                {

                                    if (ar.Tag.ToUpper() == attbName.ToUpper())
                                    {
                                        //ed.WriteMessage("\n Верхний регистр: " + ar.Tag.ToUpper());
                                        //ed.WriteMessage("\n Имя атрибута:    " + attbName.ToUpper());
                                        ar.UpgradeOpen();
                                        ar.TextString = attbValue;
                                        //ed.WriteMessage("\n Обновляемое значение: " + ar.TextString);
                                        ar.DowngradeOpen();

                                        changedCount++;
                                    }
                                }
                            }
                            changedCount += UpdateAttributesInBlock(br.BlockTableRecord, blockName, attbName, attbValue);
                        }
                    }
                }
                tr.Commit();
            }
            return changedCount;
        }
    }


    public sealed class Technical : IExtensionApplication
    {
        static Document doc = Application.DocumentManager.MdiActiveDocument;
        Editor editor = doc.Editor;

        public void Initialize()
        {
            try
            {
                doc.SendStringToExecute("_TDMSRIBBON " + "\n", true, false, false);
                TDMSRIBBON();
                XREF externalReferences = new XREF();
                externalReferences.AddEvent();
                externalReferences.XREFUpdateWithoutMsg();
                //externalReferences.RefreshXREF();
                Commands UA = new Commands();
                UA.UpdateAttributeWithoutMsg();
                TitleDoc();
                //Autodesk.Windows.ComponentManager.ItemInitialized += new System.EventHandler(ComponentManager_ItemInitialized);
                //TDMSRIBBON();

                //Document doc = Application.DocumentManager.MdiActiveDocument;
                //doc.SendStringToExecute("._TDMSRIBBON" + "\n", true, false, false);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        public void Terminate()
        {
        }

        public void DivisionDrawingOnLayouts()
        {
       
        }
    
        private void ComponentManager_ItemInitialized(object sender, Autodesk.Windows.RibbonItemEventArgs e)
        {
            try
            {
                if (Autodesk.Windows.ComponentManager.Ribbon != null)
                {
                    BuildRibbonTab();
                    Autodesk.Windows.ComponentManager.ItemInitialized -= new EventHandler<RibbonItemEventArgs>(ComponentManager_ItemInitialized);
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exeption caught: " + ex);
            }
        }
        
        private void BuildRibbonTab()
        {
            try
            {
                if (!isLoaded())
                {
                    TDMSRIBBON();
                    //acadApp.SystemVariableChanged += new SystemVariableChangedEventHandler(acadApp_SystemVariableChanged);
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        
        private bool isLoaded()
        {
            bool _loaded = false;
            RibbonControl ribCntrl = Autodesk.Windows.ComponentManager.Ribbon;
            foreach (RibbonTab tab in ribCntrl.Tabs)
            {
                if (tab.Id.Equals("TDMS") & tab.Title.Equals("TDMS"))
                {
                    _loaded = true;
                    break;
                }
                else
                {
                    _loaded = false;
                }
            }
            return _loaded;
        }

        [CommandMethod("titles")]
        public static void TitleDoc()
        {
            Editor editor = doc.Editor;
            Condition Check = new Condition();
            try
            {
                if (Check.CheckPath() == true)
                {
                    if (Check.CheckTDMSProcess() == true)
                    {
                        TDMSApplication tdmsApp = new TDMSApplication();
                        TDMSObject tdmsObj = null;
                        Document acDoc = Application.DocumentManager.MdiActiveDocument;
                        object obj = Application.GetSystemVariable("DWGTITLED");
                        string strDwgName = acDoc.Name;

                        //Возвращение пути к файлу чертежа
                        string guid = strDwgName;

                        //парсинг пути, получение актуального GUID объекта
                        string parseGuid = null;
                        Regex reg = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                        MatchCollection mc = reg.Matches(guid);
                        foreach (Match mat in mc)
                        {
                            parseGuid += mat.Value.ToString();
                        }
                        parseGuid = parseGuid.Remove(0, 38);
                        //Получение объекта по GUID
                        tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);
                        bool lockOwner = tdmsObj.Permissions.LockOwner;
                        //Заблокирован ли чертёж текущим пользователем
                        if (lockOwner)
                        {
                            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Window.Text += " TDMS - редактирование!";
                        }
                        else
                        {
                            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Window.Text += " TDMS - просмотр!";
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }

        [CommandMethod("OpenHelp")]
        public static void OpenHelp()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                System.Diagnostics.Process.Start(@"V:\_HELP\");
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught. Path not found " + ex);
            }
        }
        //Атрибут Автокада, ссылающийся на функцию, должен иметь формат:
        //CMD_SYSLIB - имя модуля
        //GetTitle - имя функции
        //далее идут имена параметров
        //Вызов происходит так 
        //attrValue = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetDataFromObjForAcad", tdmsObj, "A_STAGE_CLSF");
        //запускаем создание атрибута из функции, с разными параметрами (метод перегружен!!!)
        //проверяем, запущен ли процесс "TDMS.exe" возвращаем GUID чертежа, запускаем создание атрибута и присваиваем ему значение из объекта ТДМС.
        //без параметра
        public static void AddMultilineAttributeFunction(double x, double y, double height, double width, double widthFactor, double rotate, double oblique, string moduleName, string functionName, int blockName)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            Condition Check = new Condition();
            string attrName = moduleName + "_" + functionName + "#FUNCTION";
            editor.WriteMessage("\n ADDRESS: " + attrName);
            try
            {
                if (Check.CheckPath() == true)
                {
                    if (Check.CheckTDMSProcess() == true)
                    {
                        AttributeDefinition adAttr = new AttributeDefinition();
                        string attrValue = null;

                        TDMSApplication tdmsApp = new TDMSApplication();

                        TDMSObject tdmsObj = null;
                        Document acDoc = Application.DocumentManager.MdiActiveDocument;
                        object obj = Application.GetSystemVariable("DWGTITLED");
                        string strDwgName = acDoc.Name;
                        string guid = strDwgName;
                        string parseGuid = null;

                        Regex reg = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                        MatchCollection mc = reg.Matches(guid);
                        foreach (Match mat in mc)
                        {
                            parseGuid += mat.Value.ToString();
                        }
                        parseGuid = parseGuid.Remove(0, 38);
                        editor.WriteMessage("\n ADDRESS: " + parseGuid);
                        tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);
                        editor.WriteMessage("\n ADDRESS: TDMSOBJECT;" );
                        attrValue = tdmsApp.ExecuteScript(moduleName, functionName,  tdmsObj);
                        editor.WriteMessage("\n ADDRESS: " + attrValue);
                        if (attrValue != "")
                        {
                            Creator.CreateMultilineStampAtribut(x, y, attrValue, height, width, widthFactor, rotate, oblique, attrName, blockName);
                        }
                        else
                        {
                            Creator.CreateMultilineStampAtribut(x, y, "Text", height, width, widthFactor, rotate, oblique, attrName, blockName);
                        }
                    }
                    else
                    {
                        Creator.CreateMultilineStampAtribut(x, y, "Text", height, width, widthFactor, rotate, oblique, attrName, blockName);
                    }
                }
                else
                {
                    Creator.CreateMultilineStampAtribut(x, y, "Text", height, width, widthFactor, rotate, oblique, attrName, blockName);
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        //1 параметр
        public static void AddAttributeFunction(double x, double y, double height, double widthFactor, double rotate, double oblique, string moduleName, string functionName, int blockName, string parameter)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            Condition Check = new Condition();
            string attrName = moduleName + "_" + functionName + "|" + parameter + "#FUNCTION";
            try
            {
                if (Check.CheckPath() == true)
                {
                    if (Check.CheckTDMSProcess() == true)
                    {
                        AttributeDefinition adAttr = new AttributeDefinition();
                        string attrValue = null;

                        TDMSApplication tdmsApp = new TDMSApplication();

                        TDMSObject tdmsObj = null;
                        Document acDoc = Application.DocumentManager.MdiActiveDocument;
                        object obj = Application.GetSystemVariable("DWGTITLED");
                        string strDwgName = acDoc.Name;
                        string guid = strDwgName;
                        string parseGuid = null;

                        Regex reg = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                        MatchCollection mc = reg.Matches(guid);
                        foreach (Match mat in mc)
                        {
                            parseGuid += mat.Value.ToString();
                        }
                        parseGuid = parseGuid.Remove(0, 38);

                        tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);
                        if (parameter != "GKAB_")
                        {
                            attrValue = tdmsApp.ExecuteScript(moduleName, functionName, tdmsObj, parameter);

                        }
                        else
                        {
                            attrValue = tdmsApp.ExecuteScript(moduleName, functionName, parameter);
                        }

                        if (attrValue != "")
                        {
                            Creator.CreateStampAtribut(x, y, attrValue, height, widthFactor, rotate, oblique, attrName, blockName);
                        }
                        else
                        {
                            Creator.CreateStampAtribut(x, y, "Text", height, widthFactor, rotate, oblique, attrName, blockName);
                        }
                    }
                    else
                    {
                        Creator.CreateStampAtribut(x, y, "Text", height, widthFactor, rotate, oblique, attrName, blockName);
                    }
                }
                else
                {
                    Creator.CreateStampAtribut(x, y, "Text", height, widthFactor, rotate, oblique, attrName, blockName);
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        // 2 дополнительных параметра
        public static void AddAttributeFunction(double x, double y, double height, double widthFactor, double rotate, double oblique, string moduleName, string functionName, int blockName, string parameter, string parameter2)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            Condition Check = new Condition();
            string attrName = moduleName + "_" + functionName + "|" + parameter + "|" + parameter2 + "#FUNCTION";
            editor.WriteMessage("\n" + attrName);
            try
            {
                if (Check.CheckPath() == true)
                {
                    if (Check.CheckTDMSProcess() == true)
                    {
                        AttributeDefinition adAttr = new AttributeDefinition();
                        string attrValue = null;
                        TDMSApplication tdmsApp = new TDMSApplication();
                        TDMSObject tdmsObj = null;
                        Document acDoc = Application.DocumentManager.MdiActiveDocument;
                        object obj = Application.GetSystemVariable("DWGTITLED");
                        string strDwgName = acDoc.Name;
                        string guid = strDwgName;
                        string parseGuid = null;

                        Regex reg = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                        MatchCollection mc = reg.Matches(guid);
                        foreach (Match mat in mc)
                        {
                            parseGuid += mat.Value.ToString();
                        }
                        parseGuid = parseGuid.Remove(0, 38);

                        tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);

                        attrValue = tdmsApp.ExecuteScript(moduleName, functionName, tdmsObj, parameter, parameter2);

                        editor.WriteMessage("\n Значение" + attrValue);
                        if (attrValue != "")
                        {
                            Creator.CreateStampAtribut(x, y, attrValue, height, widthFactor, rotate, oblique, attrName, blockName);
                        }
                        else
                        {
                            Creator.CreateStampAtribut(x, y, "Text", height, widthFactor, rotate, oblique, attrName, blockName);
                        }
                    }
                    else
                    {
                        Creator.CreateStampAtribut(x, y, "Text", height, widthFactor, rotate, oblique, attrName, blockName);
                    }
                }
                else
                {
                    Creator.CreateStampAtribut(x, y, "Text", height, widthFactor, rotate, oblique, attrName, blockName);
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception" + ex);
            }
        }
        //Добавляем обычный атрибут НЕ ФУНКЦИЯ!
        //проверяем, запущен ли процесс "TDMS.exe" возвращаем GUID чертежа, запускаем создание атрибута и присваиваем ему значение из объекта ТДМС.
        public static void AddAttribute(double x, double y, double height, double widthFactor, double rotate, double oblique, string attrName, int blockName)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            Condition Check = new Condition();
            try
            {
                if (Check.CheckPath() == true)
                {
                    if (Check.CheckTDMSProcess() == true)
                    {
                        AttributeDefinition adAttr = new AttributeDefinition();
                        string attrValue = null;
                        TDMSApplication tdmsApp = new TDMSApplication();
                        TDMSObject tdmsObj = null;
                        Document acDoc = Application.DocumentManager.MdiActiveDocument;
                        object obj = Application.GetSystemVariable("DWGTITLED");
                        string strDwgName = acDoc.Name;
                        string guid = strDwgName;
                        string parseGuid = null;
                        Regex reg = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                        MatchCollection mc = reg.Matches(guid);
                        foreach (Match mat in mc)
                        {
                            parseGuid += mat.Value.ToString();
                        }
                        parseGuid = parseGuid.Remove(0, 38);
                        tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);
                        attrValue = tdmsObj.Attributes[attrName].Value;
                        if (attrValue != "")
                        {
                            Creator.CreateStampAtribut(x, y, attrValue, height, widthFactor, rotate, oblique, attrName, blockName);
                        }
                        else
                        {
                            Creator.CreateStampAtribut(x, y, "Text", height, widthFactor, rotate, oblique, attrName, blockName);
                        }
                    }
                    else
                    {
                        Creator.CreateStampAtribut(x, y, "Text", height, widthFactor, rotate, oblique, attrName, blockName);
                    }
                }
                else
                {
                    Creator.CreateStampAtribut(x, y, "Text", height, widthFactor, rotate, oblique, attrName, blockName);
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        // Многострочный атрибут без параметра
        public static void AddAttributeMultiline(double x, double y, double height, double width, double widthFactor, double rotate, double oblique, string attrName, int blockName)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            Condition Check = new Condition();
            try
            {
                if (Check.CheckPath() == true)
                {
                    if (Check.CheckTDMSProcess() == true)
                    {
                        AttributeDefinition adAttr = new AttributeDefinition();
                        string attrValue = null;
                        TDMSApplication tdmsApp = new TDMSApplication();
                        TDMSObject tdmsObj = null;
                        Document acDoc = Application.DocumentManager.MdiActiveDocument;
                        object obj = Application.GetSystemVariable("DWGTITLED");
                        string strDwgName = acDoc.Name;
                        string guid = strDwgName;
                        string parseGuid = null;
                        Regex reg = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                        MatchCollection mc = reg.Matches(guid);
                        foreach (Match mat in mc)
                        {
                            parseGuid += mat.Value.ToString();
                        }
                        parseGuid = parseGuid.Remove(0, 38);
                        tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);
                        attrValue = tdmsObj.Attributes[attrName].Value;
                        if (attrValue != "")
                        {
                            Creator.CreateMultilineStampAtribut(x, y, attrValue, height, width, widthFactor, rotate, oblique, attrName, blockName);
                        }
                        else
                        {
                            Creator.CreateMultilineStampAtribut(x, y, "Text", height, width, widthFactor, rotate, oblique, attrName, blockName);
                        }
                    }
                    else
                    {
                        Creator.CreateMultilineStampAtribut(x, y, "Text", height, width, widthFactor, rotate, oblique, attrName, blockName);
                    }
                }
                else
                {
                    Creator.CreateMultilineStampAtribut(x, y, "Text", height, width, widthFactor, rotate, oblique, attrName, blockName);
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        //Обновление атрибутов в уже вставленных блоках. Атрибуты берутся из ТДМС.
        public void RefreshAttribute(string attrName)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor editor = doc.Editor;
            string blockName = "ATTRBLK";

            AttributeDefinition adAttr = new AttributeDefinition();
            string attrValue = null;
            TDMSApplication tdmsApp = new TDMSApplication();
            TDMSObject tdmsObj = null;
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            try
            {
                object obj = Application.GetSystemVariable("DWGTITLED");
                string strDwgName = acDoc.Name;
                string guid = strDwgName;
                string parseGuid = null;
                Regex reg = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                MatchCollection mc = reg.Matches(guid);
                foreach (Match mat in mc)
                {
                    parseGuid += mat.Value.ToString();
                }
                parseGuid = parseGuid.Remove(0, 38);
                tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);
                attrValue = tdmsObj.Attributes[attrName].Value;
                if (attrValue != "")
                {
                    Commands refresh = new Commands();
                    refresh.UpdateAttributesInDatabase(db, blockName, attrName, attrValue);
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }

        //Обновление атрибута - функции без параметров
        public void RefreshAttribute(string moduleName, string functionName)
        {
            string attrName = moduleName + "_" + functionName + "#FUNCTION";
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor editor = doc.Editor;
            string blockName = "ATTRBLK";
            AttributeDefinition adAttr = new AttributeDefinition();
            string attrValue = null;
            TDMSApplication tdmsApp = new TDMSApplication();
            TDMSObject tdmsObj = null;
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            try
            {
                object obj = Application.GetSystemVariable("DWGTITLED");
                string strDwgName = acDoc.Name;
                string guid = strDwgName;
                string parseGuid = null;
                Regex reg = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                MatchCollection mc = reg.Matches(guid);
                foreach (Match mat in mc)
                {
                    parseGuid += mat.Value.ToString();
                }
                parseGuid = parseGuid.Remove(0, 38);
                tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);

                attrValue = tdmsApp.ExecuteScript(moduleName, functionName, tdmsObj);

                if (attrValue != "")
                {
                    Commands refresh = new Commands();
                    refresh.UpdateAttributesInDatabase(db, blockName, attrName, attrValue);
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }

        //Обновление атрибута - функции с 1 параметром
        public void RefreshAttribute(string moduleName, string functionName, string parameter)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor editor = doc.Editor;

            string blockName = "ATTRBLK";
            AttributeDefinition adAttr = new AttributeDefinition();
            string attrValue = null;
            TDMSApplication tdmsApp = new TDMSApplication();
            TDMSObject tdmsObj = null;
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            try
            {
                object obj = Application.GetSystemVariable("DWGTITLED");
                string attrName = moduleName + "_" + functionName + "|" + parameter + "#FUNCTION";
                string strDwgName = acDoc.Name;
                string guid = strDwgName;
                string parseGuid = null;
                Regex reg = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                MatchCollection mc = reg.Matches(guid);
                foreach (Match mat in mc)
                {
                    parseGuid += mat.Value.ToString();
                }
                parseGuid = parseGuid.Remove(0, 38);
                tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);

                if (parameter != "GKAB_")
                {
                    attrValue = tdmsApp.ExecuteScript(moduleName, functionName, tdmsObj, parameter);
                }
                else
                {
                    attrValue = tdmsApp.ExecuteScript(moduleName, functionName, parameter);
                }

                if (attrValue != "")
                {
                    Commands refresh = new Commands();
                    refresh.UpdateAttributesInDatabase(db, blockName, attrName, attrValue);
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }

        //Обновление атрибута - функции с 2 параметрами
        public void RefreshAttribute(string moduleName, string functionName, string parameter, string parameter2)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                string attrName = moduleName + "_" + functionName + "|" + parameter + "|" + parameter2 + "#FUNCTION";

                editor.WriteMessage("\n имя атрибута: " + attrName);

                Document doc = Application.DocumentManager.MdiActiveDocument;

                Database db = doc.Database;

                string blockName = "ATTRBLK";

                AttributeDefinition adAttr = new AttributeDefinition();
                string attrValue = null;
                TDMSApplication tdmsApp = new TDMSApplication();
                TDMSObject tdmsObj = null;
                Document acDoc = Application.DocumentManager.MdiActiveDocument;
                object obj = Application.GetSystemVariable("DWGTITLED");
                string strDwgName = acDoc.Name;
                string guid = strDwgName;
                string parseGuid = null;

                Regex reg = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                MatchCollection mc = reg.Matches(guid);
                foreach (Match mat in mc)
                {
                    parseGuid += mat.Value.ToString();
                }
                parseGuid = parseGuid.Remove(0, 38);
                tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);

                attrValue = tdmsApp.ExecuteScript(moduleName, functionName, tdmsObj, parameter, parameter2);
                editor.WriteMessage("\n Значение: " + attrValue);
                if (attrValue != "")
                {
                    Commands refresh = new Commands();
                    refresh.UpdateAttributesInDatabase(db, blockName, attrName, attrValue);
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
       
        //Добавляем логотип
        public void AddLogo(double x, double y)
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Editor editor = acDoc.Editor;
            try
            {
                Database acCurDb = acDoc.Database;
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.PaperSpace], OpenMode.ForWrite) as BlockTableRecord;

                    int roomNumPoint = 0;
                    //////////////////////////////////////////////////////////////////////////////////////////
                    Polyline Poly = new Polyline();
                    Poly.SetDatabaseDefaults();
                    Poly.AddVertexAt(roomNumPoint, new Point2d(x - 48.258, y + 6.215), 0, 0, 0);
                    Poly.AddVertexAt(++roomNumPoint, new Point2d(x - 44.240, y + 11.958), 0, 0, 0);
                    Poly.AddVertexAt(++roomNumPoint, new Point2d(x - 48.258, y + 11.958), 0, 0, 0);
                    Poly.Closed = true;
                    Poly.LineWeight = 0;
                    Poly.ConstantWidth = 0;
                    Poly.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly);
                    acTrans.AddNewlyCreatedDBObject(Poly, true);

                    ObjectIdCollection acObjIdColl = new ObjectIdCollection();
                    acObjIdColl.Add(Poly.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    roomNumPoint = 0;
                    Polyline Poly1 = new Polyline();
                    Poly1.SetDatabaseDefaults();
                    Poly1.AddVertexAt(roomNumPoint, new Point2d(x - 45.346, y + 9.432), 0, 0, 0);
                    Poly1.AddVertexAt(++roomNumPoint, new Point2d(x - 47.597, y + 6.215), 0, 0, 0);
                    Poly1.AddVertexAt(++roomNumPoint, new Point2d(x - 45.346, y + 6.215), 0, 0, 0);
                    Poly1.Closed = true;
                    Poly1.LineWeight = 0;
                    Poly1.ConstantWidth = 0;
                    Poly1.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly1);
                    acTrans.AddNewlyCreatedDBObject(Poly1, true);

                    ObjectIdCollection acObjIdColl1 = new ObjectIdCollection();
                    acObjIdColl1.Add(Poly1.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    roomNumPoint = 0;
                    Polyline Poly2 = new Polyline();
                    Poly2.SetDatabaseDefaults();
                    Poly2.AddVertexAt(roomNumPoint, new Point2d(x - 45.346, y + 2.989), 0, 0, 0);
                    Poly2.AddVertexAt(++roomNumPoint, new Point2d(x - 45.346, y + 4.673), 0, 0, 0);
                    Poly2.AddVertexAt(++roomNumPoint, new Point2d(x - 48.258, y + 4.673), 0, 0, 0);
                    Poly2.AddVertexAt(++roomNumPoint, new Point2d(x - 48.258, y + 2.989), 0, 0, 0);
                    Poly2.Closed = true;
                    Poly2.LineWeight = 0;
                    Poly2.ConstantWidth = 0;
                    Poly2.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly2);
                    acTrans.AddNewlyCreatedDBObject(Poly2, true);

                    ObjectIdCollection acObjIdColl2 = new ObjectIdCollection();
                    acObjIdColl2.Add(Poly2.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    roomNumPoint = 0;
                    Polyline Poly3 = new Polyline();
                    Poly3.SetDatabaseDefaults();
                    Poly3.AddVertexAt(roomNumPoint, new Point2d(x - 39.360, y + 2.989), 0, 0, 0);
                    Poly3.AddVertexAt(++roomNumPoint, new Point2d(x - 39.360, y + 4.673), 0, 0, 0);
                    Poly3.AddVertexAt(++roomNumPoint, new Point2d(x - 42.272, y + 4.673), 0, 0, 0);
                    Poly3.AddVertexAt(++roomNumPoint, new Point2d(x - 42.272, y + 6.215), 0, 0, 0);
                    Poly3.AddVertexAt(++roomNumPoint, new Point2d(x - 38.253, y + 11.958), 0, 0, 0);
                    Poly3.AddVertexAt(++roomNumPoint, new Point2d(x - 43.347, y + 11.958), 0, 0, 0);
                    Poly3.AddVertexAt(++roomNumPoint, new Point2d(x - 43.347, y + 6.215), 0, 0, 0);
                    Poly3.AddVertexAt(++roomNumPoint, new Point2d(x - 42.608, y + 6.215), 0, 0, 0);
                    Poly3.AddVertexAt(++roomNumPoint, new Point2d(x - 42.608, y + 4.673), 0, 0, 0);
                    Poly3.AddVertexAt(++roomNumPoint, new Point2d(x - 43.347, y + 4.673), 0, 0, 0);
                    Poly3.AddVertexAt(++roomNumPoint, new Point2d(x - 43.347, y + 2.989), 0, 0, 0);
                    Poly3.Closed = true;
                    Poly3.LineWeight = 0;
                    Poly3.ConstantWidth = 0;
                    Poly3.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly3);
                    acTrans.AddNewlyCreatedDBObject(Poly3, true);

                    ObjectIdCollection acObjIdColl3 = new ObjectIdCollection();
                    acObjIdColl3.Add(Poly3.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    roomNumPoint = 0;
                    Polyline Poly4 = new Polyline();
                    Poly4.SetDatabaseDefaults();
                    Poly4.AddVertexAt(roomNumPoint, new Point2d(x - 39.360, y + 6.215), 0, 0, 0);
                    Poly4.AddVertexAt(++roomNumPoint, new Point2d(x - 39.360, y + 9.432), 0, 0, 0);
                    Poly4.AddVertexAt(++roomNumPoint, new Point2d(x - 41.610, y + 6.215), 0, 0, 0);
                    Poly4.Closed = true;
                    Poly4.LineWeight = 0;
                    Poly4.ConstantWidth = 0;
                    Poly4.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly4);
                    acTrans.AddNewlyCreatedDBObject(Poly4, true);

                    ObjectIdCollection acObjIdColl4 = new ObjectIdCollection();
                    acObjIdColl4.Add(Poly4.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    roomNumPoint = 0;
                    Polyline Poly5 = new Polyline();
                    Poly5.SetDatabaseDefaults();
                    Poly5.AddVertexAt(roomNumPoint, new Point2d(x - 36.674, y + 11.958), 0, 0, 0);
                    Poly5.AddVertexAt(++roomNumPoint, new Point2d(x - 37.361, y + 11.958), 0, 0, 0);
                    Poly5.AddVertexAt(++roomNumPoint, new Point2d(x - 37.361, y + 6.215), 0, 0, 0);
                    Poly5.AddVertexAt(++roomNumPoint, new Point2d(x - 36.674, y + 6.215), 0, 0, 0);
                    Poly5.Closed = true;
                    Poly5.LineWeight = 0;
                    Poly5.ConstantWidth = 0;
                    Poly5.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly5);
                    acTrans.AddNewlyCreatedDBObject(Poly5, true);

                    ObjectIdCollection acObjIdColl5 = new ObjectIdCollection();
                    acObjIdColl5.Add(Poly5.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    roomNumPoint = 0;
                    Polyline Poly6 = new Polyline();
                    Poly6.SetDatabaseDefaults();
                    Poly6.AddVertexAt(roomNumPoint, new Point2d(x - 36.674, y + 4.673), 0, 0, 0);
                    Poly6.AddVertexAt(++roomNumPoint, new Point2d(x - 37.361, y + 4.673), 0, 0, 0);
                    Poly6.AddVertexAt(++roomNumPoint, new Point2d(x - 37.361, y + 2.989), 0, 0, 0);
                    Poly6.AddVertexAt(++roomNumPoint, new Point2d(x - 36.674, y + 2.989), 0, 0, 0);
                    Poly6.Closed = true;
                    Poly6.LineWeight = 0;
                    Poly6.ConstantWidth = 0;
                    Poly6.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly6);
                    acTrans.AddNewlyCreatedDBObject(Poly6, true);

                    ObjectIdCollection acObjIdColl6 = new ObjectIdCollection();
                    acObjIdColl6.Add(Poly6.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    // C
                    roomNumPoint = 0;
                    Polyline Poly7 = new Polyline();
                    Poly7.SetDatabaseDefaults();
                    Poly7.AddVertexAt(roomNumPoint, new Point2d(x - 47.080, y + 13.763), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.031, y + 13.963), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.146, y + 14.006), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.230, y + 14.029), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.324, y + 14.044), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.423, y + 14.054), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.552, y + 14.055), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.693, y + 14.036), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.812, y + 13.999), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.939, y + 13.938), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 48.065, y + 13.845), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 48.160, y + 13.728), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 48.217, y + 13.614), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 48.253, y + 13.486), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 48.262, y + 13.424), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 48.262, y + 13.244), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 48.238, y + 13.117), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 48.202, y + 13.037), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 48.161, y + 12.972), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 48.097, y + 12.900), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 48.016, y + 12.826), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.899, y + 12.757), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.765, y + 12.707), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.665, y + 12.684), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.589, y + 12.676), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.394, y + 12.675), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.267, y + 12.693), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.165, y + 12.720), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.102, y + 12.750), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.082, y + 12.771), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.056, y + 12.863), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.045, y + 12.898), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.047, y + 12.921), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.067, y + 12.933), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.123, y + 12.894), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.190, y + 12.857), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.283, y + 12.817), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.361, y + 12.799), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.517, y + 12.799), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.571, y + 12.808), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.693, y + 12.856), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.727, y + 12.879), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.814, y + 12.964), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.884, y + 13.095), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.903, y + 13.163), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.920, y + 13.260), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.920, y + 13.447), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.906, y + 13.555), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.869, y + 13.642), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.825, y + 13.732), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.796, y + 13.778), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.709, y + 13.857), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.606, y + 13.912), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.520, y + 13.930), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.383, y + 13.930), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.286, y + 13.902), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.231, y + 13.871), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.171, y + 13.827), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.126, y + 13.783), 0, 0, 0);
                    Poly7.AddVertexAt(++roomNumPoint, new Point2d(x - 47.108, y + 13.761), 0, 0, 0);
                    Poly7.Closed = true;
                    Poly7.LineWeight = 0;
                    Poly7.ConstantWidth = 0;
                    Poly7.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly7);
                    acTrans.AddNewlyCreatedDBObject(Poly7, true);

                    ObjectIdCollection acObjIdColl7 = new ObjectIdCollection();
                    acObjIdColl7.Add(Poly7.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    // T
                    roomNumPoint = 0;
                    Polyline Poly8 = new Polyline();
                    Poly8.SetDatabaseDefaults();
                    Poly8.AddVertexAt(roomNumPoint, new Point2d(x - 45.370, y + 13.885), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.370, y + 13.874), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.249, y + 13.874), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.249, y + 13.864), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.171, y + 13.864), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.171, y + 13.880), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.182, y + 13.880), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.182, y + 13.990), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.170, y + 13.994), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.170, y + 14.017), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 46.227, y + 14.017), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 46.227, y + 13.990), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 46.221, y + 13.990), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 46.221, y + 13.882), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 46.231, y + 13.882), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 46.231, y + 13.855), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 46.192, y + 13.855), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 46.192, y + 13.865), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 46.113, y + 13.865), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 46.113, y + 13.875), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 46.012, y + 13.875), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 46.012, y + 13.885), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.846, y + 13.885), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.846, y + 12.694), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.766, y + 12.694), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.766, y + 12.706), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.628, y + 12.706), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.628, y + 12.694), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.537, y + 12.694), 0, 0, 0);
                    Poly8.AddVertexAt(++roomNumPoint, new Point2d(x - 45.537, y + 13.885), 0, 0, 0);
                    Poly8.Closed = true;
                    Poly8.LineWeight = 0;
                    Poly8.ConstantWidth = 0;
                    Poly8.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly8);
                    acTrans.AddNewlyCreatedDBObject(Poly8, true);

                    ObjectIdCollection acObjIdColl8 = new ObjectIdCollection();
                    acObjIdColl8.Add(Poly8.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    // У
                    roomNumPoint = 0;
                    Polyline Poly9 = new Polyline();
                    Poly9.SetDatabaseDefaults();
                    Poly9.AddVertexAt(roomNumPoint, new Point2d(x - 44.346, y + 14.026), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 44.270, y + 14.026), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 44.265, y + 14.017), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 44.069, y + 14.017), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 44.064, y + 14.026), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.979, y + 14.026), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.588, y + 13.366), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.488, y + 13.547), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.367, y + 13.793), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.259, y + 14.027), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.190, y + 14.027), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.190, y + 14.017), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.125, y + 14.017), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.125, y + 14.027), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.061, y + 14.027), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.378, y + 13.454), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.513, y + 13.187), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.757, y + 12.695), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.780, y + 12.695), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.780, y + 12.707), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.893, y + 12.707), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.893, y + 12.695), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.920, y + 12.695), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.920, y + 12.707), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.740, y + 13.045), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.820, y + 13.190), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 43.909, y + 13.341), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 44.003, y + 13.485), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 44.148, y + 13.712), 0, 0, 0);
                    Poly9.AddVertexAt(++roomNumPoint, new Point2d(x - 44.346, y + 14.026), 0, 0, 0);
                    Poly9.Closed = true;
                    Poly9.LineWeight = 0;
                    Poly9.ConstantWidth = 0;
                    Poly9.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly9);
                    acTrans.AddNewlyCreatedDBObject(Poly9, true);

                    ObjectIdCollection acObjIdColl9 = new ObjectIdCollection();
                    acObjIdColl9.Add(Poly9.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    // Д и то что внутри
                    roomNumPoint = 0;
                    Polyline Poly10 = new Polyline();
                    Poly10.SetDatabaseDefaults();
                    Poly10.AddVertexAt(roomNumPoint, new Point2d(x - 42.238, y + 12.640), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 42.238, y + 12.653), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 42.227, y + 12.653), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 42.227, y + 12.821), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 42.238, y + 12.821), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 42.238, y + 12.847), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 42.150, y + 12.847), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 41.850, y + 13.518), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 41.646, y + 14.028), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 41.600, y + 14.028), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 41.600, y + 14.017), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 41.518, y + 14.017), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 41.518, y + 14.028), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 41.472, y + 14.028), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 41.320, y + 13.648), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 40.983, y + 12.847), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 40.888, y + 12.847), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 40.888, y + 12.821), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 40.899, y + 12.821), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 40.900, y + 12.653), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 40.888, y + 12.653), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 40.888, y + 12.640), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 40.921, y + 12.640), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 40.976, y + 12.668), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 41.108, y + 12.693), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 42.032, y + 12.693), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 42.126, y + 12.677), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 42.192, y + 12.655), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 42.201, y + 12.647), 0, 0, 0);
                    Poly10.AddVertexAt(++roomNumPoint, new Point2d(x - 42.201, y + 12.640), 0, 0, 0);
                    Poly10.Closed = true;
                    Poly10.LineWeight = 0;
                    Poly10.ConstantWidth = 0;
                    Poly10.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly10);
                    acTrans.AddNewlyCreatedDBObject(Poly10, true);

                    ObjectIdCollection acObjIdColl10 = new ObjectIdCollection();
                    acObjIdColl10.Add(Poly10.ObjectId);

                    roomNumPoint = 0;
                    Polyline Poly_10 = new Polyline();
                    Poly_10.SetDatabaseDefaults();
                    Poly_10.AddVertexAt(roomNumPoint, new Point2d(x - 41.645, y + 13.639), 0, 0, 0);
                    Poly_10.AddVertexAt(++roomNumPoint, new Point2d(x - 41.980, y + 12.833), 0, 0, 0);
                    Poly_10.AddVertexAt(++roomNumPoint, new Point2d(x - 41.883, y + 12.833), 0, 0, 0);
                    Poly_10.AddVertexAt(++roomNumPoint, new Point2d(x - 41.880, y + 12.819), 0, 0, 0);
                    Poly_10.AddVertexAt(++roomNumPoint, new Point2d(x - 41.529, y + 12.819), 0, 0, 0);
                    Poly_10.AddVertexAt(++roomNumPoint, new Point2d(x - 41.335, y + 12.850), 0, 0, 0);
                    Poly_10.Closed = true;
                    Poly_10.LineWeight = 0;
                    Poly_10.ConstantWidth = 0;
                    Poly_10.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly_10);
                    acTrans.AddNewlyCreatedDBObject(Poly_10, true);

                    ObjectIdCollection acObjIdColl_10 = new ObjectIdCollection();
                    acObjIdColl_10.Add(Poly_10.ObjectId);

                    //////////////////////////////////////////////////////////////////////////////////////////
                    // И
                    roomNumPoint = 0;
                    Polyline Poly11 = new Polyline();
                    Poly11.SetDatabaseDefaults();
                    Poly11.AddVertexAt(roomNumPoint, new Point2d(x - 39.934, y + 14.027), 0, 0, 0);
                    Poly11.AddVertexAt(++roomNumPoint, new Point2d(x - 39.934, y + 12.695), 0, 0, 0);
                    Poly11.AddVertexAt(++roomNumPoint, new Point2d(x - 39.786, y + 12.695), 0, 0, 0);
                    Poly11.AddVertexAt(++roomNumPoint, new Point2d(x - 39.003, y + 13.614), 0, 0, 0);
                    Poly11.AddVertexAt(++roomNumPoint, new Point2d(x - 39.003, y + 12.695), 0, 0, 0);
                    Poly11.AddVertexAt(++roomNumPoint, new Point2d(x - 38.693, y + 12.695), 0, 0, 0);
                    Poly11.AddVertexAt(++roomNumPoint, new Point2d(x - 38.698, y + 14.028), 0, 0, 0);
                    Poly11.AddVertexAt(++roomNumPoint, new Point2d(x - 38.854, y + 14.028), 0, 0, 0);
                    Poly11.AddVertexAt(++roomNumPoint, new Point2d(x - 39.242, y + 13.550), 0, 0, 0);
                    Poly11.AddVertexAt(++roomNumPoint, new Point2d(x - 39.626, y + 13.115), 0, 0, 0);
                    Poly11.AddVertexAt(++roomNumPoint, new Point2d(x - 39.624, y + 14.028), 0, 0, 0);
                    Poly11.Closed = true;
                    Poly11.LineWeight = 0;
                    Poly11.ConstantWidth = 0;
                    Poly11.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly11);
                    acTrans.AddNewlyCreatedDBObject(Poly11, true);

                    ObjectIdCollection acObjIdColl11 = new ObjectIdCollection();
                    acObjIdColl11.Add(Poly11.ObjectId);

                    //////////////////////////////////////////////////////////////////////////////////////////
                    // Я и пустота внутри
                    roomNumPoint = 0;
                    Polyline Poly12 = new Polyline();
                    Poly12.SetDatabaseDefaults();

                    Poly12.AddVertexAt(roomNumPoint, new Point2d(x - 37.290, y + 14.016), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 36.674, y + 14.018), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 36.674, y + 12.692), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 36.987, y + 12.692), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 36.987, y + 13.319), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.374, y + 12.690), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.742, y + 12.690), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.302, y + 13.351), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.367, y + 13.358), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.438, y + 13.382), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.496, y + 13.417), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.539, y + 13.457), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.573, y + 13.499), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.601, y + 13.550), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.621, y + 13.615), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.628, y + 13.683), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.614, y + 13.780), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.569, y + 13.873), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.515, y + 13.933), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.447, y + 13.980), 0, 0, 0);
                    Poly12.AddVertexAt(++roomNumPoint, new Point2d(x - 37.375, y + 14.006), 0, 0, 0);
                    Poly12.Closed = true;
                    Poly12.LineWeight = 0;
                    Poly12.ConstantWidth = 0;
                    Poly12.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly12);
                    acTrans.AddNewlyCreatedDBObject(Poly12, true);

                    ObjectIdCollection acObjIdColl12 = new ObjectIdCollection();
                    acObjIdColl12.Add(Poly12.ObjectId);


                    roomNumPoint = 0;
                    Polyline Poly13 = new Polyline();
                    Poly13.SetDatabaseDefaults();
                    Poly13.AddVertexAt(roomNumPoint, new Point2d(x - 36.986, y + 13.918), 0, 0, 0);
                    Poly13.AddVertexAt(++roomNumPoint, new Point2d(x - 36.986, y + 13.385), 0, 0, 0);
                    Poly13.AddVertexAt(++roomNumPoint, new Point2d(x - 37.105, y + 13.385), 0, 0, 0);
                    Poly13.AddVertexAt(++roomNumPoint, new Point2d(x - 37.174, y + 13.406), 0, 0, 0);
                    Poly13.AddVertexAt(++roomNumPoint, new Point2d(x - 37.229, y + 13.446), 0, 0, 0);
                    Poly13.AddVertexAt(++roomNumPoint, new Point2d(x - 37.268, y + 13.493), 0, 0, 0);
                    Poly13.AddVertexAt(++roomNumPoint, new Point2d(x - 37.301, y + 13.583), 0, 0, 0);
                    Poly13.AddVertexAt(++roomNumPoint, new Point2d(x - 37.307, y + 13.655), 0, 0, 0);
                    Poly13.AddVertexAt(++roomNumPoint, new Point2d(x - 37.308, y + 13.709), 0, 0, 0);
                    Poly13.AddVertexAt(++roomNumPoint, new Point2d(x - 37.293, y + 13.796), 0, 0, 0);
                    Poly13.AddVertexAt(++roomNumPoint, new Point2d(x - 37.275, y + 13.839), 0, 0, 0);
                    Poly13.AddVertexAt(++roomNumPoint, new Point2d(x - 37.253, y + 13.872), 0, 0, 0);
                    Poly13.AddVertexAt(++roomNumPoint, new Point2d(x - 37.221, y + 13.891), 0, 0, 0);
                    Poly13.AddVertexAt(++roomNumPoint, new Point2d(x - 37.185, y + 13.909), 0, 0, 0);
                    Poly13.AddVertexAt(++roomNumPoint, new Point2d(x - 37.130, y + 13.918), 0, 0, 0);
                    Poly13.Closed = true;
                    Poly13.LineWeight = 0;
                    Poly13.ConstantWidth = 0;
                    Poly13.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly13);
                    acTrans.AddNewlyCreatedDBObject(Poly13, true);

                    ObjectIdCollection acObjIdColl13 = new ObjectIdCollection();
                    acObjIdColl13.Add(Poly13.ObjectId);

                    //////////////////////////////////////////////////////////////////////////////////////////
                    // S
                    roomNumPoint = 0;
                    Polyline Poly14 = new Polyline();
                    Poly14.SetDatabaseDefaults();
                    Poly14.AddVertexAt(roomNumPoint, new Point2d(x - 47.446, y + 2.263), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.512, y + 2.060), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.549, y + 2.060), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.555, y + 2.093), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.582, y + 2.136), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.641, y + 2.194), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.700, y + 2.224), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.736, y + 2.234), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.846, y + 2.234), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.915, y + 2.213), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.968, y + 2.180), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.998, y + 2.131), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.012, y + 2.086), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.012, y + 2.012), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.985, y + 1.956), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.932, y + 1.905), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.890, y + 1.875), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.829, y + 1.845), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.699, y + 1.776), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.546, y + 1.695), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.508, y + 1.661), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.449, y + 1.603), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.425, y + 1.559), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.397, y + 1.484), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.388, y + 1.445), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.388, y + 1.325), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.405, y + 1.252), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.442, y + 1.180), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.459, y + 1.151), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.536, y + 1.076), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.596, y + 1.037), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.647, y + 1.009), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.717, y + 0.985), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.783, y + 0.969), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.846, y + 0.959), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.022, y + 0.959), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.083, y + 0.967), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.157, y + 0.986), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.211, y + 1.004), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.248, y + 1.023), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.270, y + 1.034), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.207, y + 1.286), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.171, y + 1.286), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.139, y + 1.206), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.122, y + 1.173), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.087, y + 1.138), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.047, y + 1.109), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.007, y + 1.090), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.942, y + 1.070), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.824, y + 1.070), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.779, y + 1.082), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.737, y + 1.100), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.704, y + 1.126), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.670, y + 1.163), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.655, y + 1.190), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.645, y + 1.222), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.645, y + 1.308), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.656, y + 1.343), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.681, y + 1.385), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.733, y + 1.431), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.810, y + 1.477), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.924, y + 1.535), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.108, y + 1.626), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.131, y + 1.642), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.178, y + 1.689), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.224, y + 1.763), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.253, y + 1.832), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.261, y + 1.875), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.263, y + 1.994), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.252, y + 2.047), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.225, y + 2.116), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.167, y + 2.204), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 48.055, y + 2.292), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.923, y + 2.334), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.853, y + 2.345), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.700, y + 2.345), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.597, y + 2.327), 0, 0, 0);
                    Poly14.AddVertexAt(++roomNumPoint, new Point2d(x - 47.515, y + 2.300), 0, 0, 0);
                    Poly14.Closed = true;
                    Poly14.LineWeight = 0;
                    Poly14.ConstantWidth = 0;
                    Poly14.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly14);
                    acTrans.AddNewlyCreatedDBObject(Poly14, true);

                    ObjectIdCollection acObjIdColl14 = new ObjectIdCollection();
                    acObjIdColl14.Add(Poly14.ObjectId);

                    //////////////////////////////////////////////////////////////////////////////////////////
                    // T
                    roomNumPoint = 0;
                    Polyline Poly15 = new Polyline();
                    Poly15.SetDatabaseDefaults();
                    Poly15.AddVertexAt(roomNumPoint, new Point2d(x - 45.692, y + 2.173), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 45.692, y + 2.162), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 45.572, y + 2.162), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 45.572, y + 2.153), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 45.493, y + 2.153), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 45.493, y + 2.168), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 45.504, y + 2.168), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 45.504, y + 2.279), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 45.493, y + 2.283), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 45.493, y + 2.306), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.550, y + 2.306), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.550, y + 2.279), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.544, y + 2.279), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.544, y + 2.170), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.553, y + 2.170), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.553, y + 2.144), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.515, y + 2.144), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.515, y + 2.154), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.436, y + 2.154), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.436, y + 2.164), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.334, y + 2.164), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.334, y + 2.173), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.168, y + 2.173), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.168, y + 0.983), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.088, y + 0.983), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 46.088, y + 0.995), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 45.950, y + 0.995), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 45.950, y + 0.983), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 45.860, y + 0.983), 0, 0, 0);
                    Poly15.AddVertexAt(++roomNumPoint, new Point2d(x - 45.860, y + 2.173), 0, 0, 0);
                    Poly15.Closed = true;
                    Poly15.LineWeight = 0;
                    Poly15.ConstantWidth = 0;
                    Poly15.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly15);
                    acTrans.AddNewlyCreatedDBObject(Poly15, true);

                    ObjectIdCollection acObjIdColl15 = new ObjectIdCollection();
                    acObjIdColl15.Add(Poly15.ObjectId);

                    //////////////////////////////////////////////////////////////////////////////////////////
                    // U
                    roomNumPoint = 0;
                    Polyline Poly16 = new Polyline();
                    Poly16.SetDatabaseDefaults();
                    Poly16.AddVertexAt(roomNumPoint, new Point2d(x - 44.466, y + 2.307), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 44.158, y + 2.307), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 44.158, y + 1.464), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 44.126, y + 1.310), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 44.094, y + 1.238), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 44.048, y + 1.180), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.971, y + 1.125), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.858, y + 1.087), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.713, y + 1.087), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.643, y + 1.099), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.555, y + 1.134), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.475, y + 1.216), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.436, y + 1.278), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.409, y + 1.345), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.390, y + 1.440), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.383, y + 1.519), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.383, y + 2.307), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.216, y + 2.307), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.216, y + 1.470), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.236, y + 1.336), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.273, y + 1.233), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.320, y + 1.162), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.384, y + 1.094), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.469, y + 1.037), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.522, y + 1.010), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.595, y + 0.986), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.680, y + 0.969), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.741, y + 0.959), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 43.936, y + 0.959), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 44.101, y + 0.987), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 44.200, y + 1.022), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 44.281, y + 1.067), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 44.348, y + 1.126), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 44.386, y + 1.175), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 44.431, y + 1.262), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 44.451, y + 1.330), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 44.458, y + 1.366), 0, 0, 0);
                    Poly16.AddVertexAt(++roomNumPoint, new Point2d(x - 44.466, y + 1.435), 0, 0, 0);
                    Poly16.Closed = true;
                    Poly16.LineWeight = 0;
                    Poly16.ConstantWidth = 0;
                    Poly16.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly16);
                    acTrans.AddNewlyCreatedDBObject(Poly16, true);

                    ObjectIdCollection acObjIdColl16 = new ObjectIdCollection();
                    acObjIdColl16.Add(Poly16.ObjectId);


                    //////////////////////////////////////////////////////////////////////////////////////////
                    // D
                    roomNumPoint = 0;
                    Polyline Poly17 = new Polyline();
                    Poly17.SetDatabaseDefaults();
                    Poly17.AddVertexAt(roomNumPoint, new Point2d(x - 42.052, y + 2.298), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 42.052, y + 0.983), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 41.372, y + 0.983), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 41.203, y + 1.017), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 41.066, y + 1.073), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 40.957, y + 1.145), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 40.887, y + 1.207), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 40.806, y + 1.316), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 40.755, y + 1.421), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 40.723, y + 1.538), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 40.715, y + 1.603), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 40.715, y + 1.766), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 40.743, y + 1.915), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 40.797, y + 2.033), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 40.834, y + 2.088), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 40.944, y + 2.188), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 41.028, y + 2.246), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 41.120, y + 2.282), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 41.192, y + 2.299), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 41.256, y + 2.308), 0, 0, 0);
                    Poly17.AddVertexAt(++roomNumPoint, new Point2d(x - 42.052, y + 2.308), 0, 0, 0);
                    Poly17.Closed = true;
                    Poly17.LineWeight = 0;
                    Poly17.ConstantWidth = 0;
                    Poly17.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly17);
                    acTrans.AddNewlyCreatedDBObject(Poly17, true);

                    ObjectIdCollection acObjIdColl17 = new ObjectIdCollection();
                    acObjIdColl17.Add(Poly17.ObjectId);

                    roomNumPoint = 0;
                    Polyline Poly18 = new Polyline();
                    Poly18.SetDatabaseDefaults();
                    Poly18.AddVertexAt(roomNumPoint, new Point2d(x - 41.749, y + 2.204), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.749, y + 1.089), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.479, y + 1.089), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.440, y + 1.096), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.365, y + 1.118), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.273, y + 1.162), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.225, y + 1.196), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.163, y + 1.267), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.127, y + 1.334), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.095, y + 1.410), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.075, y + 1.476), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.065, y + 1.576), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.060, y + 1.677), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.065, y + 1.780), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.085, y + 1.898), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.110, y + 1.963), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.149, y + 2.035), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.235, y + 2.119), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.290, y + 2.151), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.356, y + 2.178), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.440, y + 2.196), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.454, y + 2.199), 0, 0, 0);
                    Poly18.AddVertexAt(++roomNumPoint, new Point2d(x - 41.495, y + 2.204), 0, 0, 0);
                    Poly18.Closed = true;
                    Poly18.LineWeight = 0;
                    Poly18.ConstantWidth = 0;
                    Poly18.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly18);
                    acTrans.AddNewlyCreatedDBObject(Poly18, true);

                    ObjectIdCollection acObjIdColl18 = new ObjectIdCollection();
                    acObjIdColl18.Add(Poly18.ObjectId);

                    /////////////////////////////////////////////////////////////////////////////////////////
                    // I
                    roomNumPoint = 0;
                    Polyline Poly19 = new Polyline();
                    Poly19.SetDatabaseDefaults();
                    Poly19.AddVertexAt(roomNumPoint, new Point2d(x - 39.625, y + 2.298), 0, 0, 0);
                    Poly19.AddVertexAt(++roomNumPoint, new Point2d(x - 39.315, y + 2.298), 0, 0, 0);
                    Poly19.AddVertexAt(++roomNumPoint, new Point2d(x - 39.315, y + 0.983), 0, 0, 0);
                    Poly19.AddVertexAt(++roomNumPoint, new Point2d(x - 39.625, y + 0.983), 0, 0, 0);
                    Poly19.Closed = true;
                    Poly19.LineWeight = 0;
                    Poly19.ConstantWidth = 0;
                    Poly19.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly19);
                    acTrans.AddNewlyCreatedDBObject(Poly19, true);

                    ObjectIdCollection acObjIdColl19 = new ObjectIdCollection();
                    acObjIdColl19.Add(Poly19.ObjectId);

                    /////////////////////////////////////////////////////////////////////////////////////////
                    // O
                    roomNumPoint = 0;
                    Polyline Poly20 = new Polyline();
                    Poly20.SetDatabaseDefaults();
                    Poly20.AddVertexAt(roomNumPoint, new Point2d(x - 37.550, y + 2.346), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.340, y + 2.346), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.174, y + 2.321), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.065, y + 2.283), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 36.973, y + 2.236), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 36.919, y + 2.199), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 36.808, y + 2.086), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 36.759, y + 2.021), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 36.709, y + 1.909), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 36.680, y + 1.802), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 36.674, y + 1.738), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 36.674, y + 1.572), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 36.680, y + 1.516), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 36.718, y + 1.395), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 36.767, y + 1.299), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 36.818, y + 1.235), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 36.902, y + 1.148), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.002, y + 1.075), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.095, y + 1.030), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.167, y + 1.003), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.267, y + 0.975), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.386, y + 0.957), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.598, y + 0.957), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.772, y + 0.986), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.877, y + 1.023), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 38.002, y + 1.088), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 38.065, y + 1.143), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 38.120, y + 1.198), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 38.195, y + 1.305), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 38.230, y + 1.389), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 38.256, y + 1.490), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 38.266, y + 1.554), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 38.266, y + 1.730), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 38.249, y + 1.835), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 38.202, y + 1.958), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 38.119, y + 2.077), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 38.045, y + 2.151), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.943, y + 2.225), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.835, y + 2.281), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.711, y + 2.319), 0, 0, 0);
                    Poly20.AddVertexAt(++roomNumPoint, new Point2d(x - 37.621, y + 2.336), 0, 0, 0);
                    Poly20.Closed = true;
                    Poly20.LineWeight = 0;
                    Poly20.ConstantWidth = 0;
                    Poly20.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly20);
                    acTrans.AddNewlyCreatedDBObject(Poly20, true);

                    ObjectIdCollection acObjIdColl20 = new ObjectIdCollection();
                    acObjIdColl20.Add(Poly20.ObjectId);

                    roomNumPoint = 0;
                    Polyline Poly21 = new Polyline();
                    Poly21.SetDatabaseDefaults();
                    Poly21.AddVertexAt(roomNumPoint, new Point2d(x - 37.524, y + 2.230), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.391, y + 2.230), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.306, y + 2.208), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.221, y + 2.169), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.127, y + 2.072), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.061, y + 1.973), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.035, y + 1.864), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.020, y + 1.774), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.020, y + 1.591), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.041, y + 1.463), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.077, y + 1.357), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.120, y + 1.268), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.177, y + 1.187), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.261, y + 1.123), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.324, y + 1.089), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.417, y + 1.068), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.556, y + 1.068), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.652, y + 1.093), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.733, y + 1.140), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.808, y + 1.215), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.865, y + 1.307), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.898, y + 1.392), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.917, y + 1.505), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.917, y + 1.692), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.906, y + 1.799), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.889, y + 1.884), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.867, y + 1.958), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.831, y + 2.032), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.801, y + 2.074), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.716, y + 2.160), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.640, y + 2.201), 0, 0, 0);
                    Poly21.AddVertexAt(++roomNumPoint, new Point2d(x - 37.567, y + 2.227), 0, 0, 0);
                    Poly21.Closed = true;
                    Poly21.LineWeight = 0;
                    Poly21.ConstantWidth = 0;
                    Poly21.Layer = "Defpoints";
                    acBlkTblRec.AppendEntity(Poly21);
                    acTrans.AddNewlyCreatedDBObject(Poly21, true);

                    ObjectIdCollection acObjIdColl21 = new ObjectIdCollection();
                    acObjIdColl21.Add(Poly21.ObjectId);
                    //////////////////////////////////////////////////////////////////////////////////////////
                    Hatch acHatch = new Hatch();
                    acBlkTblRec.AppendEntity(acHatch);

                    acTrans.AddNewlyCreatedDBObject(acHatch, true);

                    acHatch.SetDatabaseDefaults();
                    acHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                    acHatch.Associative = true;

                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl1);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl2);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl3);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl4);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl5);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl6);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl7);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl8);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl9);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl10);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl_10);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl11);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl12);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl13);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl14);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl15);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl16);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl17);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl18);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl19);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl20);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl21);
                    acHatch.EvaluateHatch(true);
                    acHatch.Layer = "Z-STMP";
                    acTrans.Commit();
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        //Создание вкладки TDMS на ленте
        [CommandMethod("TDMSRIBBON", CommandFlags.Transparent)]
        public void TDMSRIBBON()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            RibbonControl ribbControl = ComponentManager.Ribbon;
            try
            {
                if (ribbControl != null)
                {
                    RibbonTab ribonTab = ribbControl.FindTab("TDMS");
                    if (ribonTab != null)
                    {
                        ribbControl.Tabs.Remove(ribonTab);
                    }
                    ribonTab = new RibbonTab();
                    ribonTab.Title = "TDMS  ";
                    ribonTab.Id = "TDMS";
                    //Добавляем вкладки
                    ribbControl.Tabs.Add(ribonTab);
                    //Добавляем контент вкладок
                    addContent(ribonTab);
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        // Собственный обработчик команд
        public sealed class RibbonCommandHandler : System.Windows.Input.ICommand
        {
            public bool CanExecute(object parameter)
            {
                return true;
            }
            public event EventHandler CanExecuteChanged;
            public void Execute(object parameter)
            {
                var editor = Application.DocumentManager.MdiActiveDocument.Editor;
                try
                {
                    if (parameter is RibbonButton)
                    {
                        RibbonButton button = parameter as RibbonButton;
                        acadApp.DocumentManager.MdiActiveDocument.SendStringToExecute(button.CommandParameter + " ", true, false, true);
                    }
                }
                catch (System.Exception ex)
                {
                    editor.WriteMessage("/n Exception caught" + ex);
                }
            }
        }
        //Добавляем контент в наши панели на вкладке
        static void addContent(RibbonTab ribbTab)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                ribbTab.Panels.Add(Frame());
                ribbTab.Panels.Add(FrameArch());
                ribbTab.Panels.Add(Stamp());
                ribbTab.Panels.Add(Layer());
                ribbTab.Panels.Add(RibbonSave());
                ribbTab.Panels.Add(Optimization());
                ribbTab.Panels.Add(Etransmit());
                ribbTab.Panels.Add(UpdateAttribute());
                ribbTab.Panels.Add(ribbonXREF());
                ribbTab.Panels.Add(Helps());

                RibbonSplitButton ribbSplitBtn = new RibbonSplitButton();
                /* Для RibbonSplitButton ОБЯЗАТЕЛЬНО надо указать
                 * свойство Text, а иначе при поиске команд в автокаде
                 * будет вылетать ошибка.
                 */
                ribbSplitBtn.Text = "RibbonSplitButton";
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }

        static System.Windows.Media.Imaging.BitmapImage LoadImage(string ImageName)
        {
            return new System.Windows.Media.Imaging.BitmapImage(
                new Uri("pack://application:,,,/ACadRibbon;component/" + ImageName + ".png"));
        }

        static RibbonPanel Frame()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            RibbonToolTip rttA4H = new RibbonToolTip();
            RibbonToolTip rttA4V = new RibbonToolTip();

            RibbonToolTip rttA4H3 = new RibbonToolTip();
            RibbonToolTip rttA4H4 = new RibbonToolTip();

            RibbonToolTip rttA3H = new RibbonToolTip();
            RibbonToolTip rttA3V = new RibbonToolTip();

            RibbonToolTip rttA3H3 = new RibbonToolTip();

            RibbonToolTip rttA2H = new RibbonToolTip();
            RibbonToolTip rttA2V = new RibbonToolTip();

            RibbonToolTip rttA1H = new RibbonToolTip();
            RibbonToolTip rttA1V = new RibbonToolTip();

            RibbonToolTip rttA0H = new RibbonToolTip();
            RibbonToolTip rttA0V = new RibbonToolTip();


            RibbonButton rbA4H = new RibbonButton();
            RibbonButton rbA4V = new RibbonButton();

            RibbonButton rbA4H3 = new RibbonButton();
            RibbonButton rbA4H4 = new RibbonButton();

            RibbonButton rbA3H = new RibbonButton();
            RibbonButton rbA3V = new RibbonButton();

            RibbonButton rbA3H3 = new RibbonButton();

            RibbonButton rbA2H = new RibbonButton();
            RibbonButton rbA2V = new RibbonButton();

            RibbonButton rbA1H = new RibbonButton();
            RibbonButton rbA1V = new RibbonButton();

            RibbonButton rbA0H = new RibbonButton();
            RibbonButton rbA0V = new RibbonButton();

            RibbonSplitButton ribbSplitButton;
            RibbonPanelSource ribbPanelSource = new RibbonPanelSource();

            ribbPanelSource.Title = "Frame";

            RibbonPanel ribbPanel = new RibbonPanel();

            ribbPanel.Source = ribbPanelSource;

            ribbSplitButton = new RibbonSplitButton();
            ribbSplitButton.Text = "RibbonSplitButton";
            ribbSplitButton.Orientation = System.Windows.Controls.Orientation.Vertical;
            ribbSplitButton.Size = RibbonItemSize.Large;
            ribbSplitButton.Size = RibbonItemSize.Large;
            ribbSplitButton.ShowImage = true;
            ribbSplitButton.ShowText = true;
            // Стиль кнопки
            ribbSplitButton.ListButtonStyle = Autodesk.Private.Windows.RibbonListButtonStyle.SplitButton;
            ribbSplitButton.ResizeStyle = RibbonItemResizeStyles.NoResize;
            ribbSplitButton.ListStyle = RibbonSplitButtonListStyle.List;
            try
            {
                rttA4H.IsHelpEnabled = false;
                rttA4V.IsHelpEnabled = false;

                rttA4H3.IsHelpEnabled = false;
                rttA4H4.IsHelpEnabled = false;

                rttA3H.IsHelpEnabled = false;
                rttA3V.IsHelpEnabled = false;

                rttA3H3.IsHelpEnabled = false;

                rttA2H.IsHelpEnabled = false;
                rttA2V.IsHelpEnabled = false;

                rttA1H.IsHelpEnabled = false;
                rttA1V.IsHelpEnabled = false;

                rttA0H.IsHelpEnabled = false;
                rttA0V.IsHelpEnabled = false;


                rbA4H.CommandParameter = rttA4H.Command = "_A4H";
                rbA4H.Name = "A4H(210x297)";
                rbA4H.CommandHandler = new RibbonCommandHandler();
                rbA4H.ShowText = true;
                rbA4H.Text = "A4H(210x297)";
                rbA4H.ToolTip = rttA4H;

                rbA4V.CommandParameter = rttA4V.Command = "_A4V";
                rbA4V.Name = "A4V(297x210)";
                rbA4V.CommandHandler = new RibbonCommandHandler();
                rbA4V.ShowText = true;
                rbA4V.Text = "A4V(297x210)";
                rbA4V.ToolTip = rttA4V;

                rbA4H3.CommandParameter = rttA4H3.Command = "_A4H3";
                rbA4H3.Name = "A4H(297x630)";
                rbA4H3.CommandHandler = new RibbonCommandHandler();
                rbA4H3.ShowText = true;
                rbA4H3.Text = "A4H(297x630)";
                rbA4H3.ToolTip = rttA4H3;

                rbA4H4.CommandParameter = rttA4H4.Command = "_A4H4";
                rbA4H4.Name = "A4H(297x841)";
                rbA4H4.CommandHandler = new RibbonCommandHandler();
                rbA4H4.ShowText = true;
                rbA4H4.Text = "A4H(297x841)";
                rbA4H4.ToolTip = rttA4H4;

                rbA3H.CommandParameter = rttA3H.Command = "_A3H";
                rbA3H.Name = "A3H(297x420)";
                rbA3H.CommandHandler = new RibbonCommandHandler();
                rbA3H.ShowText = true;
                rbA3H.Text = "A3H(297x420)";
                rbA3H.ToolTip = rttA3H;

                rbA3V.CommandParameter = rttA3V.Command = "_A3V";
                rbA3V.Name = "A3V(420x297)";
                rbA3V.CommandHandler = new RibbonCommandHandler();
                rbA3V.ShowText = true;
                rbA3V.Text = "A3V(420x297)";
                rbA3V.ToolTip = rttA3V;

                rbA3H3.CommandParameter = rttA3H3.Command = "_A3H3";
                rbA3H3.Name = "A3H(420x891)";
                rbA3H3.CommandHandler = new RibbonCommandHandler();
                rbA3H3.ShowText = true;
                rbA3H3.Text = "A3H(420x891)";
                rbA3H3.ToolTip = rttA3H3;

                rbA2H.CommandParameter = rttA2H.Command = "_A2H";
                rbA2H.Name = "A2H(420x594)";
                rbA2H.CommandHandler = new RibbonCommandHandler();
                rbA2H.ShowText = true;
                rbA2H.Text = "A2H(420x594)";
                rbA2H.ToolTip = rttA2H;

                rbA2V.CommandParameter = rttA2V.Command = "_A2V";
                rbA2V.Name = "A2V(594x420)";
                rbA2V.CommandHandler = new RibbonCommandHandler();
                rbA2V.ShowText = true;
                rbA2V.Text = "A2V(594x420)";
                rbA2V.ToolTip = rttA2V;

                rbA1H.CommandParameter = rttA1H.Command = "_A1H";
                rbA1H.Name = "A1H(594x841)";
                rbA1H.CommandHandler = new RibbonCommandHandler();
                rbA1H.ShowText = true;
                rbA1H.Text = "A1H(594x841)";
                rbA1H.ToolTip = rttA1H;

                rbA1V.CommandParameter = rttA1V.Command = "_A1V";
                rbA1V.Name = "A1V(841x594)";
                rbA1V.CommandHandler = new RibbonCommandHandler();
                rbA1V.ShowText = true;
                rbA1V.Text = "A1V(841x594)";
                rbA1V.ToolTip = rttA0V;

                rbA0H.CommandParameter = rttA0H.Command = "_A0H";
                rbA0H.Name = "A0H(841x1189)";
                rbA0H.CommandHandler = new RibbonCommandHandler();
                rbA0H.ShowText = true;
                rbA0H.Text = "A0H(841x1189)";
                rbA0H.ToolTip = rttA0H;

                rbA0V.CommandParameter = rttA0V.Command = "_A0V";
                rbA0V.Name = "A0V(1189x841)";
                rbA0V.CommandHandler = new RibbonCommandHandler();
                rbA0V.ShowText = true;
                rbA0V.Text = "A0V(1189x841)";
                rbA0V.ToolTip = rttA0V;

            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
            //ribbSplitButton.Current = rbA4H;

            //Добавляем кнопку на раскрывающуюся панель
            ribbSplitButton.Items.Add(rbA4H);
            ribbSplitButton.Items.Add(rbA4V);

            ribbSplitButton.Items.Add(rbA4H3);
            ribbSplitButton.Items.Add(rbA4H4);

            ribbSplitButton.Items.Add(rbA3H);
            ribbSplitButton.Items.Add(rbA3V);

            ribbSplitButton.Items.Add(rbA3H3);

            ribbSplitButton.Items.Add(rbA2H);
            ribbSplitButton.Items.Add(rbA2V);

            ribbSplitButton.Items.Add(rbA1H);
            ribbSplitButton.Items.Add(rbA1V);

            ribbSplitButton.Items.Add(rbA0H);
            ribbSplitButton.Items.Add(rbA0V);

            //ribbPanelSource.Items.Add(ribbButton);
            //Добавляем раскрывающуюся панель на панель вкладки
            ribbPanelSource.Items.Add(ribbSplitButton);
            return ribbPanel;
        }

        static RibbonPanel FrameArch()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            RibbonToolTip rttA4HA = new RibbonToolTip();
            RibbonToolTip rttA4VA = new RibbonToolTip();

            RibbonToolTip rttA3HA = new RibbonToolTip();
            RibbonToolTip rttA3VA = new RibbonToolTip();

            RibbonToolTip rttA2HA = new RibbonToolTip();
            RibbonToolTip rttA2VA = new RibbonToolTip();

            RibbonToolTip rttA1HA = new RibbonToolTip();
            RibbonToolTip rttA1VA = new RibbonToolTip();

            RibbonToolTip rttA0HA = new RibbonToolTip();
            RibbonToolTip rttA0VA = new RibbonToolTip();


            RibbonButton rbA4HA = new RibbonButton();
            RibbonButton rbA4VA = new RibbonButton();

            RibbonButton rbA3HA = new RibbonButton();
            RibbonButton rbA3VA = new RibbonButton();

            RibbonButton rbA2HA = new RibbonButton();
            RibbonButton rbA2VA = new RibbonButton();

            RibbonButton rbA1HA = new RibbonButton();
            RibbonButton rbA1VA = new RibbonButton();

            RibbonButton rbA0HA = new RibbonButton();
            RibbonButton rbA0VA = new RibbonButton();

            RibbonSplitButton ribbSplitButtonArch;
            RibbonPanelSource ribbPanelSourceArch = new RibbonPanelSource();

            ribbPanelSourceArch.Title = "FrameArch";

            RibbonPanel ribbPanel = new RibbonPanel();

            ribbPanel.Source = ribbPanelSourceArch;

            ribbSplitButtonArch = new RibbonSplitButton();
            ribbSplitButtonArch.Text = "RibbonSplitButton";
            ribbSplitButtonArch.Orientation = System.Windows.Controls.Orientation.Vertical;
            ribbSplitButtonArch.Size = RibbonItemSize.Large;
            ribbSplitButtonArch.Size = RibbonItemSize.Large;
            ribbSplitButtonArch.ShowImage = true;
            ribbSplitButtonArch.ShowText = true;
            // Стиль кнопки
            ribbSplitButtonArch.ListButtonStyle = Autodesk.Private.Windows.RibbonListButtonStyle.SplitButton;
            ribbSplitButtonArch.ResizeStyle = RibbonItemResizeStyles.NoResize;
            ribbSplitButtonArch.ListStyle = RibbonSplitButtonListStyle.List;

            try
            {
                rttA4HA.IsHelpEnabled = false;
                rttA4VA.IsHelpEnabled = false;

                rttA3HA.IsHelpEnabled = false;
                rttA3VA.IsHelpEnabled = false;

                rttA2HA.IsHelpEnabled = false;
                rttA2VA.IsHelpEnabled = false;

                rttA1HA.IsHelpEnabled = false;
                rttA1VA.IsHelpEnabled = false;

                rttA0HA.IsHelpEnabled = false;
                rttA0VA.IsHelpEnabled = false;
                ///////////////////////////////////////////////////////
                rbA4HA.CommandParameter = rttA4HA.Command = "_A4HA";
                rbA4HA.Name = "A4HA(210x297)";
                rbA4HA.CommandHandler = new RibbonCommandHandler();
                rbA4HA.ShowText = true;
                rbA4HA.Text = "A4HA(210x297)";
                rbA4HA.ToolTip = rttA4HA;

                rbA4VA.CommandParameter = rttA4VA.Command = "_A4VA";
                rbA4VA.Name = "A4VA(297x210)";
                rbA4VA.CommandHandler = new RibbonCommandHandler();
                rbA4VA.ShowText = true;
                rbA4VA.Text = "A4VA(297x210)";
                rbA4VA.ToolTip = rttA4VA;
                ///////////////////////////////////////////////////////
                rbA3HA.CommandParameter = rttA3HA.Command = "_A3HA";
                rbA3HA.Name = "A3HA(297x420)";
                rbA3HA.CommandHandler = new RibbonCommandHandler();
                rbA3HA.ShowText = true;
                rbA3HA.Text = "A3HA(297x420)";
                rbA3HA.ToolTip = rttA3HA;

                rbA3VA.CommandParameter = rttA3VA.Command = "_A3VA";
                rbA3VA.Name = "A3VA(420x297)";
                rbA3VA.CommandHandler = new RibbonCommandHandler();
                rbA3VA.ShowText = true;
                rbA3VA.Text = "A3VA(420x297)";
                rbA3VA.ToolTip = rttA3VA;
                ///////////////////////////////////////////////////////
                rbA2HA.CommandParameter = rttA2HA.Command = "_A2HA";
                rbA2HA.Name = "A2HA(420x594)";
                rbA2HA.CommandHandler = new RibbonCommandHandler();
                rbA2HA.ShowText = true;
                rbA2HA.Text = "A2HA(420x594)";
                rbA2HA.ToolTip = rttA2HA;

                rbA2VA.CommandParameter = rttA2VA.Command = "_A2VA";
                rbA2VA.Name = "A2VA(594x420)";
                rbA2VA.CommandHandler = new RibbonCommandHandler();
                rbA2VA.ShowText = true;
                rbA2VA.Text = "A2VA(594x420)";
                rbA2VA.ToolTip = rttA2VA;
                ///////////////////////////////////////////////////////
                rbA1HA.CommandParameter = rttA1HA.Command = "_A1HA";
                rbA1HA.Name = "A1HA(594x841)";
                rbA1HA.CommandHandler = new RibbonCommandHandler();
                rbA1HA.ShowText = true;
                rbA1HA.Text = "A1HA(594x841)";
                rbA1HA.ToolTip = rttA1HA;

                rbA1VA.CommandParameter = rttA1VA.Command = "_A1VA";
                rbA1VA.Name = "A1VA(841x594)";
                rbA1VA.CommandHandler = new RibbonCommandHandler();
                rbA1VA.ShowText = true;
                rbA1VA.Text = "A1VA(841x594)";
                rbA1VA.ToolTip = rttA1VA;
                ///////////////////////////////////////////////////////
                rbA0HA.CommandParameter = rttA0HA.Command = "_A0HA";
                rbA0HA.Name = "A0HA(841x1189)";
                rbA0HA.CommandHandler = new RibbonCommandHandler();
                rbA0HA.ShowText = true;
                rbA0HA.Text = "A0HA(841x1189)";
                rbA0HA.ToolTip = rttA0HA;

                rbA0VA.CommandParameter = rttA0VA.Command = "_A0VA";
                rbA0VA.Name = "A0VA(1189x841)";
                rbA0VA.CommandHandler = new RibbonCommandHandler();
                rbA0VA.ShowText = true;
                rbA0VA.Text = "A0VA(1189x841)";
                rbA0VA.ToolTip = rttA0VA;
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
            ///////////////////////////////////////////////////////
            //ribbSplitButton.Current = rbA4H;

            //Добавляем кнопку на раскрывающуюся панель
            ribbSplitButtonArch.Items.Add(rbA4HA);
            ribbSplitButtonArch.Items.Add(rbA4VA);

            ribbSplitButtonArch.Items.Add(rbA3HA);
            ribbSplitButtonArch.Items.Add(rbA3VA);

            ribbSplitButtonArch.Items.Add(rbA2HA);
            ribbSplitButtonArch.Items.Add(rbA2VA);

            ribbSplitButtonArch.Items.Add(rbA1HA);
            ribbSplitButtonArch.Items.Add(rbA1VA);

            ribbSplitButtonArch.Items.Add(rbA0HA);
            ribbSplitButtonArch.Items.Add(rbA0VA);

            //ribbPanelSource.Items.Add(ribbButton);
            //Добавляем раскрывающуюся панель на панель вкладки
            ribbPanelSourceArch.Items.Add(ribbSplitButtonArch);
            return ribbPanel;
        }
        static RibbonPanel Stamp()
        {
            RibbonToolTip rttStampK = new RibbonToolTip();
            RibbonToolTip rttStampAR = new RibbonToolTip();
            RibbonToolTip rttStampKAB = new RibbonToolTip();

            RibbonButton rbStampK = new RibbonButton();
            RibbonButton rbStampAR = new RibbonButton();
            RibbonButton rbStampKAB = new RibbonButton();

            RibbonSplitButton ribbSplitButton;

            RibbonPanelSource ribbPanelSource = new RibbonPanelSource();
            ribbPanelSource.Title = "Stamp";
            RibbonPanel ribbPanel = new RibbonPanel();
            ribbPanel.Source = ribbPanelSource;

            ribbSplitButton = new RibbonSplitButton();
            ribbSplitButton.Text = "RibbonSplitButton";
            ribbSplitButton.Orientation = System.Windows.Controls.Orientation.Vertical;
            ribbSplitButton.Size = RibbonItemSize.Large;
            ribbSplitButton.Size = RibbonItemSize.Large;
            ribbSplitButton.ShowImage = true;
            ribbSplitButton.ShowText = true;

            ribbSplitButton.ListButtonStyle = Autodesk.Private.Windows.RibbonListButtonStyle.SplitButton;
            ribbSplitButton.ResizeStyle = RibbonItemResizeStyles.NoResize;
            ribbSplitButton.ListStyle = RibbonSplitButtonListStyle.List;

            rttStampK.IsHelpEnabled = false;
            rttStampAR.IsHelpEnabled = false;
            rttStampKAB.IsHelpEnabled = false;

            rbStampK.CommandParameter = rttStampK.Command = "_STAMPK";
            rbStampK.Name = "Constructors";
            rbStampK.CommandHandler = new RibbonCommandHandler();
            rbStampK.ShowText = true;
            rbStampK.Text = "Constructors";
            rbStampK.ToolTip = rttStampK;

            rbStampAR.CommandParameter = rttStampAR.Command = "_STAMPAR";
            rbStampAR.Name = "Architects";
            rbStampAR.CommandHandler = new RibbonCommandHandler();
            rbStampAR.ShowText = true;
            rbStampAR.Text = "Architects";
            rbStampAR.ToolTip = rttStampAR;

            rbStampKAB.CommandParameter = rttStampKAB.Command = "_STAMPKAB";
            rbStampKAB.Name = "Constructor M";
            rbStampKAB.CommandHandler = new RibbonCommandHandler();
            rbStampKAB.ShowText = true;
            rbStampKAB.Text = "Constructor M";
            rbStampKAB.ToolTip = rttStampKAB;

            //Добавляем кнопку на раскрывающуюся панель
            ribbSplitButton.Items.Add(rbStampK);
            ribbSplitButton.Items.Add(rbStampAR);
            ribbSplitButton.Items.Add(rbStampKAB);

            //Добавляем раскрывающуюся панель на панель вкладки
            ribbPanelSource.Items.Add(ribbSplitButton);
            return ribbPanel;
        }

        static RibbonPanel Layer()
        {
            RibbonToolTip rttLayerArchitect = new RibbonToolTip();
            RibbonToolTip rttLayerConstruct = new RibbonToolTip();

            RibbonButton rbLayerArchitect = new RibbonButton();
            RibbonButton rbLayerConstruct = new RibbonButton();

            RibbonSplitButton ribbSplitButton;

            RibbonPanelSource ribbPanelSource = new RibbonPanelSource();
            ribbPanelSource.Title = "Layers";
            RibbonPanel ribbPanel = new RibbonPanel();
            ribbPanel.Source = ribbPanelSource;

            ribbSplitButton = new RibbonSplitButton();
            ribbSplitButton.Text = "RibbonSplitButton";
            ribbSplitButton.Orientation = System.Windows.Controls.Orientation.Vertical;
            ribbSplitButton.Size = RibbonItemSize.Large;
            ribbSplitButton.Size = RibbonItemSize.Large;
            ribbSplitButton.ShowImage = true;
            ribbSplitButton.ShowText = true;

            ribbSplitButton.ListButtonStyle = Autodesk.Private.Windows.RibbonListButtonStyle.SplitButton;
            ribbSplitButton.ResizeStyle = RibbonItemResizeStyles.NoResize;
            ribbSplitButton.ListStyle = RibbonSplitButtonListStyle.List;

            rttLayerArchitect.IsHelpEnabled = false;
            rttLayerConstruct.IsHelpEnabled = false;

            rbLayerArchitect.CommandParameter = rttLayerArchitect.Command = "_LAYERARCHITECT";
            rbLayerArchitect.Name = "Layers for architects";
            rbLayerArchitect.CommandHandler = new RibbonCommandHandler();
            rbLayerArchitect.ShowText = true;
            rbLayerArchitect.Text = "Layers for architects";
            rbLayerArchitect.ToolTip = rttLayerArchitect;

            rbLayerConstruct.CommandParameter = rttLayerConstruct.Command = "_LAYERCONSTRUCT";
            rbLayerConstruct.Name = "Layers for constructors";
            rbLayerConstruct.CommandHandler = new RibbonCommandHandler();
            rbLayerConstruct.ShowText = true;
            rbLayerConstruct.Text = "Layers for constructors";
            rbLayerConstruct.ToolTip = rttLayerConstruct;

            //Добавляем кнопку на раскрывающуюся панель
            ribbSplitButton.Items.Add(rbLayerArchitect);
            ribbSplitButton.Items.Add(rbLayerConstruct);

            //Добавляем раскрывающуюся панель на панель вкладки
            ribbPanelSource.Items.Add(ribbSplitButton);
            return ribbPanel;
        }

        static RibbonPanel RibbonSave()
        {

            RibbonToolTip rttSave = new RibbonToolTip();
            RibbonToolTip rttSaveAndClose = new RibbonToolTip();
            RibbonToolTip rttCloseAndDiscard = new RibbonToolTip();
            RibbonToolTip rttSaveDrawingWithXREF = new RibbonToolTip();

            RibbonButton rbSave = new RibbonButton();
            RibbonButton rbSaveAndClose = new RibbonButton();
            RibbonButton rbCloseAndDiscard = new RibbonButton();
            RibbonButton rbSaveDrawingWithXREF = new RibbonButton();


            RibbonSplitButton ribbSplitButton;

            RibbonPanelSource ribbPanelSource = new RibbonPanelSource();
            ribbPanelSource.Title = "Save";
            RibbonPanel ribbPanel = new RibbonPanel();
            ribbPanel.Source = ribbPanelSource;

            ribbSplitButton = new RibbonSplitButton();
            ribbSplitButton.Text = "RibbonSplitButton";
            ribbSplitButton.Orientation = System.Windows.Controls.Orientation.Vertical;
            ribbSplitButton.Size = RibbonItemSize.Large;
            ribbSplitButton.Size = RibbonItemSize.Large;
            ribbSplitButton.ShowImage = true;
            ribbSplitButton.ShowText = true;

            ribbSplitButton.ListButtonStyle = Autodesk.Private.Windows.RibbonListButtonStyle.SplitButton;
            ribbSplitButton.ResizeStyle = RibbonItemResizeStyles.NoResize;
            ribbSplitButton.ListStyle = RibbonSplitButtonListStyle.List;

            rttSave.IsHelpEnabled = false;
            rttSaveAndClose.IsHelpEnabled = false;
            rttCloseAndDiscard.IsHelpEnabled = false;
            rttSaveDrawingWithXREF.IsHelpEnabled = false;

            rbSave.CommandParameter = rttSave.Command = "_SAVEACTIVEDRAWING";
            rbSave.Name = "Save drawing";
            rbSave.CommandHandler = new RibbonCommandHandler();
            rbSave.ShowText = true;
            rbSave.Text = "Save drawing";
            rbSave.ToolTip = rttSave;

            rbSaveAndClose.CommandParameter = rttSaveAndClose.Command = "_SAVEANDCLOSE";
            rbSaveAndClose.Name = "Save and close drawing";
            rbSaveAndClose.CommandHandler = new RibbonCommandHandler();
            rbSaveAndClose.ShowText = true;
            rbSaveAndClose.Text = "Save and close drawing";
            rbSaveAndClose.ToolTip = rttSaveAndClose;

            rbCloseAndDiscard.CommandParameter = rttCloseAndDiscard.Command = "_CLOSEANDDISCARD";
            rbCloseAndDiscard.Name = "Close and discard changes";
            rbCloseAndDiscard.CommandHandler = new RibbonCommandHandler();
            rbCloseAndDiscard.ShowText = true;
            rbCloseAndDiscard.Text = "Close and discard changes";
            rbCloseAndDiscard.ToolTip = rttCloseAndDiscard;

            rbSaveDrawingWithXREF.CommandParameter = rttSaveDrawingWithXREF.Command = "_REFRESHMAP";
            rbSaveDrawingWithXREF.Name = "Save drawing with Xref in TDMS";
            rbSaveDrawingWithXREF.CommandHandler = new RibbonCommandHandler();
            rbSaveDrawingWithXREF.ShowText = true;
            rbSaveDrawingWithXREF.Text = "Save drawing with Xref in TDMS";
            rbSaveDrawingWithXREF.ToolTip = rttSaveDrawingWithXREF;

            //Добавляем кнопку на раскрывающуюся панель
            ribbSplitButton.Items.Add(rbSave);
            ribbSplitButton.Items.Add(rbSaveAndClose);
            ribbSplitButton.Items.Add(rbCloseAndDiscard);
            ribbSplitButton.Items.Add(rbSaveDrawingWithXREF);

            //Добавляем раскрывающуюся панель на панель вкладки
            ribbPanelSource.Items.Add(ribbSplitButton);
            return ribbPanel;
        }

        static RibbonPanel Optimization()
        {
            RibbonToolTip rttOptimization = new RibbonToolTip();
            RibbonButton rbOptimization = new RibbonButton();
            RibbonSplitButton ribbSplitButton;

            RibbonPanelSource ribbPanelSource = new RibbonPanelSource();
            ribbPanelSource.Title = "Optimization";
            RibbonPanel ribbPanel = new RibbonPanel();
            ribbPanel.Source = ribbPanelSource;

            ribbSplitButton = new RibbonSplitButton();
            ribbSplitButton.Text = "Optimization";
            ribbSplitButton.Orientation = System.Windows.Controls.Orientation.Vertical;
            ribbSplitButton.Size = RibbonItemSize.Large;
            ribbSplitButton.Size = RibbonItemSize.Large;
            ribbSplitButton.ShowImage = true;
            ribbSplitButton.ShowText = true;

            ribbSplitButton.ListButtonStyle = Autodesk.Private.Windows.RibbonListButtonStyle.SplitButton;
            ribbSplitButton.ResizeStyle = RibbonItemResizeStyles.NoResize;
            ribbSplitButton.ListStyle = RibbonSplitButtonListStyle.List;

            rttOptimization.IsHelpEnabled = false;

            rbOptimization.CommandParameter = rttOptimization.Command = "_CustomVar";
            rbOptimization.Name = "Optimization";
            rbOptimization.CommandHandler = new RibbonCommandHandler();
            rbOptimization.ShowText = true;
            rbOptimization.Text = "Optimization";
            rbOptimization.ToolTip = rttOptimization;

            //Добавляем кнопку на раскрывающуюся панель
            ribbSplitButton.Items.Add(rbOptimization);
            //Добавляем раскрывающуюся панель на панель вкладки
            ribbPanelSource.Items.Add(ribbSplitButton);
            return ribbPanel;
        }

        static RibbonPanel Etransmit()
        {
            RibbonToolTip rttEtransmit = new RibbonToolTip();
            RibbonButton rbEtransmit = new RibbonButton();
            RibbonSplitButton ribbSplitButton;

            RibbonPanelSource ribbPanelSource = new RibbonPanelSource();
            ribbPanelSource.Title = "Etransmit";
            RibbonPanel ribbPanel = new RibbonPanel();
            ribbPanel.Source = ribbPanelSource;

            ribbSplitButton = new RibbonSplitButton();
            ribbSplitButton.Text = "Etransmit";
            ribbSplitButton.Orientation = System.Windows.Controls.Orientation.Vertical;
            ribbSplitButton.Size = RibbonItemSize.Large;
            ribbSplitButton.Size = RibbonItemSize.Large;
            ribbSplitButton.ShowImage = true;
            ribbSplitButton.ShowText = true;

            ribbSplitButton.ListButtonStyle = Autodesk.Private.Windows.RibbonListButtonStyle.SplitButton;
            ribbSplitButton.ResizeStyle = RibbonItemResizeStyles.NoResize;
            ribbSplitButton.ListStyle = RibbonSplitButtonListStyle.List;

            rttEtransmit.IsHelpEnabled = false;

            rbEtransmit.CommandParameter = rttEtransmit.Command = "_PACKAGE";
            rbEtransmit.Name = "Etransmit";
            rbEtransmit.CommandHandler = new RibbonCommandHandler();
            rbEtransmit.ShowText = true;
            rbEtransmit.Text = "Etransmit";
            rbEtransmit.ToolTip = rttEtransmit;

            //Добавляем кнопку на раскрывающуюся панель
            ribbSplitButton.Items.Add(rbEtransmit);
            //Добавляем раскрывающуюся панель на панель вкладки
            ribbPanelSource.Items.Add(ribbSplitButton);
            return ribbPanel;
        }
        static RibbonPanel ribbonXREF()
        {
            RibbonToolTip rttXrefTDMS = new RibbonToolTip();

            RibbonButton rbXrefTDMSDWG = new RibbonButton();
            RibbonButton rbXrefTDMSIMG = new RibbonButton();
            RibbonButton rbXrefSynchronize = new RibbonButton();

            RibbonSplitButton ribbSplitButton;

            RibbonPanelSource ribbPanelSource = new RibbonPanelSource();
            ribbPanelSource.Title = "Xref";
            RibbonPanel ribbPanel = new RibbonPanel();
            ribbPanel.Source = ribbPanelSource;

            ribbSplitButton = new RibbonSplitButton();
            ribbSplitButton.Text = "Xref";
            ribbSplitButton.Orientation = System.Windows.Controls.Orientation.Vertical;
            ribbSplitButton.Size = RibbonItemSize.Large;
            ribbSplitButton.ShowImage = true;
            ribbSplitButton.ShowText = true;

            ribbSplitButton.ListButtonStyle = Autodesk.Private.Windows.RibbonListButtonStyle.SplitButton;
            ribbSplitButton.ResizeStyle = RibbonItemResizeStyles.NoResize;
            ribbSplitButton.ListStyle = RibbonSplitButtonListStyle.List;

            rttXrefTDMS.IsHelpEnabled = false;

            rbXrefTDMSDWG.CommandParameter = rttXrefTDMS.Command = "_TDMSXREFDWG";
            rbXrefTDMSDWG.Name = "Xref DWG from TDMS";
            rbXrefTDMSDWG.CommandHandler = new RibbonCommandHandler();
            rbXrefTDMSDWG.ShowText = true;
            rbXrefTDMSDWG.Text = "Xref DWG from TDMS";
            rbXrefTDMSDWG.ToolTip = rttXrefTDMS;

            rbXrefTDMSIMG.CommandParameter = rttXrefTDMS.Command = "_TDMSXREFIMG";
            rbXrefTDMSIMG.Name = "Xref IMG from TDMS";
            rbXrefTDMSIMG.CommandHandler = new RibbonCommandHandler();
            rbXrefTDMSIMG.ShowText = true;
            rbXrefTDMSIMG.Text = "Xref IMG from TDMS";
            rbXrefTDMSIMG.ToolTip = rttXrefTDMS;

            rbXrefSynchronize.CommandParameter = rttXrefTDMS.Command = "_XREFSYNCHRO";
            rbXrefSynchronize.Name = "Xref synhronize with TDMS";
            rbXrefSynchronize.CommandHandler = new RibbonCommandHandler();
            rbXrefSynchronize.ShowText = true;
            rbXrefSynchronize.Text = "Xref synhronize with TDMS";
            rbXrefSynchronize.ToolTip = rttXrefTDMS;

            //Добавляем кнопку на раскрывающуюся панель
            ribbSplitButton.Items.Add(rbXrefTDMSDWG);
            ribbSplitButton.Items.Add(rbXrefTDMSIMG);
            ribbSplitButton.Items.Add(rbXrefSynchronize);
            //Добавляем раскрывающуюся панель на панель вкладки
            ribbPanelSource.Items.Add(ribbSplitButton);
            return ribbPanel;
        }
        static RibbonPanel UpdateAttribute()
        {
            RibbonToolTip rttUpdateAttribute = new RibbonToolTip();
            RibbonButton rbUpdateAttribute = new RibbonButton();
            RibbonButton rbUpdateXREF = new RibbonButton();
            RibbonSplitButton ribbSplitButton;

            RibbonPanelSource ribbPanelSource = new RibbonPanelSource();
            ribbPanelSource.Title = "Update";
            RibbonPanel ribbPanel = new RibbonPanel();
            ribbPanel.Source = ribbPanelSource;

            ribbSplitButton = new RibbonSplitButton();
            ribbSplitButton.Text = "Update";
            ribbSplitButton.Orientation = System.Windows.Controls.Orientation.Vertical;
            ribbSplitButton.Size = RibbonItemSize.Large;
            ribbSplitButton.Size = RibbonItemSize.Large;
            ribbSplitButton.ShowImage = true;
            ribbSplitButton.ShowText = true;

            ribbSplitButton.ListButtonStyle = Autodesk.Private.Windows.RibbonListButtonStyle.SplitButton;
            ribbSplitButton.ResizeStyle = RibbonItemResizeStyles.NoResize;
            ribbSplitButton.ListStyle = RibbonSplitButtonListStyle.List;

            rttUpdateAttribute.IsHelpEnabled = false;

            rbUpdateAttribute.CommandParameter = rttUpdateAttribute.Command = "_UA";
            rbUpdateAttribute.Name = "Update attribute";
            rbUpdateAttribute.CommandHandler = new RibbonCommandHandler();
            rbUpdateAttribute.ShowText = true;
            rbUpdateAttribute.Text = "Update attribute";
            rbUpdateAttribute.ToolTip = rttUpdateAttribute;

            rbUpdateXREF.CommandParameter = rttUpdateAttribute.Command = "_XREFUpdate";
            rbUpdateXREF.Name = "Update XREF";
            rbUpdateXREF.CommandHandler = new RibbonCommandHandler();
            rbUpdateXREF.ShowText = true;
            rbUpdateXREF.Text = "Update XREF";
            rbUpdateXREF.ToolTip = rttUpdateAttribute;

            //Добавляем кнопку на раскрывающуюся панель
            ribbSplitButton.Items.Add(rbUpdateAttribute);
            ribbSplitButton.Items.Add(rbUpdateXREF);
            //Добавляем раскрывающуюся панель на панель вкладки
            ribbPanelSource.Items.Add(ribbSplitButton);
            return ribbPanel;
        }
        static RibbonPanel Helps()
        {
            RibbonToolTip rttHelps = new RibbonToolTip();
            RibbonButton rbHelps = new RibbonButton();
            RibbonSplitButton ribbSplitButton;

            RibbonPanelSource ribbPanelSource = new RibbonPanelSource();
            ribbPanelSource.Title = "Shared HELP";
            RibbonPanel ribbPanel = new RibbonPanel();
            ribbPanel.Source = ribbPanelSource;

            ribbSplitButton = new RibbonSplitButton();
            ribbSplitButton.Text = "Shared HELP";
            ribbSplitButton.Orientation = System.Windows.Controls.Orientation.Vertical;
            ribbSplitButton.Size = RibbonItemSize.Large;
            ribbSplitButton.Size = RibbonItemSize.Large;
            ribbSplitButton.ShowImage = true;
            ribbSplitButton.ShowText = true;

            ribbSplitButton.ListButtonStyle = Autodesk.Private.Windows.RibbonListButtonStyle.SplitButton;
            ribbSplitButton.ResizeStyle = RibbonItemResizeStyles.NoResize;
            ribbSplitButton.ListStyle = RibbonSplitButtonListStyle.List;

            rttHelps.IsHelpEnabled = false;

            rbHelps.CommandParameter = rttHelps.Command = "_OPENHELP";
            rbHelps.Name = "Shared HELP";
            rbHelps.CommandHandler = new RibbonCommandHandler();
            rbHelps.ShowText = true;
            rbHelps.Text = "Shared HELP";
            rbHelps.ToolTip = rttHelps;

            //Добавляем кнопку на раскрывающуюся панель
            ribbSplitButton.Items.Add(rbHelps);
            //Добавляем раскрывающуюся панель на панель вкладки
            ribbPanelSource.Items.Add(ribbSplitButton);
            return ribbPanel;
        }
        
        //Регистрация приложения в системе, не придётся постоянно вызывать Netload для загрузки dll.
        [CommandMethod("RegisterTDMSApp")]
        public void RegisterTDMSApp()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                // Из реестра получаем ключ AutoCAD
                string sProdKey = Autodesk.AutoCAD.DatabaseServices.HostApplicationServices.Current.UserRegistryProductRootKey;
                string sAppName = "MyApp";
                RegistryKey regAcadProdKey = Registry.CurrentUser.OpenSubKey(sProdKey);
                RegistryKey regAcadAppKey = regAcadProdKey.OpenSubKey("Applications", true);
                // Проверяем, существует ли в ветке реестра раздел "MyApp"
                string[] subKeys = regAcadAppKey.GetSubKeyNames();
                foreach (string subKey in subKeys)
                {
                    // Если приложение уже зарегистрировано - завершаем работу
                    if (subKey.Equals(sAppName))
                    {
                        regAcadAppKey.Close();
                        return;
                    }
                }
                // Получаем строковое представление пути к каталогу, в котором находится текущий модуль
                string sAssemblyPath = Assembly.GetExecutingAssembly().Location;
                // Регистрируем приложение
                RegistryKey regAppAddInKey = regAcadAppKey.CreateSubKey(sAppName);
                regAppAddInKey.SetValue("DESCRIPTION", sAppName, RegistryValueKind.String);
                regAppAddInKey.SetValue("LOADCTRLS", 2, RegistryValueKind.DWord);
                regAppAddInKey.SetValue("LOADER", sAssemblyPath, RegistryValueKind.String);
                regAppAddInKey.SetValue("MANAGED", 1, RegistryValueKind.DWord);
                regAcadAppKey.Close();
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        ////Очистка регистров в системе, вызвать Netload для загрузки dll.
        [CommandMethod("UnregisterTDMSApp")]
        public void UnregisterTDMSApp()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                // Из реестра получаем ключ AutoCAD
                string sProdKey = Autodesk.AutoCAD.DatabaseServices.HostApplicationServices.Current.UserRegistryProductRootKey;
                string sAppName = "MyApp";
                RegistryKey regAcadProdKey = Registry.CurrentUser.OpenSubKey(sProdKey);
                RegistryKey regAcadAppKey = regAcadProdKey.OpenSubKey("Applications", true);
                // Удаляем ключ приложения
                regAcadAppKey.DeleteSubKeyTree(sAppName);
                regAcadAppKey.Close();
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        /////////////////////////////////////////////////////////////////////////////////

        public static string GetUserName()
        {
            System.Security.Principal.WindowsIdentity win = null;
            win = System.Security.Principal.WindowsIdentity.GetCurrent();
            return win.Name.Substring(win.Name.IndexOf("\\") + 1);
        }


        ///////////////////////////////////////////////////////////////////////////////////

        ///////////////////////////////////////////////////////////////////////////////////

        // не архивировать
        [CommandMethod("PACKAGE", CommandFlags.NoBlockEditor)]
        public static void Etrans()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                SaveOptions Save = new SaveOptions();
                Save.StartSaveActiveDrawing();
                Document doc = Application.DocumentManager.MdiActiveDocument;
                //string folderName = @"c:\";
                //string pathString = System.IO.Path.Combine(folderName, "SubFolder");
                //System.IO.Directory.CreateDirectory(pathString);
                //Registry.CurrentUser.CreateSubKey("Software\\Autodesk\\AutoCAD\\R19.0\\ACAD-B001:409\\ETransmit\\setups\\Enterprise"); //Для AutoCAD 2013
                Registry.CurrentUser.CreateSubKey("Software\\Autodesk\\AutoCAD\\R19.0\\ACAD-B004:409\\ETransmit\\setups\\Enterprise"); //Для AutoCAD Architecture 2013
                RegistryKey myKey = Registry.CurrentUser.OpenSubKey("Software\\Autodesk\\AutoCAD\\R19.0\\ACAD-B004:409\\ETransmit\\setups\\Enterprise", true);
                myKey.SetValue("AEC_EXPLODE_DWG", "1", RegistryValueKind.DWord);
                myKey.SetValue("AEC_INCLUDE_DWTS", "0", RegistryValueKind.DWord);
                myKey.SetValue("AEC_INCLUDE_PROJ_INFO", "0", RegistryValueKind.DWord);
                myKey.SetValue("AEC_TEMP_SETUP", "0", RegistryValueKind.DWord);
                
                myKey.SetValue("BindType", "1", RegistryValueKind.DWord);
                myKey.SetValue("BindXref", "1", RegistryValueKind.DWord);
                myKey.SetValue("Description", "", RegistryValueKind.String);
                myKey.SetValue("DestFile", "", RegistryValueKind.String);
                myKey.SetValue("DestFileAction", "2", RegistryValueKind.DWord);
                myKey.SetValue("DestFolder", "D:\\ETRANSMIT", RegistryValueKind.String);
                myKey.SetValue("FilePathOption", "1", RegistryValueKind.DWord);

                myKey.SetValue("IncludeDataLinkFile", "0", RegistryValueKind.DWord);
                myKey.SetValue("IncludeFont", "0", RegistryValueKind.DWord);
                myKey.SetValue("IncludeMaterialTextures", "0", RegistryValueKind.DWord);
                myKey.SetValue("IncludeFotometricWebFile", "0", RegistryValueKind.DWord);
                myKey.SetValue("IncludeSSFiles", "1", RegistryValueKind.DWord);
                myKey.SetValue("IncludeUnloadedReferences", "0", RegistryValueKind.DWord);
                myKey.SetValue("Name", "Enterprise", RegistryValueKind.String);
                myKey.SetValue("PackageType", "1", RegistryValueKind.DWord);
                myKey.SetValue("PurgeDatabase", "1", RegistryValueKind.DWord);
                myKey.SetValue("RootFolder", "", RegistryValueKind.String);
                myKey.SetValue("SaveDrawingFormat", "4", RegistryValueKind.DWord);
                myKey.SetValue("SendMail", "0", RegistryValueKind.DWord);
                myKey.SetValue("SetPlotterNone", "1", RegistryValueKind.DWord);
                myKey.SetValue("UsePassword", "0", RegistryValueKind.DWord);
                myKey.SetValue("VisualFidelity", "1", RegistryValueKind.DWord);

                doc.SendStringToExecute("._-ETRANSMIT" + "\n" + "CHoose setup" + "\n" + "Enterprise" + "\n" + "Create" + "\n", true, false, false);

                doc.CloseAndDiscard();
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught", ex);
            }
        }
    }
    public sealed class LayoutWorks : IExtensionApplication
    {
        public void Initialize()
        {
        }
        public void Terminate()
        {
        }

        [CommandMethod("ListLayouts", CommandFlags.Session)]
        public void ListLayouts()
        {
            DocumentCollection acDocMgr = Application.DocumentManager;
            SaveOptions SaveDwg = new SaveOptions();
            string path = @"D:\_TDMSLAYOUT";
            
            List<string> lst = new List<string>();
            
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            
            string NameLayout;
            
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                DBDictionary lays = acTrans.GetObject(acCurDb.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

                acDoc.Editor.WriteMessage("\nLayouts:");
               
                foreach (DBDictionaryEntry item in lays)
                {
                    acDoc.Editor.WriteMessage("\n Key  " + item.Key);
                    acDoc.Editor.WriteMessage("\n Value: " + item.Value);
                    acDoc.Editor.WriteMessage("\n m_key: " + item.m_key);
                    acDoc.Editor.WriteMessage("\n m_value: " + item.m_value);
                    SaveDwg.LocalSave(item.Key, path);
                    lst.Add(item.Key);
                }
                for (int a = 0; a < lst.Count-1; a++) 
                {
                    NameLayout = lst[a];
                    acDoc.Editor.WriteMessage("\n cout: " + NameLayout);
                    if (File.Exists("d:\\_TDMSLAYOUT\\" + NameLayout + ".dwg"))
                    {
                        acDoc = Application.DocumentManager.Open("d:\\_TDMSLAYOUT\\" + NameLayout + ".dwg", false);
                        //DelLayout();
                    }
                    else
                    {
                        acDoc.Editor.WriteMessage("\n File does not exist");
                    }
                       
                }
                acDoc.Editor.WriteMessage("\nFile save succesfull");
                acTrans.Commit();
            }
        }

        [CommandMethod("CreateLayout")]
        public void CreateLayout()
        {
            // Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            // Get the layout and plot settings of the named pagesetup
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Reference the Layout Manager
                LayoutManager acLayoutMgr = LayoutManager.Current;
                // Create the new layout with default settings
                ObjectId objID = acLayoutMgr.CreateLayout("newLayout");
                // Open the layout
                Layout acLayout = acTrans.GetObject(objID,
                                                    OpenMode.ForRead) as Layout;
                // Set the layout current if it is not already
                if (acLayout.TabSelected == false)
                {
                    acLayoutMgr.CurrentLayout = acLayout.LayoutName;
                }
                // Output some information related to the layout object
                acDoc.Editor.WriteMessage("\nTab Order: " + acLayout.TabOrder +
                                          "\nTab Selected: " + acLayout.TabSelected +
                                          "\nBlock Table Record ID: " +
                                          acLayout.BlockTableRecordId.ToString());
                // Save the changes made
                acTrans.Commit();
            }
        }

        [CommandMethod("ImportLayout")]
        public void ImportLayout()
        {
            // Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;

            //PromptStringOptions InputLayoutName = new PromptStringOptions("\n Input Layout Name: ");
            //InputLayoutName.AllowSpaces = true;
            //PromptResult LN = acDoc.Editor.GetString(InputLayoutName);
            //string layoutName = LN.StringResult;

            //PromptStringOptions InputPathName = new PromptStringOptions("\n Input height: ");
            //InputPathName.AllowSpaces = true;
            //PromptResult PN = acDoc.Editor.GetString(InputLayoutName);
            //string filename = PN.StringResult;
            
            // Specify the layout name and drawing file to work with
            string layoutName = "А3_2разрез";
            string filename = "d:\\Test\\Разрезы\\Разрез 2-2.dwg";
            // Create a new database object and open the drawing into memory
            Database acExDb = new Database(false, true);
            acExDb.ReadDwgFile(filename, FileOpenMode.OpenForReadAndAllShare, true, "");
            // Create a transaction for the external drawing
            using (Transaction acTransEx = acExDb.TransactionManager.StartTransaction())
            {
                // Get the layouts dictionary
                DBDictionary layoutsEx =
                    acTransEx.GetObject(acExDb.LayoutDictionaryId,
                                        OpenMode.ForRead) as DBDictionary;
                // Check to see if the layout exists in the external drawing
                if (layoutsEx.Contains(layoutName) == true)
                {
                    // Get the layout and block objects from the external drawing
                    Layout layEx =
                        layoutsEx.GetAt(layoutName).GetObject(OpenMode.ForRead) as Layout;
                    BlockTableRecord blkBlkRecEx =
                        acTransEx.GetObject(layEx.BlockTableRecordId,
                                            OpenMode.ForRead) as BlockTableRecord;
                    // Get the objects from the block associated with the layout
                    ObjectIdCollection idCol = new ObjectIdCollection();
                    foreach (ObjectId id in blkBlkRecEx)
                    {
                        idCol.Add(id);
                    }
                    // Create a transaction for the current drawing
                    using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                    {
                        // Get the block table and create a new block
                        // then copy the objects between drawings
                        BlockTable blkTbl =
                            acTrans.GetObject(acCurDb.BlockTableId,
                                              OpenMode.ForWrite) as BlockTable;
                        using (BlockTableRecord blkBlkRec = new BlockTableRecord())
                        {
                            int layoutCount = layoutsEx.Count - 1;
                            blkBlkRec.Name = "*Paper_Space" + layoutCount.ToString();
                            blkTbl.Add(blkBlkRec);
                            acTrans.AddNewlyCreatedDBObject(blkBlkRec, true);
                            acExDb.WblockCloneObjects(idCol,
                                                      blkBlkRec.ObjectId,
                                                      new IdMapping(),
                                                      DuplicateRecordCloning.Ignore,
                                                      false);
                            // Create a new layout and then copy properties between drawings
                            DBDictionary layouts =
                                acTrans.GetObject(acCurDb.LayoutDictionaryId,
                                                  OpenMode.ForWrite) as DBDictionary;
                            using (Layout lay = new Layout())
                            {
                                lay.LayoutName = layoutName;
                                lay.AddToLayoutDictionary(acCurDb, blkBlkRec.ObjectId);
                                acTrans.AddNewlyCreatedDBObject(lay, true);
                                lay.CopyFrom(layEx);

                                DBDictionary plSets =
                                    acTrans.GetObject(acCurDb.PlotSettingsDictionaryId,
                                                      OpenMode.ForRead) as DBDictionary;

                                // Check to see if a named page setup was assigned to the layout,
                                // if so then copy the page setup settings
                                if (lay.PlotSettingsName != "")
                                {
                                    // Check to see if the page setup exists
                                    if (plSets.Contains(lay.PlotSettingsName) == false)
                                    {
                                        plSets.UpgradeOpen();

                                        using (PlotSettings plSet = new PlotSettings(lay.ModelType))
                                        {
                                            plSet.PlotSettingsName = lay.PlotSettingsName;
                                            plSet.AddToPlotSettingsDictionary(acCurDb);
                                            acTrans.AddNewlyCreatedDBObject(plSet, true);

                                            DBDictionary plSetsEx =
                                                acTransEx.GetObject(acExDb.PlotSettingsDictionaryId,
                                                                    OpenMode.ForRead) as DBDictionary;

                                            PlotSettings plSetEx =
                                                plSetsEx.GetAt(lay.PlotSettingsName).GetObject(
                                                               OpenMode.ForRead) as PlotSettings;

                                            plSet.CopyFrom(plSetEx);
                                        }
                                    }
                                }
                            }
                        }
                        // Regen the drawing to get the layout tab to display
                        acDoc.Editor.Regen();

                        // Save the changes made
                        acTrans.Commit();
                    }
                }
                else
                {
                    // Display a message if the layout could not be found in the specified drawing
                    acDoc.Editor.WriteMessage("\nLayout '" + layoutName +
                                              "' could not be imported from '" + filename + "'.");
                }
                // Discard the changes made to the external drawing file
                acTransEx.Abort();
            }
            // Close the external drawing file
            acExDb.Dispose();
        }

        [CommandMethod("EraseAllLayouts")]
        public static void EraseAllLayouts()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {

                // ACAD_LAYOUT dictionary.
                DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

                // Iterate dictionary entries.
                foreach (DBDictionaryEntry de in layoutDict)
                {
                    string layoutName = de.m_value.ToString();
                    if (layoutName != "Model")
                    {
                        LayoutManager.Current.DeleteLayout(layoutName); // Delete layout.
                    }
                }
            }
            ed.Regen();   // Updates AutoCAD GUI to relect changes.
        }

        [CommandMethod("dellt")]
        public void DelLayout()
        {
            string laytName = "Work";

            LayoutManager laytmgr = LayoutManager.Current;
            try
            {
                laytmgr.DeleteLayout(laytName);

                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.Regen();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage
                    (ex.Message + "\n" + ex.StackTrace);
            }
        }

        //[CommandMethod("EXPORTLAYOUTS")]
        //public void ExportLayouts()
        //{
        //    Document acDwg = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        //    Editor editor = acDwg.Editor;
        //    DocumentLock acDocLock = null;
        //    int i = 0;
        //    Database acDB = null;
        //    Transaction acTrans = null;
        //    BlockTable acBT = null;
        //    BlockTableRecord acBTR = null;
        //    SymbolTableEnumerator acBTRE = null;
        //    Layout acLayout = null;
        //    Layout acLayoutToExport = null;
        //    List<String> liExportedLayouts = new List<string>();
        //    FileInfo fiDrawing = new FileInfo(acDwg.Name);
        //    try
        //    {  
        //        acDocLock = acDwg.LockDocument();
        //        acDB = HostApplicationServices.WorkingDatabase;

        //        acBT = (BlockTable)(acDB.BlockTableId.GetObject(OpenMode.ForWrite));
        //        acBTRE = acBT.GetEnumerator();

        //        while (acBTRE.MoveNext())
        //        {
        //            if (acBTRE.Current.IsErased == false)
        //            {
        //                acBTR = (BlockTableRecord)(acBTRE.Current.GetObject(OpenMode.ForWrite));
        //                if (acBTR.IsLayout)
        //                {
        //                    acLayout = (Layout)(acBTR.LayoutId.GetObject(OpenMode.ForWrite));
        //                    if (acLayout.ModelType == false)
        //                    {
        //                        i += 1;
        //                    }
        //                }
        //            }
        //        }
        //        LayoutManager.Current.CurrentLayout = "Model";

        //        for (int j = 1; j <= i; j++)
        //        {
        //            Database acTmpDatabase = new Database(true, false);
        //            acLayoutToExport = null;

        //            acTmpDatabase = acDB.Wblock();
        //            acTrans = acTmpDatabase.TransactionManager.StartTransaction();

        //            using (acTrans)
        //            {
        //                acBT = (BlockTable)(acTrans.GetObject(acTmpDatabase.BlockTableId, OpenMode.ForRead));
        //                acBTRE = acBT.GetEnumerator();

        //                while (acBTRE.MoveNext())
        //                {
        //                    if (acBTRE.Current.IsErased == false)
        //                    {
        //                        acBTR = (BlockTableRecord)(acTrans.GetObject(acBTRE.Current, OpenMode.ForRead));
        //                        if (acBTR.IsLayout)
        //                        {
        //                            acLayout = (Layout)(acTrans.GetObject(acBTR.LayoutId, OpenMode.ForWrite));
        //                            if (acLayout.ModelType == false)
        //                            {
        //                                if (acLayoutToExport == null)
        //                                {
        //                                    if (liExportedLayouts.Contains(acLayout.LayoutName) == false)
        //                                    {
        //                                        acLayoutToExport = acLayout;
        //                                        liExportedLayouts.Add(acLayout.LayoutName);
        //                                        acLayout.DowngradeOpen();
        //                                        continue;
        //                                    }
        //                                    else
        //                                        acLayout.Erase();
        //                                }
        //                                else
        //                                    acLayout.Erase();
        //                            }
        //                        }
        //                    }
        //                }
        //                acTrans.Commit();
        //            }
        //            acTmpDatabase.CloseInput(true);

        //            if (acLayoutToExport != null)
        //            {

        //                acTmpDatabase.SaveAs("d:\\Test\\" + acLayout.LayoutName, DwgVersion.Current);
        //            }
        //            else
        //            {
        //                acTmpDatabase.SaveAs("d:\\Test\\" + acLayout.LayoutName, DwgVersion.Current); ;
        //            }
        //        }
        //    }
        //    catch (System.Exception ex)
        //    {
        //        editor.WriteMessage(ex.Message + "/n" + ex.StackTrace + "/n");
        //    }
        //    finally
        //    {
        //        if (acDocLock != null)
        //        {
        //            acDocLock.Dispose();
        //        }
        //    }
        //}
    }
    public static class ExportLayoutsCommands
    {
        /// <summary>
        /// Automates the EXPORTLAYOUT command to export
        /// all paper space layouts to .DWG files.
        ///
        /// In this example, we export each layout to
        /// a drawing file in the same location as the
        /// current drawing, wherein each file has the
        /// name "<dwgname>_<layoutname>.dwg".
        ///
        /// This is not a functionally-complete example:
        ///
        /// No checking is done to see if any of the
        /// files already exist, and existing files
        /// are overwritten without warning or error.
        ///
        /// No checking is done to detect if an existing
        /// file exists and is in-use by another user, or
        /// cannot be overwritten for any other reason.
        ///
        /// No checking is done to ensure that the user
        /// has sufficient rights to write files in the
        /// target location.
        ///
        /// You can and should deal with any or all of
        /// the above as per your own requirements.
        ///
        /// </summary>

        [CommandMethod("EXPORTLAYOUTS")]
        public static void ExportLayoutsCommand()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var editor = doc.Editor;
            try
            {
                if ((short)Application.GetSystemVariable("DWGTITLED") == 0)
                {
                    editor.WriteMessage(
                       "\nCommand cannot be used on an unnamed drawing"
                    );
                    return;
                }
                string format =
                   Path.Combine(
                      Path.GetDirectoryName(doc.Name),
                      Path.GetFileNameWithoutExtension(doc.Name))
                   + "_{0}.dwg";

                string[] names = null;

                using (Transaction tr = doc.TransactionManager.StartTransaction())
                {
                    // Get the localized name of the model tab:
                    BlockTableRecord btr = (BlockTableRecord)
                       SymbolUtilityServices.GetBlockModelSpaceId(db)
                          .GetObject(OpenMode.ForRead);
                    Layout layout = (Layout)
                       btr.LayoutId.GetObject(OpenMode.ForRead);
                    string model = layout.LayoutName;
                    // Open the Layout dictionary:
                    IDictionary layouts = (IDictionary)
                       db.LayoutDictionaryId.GetObject(OpenMode.ForRead);
                    // Get the names and ids of all paper space layouts into a list:
                    names = layouts.Keys.Cast<string>()
                       .Where(name => name != model).ToArray();

                    tr.Commit();
                }

                int cmdecho = 0;
#if DEBUG
                cmdecho = 1;
#endif
                using (new ManagedSystemVariable("CMDECHO", cmdecho))
                using (new ManagedSystemVariable("CMDDIA", 0))
                using (new ManagedSystemVariable("FILEDIA", 0))
                using (new ManagedSystemVariable("CTAB"))
                {
                    foreach (string name in names)
                    {
                        string filename = string.Format(format, name);
                        editor.WriteMessage("\nExporting {0}\n", filename);
                        Application.SetSystemVariable("CTAB", name);
                        editor.Command("._EXPORTLAYOUT", filename);
                    }
                }
            }
            catch (System.Exception ex)
            {
#if DEBUG
                editor.WriteMessage(ex.ToString());
#else
           throw ex;
#endif
            }
        }

        /// <summary>
        ///
        /// Doesn't use the command line, requires AutoCAD R12
        /// or later and a reference to AcExportLayout.dll:
        ///
        /// This version can be used from the application context,
        /// which can make it easier to use in a batch process that
        /// exports layouts of many files.
        ///
        /// The example also shows how to use the AcExportLayout
        /// component to export a layout to an in-memory Database
        /// without creating a drawing file.
        ///
        /// </summary>

        [CommandMethod("EXPORTLAYOUTS2", CommandFlags.Session)]
        public static void ExportLayouts2()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var editor = doc.Editor;
            try
            {
                if ((short)Application.GetSystemVariable("DWGTITLED") == 0)
                {
                    editor.WriteMessage(
                       "\nCommand cannot be used on an unnamed drawing"
                    );
                    return;
                }
                string format =
                   Path.Combine(
                      Path.GetDirectoryName(doc.Name),
                      Path.GetFileNameWithoutExtension(doc.Name))
                   + "_{0}.dwg";

                Dictionary<string, ObjectId> layouts = null;

                using (doc.LockDocument())
                {
                    using (Transaction tr = doc.TransactionManager.StartTransaction())
                    {
                        // Get the localized name of the model tab:
                        BlockTableRecord btr = (BlockTableRecord)
                           SymbolUtilityServices.GetBlockModelSpaceId(db)
                              .GetObject(OpenMode.ForRead);
                        Layout layout = (Layout)
                           btr.LayoutId.GetObject(OpenMode.ForRead);
                        string model = layout.LayoutName;
                        // Open the Layout dictionary:
                        IDictionary layoutDictionary = (IDictionary)
                           db.LayoutDictionaryId.GetObject(OpenMode.ForRead);
                        // Get the names and ids of all paper space layouts
                        // into a Dictionary<string,ObjectId>:
                        layouts = layoutDictionary.Cast<DictionaryEntry>()
                           .Where(e => ((string)e.Key) != model)
                           .ToDictionary(
                              e => (string)e.Key,
                              e => (ObjectId)e.Value);

                        tr.Commit();
                    }

                    /// Get the export layout 'engine':
                    Autodesk.AutoCAD.ExportLayout.Engine engine =
                       Autodesk.AutoCAD.ExportLayout.Engine.Instance();

                    using (new ManagedSystemVariable("CTAB"))
                    {
                        foreach (var entry in layouts)
                        {
                            string filename = string.Format(format, entry.Key);
                            editor.WriteMessage("\nExporting {0} => {1}\n", entry.Key, filename);
                            Application.SetSystemVariable("CTAB", entry.Key);
                            using (Database database = engine.ExportLayout(entry.Value))
                            {
                                if (engine.EngineStatus == AcExportLayout.ErrorStatus.Succeeded)
                                {
                                    database.SaveAs(filename, DwgVersion.Newest);
                                }
                                else
                                {
                                    editor.WriteMessage("\nExportLayout failed: ",
                                       engine.EngineStatus.ToString());
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
#if DEBUG
                editor.WriteMessage(ex.ToString());
#else
           throw ex;
#endif
            }
        }
    }

    public static class EditorInputExtensionMethods
    {
        public static PromptStatus Command(this Editor editor, params object[] args)
        {
            if (editor == null)
                throw new ArgumentNullException("editor");
            return runCommand(editor, args);
        }

        static Func<Editor, object[], PromptStatus> runCommand = GenerateRunCommand();

        static Func<Editor, object[], PromptStatus> GenerateRunCommand()
        {
            MethodInfo method = typeof(Editor).GetMethod("RunCommand",
               BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
            var instance = Expression.Parameter(typeof(Editor), "instance");
            var args = Expression.Parameter(typeof(object[]), "args");
            return Expression.Lambda<Func<Editor, object[], PromptStatus>>(
               Expression.Call(instance, method, args), instance, args)
                  .Compile();
        }
    }


    /// <summary>
    /// Automates saving/changing/restoring system variables
    /// </summary>

    public class ManagedSystemVariable : IDisposable
    {
        string name = null;
        object oldval = null;

        public ManagedSystemVariable(string name, object value)
            : this(name)
        {
            Application.SetSystemVariable(name, value);
        }

        public ManagedSystemVariable(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("name");
            this.name = name;
            this.oldval = Application.GetSystemVariable(name);
        }

        public void Dispose()
        {
            if (oldval != null)
            {
                object temp = oldval;
                oldval = null;
                Application.SetSystemVariable(name, temp);
            }
        }
    }
    public sealed class Layers : IExtensionApplication
    {
        public void Initialize()
        {
        }
        public void Terminate()
        {
        }
        //Добавляются слои для архитекторов гр.Смолина. Создано на основе файла-шаблона С44 A2013.dwt
        [CommandMethod("LayerArchitect")]
        public static void LayerArchitect()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            try
            {
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    LayerTable acLyrTbl;
                    acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead) as LayerTable;
                    string[] sLayerNames = new string[125];
                    sLayerNames[0] = "A-00-HELP";
                    sLayerNames[1] = "A-01-GRID";
                    sLayerNames[2] = "A-01-GRID-DIMM";
                    sLayerNames[3] = "A-01-GRID-MARK";
                    sLayerNames[4] = "A-02-EXPL";
                    sLayerNames[5] = "A-02-LABL";
                    sLayerNames[6] = "A-02-NOTE";
                    sLayerNames[7] = "A-02-REMK";
                    sLayerNames[8] = "A-02-REVS";
                    sLayerNames[9] = "A-02-TABL";
                    sLayerNames[10] = "A-02-TABL-CONN";
                    sLayerNames[11] = "A-02-TABL-DOOR";
                    sLayerNames[12] = "A-02-TABL-WNDW";
                    sLayerNames[13] = "A-02-TEXT";
                    sLayerNames[14] = "A-02-TEXT-NOPR";
                    sLayerNames[15] = "A-03-CONN-MARK";
                    sLayerNames[16] = "A-03-DETL-MARK";
                    sLayerNames[17] = "A-03-DOOR-MARK";
                    sLayerNames[18] = "A-03-ELVT-MARK";
                    sLayerNames[19] = "A-03-FIRE-MARK";
                    sLayerNames[20] = "A-03-FLOR-MARK";
                    sLayerNames[21] = "A-03-HOLE-AПT-MARK";
                    sLayerNames[22] = "A-03-HOLE-BK-MARK";
                    sLayerNames[23] = "A-03-HOLE-OB-MARK";
                    sLayerNames[24] = "A-03-HOLE-OBK-MARK";
                    sLayerNames[25] = "A-03-HOLE-CC-MARK";
                    sLayerNames[26] = "A-03-HOLE-ЭС-MARK";
                    sLayerNames[27] = "A-03-MARK";
                    sLayerNames[28] = "A-03-MARK-AREA";
                    sLayerNames[29] = "A-03-MARK-AREA-UNDLN";
                    sLayerNames[30] = "A-03-MARK-ELEV";
                    sLayerNames[31] = "A-03-MARK-ROOM";
                    sLayerNames[32] = "A-03-METL-MARK";
                    sLayerNames[33] = "A-03-OPEN-MARK";
                    sLayerNames[34] = "A-03-RAIL-MARK";
                    sLayerNames[35] = "A-03-SECT-MARK";
                    sLayerNames[36] = "A-03-STAI-MARK";
                    sLayerNames[37] = "A-03-WALL-MARK";
                    sLayerNames[38] = "A-03-WNDW-MARK";
                    sLayerNames[39] = "A-04-DIMM-DEMOL";
                    sLayerNames[40] = "A-04-DIMM-DETL";
                    sLayerNames[41] = "A-04-DIMM-AP";
                    sLayerNames[42] = "A-04-DIMM-KP";
                    sLayerNames[43] = "A-04-DIMM-ELEV";
                    sLayerNames[44] = "A-04-LEVEL";
                    sLayerNames[45] = "A-04-MEAS";
                    sLayerNames[46] = "A-05-DEKO";
                    sLayerNames[47] = "A-05-DOOR";
                    sLayerNames[48] = "A-05-DOOR-ARH";
                    sLayerNames[49] = "A-05-DOOR-R";
                    sLayerNames[50] = "A-05-ELVT";
                    sLayerNames[51] = "A-05-EQPM";
                    sLayerNames[52] = "A-05-FURN";
                    sLayerNames[53] = "A-05-OPEN";
                    sLayerNames[54] = "A-05-RAIL";
                    sLayerNames[55] = "A-05-RAMP";
                    sLayerNames[56] = "A-05-SANT";
                    sLayerNames[57] = "A-05-SCAC";
                    sLayerNames[58] = "A-05-STAGE";
                    sLayerNames[59] = "A-05-STAI";
                    sLayerNames[60] = "A-05-STGL";
                    sLayerNames[61] = "A-05-WNDW";
                    sLayerNames[62] = "A-05-WNDW-DETL";
                    sLayerNames[63] = "A-06-CEIL";
                    sLayerNames[64] = "A-06-CLMN";
                    sLayerNames[65] = "A-06-CNST";
                    sLayerNames[66] = "A-06-COAT";
                    sLayerNames[67] = "A-06-FINS";
                    sLayerNames[68] = "A-06-FLOR";
                    sLayerNames[69] = "A-06-FRAM";
                    sLayerNames[70] = "A-06-INSU";
                    sLayerNames[71] = "A-06-PART";
                    sLayerNames[72] = "A-06-PART-BREAK";
                    sLayerNames[73] = "A-06-SLAB";
                    sLayerNames[74] = "A-06-SLAB-HOLE";
                    sLayerNames[75] = "A-06-STAI";
                    sLayerNames[76] = "A-06-STEEL";
                    sLayerNames[77] = "A-06-WALL";
                    sLayerNames[78] = "A-06-WALL-BRICK";
                    sLayerNames[79] = "A-06-WALL-CONCR";
                    sLayerNames[80] = "A-06-WALL-DEMOL";
                    sLayerNames[81] = "A-06-WALL-EXST";
                    sLayerNames[82] = "A-06-WALL-HOLE";
                    sLayerNames[83] = "A-06-WALL-METL";
                    sLayerNames[84] = "A-06-WALL-NEW";
                    sLayerNames[85] = "A-06-WTPR";
                    sLayerNames[86] = "A-07-LINE";
                    sLayerNames[87] = "A-07-AUXIL";
                    sLayerNames[88] = "A-07-LINE-CEIL-ARC";
                    sLayerNames[89] = "A-07-LINE-OVRHD";
                    sLayerNames[90] = "A-07-LINE-PRGCT";
                    sLayerNames[91] = "A-07-RASTR";
                    sLayerNames[92] = "A-08-DETL-BRICK";
                    sLayerNames[93] = "A-08-DETL-DEKO";
                    sLayerNames[94] = "A-08-DETL-GLASS";
                    sLayerNames[95] = "A-09-MASKA";
                    sLayerNames[96] = "A-09-PATT";
                    sLayerNames[97] = "A-09-PATT-BRICK";
                    sLayerNames[98] = "A-09-PATT-CONCR";
                    sLayerNames[99] = "A-09-PATT-DEKO";
                    sLayerNames[100] = "A-09-PATT-DEMOL";
                    sLayerNames[101] = "A-09-PATT-FACIN";
                    sLayerNames[102] = "A-09-PATT-GLASS";
                    sLayerNames[103] = "A-09-PATT-INSUL";
                    sLayerNames[104] = "A-09-PATT-OTHER";
                    sLayerNames[105] = "A-09-PATT-STEEL";
                    sLayerNames[106] = "A-09-PATT-WOOD";
                    sLayerNames[107] = "A-09-PATT";
                    sLayerNames[108] = "A-10-AREA-ICUT";
                    sLayerNames[109] = "A-10-CEIL-ADD";
                    sLayerNames[110] = "A-10-ZONE-GREEN";
                    sLayerNames[111] = "A-10-ZONE-PAVE";
                    sLayerNames[112] = "A-10-ZONE-PLAN";
                    sLayerNames[113] = "A-10-ZONE-SECT";
                    sLayerNames[114] = "A-10-ZONE-TEXT";
                    sLayerNames[115] = "Defpoints";
                    sLayerNames[116] = "NO-PRINT";
                    sLayerNames[117] = "Z-DIMM";
                    sLayerNames[118] = "Z-PATT";
                    sLayerNames[119] = "Z-RAST";
                    sLayerNames[120] = "Z-SHBD";
                    sLayerNames[121] = "Z-STMP";
                    sLayerNames[122] = "Z-TABL";
                    sLayerNames[123] = "Z-TEXT";
                    sLayerNames[124] = "Z-VPRT";

                    //Массив цветовых палитр слоёв
                    Autodesk.AutoCAD.Colors.Color[] acColors = new Autodesk.AutoCAD.Colors.Color[125];
                    acColors[0] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 213);
                    acColors[1] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[2] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 142);
                    acColors[3] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 142);
                    acColors[4] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[5] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 230);
                    acColors[6] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 5);
                    acColors[7] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 232);
                    acColors[8] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 210);
                    acColors[9] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[10] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 111);
                    acColors[11] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 2);
                    acColors[12] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 133);
                    acColors[13] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 171);
                    acColors[14] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 60);
                    acColors[15] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 127);
                    acColors[16] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 102);
                    acColors[17] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 84);
                    acColors[18] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 195);
                    acColors[19] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 232);
                    acColors[20] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 73);
                    acColors[21] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 44);
                    acColors[22] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 122);
                    acColors[23] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 153);
                    acColors[24] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 173);
                    acColors[25] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 203);
                    acColors[26] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 227);
                    acColors[27] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[28] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 60);
                    acColors[29] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 61);
                    acColors[30] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 150);
                    acColors[31] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 115);
                    acColors[32] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 136);
                    acColors[33] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[34] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 204);
                    acColors[35] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 230);
                    acColors[36] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 51);
                    acColors[37] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 66);
                    acColors[38] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 144);
                    acColors[39] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[40] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 2);
                    acColors[41] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 153);
                    acColors[42] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 234);
                    acColors[43] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 1);
                    acColors[44] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 12);
                    acColors[45] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 126);
                    acColors[46] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 33);
                    acColors[47] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 51);
                    acColors[48] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 41);
                    acColors[49] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 230);
                    acColors[50] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 195);
                    acColors[51] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 201);
                    acColors[52] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 63);
                    acColors[53] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[54] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 123);
                    acColors[55] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 52);
                    acColors[56] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 163);
                    acColors[57] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 32);
                    acColors[58] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 21);
                    acColors[59] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 50);
                    acColors[60] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 141);
                    acColors[61] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 133);
                    acColors[62] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 141);
                    acColors[63] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 163);
                    acColors[64] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 135);
                    acColors[65] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[66] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 2);
                    acColors[67] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[68] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[69] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 111);
                    acColors[70] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 41);
                    acColors[71] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 32);
                    acColors[72] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 13);
                    acColors[73] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[74] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 15);
                    acColors[75] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 51);
                    acColors[76] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 181);
                    acColors[77] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 42);
                    acColors[78] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 42);
                    acColors[79] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 130);
                    acColors[80] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 2);
                    acColors[81] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 143);
                    acColors[82] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 213);
                    acColors[83] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 230);
                    acColors[84] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 53);
                    acColors[85] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 20);
                    acColors[86] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[87] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 71);
                    acColors[88] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[89] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 151);
                    acColors[90] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[91] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[92] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 22);
                    acColors[93] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 43);
                    acColors[94] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 121);
                    acColors[95] = Autodesk.AutoCAD.Colors.Color.FromRgb(38, 38, 38);
                    acColors[96] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[97] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 34);
                    acColors[98] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 126);
                    acColors[99] = Autodesk.AutoCAD.Colors.Color.FromRgb(144, 132, 111);
                    acColors[100] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[101] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[102] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 147);
                    acColors[103] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[104] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[105] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 146);
                    acColors[106] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 35);
                    acColors[107] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 32);
                    acColors[108] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 5);
                    acColors[109] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 106);
                    acColors[110] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 54);
                    acColors[111] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 240);
                    acColors[112] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 23);
                    acColors[113] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 214);
                    acColors[114] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 51);
                    acColors[115] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 103);
                    acColors[116] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 145);
                    acColors[117] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 242);
                    acColors[118] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[119] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[120] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[121] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[122] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[123] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 152);
                    acColors[124] = Autodesk.AutoCAD.Colors.Color.FromRgb(163, 224, 194);

                    //acColors[2] = Autodesk.AutoCAD.Colors.Color.FromNames("PANTONE Yellow 0131 C", "PANTONE(R) pastel coated");
                    string[] LayerDescription = new string[125];
                    LayerDescription[0] = "Рабочий слой";
                    LayerDescription[1] = "Сетка координационных осей";
                    LayerDescription[2] = "Сетка координационных осей Размеры";
                    LayerDescription[3] = "Сетка координационных осей Маркер";
                    LayerDescription[4] = "Экспликация";
                    LayerDescription[5] = "Условные обозначения";
                    LayerDescription[6] = "примечания";
                    LayerDescription[7] = "Замечания и уточнения";
                    LayerDescription[8] = "Изменения и ревизии";
                    LayerDescription[9] = "Таблицы и спецификации";
                    LayerDescription[10] = "Таблицы и спецификации перемычки";
                    LayerDescription[11] = "Таблицы и спецификации Двери";
                    LayerDescription[12] = "Таблицы и спецификации Окна";
                    LayerDescription[13] = "Текст";
                    LayerDescription[14] = "Текст Непечатаемый";
                    LayerDescription[15] = "Маркер Перемычка";
                    LayerDescription[16] = "Маркер Узлы и фрагменты";
                    LayerDescription[17] = "Маркер Двери";
                    LayerDescription[18] = "Маркер Лифты";
                    LayerDescription[19] = "Маркер Огнестойкость";
                    LayerDescription[20] = "Маркер Полы";
                    LayerDescription[21] = "Маркер Отверстие АПТ";
                    LayerDescription[22] = "Маркер Отверстие ВК";
                    LayerDescription[23] = "Маркер Отверстие ОВ";
                    LayerDescription[24] = "Маркер Отверстие ОВК";
                    LayerDescription[25] = "Маркер Отверстие СС";
                    LayerDescription[26] = "Маркер Отверстие ЭС";
                    LayerDescription[27] = "Маркер";
                    LayerDescription[28] = "Маркер Площади Помещения";
                    LayerDescription[29] = "Линия подчерка марки площади";
                    LayerDescription[30] = "Маркер Уровень";
                    LayerDescription[31] = "Марка номера помещения";
                    LayerDescription[32] = "Маркер Металл и другие констр. элементы";
                    LayerDescription[33] = "Маркер Проёма";
                    LayerDescription[34] = "Маркер Ограждения и другие арх. элементы";
                    LayerDescription[35] = "Маркер Разрез";
                    LayerDescription[36] = "Маркер Лестницы и лифты";
                    LayerDescription[37] = "Маркер Типы стен";
                    LayerDescription[38] = "Маркер Окна Витражи";
                    LayerDescription[39] = "Размеры - Демонтаж";
                    LayerDescription[40] = "Размеры - Детали";
                    LayerDescription[41] = "Размеры - АР";
                    LayerDescription[42] = "Размеры - КР";
                    LayerDescription[43] = "Отметка плана";
                    LayerDescription[44] = "Отметка уровня";
                    LayerDescription[45] = "Размеры - Обмеры";
                    LayerDescription[46] = "Декор";
                    LayerDescription[47] = "Двери";
                    LayerDescription[48] = "Двери архитекторы";
                    LayerDescription[49] = "Двери реставраторы";
                    LayerDescription[50] = "Лифты";
                    LayerDescription[51] = "Оборудование";
                    LayerDescription[52] = "Оборудование, мебель";
                    LayerDescription[53] = "Проем";
                    LayerDescription[54] = "Ограждения";
                    LayerDescription[55] = "Пандус";
                    LayerDescription[56] = "Сантехнические приборы";
                    LayerDescription[57] = "Помещения";
                    LayerDescription[58] = "Ступени";
                    LayerDescription[59] = "Лестницы";
                    LayerDescription[60] = "Витражи";
                    LayerDescription[61] = "Окна";
                    LayerDescription[62] = "Окна и витражи";
                    LayerDescription[63] = "Потолки";
                    LayerDescription[64] = "Колонны, пилоны";
                    LayerDescription[65] = "Различные конструкции";
                    LayerDescription[66] = "Кровля";
                    LayerDescription[67] = "Отделка";
                    LayerDescription[68] = "Полы";
                    LayerDescription[69] = "Балки, фермы";
                    LayerDescription[70] = "Утеплитель";
                    LayerDescription[71] = "Перегородка";
                    LayerDescription[72] = "Перегородка кирпич";
                    LayerDescription[73] = "Перекрытия";
                    LayerDescription[74] = "Перекрытие отверстие";
                    LayerDescription[75] = "Лестницы";
                    LayerDescription[76] = "Металлоконструкции";
                    LayerDescription[77] = "Стены";
                    LayerDescription[78] = "Стены кирпич";
                    LayerDescription[79] = "Стены железобетон";
                    LayerDescription[80] = "Стены демонтаж";
                    LayerDescription[81] = "Стены, Существующий";
                    LayerDescription[82] = "Стены отверстие";
                    LayerDescription[83] = "Стены, Усиление";
                    LayerDescription[84] = "Стены, Возводимый";
                    LayerDescription[85] = "Гидроизоляция";
                    LayerDescription[86] = "Линии";
                    LayerDescription[87] = "Линии вспомогательные";
                    LayerDescription[88] = "Своды";
                    LayerDescription[89] = "Линии над головой";
                    LayerDescription[90] = "Линии в проекции";
                    LayerDescription[91] = "Растровые изображения";
                    LayerDescription[92] = "Кирпич фасад";
                    LayerDescription[93] = "Штукатурка фасад";
                    LayerDescription[94] = "Стекло фасад";
                    LayerDescription[95] = "Маска";
                    LayerDescription[96] = "Штриховка";
                    LayerDescription[97] = "Штриховка кирпич";
                    LayerDescription[98] = "Штриховка железобетон";
                    LayerDescription[99] = "Штриховка декор";
                    LayerDescription[100] = "Штриховка демонтаж";
                    LayerDescription[101] = "Штриховка облицовка";
                    LayerDescription[102] = "Штриховка стекло";
                    LayerDescription[103] = "Штриховка утеплитель";
                    LayerDescription[104] = "Штриховка другие материалы";
                    LayerDescription[105] = "Штриховка металл";
                    LayerDescription[106] = "Штриховка дерево";
                    LayerDescription[107] = "Зона Площади помещений, На линии сечения";
                    LayerDescription[108] = "Зона подвесных потолков";
                    LayerDescription[109] = "Зона Озеленение";
                    LayerDescription[110] = "Зона Мощение";
                    LayerDescription[111] = "Зона План";
                    LayerDescription[112] = "Зона Сечение";
                    LayerDescription[113] = "Зона Текст";
                    LayerDescription[114] = "Default Non-Plotter Layer";
                    LayerDescription[115] = "Вспомогательный слой (не печатается)";
                    LayerDescription[116] = "Размеры";
                    LayerDescription[117] = "Примечания";
                    LayerDescription[118] = "Штриховка";
                    LayerDescription[119] = "Растровое изображение";
                    LayerDescription[120] = "Рамка";
                    LayerDescription[121] = "Штамп";
                    LayerDescription[122] = "Таблица";
                    LayerDescription[123] = "Текст";
                    LayerDescription[124] = "Видовой экран";

                    int nCnt = 0;
                    LayerTableRecord acLyrTblRec;
                    foreach (string sLayerName in sLayerNames)
                    {
                        if (acLyrTbl.Has(sLayerName) == false)
                        {
                            acLyrTblRec = new LayerTableRecord();
                            acLyrTblRec.Name = sLayerName;
                            if (acLyrTbl.IsWriteEnabled == false) acLyrTbl.UpgradeOpen();

                            acLyrTbl.Add(acLyrTblRec);
                            acTrans.AddNewlyCreatedDBObject(acLyrTblRec, true);
                        }
                        else
                        {
                            acLyrTblRec = acTrans.GetObject(acLyrTbl[sLayerName], OpenMode.ForWrite) as LayerTableRecord;
                        }
                        acLyrTblRec.Color = acColors[nCnt];

                        acLyrTblRec.Description = LayerDescription[nCnt];

                        nCnt += 1;
                    }
                    acTrans.Commit();
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }

        public static void CreateLayerStampAndFrame()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                Document acDoc = Application.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;

                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    LayerTable acLyrTbl;
                    acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead) as LayerTable;
                    string[] sLayerNames = new string[2];

                    sLayerNames[0] = "Z-TEXT";
                    sLayerNames[1] = "Z-STMP";

                    Autodesk.AutoCAD.Colors.Color[] acColors = new Autodesk.AutoCAD.Colors.Color[2];
                    acColors[0] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[1] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);

                    Autodesk.AutoCAD.DatabaseServices.LineWeight[] acLineWeight = new Autodesk.AutoCAD.DatabaseServices.LineWeight[2];

                    acLineWeight[0] = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight030;
                    acLineWeight[1] = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight050;

                    int nCnt = 0;
                    LayerTableRecord acLyrTblRec;

                    foreach (string sLayerName in sLayerNames)
                    {
                        if (acLyrTbl.Has(sLayerName) == false)
                        {
                            acLyrTblRec = new LayerTableRecord();
                            acLyrTblRec.Name = sLayerName;
                            if (acLyrTbl.IsWriteEnabled == false) acLyrTbl.UpgradeOpen();

                            acLyrTbl.Add(acLyrTblRec);
                            acTrans.AddNewlyCreatedDBObject(acLyrTblRec, true);
                        }
                        else
                        {
                            acLyrTblRec = acTrans.GetObject(acLyrTbl[sLayerName], OpenMode.ForWrite) as LayerTableRecord;
                        }
                        acLyrTblRec.Color = acColors[nCnt];
                        acLyrTblRec.LineWeight = acLineWeight[nCnt];
                        nCnt += 1;
                    }
                    acTrans.Commit();
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        //Добавляются слои для конструкторов создано на основе файла-шаблона С44 К(2013).dwt
        [CommandMethod("LayerConstruct")]
        public static void LayerConstruct()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                Document acDoc = Application.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;

                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    LayerTable acLyrTbl;
                    acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead) as LayerTable;
                    string[] sLayerNames = new string[76];
                    sLayerNames[0] = "S-AXIS";
                    sLayerNames[1] = "S-BEAM";
                    sLayerNames[2] = "S-BEAM-DIMM";
                    sLayerNames[3] = "S-BEAM-PATT";
                    sLayerNames[4] = "S-BEAM-TEXT";
                    sLayerNames[5] = "S-CLMN";
                    sLayerNames[6] = "S-CLMN-DIMM";
                    sLayerNames[7] = "S-CLMN-PATT";
                    sLayerNames[8] = "S-CLMN-TEXT";
                    sLayerNames[9] = "S-DETL";
                    sLayerNames[10] = "S-DETL-DIMM";
                    sLayerNames[11] = "S-DETL-PATT";
                    sLayerNames[12] = "S-DETL-TEXT";
                    sLayerNames[13] = "S-DIMM";
                    sLayerNames[14] = "S-FUND";
                    sLayerNames[15] = "S-FUND-DIMM";
                    sLayerNames[16] = "S-FUND-PATT";
                    sLayerNames[17] = "S-FUND-PILE";
                    sLayerNames[18] = "S-FUND-TEXT";
                    sLayerNames[19] = "S-GRID";
                    sLayerNames[20] = "S-GRID-ANNO";
                    sLayerNames[21] = "S-GRID-DIMM";
                    sLayerNames[22] = "S-HOLE";
                    sLayerNames[23] = "S-HOLE-DIMM";
                    sLayerNames[24] = "S-HOLE-ELCT";
                    sLayerNames[25] = "S-HOLE-ELCT-DIMM";
                    sLayerNames[26] = "S-HOLE-ELCT-TEXT";
                    sLayerNames[27] = "S-HOLE-HVAC";
                    sLayerNames[28] = "S-HOLE-HVAC-DIMM";
                    sLayerNames[29] = "S-HOLE-HVAC-TEXT";
                    sLayerNames[30] = "S-HOLE-TEXT";
                    sLayerNames[31] = "S-METL";
                    sLayerNames[32] = "S-METL-DETL";
                    sLayerNames[33] = "S-METL-DIMM";
                    sLayerNames[34] = "S-METL-TEXT";
                    sLayerNames[35] = "S-NOTE";
                    sLayerNames[36] = "S-PATT";
                    sLayerNames[37] = "S-RAIL";
                    sLayerNames[38] = "S-RBAR-DWGR";
                    sLayerNames[39] = "S-RBAR-DWVR";
                    sLayerNames[40] = "S-RBAR-DWZONE";
                    sLayerNames[41] = "S-RBAR-UPGR";
                    sLayerNames[42] = "S-RBAR-UPVR";
                    sLayerNames[43] = "S-RBAR-UPZONE";
                    sLayerNames[44] = "S-SECT";
                    sLayerNames[45] = "S-SIMB";
                    sLayerNames[46] = "S-SLAB";
                    sLayerNames[47] = "S-SLAB-DIMM";
                    sLayerNames[48] = "S-SLAB-EDGE";
                    sLayerNames[49] = "S-SLAB-PATT";
                    sLayerNames[50] = "S-SLAB-TEXT";
                    sLayerNames[51] = "S-SPAC";
                    sLayerNames[52] = "S-STAI";
                    sLayerNames[53] = "S-STAI-DIMM";
                    sLayerNames[54] = "S-STAI-FLIGHT";
                    sLayerNames[55] = "S-STAI-STEP";
                    sLayerNames[56] = "S-STAI-TEXT";
                    sLayerNames[57] = "S-SWLL";
                    sLayerNames[58] = "S-SWLL-DIMM";
                    sLayerNames[59] = "S-SWLL-PATT";
                    sLayerNames[60] = "S-SWLL-TEXT";
                    sLayerNames[61] = "S-TABL";
                    sLayerNames[62] = "S-TEXT";
                    sLayerNames[63] = "S-WALL";
                    sLayerNames[64] = "S-WALL-PATT";
                    sLayerNames[65] = "S-WALL-TEXT";
                    sLayerNames[66] = "Vport";
                    sLayerNames[67] = "Z-DIMM";
                    sLayerNames[68] = "Z-NOTE";
                    sLayerNames[69] = "Z-RAST";
                    sLayerNames[70] = "Z-SHBD";
                    sLayerNames[71] = "Z-STMP";
                    sLayerNames[72] = "Z-TABL";
                    sLayerNames[73] = "Z-TEXT";
                    sLayerNames[74] = "Z-VPRT";
                    sLayerNames[75] = "Defpoints";

                    //Массив цветовых палитр слоёв
                    Autodesk.AutoCAD.Colors.Color[] acColors = new Autodesk.AutoCAD.Colors.Color[76];
                    acColors[0] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 213);
                    acColors[1] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[2] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 142);
                    acColors[3] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 142);
                    acColors[4] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[5] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 230);
                    acColors[6] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 5);
                    acColors[7] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 232);
                    acColors[8] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 210);
                    acColors[9] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[10] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 111);
                    acColors[11] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 2);
                    acColors[12] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 133);
                    acColors[13] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 171);
                    acColors[14] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 60);
                    acColors[15] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 127);
                    acColors[16] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 102);
                    acColors[17] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 84);
                    acColors[18] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 195);
                    acColors[19] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 232);
                    acColors[20] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 73);
                    acColors[21] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 44);
                    acColors[22] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 122);
                    acColors[23] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 153);
                    acColors[24] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 173);
                    acColors[25] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 203);
                    acColors[26] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 227);
                    acColors[27] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[28] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 60);
                    acColors[29] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 61);
                    acColors[30] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 150);
                    acColors[31] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 115);
                    acColors[32] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 136);
                    acColors[33] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[34] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 204);
                    acColors[35] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 230);
                    acColors[36] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 51);
                    acColors[37] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 66);
                    acColors[38] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 144);
                    acColors[39] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[40] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 2);
                    acColors[41] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 153);
                    acColors[42] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 234);
                    acColors[43] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 1);
                    acColors[44] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 12);
                    acColors[45] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 126);
                    acColors[46] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 33);
                    acColors[47] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 51);
                    acColors[48] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 41);
                    acColors[49] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 230);
                    acColors[50] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 195);
                    acColors[51] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 201);
                    acColors[52] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 63);
                    acColors[53] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[54] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 123);
                    acColors[55] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 52);
                    acColors[56] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 163);
                    acColors[57] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 32);
                    acColors[58] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 21);
                    acColors[59] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 50);
                    acColors[60] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 141);
                    acColors[61] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 133);
                    acColors[62] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 141);
                    acColors[63] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 163);
                    acColors[64] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 135);
                    acColors[65] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[66] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 2);
                    acColors[67] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[68] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[69] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 111);
                    acColors[70] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 41);
                    acColors[71] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 32);
                    acColors[72] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 13);
                    acColors[73] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[74] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 15);
                    acColors[75] = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 51);

                    //acColors[2] = Autodesk.AutoCAD.Colors.Color.FromNames("PANTONE Yellow 0131 C", "PANTONE(R) pastel coated");
                    string[] LayerDescription = new string[76];
                    LayerDescription[0] = "Оси конструкций";
                    LayerDescription[1] = "Балка";
                    LayerDescription[2] = "Балка - размеры";
                    LayerDescription[3] = "Балка - штриховка";
                    LayerDescription[4] = "Балка - текст";
                    LayerDescription[5] = "Колонна";
                    LayerDescription[6] = "Колонна - размеры";
                    LayerDescription[7] = "Колонна - штриховка";
                    LayerDescription[8] = "Колонна - текст";
                    LayerDescription[9] = "Узлы/Детали";
                    LayerDescription[10] = "Узлы - размеры";
                    LayerDescription[11] = "Узлы и штриховка";
                    LayerDescription[12] = "Узлы - текст";
                    LayerDescription[13] = "Размеры";
                    LayerDescription[14] = "Фундамент";
                    LayerDescription[15] = "Фундамент - размеры";
                    LayerDescription[16] = "Фундамент - штриховка";
                    LayerDescription[17] = "Фундамент - сваи забивные и буронабивные";
                    LayerDescription[18] = "Фундамент - текст";
                    LayerDescription[19] = "Сетка координационных осей";
                    LayerDescription[20] = "Выносные элементы сетки координат";
                    LayerDescription[21] = "Оси - размеры";
                    LayerDescription[22] = "Отверстия";
                    LayerDescription[23] = "Отверстия - размеры";
                    LayerDescription[24] = "Отверстия для электрики";
                    LayerDescription[25] = "Отверстия для электрики - размеры";
                    LayerDescription[26] = "Отверстия для электрики - текст";
                    LayerDescription[27] = "Отверсия ОВ и ВК";
                    LayerDescription[28] = "Отверсия ОВ и ВК - размеры";
                    LayerDescription[29] = "Отверсия ОВ и ВК - текст";
                    LayerDescription[30] = "Отверсия - текст";
                    LayerDescription[31] = "Металлы";
                    LayerDescription[32] = "Металлы/узлы";
                    LayerDescription[33] = "Металл - размеры";
                    LayerDescription[34] = "Металл - текст";
                    LayerDescription[35] = "Примечания, ссылки";
                    LayerDescription[36] = "Штриховки";
                    LayerDescription[37] = "Ограждения";
                    LayerDescription[38] = "Армирование нижнее горизонтальное";
                    LayerDescription[39] = "Армирование нижнее вертикальное";
                    LayerDescription[40] = "Армирование нижнее ЗОНА";
                    LayerDescription[41] = "Армирование верхнее горизонтальное";
                    LayerDescription[42] = "Армирование верхнее вертикальное";
                    LayerDescription[43] = "Армирование верхнее ЗОНА";
                    LayerDescription[44] = "Разрез/Сечение";
                    LayerDescription[45] = "Символы, обозначения";
                    LayerDescription[46] = "Плиты";
                    LayerDescription[47] = "Плиты - размеры";
                    LayerDescription[48] = "Контур плиты";
                    LayerDescription[49] = "Плиты - штриховка";
                    LayerDescription[50] = "Плиты - текст";
                    LayerDescription[51] = "Помещения";
                    LayerDescription[52] = "Лестница";
                    LayerDescription[53] = "Лестница - размеры";
                    LayerDescription[54] = "Лестница - марш";
                    LayerDescription[55] = "Лестница - ступени";
                    LayerDescription[56] = "Лестница - текст";
                    LayerDescription[57] = "Стена несущая";
                    LayerDescription[58] = "Стена несущая - размеры";
                    LayerDescription[59] = "Стена несущая - штриховка";
                    LayerDescription[60] = "Стена несущая - текст";
                    LayerDescription[61] = "Спецификации и таблицы";
                    LayerDescription[62] = "Общие пояснения и спецификации";
                    LayerDescription[63] = "Стена конструктивная ненесущая или диафрагма";
                    LayerDescription[64] = "Стена конструктивная ненесущая - штриховка";
                    LayerDescription[65] = "Стена конструктивная - текст";
                    LayerDescription[66] = "";
                    LayerDescription[67] = "Размеры";
                    LayerDescription[68] = "Примечания";
                    LayerDescription[69] = "Растровое изображение";
                    LayerDescription[70] = "Рамка";
                    LayerDescription[71] = "Штамп";
                    LayerDescription[72] = "Таблица";
                    LayerDescription[73] = "Текст";
                    LayerDescription[74] = "Видовой экран";
                    LayerDescription[75] = "";

                    int nCnt = 0;
                    LayerTableRecord acLyrTblRec;
                    foreach (string sLayerName in sLayerNames)
                    {
                        if (acLyrTbl.Has(sLayerName) == false)
                        {
                            acLyrTblRec = new LayerTableRecord();
                            acLyrTblRec.Name = sLayerName;
                            if (acLyrTbl.IsWriteEnabled == false) acLyrTbl.UpgradeOpen();

                            acLyrTbl.Add(acLyrTblRec);
                            acTrans.AddNewlyCreatedDBObject(acLyrTblRec, true);
                        }
                        else
                        {
                            acLyrTblRec = acTrans.GetObject(acLyrTbl[sLayerName], OpenMode.ForWrite) as LayerTableRecord;
                        }
                        acLyrTblRec.Color = acColors[nCnt];
                        acLyrTblRec.Description = LayerDescription[nCnt];
                        nCnt += 1;
                    }
                    acTrans.Commit();
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exeption caught" + ex);
            }
        }
    }
    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    /// Класс CustomVar предназначен для установки внутренних системных переменных и дополнительных настроек в оптималный режим работы
    public sealed class CustomVar : IExtensionApplication
    {
        public void Initialize()
        {}
        public void Terminate()
        {}

        [CommandMethod("CustomVar")]
        public void StartVariables()
        {
            Variables();
        }
        private void Variables()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                Application.SetSystemVariable("zoomfactor", 100); 
                Application.SetSystemVariable("xrefnotify", 0);
                Application.SetSystemVariable("whipthread", 3);
                Application.SetSystemVariable("whiparc", 0);
                Application.SetSystemVariable("vtfps", 1);
                //Application.SetSystemVariable("treemax", 15000000);
                //Application.SetSystemVariable("treedeprh", 3020);
                //Application.SetSystemVariable("savetime", 10);
                //Application.SetSystemVariable("ribbonbgload", 1);
                //Application.SetSystemVariable("regenmode", 1);
                //Application.SetSystemVariable("rastertreshold", 60);
                //Application.SetSystemVariable("proxygraphics", 1);
                //Application.SetSystemVariable("plquiet", 1);
                Application.SetSystemVariable("openpartial", 0);
                Application.SetSystemVariable("maxactvp", 16);
                Application.SetSystemVariable("lockui", 8);
                //Application.SetSystemVariable("layoutregenctrl", 1);
                //Application.SetSystemVariable("largeobjectsupport", 1);
                //Application.SetSystemVariable("isavepercent", 0);
                //Application.SetSystemVariable("isavebak", 1);
                //Application.SetSystemVariable("intelligentupdate", 20);
                //Application.SetSystemVariable("imagehlt", 0);
                Application.SetSystemVariable("highlight", 1);
                Application.SetSystemVariable("hideprecision", 0);
                Application.SetSystemVariable("gripobjlimit", 2);
                Application.SetSystemVariable("dragp1", 10000);
                Application.SetSystemVariable("dragp2", 10);
                Application.SetSystemVariable("cmdinputhistorymax", 10);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("Exception caught " + ex);
            }       
        }
    }
    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    ///Класс help предназначен для вывода справочной информации
    public sealed class HELP : IExtensionApplication
    {
        public void Initialize()
        {}
        public void Terminate()
        {}
        /////////////////////////////////////////////////////////////////////////////////////////////////////////////
        // Команда, выдающая в консоль информацию о разработчике и список используемых комманд
        [CommandMethod("HelpCommands")]
        public void StartCreator()
        {
            this.Creator();
        }
        private void Creator()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                editor.WriteMessage("_This plugin to developed by Lev Kozhaev aka Sfinks in 2013._" + "\n");
                editor.WriteMessage("_A1H" + "\n");
                editor.WriteMessage("_A1V" + "\n");
                editor.WriteMessage("_A2H" + "\n");
                editor.WriteMessage("_A2V" + "\n");
                editor.WriteMessage("_A3H" + "\n");
                editor.WriteMessage("_A3V" + "\n");
                editor.WriteMessage("_A3H3" + "\n");
                editor.WriteMessage("_A4H" + "\n");
                editor.WriteMessage("_A4V" + "\n");
                editor.WriteMessage("_A4H3" + "\n");
                editor.WriteMessage("_A4H4" + "\n");

                editor.WriteMessage("_StampAR" + "\n");
                editor.WriteMessage("_StampK" + "\n");
                editor.WriteMessage("_StampKAB" + "\n");

                editor.WriteMessage("_RefreshXREF" + "\n");
                editor.WriteMessage("_FindImages" + "\n");
                editor.WriteMessage("_AddEvent" + "\n");
                editor.WriteMessage("_XREFUpdate" + "\n");
                editor.WriteMessage("_TDMSXREFIMG" + "\n");
                editor.WriteMessage("_TDMSXREFDWG" + "\n");

                editor.WriteMessage("_titles" + "\n");
                editor.WriteMessage("_UA" + "\n");

                editor.WriteMessage("_TDMSRIBBON" + "\n");

                editor.WriteMessage("_SaveActiveDrawing" + "\n");
                editor.WriteMessage("_SaveAndClose" + "\n");
                editor.WriteMessage("_CloseAndDiscard" + "\n");

                editor.WriteMessage("_RegisterTDMSApp" + "\n");
                editor.WriteMessage("_UnregisterTDMSApp" + "\n");

                editor.WriteMessage("ETR" + "\n");
                editor.WriteMessage("LayerArchitect" + "\n");
                editor.WriteMessage("LayerConstruct" + "\n");
                editor.WriteMessage("CustomVar" + "\n");
                editor.WriteMessage("HelpCommands" + "\n");
                editor.WriteMessage("OpenHelp" + "\n");
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
    }
    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    public sealed class Creator : IExtensionApplication
    {
        public void Initialize()
        { }
        public void Terminate()
        { }
        
        //Возвращение названия Layout и распечатка в LineText   
        public static void NameDrawingStamp(double x, double y, double z)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                Document acDoc = Application.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.PaperSpace], OpenMode.ForWrite) as BlockTableRecord;

                    LayoutManager acLayoutMgr;
                    acLayoutMgr = LayoutManager.Current;
                    Layout acLayout;
                    acLayout = acTrans.GetObject(acLayoutMgr.GetLayoutId(acLayoutMgr.CurrentLayout), OpenMode.ForRead) as Layout;

                    MText acMText = new MText();
                    acMText.SetDatabaseDefaults();
                    acMText.Location = new Point3d(x, y, z);

                    acMText.Width = 68;
                    acMText.TextHeight = 2.25;

                    acMText.Contents = "\\pxqc;" + System.Convert.ToString(acLayout.LayoutName);
                    acMText.Layer = "Z-TEXT";
                    acBlkTblRec.AppendEntity(acMText);
                    acTrans.AddNewlyCreatedDBObject(acMText, true);
                    acTrans.Commit();
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        // OOO "Архитектурное бюро "Студия 44"
        public static void NameStudio(double x, double y, double z)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                Document acDoc = Application.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.PaperSpace], OpenMode.ForWrite) as BlockTableRecord;

                    MText acMText = new MText();
                    acMText.SetDatabaseDefaults();
                    acMText.Location = new Point3d(x, y, z);
                    acMText.Width = 36.838;
                    acMText.Height = 10.034;
                    acMText.TextHeight = 2.5;
                    acMText.LineSpaceDistance = 4.167;
                    acMText.LineSpacingFactor = 1.0;

                    acMText.Contents = "\\pxqc;" + "\"OOO Архитектурное бюро \"Студия 44\"";
                    acMText.Layer = "Z-TEXT";
                    acBlkTblRec.AppendEntity(acMText);
                    acTrans.AddNewlyCreatedDBObject(acMText, true);
                    acTrans.Commit();
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        //[CommandMethod("RefreshAttribut", CommandFlags.NoBlockEditor)]
        public static void RefreshAttribut(string blockName, string text)
        {
            Database dbCurrent = Application.DocumentManager.MdiActiveDocument.Database;
            Editor edCurrent = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                using (Transaction trAdding = dbCurrent.TransactionManager.StartTransaction())
                {
                    BlockTable btTable = (BlockTable)trAdding.GetObject(dbCurrent.BlockTableId, OpenMode.ForRead);
                    AttributeDefinition adAttr = new AttributeDefinition();

                    try
                    {
                        if (btTable.Has(blockName))
                            adAttr.TextString = text;
                    }
                    catch (System.Exception ex)
                    {
                        edCurrent.WriteMessage("\nInvalid block name." + ex);
                    }
                    trAdding.Commit();
                }
            }
            catch (System.Exception ex)
            {
                edCurrent.WriteMessage("\n Exception caught" + ex);
            }
        }
        //Создаём атрибут штампа, создавать можно как напрямую, так и через AddAttribute(там происходит запрос значений аттрибутов из ТДМС)
        public static void CreateStampAtribut(double x, double y, string text, double height, double widthFactor, double rotate, double oblique, string attrName, int blockName)
        {
            Database dbCurrent = Application.DocumentManager.MdiActiveDocument.Database;
            Editor edCurrent = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                    using (Transaction trAdding = dbCurrent.TransactionManager.StartTransaction())
                    {
                        BlockTable btTable = (BlockTable)trAdding.GetObject(dbCurrent.BlockTableId, OpenMode.ForRead);
                        string strBlockName = "ATTRBLK" + System.Convert.ToString(blockName);
                        try
                        {
                            if (btTable.Has(strBlockName))
                                edCurrent.WriteMessage("\n A block with this name already exist.");
                        }
                        catch
                        {
                            edCurrent.WriteMessage("\n Invalid block name.");
                        }

                        AttributeDefinition adAttr = new AttributeDefinition();
                        
                        adAttr.Position = new Point3d(x, y, 0);
                        adAttr.WidthFactor = widthFactor;
                        adAttr.Height = height;
                        adAttr.Rotation = rotate;
                        adAttr.Oblique = oblique;
                        adAttr.Tag = attrName;

                        BlockTableRecord btrRecord = new BlockTableRecord();
                        btrRecord.Name = strBlockName;
                        btTable.UpgradeOpen();

                        ObjectId btrID = btTable.Add(btrRecord);

                        trAdding.AddNewlyCreatedDBObject(btrRecord, true);

                        btrRecord.AppendEntity(adAttr);
                        trAdding.AddNewlyCreatedDBObject(adAttr, true);

                        BlockTableRecord btrPapperSpace = (BlockTableRecord)trAdding.GetObject(btTable[BlockTableRecord.PaperSpace], OpenMode.ForWrite);

                        BlockReference brRefBlock = new BlockReference(Point3d.Origin, btrID);

                        btrPapperSpace.AppendEntity(brRefBlock);
                        trAdding.AddNewlyCreatedDBObject(brRefBlock, true);

                        //задаём значение атрибута
                        AttributeReference arAttr = new AttributeReference();
                        arAttr.SetAttributeFromBlock(adAttr, brRefBlock.BlockTransform);
                       
                        arAttr.TextString = text;
                        arAttr.Layer = "Z-TEXT";

                        brRefBlock.AttributeCollection.AppendAttribute(arAttr);

                        trAdding.AddNewlyCreatedDBObject(arAttr, true);

                        trAdding.Commit();
                    }
            }
            catch (System.Exception ex)
            {
                edCurrent.WriteMessage("\n Exception caught " + ex);
            }
        }
        public static void CreateMultilineStampAtributNameDrawing(double x, double y, double height, double width, double widthFactor, double rotate, double oblique, string attrName, int blockName)
        {
            var docCurrent = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var dbCurrent = docCurrent.Database;
            var edCurrent = docCurrent.Editor;
            string text;
            BlockTableRecord btrRecord = null;
            ObjectId btrID = ObjectId.Null;
            try
            {
                docCurrent.TransactionManager.EnableGraphicsFlush(true);
                using (Transaction trAdding = dbCurrent.TransactionManager.StartTransaction())
                {
                    LayoutManager acLayoutMgr;
                    acLayoutMgr = LayoutManager.Current;
                    Layout acLayout;
                    acLayout = trAdding.GetObject(acLayoutMgr.GetLayoutId(acLayoutMgr.CurrentLayout), OpenMode.ForRead) as Layout;

                    text = "\\pxqc;" + System.Convert.ToString(acLayout.LayoutName);

                    BlockTable btTable = (BlockTable)trAdding.GetObject(dbCurrent.BlockTableId, OpenMode.ForRead);
                    string strBlockName = "ATTRBLK" + System.Convert.ToString(blockName);
                    try
                    {
                        if (btTable.Has(strBlockName))
                            edCurrent.WriteMessage("\n A block with this name already exist.");
                        else
                        {
                            AttributeDefinition adAttr = new AttributeDefinition();
                            adAttr.Position = new Point3d(x, y, 0);
                            adAttr.WidthFactor = widthFactor;
           
                            adAttr.Height = height;
                            adAttr.Rotation = rotate;
                            adAttr.Oblique = oblique;
                            adAttr.Justify = AttachmentPoint.MiddleCenter;
                            adAttr.Tag = attrName;
                            btTable.UpgradeOpen();
                            btrRecord = new BlockTableRecord();
                            btrRecord.Name = strBlockName;
                            btTable.Add(btrRecord);
                            trAdding.AddNewlyCreatedDBObject(btrRecord, true);
                            btTable.DowngradeOpen();
                            adAttr.IsMTextAttributeDefinition = true;
                            adAttr.MTextAttributeDefinition.Contents = text;
                            adAttr.MTextAttributeDefinition.Width = width;
                            btrRecord.AppendEntity(adAttr);
                            trAdding.AddNewlyCreatedDBObject(adAttr, true);
                        }
                    }
                    catch
                    {
                        edCurrent.WriteMessage("\n Invalid block name.");
                    }
                    BlockTableRecord btrPapperSpace = (BlockTableRecord)trAdding.GetObject(btTable[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
                    BlockReference brRefBlock = new BlockReference(new Point3d(x, y, 0), btrRecord.ObjectId);
                    btrPapperSpace.AppendEntity(brRefBlock);
                    trAdding.AddNewlyCreatedDBObject(brRefBlock, true);
                    //задаём значение атрибута
                    foreach (ObjectId id in btrRecord)
                    {
                        DBObject obj = id.GetObject(OpenMode.ForRead);
                        AttributeDefinition adAttr = obj as AttributeDefinition;
                        if ((adAttr != null) && (!adAttr.Constant))
                        {
                            AttributeReference arAttr = new AttributeReference();
                            arAttr.SetAttributeFromBlock(adAttr, brRefBlock.BlockTransform);
                            if (arAttr.IsMTextAttribute)
                            {
                                arAttr.Layer = "Z-TEXT";

                                MText mText = arAttr.MTextAttribute;
                                mText.Width = width;
                                mText.Contents = text;

                                arAttr.MTextAttribute = mText;
                                //arAttr.TextString = mText;
                                arAttr.MTextAttribute.Width = 118;
                                arAttr.Justify = AttachmentPoint.MiddleCenter;
                                arAttr.UpdateMTextAttribute();
                                brRefBlock.RecordGraphicsModified(true);
                            }
                            else
                            {
                                // add usual attribute reference here
                            }
                            brRefBlock.AttributeCollection.AppendAttribute(arAttr);
                            trAdding.AddNewlyCreatedDBObject(arAttr, true);
                        }
                        trAdding.TransactionManager.QueueForGraphicsFlush();
                    }
                    docCurrent.TransactionManager.FlushGraphics();
                    trAdding.Commit();
                }
            }
            catch (System.Exception ex)
            {
                edCurrent.WriteMessage("\n{0}\n{1}", ex.Message, ex.StackTrace);
            }
            finally
            {
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\n\t---\tSee results\t---\n");
            }
        }
        public static void CreateMultilineStampAtribut(double x, double y, string text, double height, double width, double widthFactor, double rotate, double oblique, string attrName, int blockName)
        {
            var docCurrent = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var dbCurrent = docCurrent.Database;
            var edCurrent = docCurrent.Editor;
            BlockTableRecord btrRecord = null;
            ObjectId btrID = ObjectId.Null;
            try
            {
                docCurrent.TransactionManager.EnableGraphicsFlush(true);
                using (Transaction trAdding = dbCurrent.TransactionManager.StartTransaction())
                {
                    BlockTable btTable = (BlockTable)trAdding.GetObject(dbCurrent.BlockTableId, OpenMode.ForRead);
                    string strBlockName = "ATTRBLK" + System.Convert.ToString(blockName);
                    try
                    {
                        if (btTable.Has(strBlockName))
                            edCurrent.WriteMessage("\n A block with this name already exist.");
                        else
                        {
                            AttributeDefinition adAttr = new AttributeDefinition();
                            adAttr.Position = new Point3d(x, y, 0);
                            adAttr.WidthFactor = widthFactor;
                            
                            adAttr.Height = height;
                            adAttr.Rotation = rotate;
                            adAttr.Oblique = oblique;
                            adAttr.Justify = AttachmentPoint.MiddleCenter;
                            adAttr.Tag = attrName;
                            btTable.UpgradeOpen();
                            btrRecord = new BlockTableRecord();
                            btrRecord.Name = strBlockName;
                            btTable.Add(btrRecord);
                            trAdding.AddNewlyCreatedDBObject(btrRecord, true);
                            btTable.DowngradeOpen();
                            adAttr.IsMTextAttributeDefinition = true;
                            adAttr.MTextAttributeDefinition.Contents = text;
                            adAttr.MTextAttributeDefinition.Width = width;
                            btrRecord.AppendEntity(adAttr);
                            trAdding.AddNewlyCreatedDBObject(adAttr, true);
                        }
                    }
                    catch
                    {
                        edCurrent.WriteMessage("\n Invalid block name.");
                    }
                    BlockTableRecord btrPapperSpace = (BlockTableRecord)trAdding.GetObject(btTable[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
                    BlockReference brRefBlock = new BlockReference(new Point3d(x, y, 0), btrRecord.ObjectId);
                    btrPapperSpace.AppendEntity(brRefBlock);
                    trAdding.AddNewlyCreatedDBObject(brRefBlock, true);
                    //задаём значение атрибута
                    foreach (ObjectId id in btrRecord)
                    {
                        DBObject obj = id.GetObject(OpenMode.ForRead);
                        AttributeDefinition adAttr = obj as AttributeDefinition;
                        if ((adAttr != null) && (!adAttr.Constant))
                        {
                            AttributeReference arAttr = new AttributeReference();
                            arAttr.SetAttributeFromBlock(adAttr, brRefBlock.BlockTransform);
                            if (arAttr.IsMTextAttribute)
                            {
                                arAttr.Layer = "Z-TEXT";

                                MText mText = arAttr.MTextAttribute;
                                mText.Width = width;
                                mText.Contents = text;
                                
                                arAttr.MTextAttribute = mText;
                                //arAttr.TextString = mText;
                                arAttr.MTextAttribute.Width = width;
                                arAttr.Justify = AttachmentPoint.MiddleCenter;
                                arAttr.UpdateMTextAttribute();
                                brRefBlock.RecordGraphicsModified(true);
                            }
                            else
                            {
                                // add usual attribute reference here
                            }
                            brRefBlock.AttributeCollection.AppendAttribute(arAttr);
                            trAdding.AddNewlyCreatedDBObject(arAttr, true);
                        }
                        trAdding.TransactionManager.QueueForGraphicsFlush();
                    }
                    docCurrent.TransactionManager.FlushGraphics();
                    trAdding.Commit();
                }
            }
            catch (System.Exception ex)
            {
                edCurrent.WriteMessage("\n{0}\n{1}", ex.Message, ex.StackTrace);
            }
            finally
            {
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\n\t---\tSee results\t---\n");
            }
        }
        //Создаём текст в штампе
        public static void CreateTextStamp(double x, double y, string text, double height, double widthFactor, double rotate, double oblique)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                Document acDoc = Application.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;

                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.PaperSpace], OpenMode.ForWrite) as BlockTableRecord;

                    DBText acText = new DBText();
                    acText.SetDatabaseDefaults();
                    acText.Position = new Point3d(x, y, 0);
                    acText.Height = height;
                    acText.Rotation = rotate;
                    acText.WidthFactor = widthFactor;
                    acText.TextString = text;
                    acText.Oblique = oblique;
                    acText.Layer = "Z-TEXT";
                    acBlkTblRec.AppendEntity(acText);
                    acTrans.AddNewlyCreatedDBObject(acText, true);

                    acTrans.Commit();
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        //генерация обычного текста
        public static void CreateText(string nameObject, float x, float y, string _area)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                // получаем актуальный документ (тот который открыт в данный момент времени) и базу данных
                Document acDoc = Application.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;
                // Начало транзакций
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl; // Открывается блок таблиц для чтения
                    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord acBlkTblRec; // Открыть блок таблиц для чтения и записи в модели пространства
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    DBText acText = new DBText();
                    acText.SetDatabaseDefaults();
                    acText.Position = new Point3d(x - 1, y + 4, 0);
                    acText.Height = 2.0;
                    acText.TextString = System.Convert.ToString(nameObject);
                    acBlkTblRec.AppendEntity(acText);
                    acTrans.AddNewlyCreatedDBObject(acText, true);

                    DBText acTextArea = new DBText();
                    acTextArea.SetDatabaseDefaults();
                    acTextArea.Position = new Point3d(x - 1, y, 0);
                    acTextArea.Height = 2.5;
                    acTextArea.TextString = System.Convert.ToString(_area);

                    acBlkTblRec.AppendEntity(acTextArea);
                    acTrans.AddNewlyCreatedDBObject(acTextArea, true);

                    acTrans.Commit();
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
    }
    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    //Класс, содержащий методы и конструкторы для генерации рамок и штампов
    public sealed class FrameStamp : IExtensionApplication
    {
        static int Rand(int min, int max)
        {
            int result = 0;
            Random r = new Random();
            return result = r.Next(min, max + 1);
        }
        static int blk = Rand(0, 10000);

        private const string rab = "Разработал";
        private const string search = "Проверил";
        private const string GIP = "ГИП";
        private const string GAP = "ГAП";
        private const string normal = "Н.контроль";
        private const string IZM = "Изм.";
        private const string coluch = "Кол.уч.";
        private const string Paper = "Лист";
        private const string Ndoc = "N док.";
        private const string signature = "Подпись";
        private const string date = "Дата";
        private const string stage = "Стадия";
        private const string Papers = "Листов";
        private const string rukGroup = "Рук. группы";
        private const string glConstPr = "Гл.констр.пр.";
        private const string glConstAB = "Гл.констр.\"AБ\"";

        //Атрибуты штампа
        private const string nameAtrCode = "A_OBOZN_DOC"; //шифр объекта
        private const string nameAtrAddress = "GETOPADDRESS"; //адрес
        private const string nameAtrProjectObj = "A_DESIGN_OBJ_REF"; //название

        private const string AttNameDrawing = "ATT_NAME_DRAWING"; //многострочный атрибут для ввода наименования чертежа, в дальнейшем используется для именования листов после разбивки.

        private const string nameAtrList = "A_PAGE_NUM";
        private const string nameAtrLists = "A_PAGE_COUNT";

        private const string moduleName = "CMD_SYSLIB";
        private const string _GDFOFA = "GETDATAFROMOBJFORACAD";
        private const string _GDFARTFA = "GETDATAFROMACTIVEROUTETABLEFORACAD";
        private const string _GDFSPFA = "GETDATAFROMSYSPROPSFORACAD";

        public void Initialize(){}
        public void Terminate(){}

        //добавляем заливку в архитектурные штампы
        private void AddSolid(int height, int width, double pointX, double pointY)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                Document acDoc = Application.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    double fontSize = 0;
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.PaperSpace], OpenMode.ForWrite) as BlockTableRecord;
                    int roomNumPoint = 0;
                    Polyline Poly = new Polyline();
                    Polyline InContur = new Polyline();

                    Poly.SetDatabaseDefaults();

                    //Нижний контур 
                    Poly.AddVertexAt(roomNumPoint, new Point2d(pointX - 5, pointY + 5), 0, fontSize, fontSize);
                    Poly.AddVertexAt(++roomNumPoint, new Point2d(pointX - width + 20, pointY + 5), 0, fontSize, fontSize);
                    Poly.AddVertexAt(++roomNumPoint, new Point2d(pointX - width + 20, pointY + 20), 0, fontSize, fontSize);
                    Poly.AddVertexAt(++roomNumPoint, new Point2d(pointX - 5, pointY + 20), 0, fontSize, fontSize);
                    Poly.Closed = true;
                    Poly.Layer = "Z-STMP";
                    acBlkTblRec.AppendEntity(Poly);
                    acTrans.AddNewlyCreatedDBObject(Poly, true);

                    //Верхний контур
                    roomNumPoint = 0;
                    fontSize = 0;
                    InContur.AddVertexAt(roomNumPoint, new Point2d(pointX - 5, pointY + height - 25), 0, fontSize, fontSize);
                    InContur.AddVertexAt(++roomNumPoint, new Point2d(pointX - width + 20, pointY + height - 25), 0, fontSize, fontSize);
                    InContur.AddVertexAt(++roomNumPoint, new Point2d(pointX - width + 20, pointY + height - 5), 0, fontSize, fontSize);
                    InContur.AddVertexAt(++roomNumPoint, new Point2d(pointX - 5, pointY + height - 5), 0, fontSize, fontSize);
                    InContur.Closed = true;
                    InContur.Layer = "Z-STMP";
                    acBlkTblRec.AppendEntity(InContur);
                    acTrans.AddNewlyCreatedDBObject(InContur, true);

                    ObjectIdCollection acObjIdColl = new ObjectIdCollection();
                    acObjIdColl.Add(Poly.ObjectId);
                    ObjectIdCollection acObjIdColl1 = new ObjectIdCollection();
                    acObjIdColl1.Add(InContur.ObjectId);

                    Hatch acHatch = new Hatch();
                    acBlkTblRec.AppendEntity(acHatch);

                    acTrans.AddNewlyCreatedDBObject(acHatch, true);

                    acHatch.SetDatabaseDefaults();
                    acHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                    acHatch.Associative = true;

                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl1);
                   
                    acHatch.EvaluateHatch(true);
                    acHatch.Layer = "Z-STMP";
                    acHatch.ColorIndex = 9;
                    acTrans.Commit();
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }

        //создаём форматную рамку с боковыми штампами и текстом в штампах
        private void FrameSize(double height, double width, double pointX, double pointY)
        {
             var editor = Application.DocumentManager.MdiActiveDocument.Editor;
             try
             {
                 Document acDoc = Application.DocumentManager.MdiActiveDocument;
                 Database acCurDb = acDoc.Database;
                 using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                 {
                     double fontSize = 0.25;
                     BlockTable acBlkTbl;
                     acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                     BlockTableRecord acBlkTblRec;
                     acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.PaperSpace], OpenMode.ForWrite) as BlockTableRecord;
                     int roomNumPoint = 0;
                     Polyline Poly = new Polyline();
                     Polyline InContur = new Polyline();
                     Polyline Number = new Polyline();
                     Polyline sideStamp = new Polyline();

                     Poly.SetDatabaseDefaults();

                     //Внешний контур 
                     Poly.AddVertexAt(roomNumPoint, new Point2d(pointX, pointY), 0, fontSize, fontSize);
                     Poly.AddVertexAt(++roomNumPoint, new Point2d(pointX - width, pointY), 0, fontSize, fontSize);
                     Poly.AddVertexAt(++roomNumPoint, new Point2d(pointX - width, pointY + height), 0, fontSize, fontSize);
                     Poly.AddVertexAt(++roomNumPoint, new Point2d(pointX, pointY + height), 0, fontSize, fontSize);
                     Poly.Closed = true;
                     Poly.Layer = "Z-STMP";
                     acBlkTblRec.AppendEntity(Poly);
                     acTrans.AddNewlyCreatedDBObject(Poly, true);

                     //Внутренний контур
                     roomNumPoint = 0;
                     fontSize = 0.7;
                     InContur.AddVertexAt(roomNumPoint, new Point2d(pointX - 5, pointY + 5), 0, fontSize, fontSize);
                     InContur.AddVertexAt(++roomNumPoint, new Point2d(pointX - width + 20, pointY + 5), 0, fontSize, fontSize);
                     InContur.AddVertexAt(++roomNumPoint, new Point2d(pointX - width + 20, pointY + height - 5), 0, fontSize, fontSize);
                     InContur.AddVertexAt(++roomNumPoint, new Point2d(pointX - 5, pointY + height - 5), 0, fontSize, fontSize);
                     InContur.Closed = true;
                     InContur.Layer = "Z-STMP";
                     acBlkTblRec.AppendEntity(InContur);
                     acTrans.AddNewlyCreatedDBObject(InContur, true);

                     //Контур в правом верхнем углу, для сквозной нумерации
                     roomNumPoint = 0;
                     fontSize = 0.4;
                     Number.AddVertexAt(roomNumPoint, new Point2d(pointX - 5, pointY + height - 12), 0, fontSize, fontSize);
                     Number.AddVertexAt(++roomNumPoint, new Point2d(pointX - 15, pointY + height - 12), 0, fontSize, fontSize);
                     Number.AddVertexAt(++roomNumPoint, new Point2d(pointX - 15, pointY + height - 5), 0, fontSize, fontSize);
                     Number.Layer = "Z-STMP";
                     acBlkTblRec.AppendEntity(Number);
                     acTrans.AddNewlyCreatedDBObject(Number, true);

                     //Боковой штамп
                     roomNumPoint = 0;
                     fontSize = 0.7;
                     sideStamp.AddVertexAt(roomNumPoint, new Point2d(pointX - width + 20, pointY + 5), 0, fontSize, fontSize);
                     sideStamp.AddVertexAt(++roomNumPoint, new Point2d(pointX - width + 8, pointY + 5), 0, fontSize, fontSize);
                     sideStamp.AddVertexAt(++roomNumPoint, new Point2d(pointX - width + 8, pointY + 90), 0, fontSize, fontSize);
                     sideStamp.AddVertexAt(++roomNumPoint, new Point2d(pointX - width + 20, pointY + 90), 0, fontSize, fontSize);
                     sideStamp.Layer = "Z-STMP";
                     acBlkTblRec.AppendEntity(sideStamp);
                     acTrans.AddNewlyCreatedDBObject(sideStamp, true);

                     //Боковой штамп
                     CreateStamp(pointX - width + 20, pointY + 30, pointX - width + 8, pointY + 30, fontSize);
                     CreateStamp(pointX - width + 20, pointY + 65, pointX - width + 8, pointY + 65, fontSize);
                     CreateStamp(pointX - width + 20, pointY + 90, pointX - width + 8, pointY + 90, fontSize);
                     CreateStamp(pointX - width + 13, pointY + 5, pointX - width + 13, pointY + 90, fontSize);

                     //Штамп согласования
                     fontSize = 0.25;
                     CreateStamp(pointX - width + 20, pointY + 90, pointX - width, pointY + 90, fontSize);
                     CreateStamp(pointX - width + 20, pointY + 155, pointX - width, pointY + 155, fontSize);
                     CreateStamp(pointX - width + 20, pointY + 110, pointX - width + 5, pointY + 110, fontSize);
                     CreateStamp(pointX - width + 20, pointY + 130, pointX - width + 5, pointY + 130, fontSize);
                     CreateStamp(pointX - width + 20, pointY + 145, pointX - width + 5, pointY + 145, fontSize);

                     CreateStamp(pointX - width + 15, pointY + 90, pointX - width + 15, pointY + 155, fontSize);
                     CreateStamp(pointX - width + 10, pointY + 90, pointX - width + 10, pointY + 155, fontSize);
                     CreateStamp(pointX - width + 5, pointY + 90, pointX - width + 5, pointY + 155, fontSize);

                     acTrans.Commit();
                 }
             }
             catch (System.Exception ex)
             {
                 editor.WriteMessage("\n Exception caught" + ex);
             }
        }

        //создаём форматную рамку с боковыми штампами и текстом в штампах
        private void FrameSizeArcitect(double height, double width, double pointX, double pointY)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                Document acDoc = Application.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    double fontSize = 0.25;
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.PaperSpace], OpenMode.ForWrite) as BlockTableRecord;
                    int roomNumPoint = 0;
                    Polyline Poly = new Polyline();
                    Polyline InContur = new Polyline();

                    Poly.SetDatabaseDefaults();

                    //Внешний контур 
                    Poly.AddVertexAt(roomNumPoint, new Point2d(pointX, pointY), 0, fontSize, fontSize);
                    Poly.AddVertexAt(++roomNumPoint, new Point2d(pointX - width, pointY), 0, fontSize, fontSize);
                    Poly.AddVertexAt(++roomNumPoint, new Point2d(pointX - width, pointY + height), 0, fontSize, fontSize);
                    Poly.AddVertexAt(++roomNumPoint, new Point2d(pointX, pointY + height), 0, fontSize, fontSize);
                    Poly.Closed = true;
                    Poly.Layer = "Z-STMP";
                    acBlkTblRec.AppendEntity(Poly);
                    acTrans.AddNewlyCreatedDBObject(Poly, true);

                    //Внутренний контур
                    roomNumPoint = 0;
                    fontSize = 0.7;
                    InContur.AddVertexAt(roomNumPoint, new Point2d(pointX - 5, pointY + 5), 0, fontSize, fontSize);
                    InContur.AddVertexAt(++roomNumPoint, new Point2d(pointX - width + 20, pointY + 5), 0, fontSize, fontSize);
                    InContur.AddVertexAt(++roomNumPoint, new Point2d(pointX - width + 20, pointY + height - 5), 0, fontSize, fontSize);
                    InContur.AddVertexAt(++roomNumPoint, new Point2d(pointX - 5, pointY + height - 5), 0, fontSize, fontSize);
                    InContur.Closed = true;
                    InContur.Layer = "Z-STMP";
                    acBlkTblRec.AppendEntity(InContur);
                    acTrans.AddNewlyCreatedDBObject(InContur, true);

                    acTrans.Commit();
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }

        private void CreateStamp(double X1, double Y1, double X2, double Y2, double fontSize)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                Document acDoc = Application.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.PaperSpace], OpenMode.ForWrite) as BlockTableRecord;
                    int roomNumPoint = 0;
                    Polyline Poly = new Polyline();
                    Poly.SetDatabaseDefaults();
                    Poly.AddVertexAt(roomNumPoint, new Point2d(X1, Y1), 0, fontSize, fontSize);
                    Poly.AddVertexAt(++roomNumPoint, new Point2d(X2, Y2), 0, fontSize, fontSize);
                    Poly.Layer = "Z-STMP";
                    acBlkTblRec.AppendEntity(Poly);
                    acTrans.AddNewlyCreatedDBObject(Poly, true);
                    acTrans.Commit();
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        
        //конструктор стандартных рамок по формату
        private void FrameBuilder(int height, int width)
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                Document acDoc = Application.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    FrameStamp CreateStamp = new FrameStamp();
                    PromptPointOptions pointOptions = new PromptPointOptions("УКАЖИТЕ ТОЧКУ: ");
                    PromptPointResult pointResult = editor.GetPoint(pointOptions);
                    Point3d PointFrame = pointResult.Value;
                    Layers.CreateLayerStampAndFrame();
                    CreateStamp.FrameSize(height, width, PointFrame.X, PointFrame.Y);
                    SideStampText(height, width, PointFrame.X, PointFrame.Y);
                    acTrans.Commit();
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }

        }
        //конструктор архитектурных рамок по формату
        private void FrameBuilderArch(int height, int width)
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                Creator text = new Creator();

                string buro = "\"Архитектурное бюро \"Студия 44\"";
                string nameDrawing = "Наименование чертежа";
                string nameObject = "Наименование объекта";

                Document acDoc = Application.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    FrameStamp CreateStamp = new FrameStamp();
                    Technical logo = new Technical();

                    PromptPointOptions pointOptions = new PromptPointOptions("УКАЖИТЕ ТОЧКУ: ");
                    PromptPointResult pointResult = ed.GetPoint(pointOptions);
                    Point3d PointFrame = pointResult.Value;
                    Layers.CreateLayerStampAndFrame();

                    CreateStamp.AddSolid(height, width, PointFrame.X, PointFrame.Y);
                    CreateStamp.FrameSizeArcitect(height, width, PointFrame.X, PointFrame.Y);

                    logo.AddLogo(PointFrame.X - width + 74, PointFrame.Y + 5);

                    Creator.CreateTextStamp(PointFrame.X - width + 42, PointFrame.Y + 11, buro, 3, 1, 0, 0);
                    Creator.CreateTextStamp(PointFrame.X - 56, PointFrame.Y + 11, nameDrawing, 3, 1, 0, 0);
                    Creator.CreateTextStamp(PointFrame.X - width + 25, PointFrame.Y + height - 16, nameObject, 3, 1, 0, 0);

                    acTrans.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\n Exception caught" + ex);
            }
        }

        //Рамка А4 горизонтальная
        [CommandMethod("A4H", CommandFlags.NoTileMode)]
        public static void A4H()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                FrameStamp Frame = new FrameStamp();
                Frame.FrameBuilder(210, 297);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        //Рамка А4 вертикальная
        [CommandMethod("A4V", CommandFlags.NoTileMode)]
        public static void A4V()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                FrameStamp Frame = new FrameStamp();
                Frame.FrameBuilder(297, 210);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        //Рамка A4 x 3 
        [CommandMethod("A4H3", CommandFlags.NoTileMode)]
        public static void A4H3()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                FrameStamp Frame = new FrameStamp();
                Frame.FrameBuilder(297, 630);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        //Рамка A4 x 4 
        [CommandMethod("A4H4", CommandFlags.NoTileMode)]
        public static void A4H4()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                FrameStamp Frame = new FrameStamp();
                Frame.FrameBuilder(297, 841);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        //Рамка А3 горизонтальная
        [CommandMethod("A3H", CommandFlags.NoTileMode)]
        public static void A3H()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                FrameStamp Frame = new FrameStamp();
                Frame.FrameBuilder(297, 420);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        //Рамка А3 вертикальная
        [CommandMethod("A3V", CommandFlags.NoTileMode)]
        public static void A3V()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                FrameStamp Frame = new FrameStamp();
                Frame.FrameBuilder(420, 297);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        //Рамка А3 x 3
        [CommandMethod("A3H3", CommandFlags.NoTileMode)]
        public static void A3H3()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                FrameStamp Frame = new FrameStamp();
                Frame.FrameBuilder(420, 891);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        //Рамка А2 горизонтальная
        [CommandMethod("A2H", CommandFlags.NoTileMode)]
        public static void A2H()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                FrameStamp Frame = new FrameStamp();
                Frame.FrameBuilder(420, 594);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        //Рамка А2 вертикальная
        [CommandMethod("A2V", CommandFlags.NoTileMode)]
        public static void A2V()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                FrameStamp Frame = new FrameStamp();
                Frame.FrameBuilder(594, 420);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        //Рамка А1 горизонтальная
        [CommandMethod("A1H", CommandFlags.NoTileMode)]
        public static void A1H()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                FrameStamp Frame = new FrameStamp();
                Frame.FrameBuilder(594, 841);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        //Рамка А1 вертикальная
        [CommandMethod("A1V", CommandFlags.NoTileMode)]
        public static void A1V()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                FrameStamp Frame = new FrameStamp();
                Frame.FrameBuilder(841, 594);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        //Рамка А0 горизонтальная
        [CommandMethod("A0H", CommandFlags.NoTileMode)]
        public static void A0H()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                FrameStamp Frame = new FrameStamp();
                Frame.FrameBuilder(841, 1189);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        //Рамка А0 вертикальная
        [CommandMethod("A0V", CommandFlags.NoTileMode)]
        public static void A0V()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                FrameStamp Frame = new FrameStamp();
                Frame.FrameBuilder(1189, 841);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }

        //Рамка А4 горизонтальная архитектурная
        [CommandMethod("A4HA", CommandFlags.NoTileMode)]
        public static void A4HA()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                FrameStamp Frame = new FrameStamp();
                Frame.FrameBuilderArch(210, 297);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        //Рамка А4 вертикальная архитектурная
        [CommandMethod("A4VA", CommandFlags.NoTileMode)]
        public static void A4VA()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                FrameStamp Frame = new FrameStamp();
                Frame.FrameBuilderArch(297, 210);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        //Рамка А3 горизонтальная архитектурная
        [CommandMethod("A3HA", CommandFlags.NoTileMode)]
        public static void A3HA()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                FrameStamp Frame = new FrameStamp();
                Frame.FrameBuilderArch(297, 420);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        //Рамка А3 вертикальная архитектурная
        [CommandMethod("A3VA", CommandFlags.NoTileMode)]
        public static void A3VA()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                FrameStamp Frame = new FrameStamp();
                Frame.FrameBuilderArch(420, 297);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        //Рамка А2 горизонтальная архитектурная
        [CommandMethod("A2HA", CommandFlags.NoTileMode)]
        public static void A2HA()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                FrameStamp Frame = new FrameStamp();
                Frame.FrameBuilderArch(420, 594);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        //Рамка А2 вертикальная архитектурная
        [CommandMethod("A2VA", CommandFlags.NoTileMode)]
        public static void A2VA()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                FrameStamp Frame = new FrameStamp();
                Frame.FrameBuilderArch(594, 420);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        //Рамка А1 горизонтальная архитектурная
        [CommandMethod("A1HA", CommandFlags.NoTileMode)]
        public static void A1HA()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                FrameStamp Frame = new FrameStamp();
                Frame.FrameBuilderArch(594, 841);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        //Рамка А1 вертикальная архитектурная
        [CommandMethod("A1VA", CommandFlags.NoTileMode)]
        public static void A1VA()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                FrameStamp Frame = new FrameStamp();
                Frame.FrameBuilderArch(841, 594);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        //Рамка А0 горизонтальная архитектурная
        [CommandMethod("A0HA", CommandFlags.NoTileMode)]
        public static void A0HA()
        {
           Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                FrameStamp Frame = new FrameStamp();
                Frame.FrameBuilderArch(841, 1189);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        //Рамка А0 горизонтальная архитектурная
        [CommandMethod("A0VA", CommandFlags.NoTileMode)]
        public static void A0VA()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                FrameStamp Frame = new FrameStamp();
                Frame.FrameBuilderArch(1189, 841);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }

        //Конструктор рамок по размерам пользователей 
        [CommandMethod("FrameCustom", CommandFlags.NoTileMode)]
        public static void FrameCustom()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            PromptIntegerResult InputHeight = editor.GetInteger("\n Input height: ");
            if (InputHeight.Status != PromptStatus.OK)
            {
                editor.WriteMessage("\n No integer was provided");
                return;
            }
            int heightFrame = System.Convert.ToInt16(InputHeight.Value.ToString());
            if (heightFrame < 210 || heightFrame > 1189)
            {
                editor.WriteMessage("\n Incorrect frame size");
                return;
            }
            PromptIntegerResult InputWidth = editor.GetInteger("\n Input width: ");
            if (InputWidth.Status != PromptStatus.OK)
            {
                editor.WriteMessage("\n No integer was provided");
                return;
            }
            int widthFrame = System.Convert.ToInt16(InputWidth.Value.ToString());
            if (widthFrame < 210 || widthFrame > 1189)
            {
                editor.WriteMessage("\n Incorrect frame size");
                return;
            }        
            FrameStamp Frame = new FrameStamp();
            Frame.FrameBuilder(heightFrame, widthFrame);
            editor.WriteMessage("\n Create custom frame: height {0} width {1}", heightFrame, widthFrame);
        }

        //Создаём текст бокового штампа
        private static void SideStampText(int height, int width, double X, double Y)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                const double correctWidth = 11.75;
                const double heightText = 2.5;
                const double widthFactor = 0.8;
                const double angle = 1.57079632679; //90 градусов
                const double oblique = 0;
                Creator.CreateTextStamp(X - width + 3.4529, Y + 91.5604, "Согласовано:", heightText, widthFactor, angle, oblique);
                Creator.CreateTextStamp(X - width + correctWidth, Y + 67.8355, "Взам. инв.N%%D", heightText, widthFactor, angle, oblique);
                Creator.CreateTextStamp(X - width + correctWidth, Y + 36.5089, "Подп. и дата", heightText, widthFactor, angle, oblique);
                Creator.CreateTextStamp(X - width + correctWidth, Y + 7.9314, "Инв.N%%D подп.", heightText, widthFactor, angle, oblique);
                Technical.AddAttribute(X - 12.6667, Y + height - 10.25, 3.5, 0.8, 0, 0, "A_TOM_PAGE_NUM", blk++);
                Technical.AddAttribute(X - width + 17.7061, Y + 12.5476, 2.5, 0.8, angle, 0, "A_ARCH_SIGN", blk++);
                Technical.AddAttributeFunction(X - width + 17.7061, Y + 73.2334, 2.5, 0.8, angle, 0, "CMD_SYSLIB", "GETDATAFROMPAROBJFORACAD", blk++, "A_INSTEAD_OF_NUM");
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }

        private static void StampBuilder(string StampType)
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                Creator text = new Creator();
                Document acDoc = Application.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;

                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    FrameStamp createStamp = new FrameStamp();

                    PromptPointOptions pointOptions = new PromptPointOptions("УКАЖИТЕ ТОЧКУ: ");
                    PromptPointResult pointResult = editor.GetPoint(pointOptions);
                    Point3d selectedPoint = pointResult.Value;
                    Layers.CreateLayerStampAndFrame();
                    //Горизонтальные линии штампа                
                    createStamp.CreateStamp(selectedPoint.X - 185, selectedPoint.Y + 55, selectedPoint.X, selectedPoint.Y + 55, 0.7);
                    createStamp.CreateStamp(selectedPoint.X - 120, selectedPoint.Y + 15, selectedPoint.X, selectedPoint.Y + 15, 0.7);
                    createStamp.CreateStamp(selectedPoint.X - 120, selectedPoint.Y + 30, selectedPoint.X, selectedPoint.Y + 30, 0.7);
                    createStamp.CreateStamp(selectedPoint.X - 120, selectedPoint.Y + 45, selectedPoint.X, selectedPoint.Y + 45, 0.7);
                    createStamp.CreateStamp(selectedPoint.X - 50, selectedPoint.Y + 25, selectedPoint.X, selectedPoint.Y + 25, 0.7);
                    int step = 5;
                    double fontSize = 0.25;
                    for (int i = 0; i < 10; ++i)
                    {
                        if (i == 5 || i == 6)
                        {
                            fontSize = 0.7;
                            createStamp.CreateStamp(selectedPoint.X - 185, selectedPoint.Y + step, selectedPoint.X - 120, selectedPoint.Y + step, fontSize);
                            fontSize = 0.25;
                        }
                        createStamp.CreateStamp(selectedPoint.X - 185, selectedPoint.Y + step, selectedPoint.X - 120, selectedPoint.Y + step, fontSize);
                        step += 5;
                    }
                    //Вертикальные линии штампа
                    step = 10;
                    fontSize = 0.7;
                    createStamp.CreateStamp(selectedPoint.X - 185, selectedPoint.Y, selectedPoint.X - 185, selectedPoint.Y + 55, fontSize);
                    for (int i = 0; i < 4; ++i)
                    {
                        if (i == 1 || i == 3)
                        {
                            createStamp.CreateStamp(selectedPoint.X - 185 + step, selectedPoint.Y, selectedPoint.X - 185 + step, selectedPoint.Y + 55, fontSize);
                        }
                        createStamp.CreateStamp(selectedPoint.X - 185 + step, selectedPoint.Y + 30, selectedPoint.X - 185 + step, selectedPoint.Y + 55, fontSize);
                        step += 10;
                    }
                    createStamp.CreateStamp(selectedPoint.X - 130, selectedPoint.Y, selectedPoint.X - 130, selectedPoint.Y + 55, fontSize);
                    createStamp.CreateStamp(selectedPoint.X - 120, selectedPoint.Y, selectedPoint.X - 120, selectedPoint.Y + 55, fontSize);
                    createStamp.CreateStamp(selectedPoint.X - 50, selectedPoint.Y, selectedPoint.X - 50, selectedPoint.Y + 30, fontSize);
                    createStamp.CreateStamp(selectedPoint.X - 35, selectedPoint.Y + 15, selectedPoint.X - 35, selectedPoint.Y + 30, fontSize);
                    createStamp.CreateStamp(selectedPoint.X - 20, selectedPoint.Y + 15, selectedPoint.X - 20, selectedPoint.Y + 30, fontSize);
                    //Создание текста в штампе
                    if (StampType != "stampar")
                    {
                        Creator.CreateTextStamp(selectedPoint.X - 184.587, selectedPoint.Y + 26.25, rab, 2.5, 0.75, 0, 0);         //Разработал
                        Creator.CreateTextStamp(selectedPoint.X - 184.587, selectedPoint.Y + 21.25, search, 2.5, 0.75, 0, 0);      //Проверил
                        Creator.CreateTextStamp(selectedPoint.X - 184.587, selectedPoint.Y + 16.25, rukGroup, 2.5, 0.75, 0, 0);    //Руководитель группы
                        Creator.CreateTextStamp(selectedPoint.X - 184.587, selectedPoint.Y + 11.25, normal, 2.5, 0.75, 0, 0);      //Нормоконтроль
                        Creator.CreateTextStamp(selectedPoint.X - 184.587, selectedPoint.Y + 6.25, glConstPr, 2.5, 0.75, 0, 0);    //Главный конструктор проекта

                        //ФАМИЛИИ (Берутся из ТДМС)
                        Technical.AddAttributeFunction(selectedPoint.X - 164.292, selectedPoint.Y + 26.25, 2.5, 0.6, 0, 0, moduleName, _GDFARTFA, blk++, "DEVELOP", "A_User"); //Разработал
                        Technical.AddAttributeFunction(selectedPoint.X - 164.292, selectedPoint.Y + 21.25, 2.5, 0.6, 0, 0, moduleName, _GDFARTFA, blk++, "CHECK", "A_User");   //Проверил
                        Technical.AddAttributeFunction(selectedPoint.X - 164.292, selectedPoint.Y + 16.25, 2.5, 0.6, 0, 0, moduleName, _GDFARTFA, blk++, "GR_HEAD", "A_User"); //Руководитель группы
                        Technical.AddAttributeFunction(selectedPoint.X - 164.292, selectedPoint.Y + 11.25, 2.5, 0.6, 0, 0, moduleName, _GDFARTFA, blk++, "NORMKL", "A_User");  //Нормоконтроль
                        Technical.AddAttributeFunction(selectedPoint.X - 164.292, selectedPoint.Y + 6.25, 2.5, 0.6, 0, 0, moduleName, _GDFARTFA, blk++, "GKP_", "A_User");     //Главный конструктор проекта

                        //ДАТЫ (Берутся из ТДМС)
                        Technical.AddAttributeFunction(selectedPoint.X - 127.943, selectedPoint.Y + 26.25, 2.5, 0.6, 0, 0, moduleName, _GDFARTFA, blk++, "DEVELOP", "A_DATE");
                        Technical.AddAttributeFunction(selectedPoint.X - 127.943, selectedPoint.Y + 21.25, 2.5, 0.6, 0, 0, moduleName, _GDFARTFA, blk++, "CHECK", "A_DATE");
                        Technical.AddAttributeFunction(selectedPoint.X - 127.943, selectedPoint.Y + 16.25, 2.5, 0.6, 0, 0, moduleName, _GDFARTFA, blk++, "GR_HEAD", "A_DATE");
                        Technical.AddAttributeFunction(selectedPoint.X - 127.943, selectedPoint.Y + 11.25, 2.5, 0.6, 0, 0, moduleName, _GDFARTFA, blk++, "NORMKL", "A_DATE");
                        Technical.AddAttributeFunction(selectedPoint.X - 127.943, selectedPoint.Y + 6.25, 2.5, 0.6, 0, 0, moduleName, _GDFARTFA, blk++, "GKP_", "A_DATE");

                        if (StampType == "stampkab")
                        {
                            //Главный конструктор AБ
                            Technical.AddAttributeFunction(selectedPoint.X - 164.292, selectedPoint.Y + 1.25, 2.5, 0.6, 0, 0, moduleName, _GDFSPFA, blk++, "GKAB_");
                            Technical.AddAttributeFunction(selectedPoint.X - 127.943, selectedPoint.Y + 1.25, 2.5, 0.6, 0, 0, moduleName, _GDFARTFA, blk++, "GKAB_", "A_DATE");
                            Creator.CreateTextStamp(selectedPoint.X - 184.587, selectedPoint.Y + 1.25, glConstAB, 2.5, 0.75, 0, 0);
                        }
                    }
                    else
                    {
                        Creator.CreateTextStamp(selectedPoint.X - 183.8, selectedPoint.Y + 26.25, rab, 2.5, 0.75, 0, 0);        //Разработал
                        Creator.CreateTextStamp(selectedPoint.X - 183.8, selectedPoint.Y + 21.25, GIP, 2.5, 0.75, 0, 0);        //ГИП
                        Creator.CreateTextStamp(selectedPoint.X - 183.8, selectedPoint.Y + 16.25, GAP, 2.5, 0.75, 0, 0);        //ГАП
                        Creator.CreateTextStamp(selectedPoint.X - 183.8, selectedPoint.Y + 11.25, search, 2.5, 0.75, 0, 0);     //Проверил
                        Creator.CreateTextStamp(selectedPoint.X - 183.8, selectedPoint.Y + 1.25, normal, 2.5, 0.75, 0, 0);      //Нормоконтроль

                        //ФАМИЛИИ (Берутся из ТДМС)
                        Technical.AddAttributeFunction(selectedPoint.X - 164.292, selectedPoint.Y + 26.25, 2.5, 0.6, 0, 0, moduleName, _GDFARTFA, blk++, "DEVELOP", "A_User"); //Разработал
                        Technical.AddAttributeFunction(selectedPoint.X - 164.292, selectedPoint.Y + 21.25, 2.5, 0.6, 0, 0, moduleName, _GDFARTFA, blk++, "GIP_", "A_User");    //ГИП
                        Technical.AddAttributeFunction(selectedPoint.X - 164.292, selectedPoint.Y + 16.25, 2.5, 0.6, 0, 0, moduleName, _GDFARTFA, blk++, "GAP_", "A_User");    //ГАП
                        Technical.AddAttributeFunction(selectedPoint.X - 164.292, selectedPoint.Y + 11.25, 2.5, 0.6, 0, 0, moduleName, _GDFARTFA, blk++, "CHECK", "A_User");   //Проверил
                        Technical.AddAttributeFunction(selectedPoint.X - 164.292, selectedPoint.Y + 1.25, 2.5, 0.6, 0, 0, moduleName, _GDFARTFA, blk++, "NORMKL", "A_User");   //Нормоконтроль

                        //ДАТЫ (Берутся из ТДМС)
                        Technical.AddAttributeFunction(selectedPoint.X - 127.943, selectedPoint.Y + 26.25, 2.5, 0.6, 0, 0, moduleName, _GDFARTFA, blk++, "DEVELOP", "A_DATE");
                        Technical.AddAttributeFunction(selectedPoint.X - 127.943, selectedPoint.Y + 21.25, 2.5, 0.6, 0, 0, moduleName, _GDFARTFA, blk++, "GIP_", "A_DATE");
                        Technical.AddAttributeFunction(selectedPoint.X - 127.943, selectedPoint.Y + 16.25, 2.5, 0.6, 0, 0, moduleName, _GDFARTFA, blk++, "GAP_", "A_DATE");
                        Technical.AddAttributeFunction(selectedPoint.X - 127.943, selectedPoint.Y + 11.25, 2.5, 0.6, 0, 0, moduleName, _GDFARTFA, blk++, "CHECK", "A_DATE");

                        Technical.AddAttributeFunction(selectedPoint.X - 127.943, selectedPoint.Y + 1.25, 2.5, 0.6, 0, 0, moduleName, _GDFARTFA, blk++, "NORMKL", "A_DATE");
                    }
                    Creator.CreateTextStamp(selectedPoint.X - 182.513, selectedPoint.Y + 31.25, IZM, 2, 0.8, 0, 0);      //Изменения
                    Creator.CreateTextStamp(selectedPoint.X - 173.745, selectedPoint.Y + 31.25, coluch, 2, 0.75, 0, 0);  //Количество участников
                    Creator.CreateTextStamp(selectedPoint.X - 163, selectedPoint.Y + 31.25, Paper, 2, 0.8, 0, 0);        //Лист
                    Creator.CreateTextStamp(selectedPoint.X - 153.23, selectedPoint.Y + 31.25, Ndoc, 2, 0.7, 0, 0);      //Номер документа
                    Creator.CreateTextStamp(selectedPoint.X - 142.61, selectedPoint.Y + 31.25, signature, 2, 0.8, 0, 0); //Подпись
                    Creator.CreateTextStamp(selectedPoint.X - 128.365, selectedPoint.Y + 31.25, date, 2, 0.8, 0, 0);     //Дата

                    Creator.CreateTextStamp(selectedPoint.X - 47.110, selectedPoint.Y + 26.17, stage, 2, 0.8, 0, 0);     //Стадия
                    Creator.CreateTextStamp(selectedPoint.X - 30.583, selectedPoint.Y + 26.17, Paper, 2, 0.8, 0, 0);     //Лист
                    Creator.CreateTextStamp(selectedPoint.X - 14.557, selectedPoint.Y + 26.17, Papers, 2, 0.8, 0, 0);    //Листов

                    //Cоздание атрибутов в штампе
                    //Середина штампа
                    Technical.AddAttributeMultiline(selectedPoint.X - 60, selectedPoint.Y + 50, 4, 118, 0.8, 0, 0, nameAtrCode, blk++);       //шифр объекта

                    Technical.AddMultilineAttributeFunction(selectedPoint.X - 60, selectedPoint.Y + 37, 2.5, 118, 0.8, 0, 0, moduleName, nameAtrAddress, blk++);   //адрес объекта

                    Technical.AddAttributeMultiline(selectedPoint.X - 85, selectedPoint.Y + 22, 2.5, 68, 0.8, 0, 0, nameAtrProjectObj, blk++); //Наименование объекта
                                                            
                    //Номера листов и стадия (Берутся из ТДМС)
                    Technical.AddAttributeFunction(selectedPoint.X - 44.1065, selectedPoint.Y + 17.8617, 4, 0.8, 0, 0, moduleName, _GDFOFA, blk++, "A_STAGE_CLSF");
                    Technical.AddAttribute(selectedPoint.X - 29.4215, selectedPoint.Y + 17.8617, 4, 0.8, 0, 0, nameAtrList, blk++);
                    Technical.AddAttribute(selectedPoint.X - 12.0115, selectedPoint.Y + 17.8617, 4, 0.8, 0, 0, nameAtrLists, blk++);
                                        
                    //Вставляем логотип
                    Technical logo = new Technical();
                    logo.AddLogo(selectedPoint.X, selectedPoint.Y);

                    Creator.CreateMultilineStampAtributNameDrawing(selectedPoint.X - 85, selectedPoint.Y + 7, 2.5, 68, 0.8, 0, 0, AttNameDrawing, blk++); //Наименование чертежа из названия layout

                    Creator.NameStudio(selectedPoint.X - 35.959, selectedPoint.Y + 11.269, 0); // Наименование в штампе OOO "Архитектурное бюро "Студия 44"

                    editor.WriteMessage(selectedPoint.ToString());
                    acTrans.Commit();
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }

        //Запуск создания форматного штампа с наименованием чертежа из названия layout
        //Для конструкторов с корректно заполненным штампом
        [CommandMethod("StampKAB", CommandFlags.NoTileMode)]
        public void StampKAB()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                StampBuilder("stampkab");
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }

        //Запуск создания форматного штампа с наименованием чертежа из названия layout
        //Для конструкторов с корректно заполненным штампом
        [CommandMethod("StampK", CommandFlags.NoTileMode)]
        public void StampK()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                StampBuilder("stampk");
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
        //Запуск создания форматного штампа с наименованием чертежа из названия layout
        //Для архитекторов с корректно заполненным штампом
        [CommandMethod("StampAR", CommandFlags.NoTileMode)]
        public void StampAR()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                StampBuilder("stampar");
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
    }

    public sealed class Condition
    {
        public void Initialize()
        { }
        public void Terminate()
        { }
        //Проверка чертежа на принадлежность ТДМС, открыт ли он из ТДМС?
        public bool CheckPath()
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            string parsePath = acDoc.Name;
            parsePath = parsePath.Remove(7);
            if (parsePath != "C:\\Temp")
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        //Обёртка с проверкой на пустой путь
        public bool StartCheckPath(string path)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                if (path == "")
                {
                    editor.WriteMessage("\n Path not found");
                    return false;
                }
                else
                {
                    return CheckPath(path);
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exeption caught (class Technical, method StartCheckPath): " + ex);
                return false;
            }
        }
        //Проверка внешних ссылок на принадлежность ТДМС, открыты ли они из ТДМС?
        private bool CheckPath(string path)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                path = path.Remove(7);
                if (path != "C:\\Temp")
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exeption caught (class Technical, method CheckPath): " + ex);
                return false;
            }
        }
        public bool CheckTDMSProcess()
        {
            Process[] process = Process.GetProcessesByName("TDMS");
            if (process.Length == 1)
            {
                return true;
            }
            return false;
        }

        public string StartParseGUID(string pathName)
        {
            return ParseGUID(pathName);
        }
        private string ParseGUID(string pathName)
        {
            string guidFromFile = pathName;
            string parseGuidFromFile = null;
            Regex regFF = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
            MatchCollection mcFF = regFF.Matches(guidFromFile);
            foreach (Match mat in mcFF)
            {
                parseGuidFromFile += mat.Value.ToString();
            }
            parseGuidFromFile = parseGuidFromFile.Remove(0, 38);
            return parseGuidFromFile;
        }
    }

    //класс включает в себя все методы для сохранения \ импорта чертежей в базе
    public class SaveOptions
    {
        public void Initialize()
        { }
        public void Terminate()
        { }

        //метод для локального сохранения чертежа по заданному пути (2 параметра, путь и имя чертежа)
        public void LocalSave(string strDwgName, string pathName)
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                string fullPath = pathName + "\\" + strDwgName;
                acDoc.Database.SaveAs(fullPath, true, DwgVersion.Current,acDoc.Database.SecurityParameters);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("Exception caught {0} {1}", ex.Message, ex.StackTrace);
            }
        }
        //метод для локального сохранения чертежа по заданному пути (1 параметр, путь + имя чертежа)
        public void LocalSave(string fullPathAndDWGName)
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                acDoc.Database.SaveAs(fullPathAndDWGName, true, DwgVersion.Current, acDoc.Database.SecurityParameters);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("Exception caught {0} {1}", ex.Message, ex.StackTrace);
            }
        }

        //метод сохранения чертежа сначала локально, затем в ТДМС
        [CommandMethod("SaveActiveDrawing", CommandFlags.Session | CommandFlags.NoBlockEditor)]
        public void StartSaveActiveDrawing()
        {
            SaveActiveDrawing();
        }
        private static void SaveActiveDrawing()
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                Condition Check = new Condition();
                string strDwgName = acDoc.Name;
                if (Check.CheckPath() == true)
                {
                    if (Check.CheckTDMSProcess() == true)
                    {
                        TDMSApplication tdmsApp = new TDMSApplication();
                        TDMSObject tdmsObj = null;
                        object obj = Application.GetSystemVariable("DWGTITLED");
                        //Возвращение пути к файлу чертежа
                        string guid = strDwgName;
                        //парсинг пути, получение актуального GUID объекта
                        string parseGuid = null;
                        Regex reg = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                        MatchCollection mc = reg.Matches(guid);
                        foreach (Match mat in mc)
                        {
                            parseGuid += mat.Value.ToString();
                        }
                        parseGuid = parseGuid.Remove(0, 38);
                        //Получение объекта по GUID
                        tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);
                        bool lockOwner = tdmsObj.Permissions.LockOwner;
                        //Заблокирован ли чертёж текущим пользователем?
                        if (lockOwner)
                        {
                            //Сохранение изменений в текущем чертеже
                            acDoc.Database.SaveAs(strDwgName, true, DwgVersion.Current, acDoc.Database.SecurityParameters);
                            acDoc.Database.CloseInput(true);
                            //загрузить в базу ТДМС, сохранить изменения
                            tdmsObj.CheckIn();
                            tdmsObj.Update();
                            Technical.TitleDoc();
                        }
                        else
                        {
                            editor.WriteMessage("Документ открыт на просмотр; Изменения не будут сохранены в TDMS!");
                            acDoc.Database.SaveAs(strDwgName, true, DwgVersion.Current, acDoc.Database.SecurityParameters);
                        }
                    }
                    else
                    {
                        Application.ShowAlertDialog("Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                        acDoc.Database.SaveAs(strDwgName, true, DwgVersion.Current, acDoc.Database.SecurityParameters);
                    }
                }
                else
                {
                    Application.ShowAlertDialog("Документ не принадлежит TDMS!");
                    acDoc.Database.SaveAs(strDwgName, true, DwgVersion.Current, acDoc.Database.SecurityParameters);
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }

        //Закрыть файл и загрузить в базу ТДМС без сохранения изменений
        [CommandMethod("CloseAndDiscard", CommandFlags.Session | CommandFlags.NoBlockEditor)]
        public void StartCloseAndDiscard()
        {
            CloseAndDiscard();
        }
        private static void CloseAndDiscard()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                Condition Check = new Condition();
                if (Check.CheckPath() == true)
                {
                    if (Check.CheckTDMSProcess() == true)
                    {
                        TDMSApplication tdmsApp = new TDMSApplication();
                        TDMSObject tdmsObj = null;
                        Document acDoc = Application.DocumentManager.MdiActiveDocument;
                        object obj = Application.GetSystemVariable("DWGTITLED");
                        string strDwgName = acDoc.Name;

                        //Возвращение пути к файлу чертежа
                        string guid = strDwgName;
                        //парсинг пути, получение актуального GUID объекта
                        string parseGuid = null;
                        Regex reg = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                        MatchCollection mc = reg.Matches(guid);
                        foreach (Match mat in mc)
                        {
                            parseGuid += mat.Value.ToString();
                        }
                        parseGuid = parseGuid.Remove(0, 38);

                        //Получение объекта по GUID
                        tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);

                        bool lockOwner = tdmsObj.Permissions.LockOwner;
                        //Заблокировани ли чертёж текущим пользователем
                        if (lockOwner)
                        {
                            //Закрыть файл и загрузить в базу ТДМС без сохранения изменений
                            tdmsObj.UnlockCheckIn(0);
                            acDoc.CloseAndDiscard();
                            tdmsObj.Update();
                        }
                        else
                        {
                            //Закрытие чертежа
                            acDoc.CloseAndDiscard();
                        }
                    }
                    else
                    {
                        Application.ShowAlertDialog("Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                    }
                }
                else
                {
                    Application.ShowAlertDialog("Документ не принадлежит TDMS!");
                    editor.WriteMessage("Документ не принадлежит TDMS!");
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }

        //Сохранение изменений в чертеже, загрузка чертежа в базу ТДМС, 
        //объект базы разблокируется, текущий чертёж закрывается.
        [CommandMethod("SaveAndClose", CommandFlags.Session | CommandFlags.NoBlockEditor)]
        public void StartSaveAndClose()
        {
            SaveAndClose();
        }
        private static void SaveAndClose()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                Condition Check = new Condition();
                Document acDoc = Application.DocumentManager.MdiActiveDocument;
                string strDwgName = acDoc.Name;
                if (Check.CheckPath() == true)
                {
                    if (Check.CheckTDMSProcess() == true)
                    {
                        TDMSApplication tdmsApp = new TDMSApplication();
                        TDMSObject tdmsObj = null;

                        object obj = Application.GetSystemVariable("DWGTITLED");

                        //Возвращение пути к файлу чертежа
                        string guid = strDwgName;

                        //парсинг пути, получение актуального GUID объекта
                        string parseGuid = null;
                        Regex reg = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                        MatchCollection mc = reg.Matches(guid);
                        foreach (Match mat in mc)
                        {
                            parseGuid += mat.Value.ToString();
                        }
                        parseGuid = parseGuid.Remove(0, 38);

                        //Получение объекта по GUID
                        tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);

                        bool lockOwner = tdmsObj.Permissions.LockOwner;
                        //Заблокировани ли чертёж текущим пользователем
                        if (lockOwner)
                        {
                            //Сохранение изменений в текущем чертеже
                            acDoc.Database.SaveAs(strDwgName, true, DwgVersion.Current, acDoc.Database.SecurityParameters);
                            acDoc.Database.CloseInput(true);
                            //Сохранение в базе и разблокировка
                            tdmsObj.UnlockCheckIn(1);
                           
                            acDoc.CloseAndDiscard();
                            tdmsObj.Update();
                        }
                        else
                        {
                            //Сохранение чертежа в базе ТДМС и его разблокировка для других пользователей
                            Application.ShowAlertDialog("Невозможно сохранить в TDMS, т.к. документ открыт на просмотр! Сохраните файл на локальном диске.");
                            Technical.TitleDoc();
                        }
                    }
                    else
                    {
                        Application.ShowAlertDialog("Невозможно сохранить в TDMS, т.к. он не запущен! Сохраните файл на локальном диске.");
                        acDoc.Database.SaveAs(strDwgName, true, DwgVersion.Current, acDoc.Database.SecurityParameters);
                    }
                }
                else
                {
                    Application.ShowAlertDialog("Документ не принадлежит TDMS!");
                    acDoc.Database.SaveAs(strDwgName, true, DwgVersion.Current, acDoc.Database.SecurityParameters);
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught" + ex);
            }
        }
    }
}
