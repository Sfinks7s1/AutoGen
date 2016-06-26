namespace Auto
{
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.EditorInput;
    using Autodesk.AutoCAD.Geometry;
    using Autodesk.AutoCAD.Runtime;
    using System;
    using System.IO;
    using System.Text.RegularExpressions;
    using System.Threading;
    using TDMS.Interop;

    /// <summary>
    /// Класс содержит методы для обработки внешних ссылок
    /// </summary>
    public class Xref
    {
        private static TDMSObject _tdmsObjXref;
        private FindFileHint _xRefDrawing;
        private FindFileHint _xRefDrawingExport;

        /// <summary>
        /// Поиск и выгрузка из ТДМС изображений
        /// </summary>
        public static void FindImages()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var acCurDb = doc.Database;
            var editor = doc.Editor;
            var externalReferences = new Xref();
            try
            {
                //Получение пути к объектам c:\temp\tdms\userGUID\
                //из пути к главному файлу вычитается GUID объекта ({39 символов + \})
                var pathMainFile = Path.GetDirectoryName(doc.Name);
                pathMainFile = pathMainFile.Remove(pathMainFile.Length - 39);

                string parseGuid = null;
                var tdmsApp = new TDMSApplication();
                var reg = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);

                using (var acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    var filterlist = new TypedValue[1];
                    filterlist[0] = new TypedValue(0, "IMAGE");
                    var filter = new SelectionFilter(filterlist);
                    var selRes = editor.SelectAll(filter);
                    if (selRes.Status != PromptStatus.OK)
                    {
                        editor.WriteMessage("\n No Images selected.");
                        externalReferences.RefreshXref();
                        return;
                    }

                    var oSs = selRes.Value;

                    for (var i = 0; i < oSs.Count; i++)
                    {
                        var oRaster = (RasterImage)acTrans.GetObject(oSs[i].ObjectId, OpenMode.ForRead);
                        var oRasterIDef = (RasterImageDef)acTrans.GetObject(oRaster.ImageDefId, OpenMode.ForRead);

                        try
                        {
                            var mc = reg.Matches(oRasterIDef.SourceFileName.ToString());
                            foreach (Match mat in mc)
                            {
                                parseGuid += mat.Value.ToString();
                            }
                            //editor.WriteMessage("\n PARSER: " + System.Convert.ToUInt32(parseGuid.Length));
                            //editor.WriteMessage("\n PARSERGUID: " + parseGuid);
                            TDMSObject tdmsObj = null;
                            switch (Convert.ToUInt32(parseGuid.Length))
                            {
                                case 38:
                                    {
                                        tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);
                                        var mainFile = tdmsObj.Files.Main;
                                        var fileNameInTdms = Path.GetFileName(mainFile.FileName).ToString();
                                        var fileNameInAutoCadXref = Path.GetFileName(oRasterIDef.SourceFileName).ToString();

                                        if (fileNameInTdms != fileNameInAutoCadXref)
                                        {
                                            parseGuid = null;
                                        }
                                        else
                                        {
                                            mainFile.CheckOut(pathMainFile + @"\" + parseGuid + @"\" + fileNameInTdms);
                                            parseGuid = null;
                                        }
                                        break;
                                    }
                                case 76:
                                    {
                                        parseGuid = parseGuid.Remove(0, 38);
                                        tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);
                                        var mainFile = tdmsObj.Files.Main;

                                        var fileNameInTdms = Path.GetFileName(mainFile.FileName).ToString();
                                        var fileNameInAutoCadXref = Path.GetFileName(oRasterIDef.SourceFileName).ToString();
                                        var directoryNameInAutoCadXref = Path.GetDirectoryName(oRasterIDef.SourceFileName.ToString());

                                        if (fileNameInTdms != fileNameInAutoCadXref)
                                        {
                                            parseGuid = null;
                                        }
                                        else
                                        {
                                            mainFile.CheckOut(directoryNameInAutoCadXref + @"\" + fileNameInTdms);
                                            parseGuid = null;
                                        }
                                        break;
                                    }
                            }
                        }
                        catch (System.Exception ex)
                        {
                            editor.WriteMessage("\n Link is out TDMS.");
                            editor.WriteMessage("\n Path Link is out TDMS: " + oRasterIDef.SourceFileName.ToString());
                            parseGuid = null;
                            externalReferences.RefreshXref();
                        }
                    }
                    acTrans.Commit();
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Поиск и выгрузка из ТДМС изображений по заданному пути
        /// </summary>
        public static void FindImages(string checkOutPath)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var acCurDb = doc.Database;
            var editor = doc.Editor;
            var externalReferences = new Xref();

            editor.WriteMessage("\n 1_CheckOutPath " + checkOutPath);

            try
            {
                string parseGuid = null;
                var tdmsApp = new TDMSApplication();
                var reg = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);

                using (var acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    var filterlist = new TypedValue[1];
                    filterlist[0] = new TypedValue(0, "IMAGE");
                    var filter = new SelectionFilter(filterlist);
                    var selRes = editor.SelectAll(filter);

                    if (selRes.Status != PromptStatus.OK)
                    {
                        editor.WriteMessage("\n No Images selected.");
                        externalReferences.RefreshXref();
                        return;
                    }

                    var oSs = selRes.Value;

                    for (var i = 0; i < oSs.Count; i++)
                    {
                        var oRaster = (RasterImage)acTrans.GetObject(oSs[i].ObjectId, OpenMode.ForRead);
                        var oRasterIDef = (RasterImageDef)acTrans.GetObject(oRaster.ImageDefId, OpenMode.ForRead);
                        var imageName = Path.GetFileName(oRasterIDef.SourceFileName.ToString());
                        editor.WriteMessage("\n 2_ImageName" + imageName);
                        try
                        {
                            var mc = reg.Matches(oRasterIDef.SourceFileName.ToString());
                            foreach (Match mat in mc)
                            {
                                parseGuid += mat.Value.ToString();
                            }

                            var tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);
                            parseGuid = null;
                            var mainFile = tdmsObj.Files.Main;

                            var directoryName = Path.GetDirectoryName(checkOutPath);

                            var newpath = directoryName + @"\XREF\" + imageName;

                            editor.WriteMessage("\n 3_newpath" + newpath);

                            mainFile.CheckOut(newpath);
                        }
                        catch (System.Exception ex)
                        {
                            editor.WriteMessage("\n Link is out TDMS.");
                            editor.WriteMessage("\n Path Link is out TDMS: " + oRasterIDef.SourceFileName.ToString());
                            parseGuid = null;
                            externalReferences.RefreshXref();
                        }
                    }
                    acTrans.Commit();
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        public static TDMSObject GetTdmsObjXref()
        {
            return _tdmsObjXref;
        }

        /// <summary>
        /// Метод реализует возможность редактирования внешней ссылки в пространстве чертежа.
        /// </summary>
        [CommandMethod("XREFEDIT", CommandFlags.NoBlockEditor)]
        public static void XRefEdit()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var editor = doc.Editor;
            var docName = doc.Name;
            try
            {
                using (var tr = doc.TransactionManager.StartTransaction())
                {
                    //выбираем объект в пространстве чертежа
                    var acSsPrompt = doc.Editor.GetSelection();
                    if (acSsPrompt.Status == PromptStatus.OK)
                    {
                        var acSSet = acSsPrompt.Value;
                        //перебираем объекты в выбранной коллекции
                        foreach (SelectedObject acSsObj in acSSet)
                        {
                            if (acSsObj != null)
                            {
                                //получаем таблицу ID объектов
                                var acEnt = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;
                                if (acEnt != null)
                                {
                                    //находим таблицу блоков
                                    if (Convert.ToString(acEnt.GetType()) == "Autodesk.AutoCAD.DatabaseServices.BlockTable")
                                    {
                                        //перемещаемся по записям в таблице блоков
                                        foreach (var id in acEnt)
                                        {
                                            var btr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForWrite);
                                            //является ли выбранный блок внешней ссылкой?
                                            if (btr.IsFromExternalReference)
                                            {
                                                editor.WriteMessage("Debug 1: " + btr.PathName + "\n");
                                                var check = new Condition();

                                                if (check.CheckPath(doc.Name))
                                                {
                                                    editor.WriteMessage("Debug 2: Проверка пути к основному файлу чертежа. Чертёж из TDMS" + "\n");

                                                    if (check.CheckTdmsProcess())
                                                    {
                                                        editor.WriteMessage("Debug 3: Проверка на запуск TDMS пройдена" + "\n");

                                                        var tdmsApp = new TDMSApplication();

                                                        string parseGuid = null;
                                                        var reg = new Regex("{.*?}", RegexOptions.IgnoreCase);

                                                        editor.WriteMessage("Debug 4: " + btr.PathName);

                                                        if (check.CheckPath(btr.PathName) | check.CheckRelativePathCut(btr.PathName) | check.CheckRelativePath(btr.PathName))
                                                        {
                                                            editor.WriteMessage("Debug 5: Проверка на путь внешней ссылки из TDMS пройдена" + "\n");

                                                            //парсинг пути, получение актуального GUID объекта
                                                            var mc = reg.Matches(btr.PathName);
                                                            editor.WriteMessage("\n a \n");
                                                            foreach (Match mat in mc)
                                                            {
                                                                parseGuid += mat.Value.ToString();
                                                            }
                                                            editor.WriteMessage("\n b \n");
                                                            editor.WriteMessage("Полученный GUID из пути внешней ссылки: " + parseGuid + "\n");

                                                            switch (Convert.ToUInt32(parseGuid.Length))
                                                            {
                                                                case 38:
                                                                    {
                                                                        editor.WriteMessage("Debug 6: Ссылка относительная" + "\n");
                                                                        try
                                                                        {
                                                                            if (tdmsApp.GetObjectByGUID(parseGuid) != null)
                                                                            {
                                                                                editor.WriteMessage("Debug 7: Объект по GUID найден в TDMS" + "\n");
                                                                                _tdmsObjXref = tdmsApp.GetObjectByGUID(parseGuid);
                                                                                var mainFile = _tdmsObjXref.Files.Main;

                                                                                var editFiles = _tdmsObjXref.Permissions.EditFiles;

                                                                                //edit files - должно вернуть значение, редактируется или нет.
                                                                                var prava = editFiles;
                                                                                editor.WriteMessage("Имеется ли доступ к объекту (На редактирование)? \n" + prava + "\n");

                                                                                if (prava.ToString().ToLower() == "tdmallow")
                                                                                {
                                                                                    //Заблокирован ли чертёж?
                                                                                    var lockOwner = _tdmsObjXref.Permissions.LockOwner;

                                                                                    if (lockOwner)
                                                                                    {
                                                                                        editor.WriteMessage("Внешняя ссылка заблокирована пользователем и редактируется");
                                                                                    }
                                                                                    else
                                                                                    {
                                                                                        //Блокировать объект
                                                                                        _tdmsObjXref.Lock(3);
                                                                                        _tdmsObjXref.Update();
                                                                                        //Открыть внешнюю ссылку на редктирование из пространства чертежа
                                                                                        doc.SendStringToExecute("_REFEDIT" + "\n", true, false, false);
                                                                                    }
                                                                                    parseGuid = null;
                                                                                }
                                                                                else
                                                                                {
                                                                                    parseGuid = null;
                                                                                    editor.WriteMessage("Доступ на редактирование внешней ссылки отсутствует. \n");
                                                                                }
                                                                            }
                                                                            else
                                                                            {
                                                                                parseGuid = null;
                                                                                editor.WriteMessage("\n File " + btr.PathName + " not found in TDMS.");
                                                                            }
                                                                        }
                                                                        catch (System.Exception ex)
                                                                        {
                                                                            parseGuid = null;
                                                                            editor.WriteMessage("\n Exception caught: " + ex);
                                                                        }
                                                                        break;
                                                                    }
                                                                case 76:
                                                                    {
                                                                        editor.WriteMessage("Debug 6: Ссылка абсолютная" + "\n");
                                                                        try
                                                                        {
                                                                            parseGuid = parseGuid.Remove(0, 38);
                                                                            if (tdmsApp.GetObjectByGUID(parseGuid) != null)
                                                                            {
                                                                                editor.WriteMessage("Debug 7: Объект по GUID найден в TDMS" + "\n");
                                                                                _tdmsObjXref = tdmsApp.GetObjectByGUID(parseGuid);

                                                                                var editFiles = _tdmsObjXref.Permissions.EditFiles;
                                                                                var prava = editFiles.ToString();
                                                                                editor.WriteMessage("Имеется ли доступ к объекту (На редактирование)? " + prava + "\n");
                                                                                if (prava.ToString().ToLower() == "tdmallow")
                                                                                {
                                                                                    //Заблокировани ли чертёж текущим пользователем
                                                                                    var lockOwner = _tdmsObjXref.Permissions.LockOwner;
                                                                                    if (lockOwner)
                                                                                    {
                                                                                        editor.WriteMessage("Внешняя ссылка заблокирована пользователем");
                                                                                    }
                                                                                    else
                                                                                    {
                                                                                        //Блокировать объект
                                                                                        _tdmsObjXref.Lock(3);
                                                                                        _tdmsObjXref.Update();
                                                                                        //Открыть внешнюю ссылку на редкатирование из пространства чертежа
                                                                                        doc.SendStringToExecute("_REFEDIT" + "\n", true, false, false);
                                                                                    }
                                                                                    parseGuid = null;
                                                                                }
                                                                                parseGuid = null;
                                                                                editor.WriteMessage("Доступ на редактирование внешней ссылки отсутствует. \n");
                                                                            }
                                                                            else
                                                                            {
                                                                                parseGuid = null;
                                                                                editor.WriteMessage("\n File " + btr.PathName + " not found in TDMS.");
                                                                            }
                                                                        }
                                                                        catch (System.Exception ex)
                                                                        {
                                                                            parseGuid = null;
                                                                            editor.WriteMessage("\n Exception caught: " + ex);
                                                                        }
                                                                        break;
                                                                    }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            editor.WriteMessage("\n Path Link is out TDMS: " + btr.PathName);
                                                            doc.SendStringToExecute("_REFEDIT" + "\n", true, false, false);
                                                            parseGuid = null;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        editor.WriteMessage("\nНевозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                                                    }
                                                }
                                                else
                                                {
                                                    editor.WriteMessage("\nДокумент не принадлежит TDMS, но ссылка будет открыта для редактирования в пространстве модели.");
                                                    //doc.SendStringToExecute("ССЫЛРЕД" + "\n" + "ОК" + "\n" + "Все" + "\n" + "Да" + "\n", true, false, false);
                                                    doc.SendStringToExecute("_REFEDIT" + "\n", true, false, false);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace);
            }
        }

        /// <summary>
        /// Регистрация реакции на событие "Создание документа"
        /// </summary>
        [CommandMethod("AddEvent")]
        public void AddEvent()
        {
            try { Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.DocumentActivated += new DocumentCollectionEventHandler(DocRefresh); }
            catch (System.Exception)
            {
                // ignored
            }
        }

        [CommandMethod("ChangeFullPathOnRelativePath")]
        public void ChangeFullPathOnRelativePath()
        {
            var editor = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor;
            var acDoc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var db = acDoc.Database;
            var check = new Condition();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                db.ResolveXrefs(true, false);
                var xg = db.GetHostDwgXrefGraph(true);
                new ObjectIdCollection();

                var xrefcount = xg.NumNodes - 1;
                if (xrefcount == 0)
                {
                }
                else
                {
                    for (var r = 1; r < (xrefcount + 1); r++)
                    {
                        var child = xg.GetXrefNode(r);

                        if (child.XrefStatus != XrefStatus.Resolved) continue;
                        var btr = (BlockTableRecord)tr.GetObject(child.BlockTableRecordId, OpenMode.ForWrite);

                        db.XrefEditEnabled = true;
                        try
                        {
                            var pathName = check.ParseGuid(Path.GetDirectoryName(btr.PathName));
                            if (pathName.Length != 0)
                            {
                                var childname = Path.GetFileName(btr.PathName);

                                var relativePath = @"..\" + pathName + @"\" + childname;

                                btr.PathName = relativePath;
                            }
                        }
                        catch (System.Exception ex) { editor.WriteMessage("\n Path Xref Not found: " + ex.Message + "\n" + ex.StackTrace); }
                    }
                }
                tr.Commit();
            }
        }

        /// <summary>
        /// Изменение путей внешних ссылок с абсолютных на относительные
        /// </summary>
        [CommandMethod("ChangeXRefPath")]
        public void ChangeXRefPathMethod()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var collection = new ObjectIdCollection();
            using (var db = new Database(false, true))
            {
                db.ReadDwgFile(@"c:\Temp\Test.dwg", FileOpenMode.OpenForReadAndWriteNoShare, false, "");
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    foreach (var btrId in bt)
                    {
                        var btr = tr.GetObject(btrId, OpenMode.ForRead)
                                                                as BlockTableRecord;
                        if (!btr.IsFromExternalReference) continue;

                        btr.UpgradeOpen();
                        var oldPath = btr.PathName;
                        var newPath = oldPath.Replace(@"C:\Temp\", ".");
                        btr.PathName = newPath;
                        collection.Add(btrId);
                        ed.WriteMessage(String.Format("{0}Old Path : {1} New Path : {2}",
                            Environment.NewLine, oldPath, newPath));
                    }
                    tr.Commit();
                }
                if (collection.Count > 0)
                {
                    db.ReloadXrefs(collection);
                }
                db.SaveAs(db.OriginalFileName, true, db.OriginalFileVersion, db.SecurityParameters);
            }
        }

        //требуется проверить данный код. Если внешние ссылки не найдены, то таким образом удалить их.
        //взято отсюда http://forums.autodesk.com/t5/NET/NET-c-unable-to-detaching-FileNotFound-xref-without-opening-the/m-p/3556778/highlight/true#M30162
        [CommandMethod("DetachXref")]
        public void detach_xref()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            var mainDrawingFile = @"C:\Temp\Test.dwg";

            var db = new Database(false, false);
            using (db)
            {
                try
                {
                    db.ReadDwgFile(mainDrawingFile, FileShare.ReadWrite, false, "");
                }
                catch (System.Exception)
                {
                    ed.WriteMessage("\nUnable to read the drawingfile.");
                    return;
                }
                var saveRequired = false;
                db.ResolveXrefs(true, false);
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var xg = db.GetHostDwgXrefGraph(true);

                    var xrefcount = xg.NumNodes;
                    for (var j = 0; j < xrefcount; j++)
                    {
                        var xrNode = xg.GetXrefNode(j);
                        var nodeName = xrNode.Name;

                        if (xrNode.XrefStatus == XrefStatus.FileNotFound)
                        {
                            var detachid = xrNode.BlockTableRecordId;

                            db.DetachXref(detachid);

                            saveRequired = true;
                            ed.WriteMessage("\nDetached successfully");

                            break;
                        }
                    }
                    tr.Commit();
                }

                if (saveRequired)
                    db.SaveAs(mainDrawingFile, DwgVersion.Current);
            }
        }

        /// <summary>
        /// Метод, проверяет флаги настроек в меню AutoCAD для настройки TDMS.
        /// </summary>
        public void DocRefresh(object senderObj, DocumentCollectionEventArgs docColDocActEvtArgs)
        {
            var prop = new Property();
            var ua = new Commands();
            var wt = new WorkTable();
            //следует вызывать только при активации документов,
            //т.к. если использовать при создании, то получатся двойные надписи в заголовках.

            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                var editor = doc.Editor;
                try
                {
                    if (prop.GetCheckUpdateAttr() == 1)
                    {
                        ua.UpdateAttributeWithoutMsg();
                        editor.WriteMessage("\n __UpdateAttr is OK \n");
                    }

                    if (prop.GetCheckUpdateXref() == 1)
                    {
                        XrefUpdate();
                        ReloadXRefs();
                        RefreshXref();
                        editor.WriteMessage("\n StartUpdateXREF is OK \n");
                    }
                    if (prop.GetCheckScheduleTable() == 1)
                    {
                        wt.UpdateTable();
                        editor.WriteMessage("\n __UpdateScheduleTable is OK \n");
                    }
                    if (prop.GetCheckSearchChangeXref() == 1)
                    {
                        editor.WriteMessage("\n __SearchChangeXREF is OK \n");
                    }
                    Technical.TitleDoc();
                }
                catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
            }
        }

        /// <summary>
        /// Выгрузка из ТДМС и сохранение главного файла и всех внешних ссылок (dwg и картинок) на HDD локально//
        /// </summary>
        [CommandMethod("EXPORTXREFFROMTDMS")]
        public void Exportxreffromtdms()
        {
            var editor = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                var check = new Condition();
                if (check.CheckTdmsProcess() == true)
                {
                    var acDoc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;

                    var filePath = SaveOptions.SaveDialog(acDoc);

                    var db = acDoc.Database;

                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        db.ResolveXrefs(true, false);

                        var xg = db.GetHostDwgXrefGraph(true);
                        var root = xg.RootNode;
                        var objcoll = new ObjectIdCollection();

                        var xrefcount = xg.NumNodes - 1;

                        if (xrefcount == 0)
                        {
                            editor.WriteMessage("\nNo xref found in drawing.");
                            acDoc.Database.SaveAs(filePath, true, DwgVersion.Current, acDoc.Database.SecurityParameters);
                        }
                        else
                        {
                            for (var r = 1; r < (xrefcount + 1); r++)
                            {
                                var child = xg.GetXrefNode(r);

                                if (child.XrefStatus == XrefStatus.Resolved)
                                {
                                    var btr = (BlockTableRecord)tr.GetObject(child.BlockTableRecordId, OpenMode.ForWrite);

                                    db.XrefEditEnabled = true;

                                    try
                                    {
                                        var childname = Path.GetFileName(child.Database.Filename);
                                        var originalRelativePath = child.Database.Filename;

                                        var directoryName = Path.GetDirectoryName(filePath);
                                        Directory.CreateDirectory(directoryName + @"\XREF\");
                                        var newpath = directoryName + @"\XREF\" + childname;

                                        File.Copy(originalRelativePath, newpath, true);

                                        btr.PathName = @".\XREF\" + childname;

                                        RefremapFromTdms(originalRelativePath, newpath); //изменение путей внешних ссылок для внешних ссылок

                                        editor.WriteMessage("\nxref old path: " + originalRelativePath + "\n");
                                        editor.WriteMessage("\nxref new path: " + newpath + "\n" + "XREF changed" + "\n");
                                    }
                                    catch (System.Exception ex) { editor.WriteMessage("\nException caught: " + "\n" + ex + "\n"); }
                                }
                            }
                        }
                        tr.Commit();
                        acDoc.Database.SaveAs(filePath, true, DwgVersion.Current, acDoc.Database.SecurityParameters);
                    }
                    FindImages(filePath);
                }
                else
                {
                    editor.WriteMessage("\n Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\nException caught: " + ex.Message + "\n" + ex.StackTrace + "\n");
            }
        }

        [CommandMethod("REFRESHMAP")]
        public void RefreshMap()
        {
            var editor = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor;

            try
            {
                var check = new Condition();
                if (check.CheckTdmsProcess() == true)
                {
                    var acDoc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
                    //путь к файлу, включая название файла, который является текущим в Autocad.
                    var filePath = acDoc.Name;

                    editor.WriteMessage("\n File Name: " + filePath);

                    var db = acDoc.Database;

                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        editor.WriteMessage("\n----------XRefs_Details----------");
                        db.ResolveXrefs(true, false);
                        var xg = db.GetHostDwgXrefGraph(true);
                        var root = xg.RootNode;
                        var objcoll = new ObjectIdCollection();

                        var tdmsApp = new TDMSApplication();
                        Autodesk.AutoCAD.ApplicationServices.Core.Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Minimized;
                        tdmsApp.Visible = false;
                        tdmsApp.Visible = true;

                        var moduleName = "CMD_SYSLIB";
                        var selectPsdFolder = "SelectPsdFolder";
                        var loadFileToPsd = "LoadFileToPSD";
                        var flMainFile = false;

                        string pathName = tdmsApp.ExecuteScript(moduleName, selectPsdFolder);
                        editor.WriteMessage("\n Path output from TDMS " + pathName);

                        string parseGuidFromFile = null;
                        var xrefcount = xg.NumNodes - 1;

                        if (xrefcount == 0)
                        {
                            var Check = new Condition();
                            parseGuidFromFile = Check.ParseGuid(pathName);
                            editor.WriteMessage("\n Parse GUID: " + parseGuidFromFile);

                            flMainFile = true;

                            editor.WriteMessage("\n No xref found in drawing");
                            acDoc.Database.SaveAs(filePath, true, DwgVersion.Current, acDoc.Database.SecurityParameters);
                            string mainFile = tdmsApp.ExecuteScript(moduleName, loadFileToPsd, filePath, parseGuidFromFile, flMainFile);
                        }
                        else
                        {
                            var Check = new Condition();
                            parseGuidFromFile = Check.ParseGuid(pathName);
                            editor.WriteMessage("\n Parse GUID: " + parseGuidFromFile);
                            for (var r = 1; r < (xrefcount + 1); r++)
                            {
                                var child = xg.GetXrefNode(r);
                                if (child.XrefStatus == XrefStatus.Resolved)
                                {
                                    var btr = (BlockTableRecord)tr.GetObject(child.BlockTableRecordId, OpenMode.ForWrite);

                                    db.XrefEditEnabled = true;
                                    try
                                    {
                                        var originalPath = btr.PathName;
                                        var childname = Path.GetFileName(originalPath);

                                        editor.WriteMessage("\n childname:         {0}", childname);
                                        editor.WriteMessage("\n moduleName:        {0}", moduleName);
                                        editor.WriteMessage("\n LoadFileToPSD:     {0}", loadFileToPsd);
                                        editor.WriteMessage("\n originalPath:      {0}", originalPath);
                                        editor.WriteMessage("\n parseGuidFromFile: {0}", parseGuidFromFile);
                                        editor.WriteMessage("\n fl_main_file:      {0}", flMainFile.ToString());

                                        string newpath = tdmsApp.ExecuteScript(moduleName, loadFileToPsd, originalPath, parseGuidFromFile, flMainFile) + @"\" + childname;

                                        btr.PathName = newpath;

                                        //изменение путей внешних ссылок для внешних ссылок
                                        Refremap(originalPath, newpath);

                                        editor.WriteMessage("\n xref old path: " + originalPath);
                                        editor.WriteMessage("\n xref new path: " + newpath + " xref fixed !!!");
                                    }
                                    catch (System.Exception ex) { editor.WriteMessage("\n Path Xref Not found: " + ex.Message + "\n" + ex.StackTrace); }
                                }
                            }
                            flMainFile = true;

                            editor.WriteMessage("\n moduleName:         {0}", moduleName);
                            editor.WriteMessage("\n LoadFileToPSD:      {0}", loadFileToPsd);
                            editor.WriteMessage("\n filePath:           {0}", filePath);
                            editor.WriteMessage("\n parseGuidFromFile:  {0}", parseGuidFromFile);
                            editor.WriteMessage("\n fl_main_file:       {0}", flMainFile.ToString());
                            string mainFile = tdmsApp.ExecuteScript(moduleName, loadFileToPsd, filePath, parseGuidFromFile, flMainFile);
                        }
                        tr.Commit();
                        acDoc.Database.SaveAs(filePath, true, DwgVersion.Current, acDoc.Database.SecurityParameters);
                    }
                    RefreshXref();
                    Autodesk.AutoCAD.ApplicationServices.Core.Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Maximized;
                }
                else { Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog("Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного."); }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// переподгрузка всех внешних ссылок
        /// </summary>
        [CommandMethod("RefreshXREF")]
        public void RefreshXref()
        {
            var editor = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                var cvPath = new CurrentVersion();
                var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;

                var myKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(cvPath.pathLanguage());
                var path = (String)(myKey.GetValue("LocalRootFolder"));

                var local = path.ToLower();
                string parseGuid = null;
                var reg = new Regex("rus", RegexOptions.IgnoreCase);
                var mc = reg.Matches(local);
                foreach (Match mat in mc)
                {
                    parseGuid = mat.Value.ToString();
                }
                if (parseGuid == "rus")
                {
                    //Для русской версии AutoCAD
                    doc.SendStringToExecute("-ССЫЛКА" + "\n" + "Обновить" + "\n" + "*" + "\n", true, false, false);
                    doc.SendStringToExecute("-ИЗОБ" + "\n" + "Обновить" + "\n" + "*" + "\n", true, false, false);
                }
                else
                {
                    //Для английской версии AutoCAD
                    doc.SendStringToExecute("-XREF" + "\n" + "R" + "\n" + "*" + "\n", true, false, false);
                    doc.SendStringToExecute("-IMAGE" + "\n" + "R" + "\n" + "*" + "\n", true, false, false);
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex + "\n" + ex.StackTrace); }
        }

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //Сохранение главного файла и объектов внешних ссылок (dwg и картинок) в TDMS (внешние ссылки в каталоге XREF в TDMS)//
        //                                      Абсолютные (полные) пути в ссылках                                           //
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //Сохранение главного файла и объектов внешних ссылок (dwg и картинок) в TDMS (внешние ссылки в каталоге XREF в TDMS)//
        //                         Относительные (..\GUID объекта\Имя файла с расширением) пути в ссылках                    //
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        [CommandMethod("RelativePath")]
        public void RelativePath()
        {
            var editor = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor;

            try
            {
                //проверка на запущенный процесс TDMS.exe
                var check = new Condition();
                if (check.CheckTdmsProcess() == true)
                {
                    var acDoc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
                    //путь к файлу, включая название файла, который является текущим в Autocad.
                    var filePath = acDoc.Name;

                    editor.WriteMessage("\n File Name: " + filePath);

                    var db = acDoc.Database;

                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        db.ResolveXrefs(true, false);
                        var xg = db.GetHostDwgXrefGraph(true);
                        var root = xg.RootNode;
                        var objcoll = new ObjectIdCollection();

                        var tdmsApp = new TDMSApplication();
                        Autodesk.AutoCAD.ApplicationServices.Core.Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Minimized;
                        tdmsApp.Visible = false;
                        tdmsApp.Visible = true;

                        const string moduleName = "CMD_SYSLIB";
                        const string selectPsdFolder = "SelectPsdFolder";
                        const string loadFileToPsd = "LoadFileToPSD";
                        var flMainFile = false;

                        string pathName = tdmsApp.ExecuteScript(moduleName, selectPsdFolder);

                        editor.WriteMessage("\n Path output from TDMS " + pathName);

                        string parseGuidFromFile = null;
                        var xrefcount = xg.NumNodes - 1;

                        if (xrefcount == 0)
                        {
                            var Check = new Condition();
                            parseGuidFromFile = Check.ParseGuid(pathName);
                            editor.WriteMessage("\n Parse GUID: " + parseGuidFromFile);

                            flMainFile = true;

                            editor.WriteMessage("\n No xref found in drawing");
                            acDoc.Database.SaveAs(filePath, true, DwgVersion.Current, acDoc.Database.SecurityParameters);
                            tdmsApp.ExecuteScript(moduleName, loadFileToPsd, filePath, parseGuidFromFile, flMainFile);
                        }
                        else
                        {
                            var Check = new Condition();
                            parseGuidFromFile = Check.ParseGuid(pathName);
                            editor.WriteMessage("\n Parse GUID: " + parseGuidFromFile);

                            for (var r = 1; r < (xrefcount + 1); r++)
                            {
                                var child = xg.GetXrefNode(r);

                                if (child.XrefStatus == XrefStatus.Resolved)
                                {
                                    var btr = (BlockTableRecord)tr.GetObject(child.BlockTableRecordId, OpenMode.ForWrite);

                                    db.XrefEditEnabled = true;
                                    try
                                    {
                                        var childname = Path.GetFileName(btr.PathName);
                                        var originalRelativePath = child.Database.Filename;

                                        //editor.WriteMessage("\n childname:            {0}", childname);
                                        //editor.WriteMessage("\n moduleName:           {0}", moduleName);
                                        //editor.WriteMessage("\n LoadFileToPSD:        {0}", loadFileToPsd);
                                        //editor.WriteMessage("\n originalRelativePath: {0}", originalRelativePath);
                                        //editor.WriteMessage("\n parseGuidFromFile:    {0}", parseGuidFromFile);
                                        //editor.WriteMessage("\n fl_main_file:         {0}", false.ToString());

                                        string newpath = tdmsApp.ExecuteScript(moduleName, loadFileToPsd, originalRelativePath, parseGuidFromFile, flMainFile) + @"\" + childname;
                                        newpath = Check.ParseGuid(newpath);

                                        var relativePath = @"..\" + newpath + @"\" + childname;

                                        btr.PathName = relativePath;

                                        //изменение путей внешних ссылок для внешних ссылок
                                        Refremap(originalRelativePath, relativePath);

                                        editor.WriteMessage("\n xref old path: " + originalRelativePath);
                                        editor.WriteMessage("\n xref relative path: " + relativePath + " xref fixed !!!");
                                    }
                                    catch (System.Exception ex) { editor.WriteMessage("\n Path Xref Not found: " + ex.Message + "\n" + ex.StackTrace); }
                                }
                            }
                            flMainFile = true;

                            //editor.WriteMessage("\n moduleName:         {0}", moduleName);
                            //editor.WriteMessage("\n LoadFileToPSD:      {0}", loadFileToPsd);
                            //editor.WriteMessage("\n filePath:           {0}", filePath);
                            //editor.WriteMessage("\n parseGuidFromFile:  {0}", parseGuidFromFile);
                            //editor.WriteMessage("\n fl_main_file:       {0}", flMainFile.ToString());
                            acDoc.Database.SaveAs(filePath, true, DwgVersion.Current, acDoc.Database.SecurityParameters);
                            tdmsApp.ExecuteScript(moduleName, loadFileToPsd, filePath, parseGuidFromFile, flMainFile);
                        }
                        tr.Commit();
                    }
                    Autodesk.AutoCAD.ApplicationServices.Core.Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Maximized;
                }
                else { Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog("Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного."); }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        [CommandMethod("ReloadXRefs")]
        public void ReloadXRefs()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var editor = doc.Editor;
            var ids = new ObjectIdCollection();
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var table = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;
                    foreach (var id in table)
                    {
                        try
                        {
                            var record = tr.GetObject(id, OpenMode.ForWrite) as BlockTableRecord;
                            if (!record.IsFromExternalReference) continue;
                            ids.Add(id);
                            db.ReloadXrefs(ids);
                        }
                        catch (System.Exception ex)
                        {
                            editor.WriteMessage("\n Link not found: " + id.Database.Filename.ToString());
                        }
                    }
                    tr.Commit();
                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                editor.WriteMessage("\n ReloadXRefs fail. \n");
            }
        }

        /// <summary>
        /// Вставка внешней ссылки. В зависимости от типа файла вызывает ту или иную функцию.
        /// </summary>
        [CommandMethod("TDMSXREFADD", CommandFlags.NoBlockEditor)]
        public void Tdmsxrefadd()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var editor = doc.Editor;
            try
            {
                var check = new Condition();
                if (check.CheckPath() == true)
                {
                    if (check.CheckTdmsProcess() == true)
                    {
                        var tdmsApp = new TDMSApplication();
                        Autodesk.AutoCAD.ApplicationServices.Core.Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Minimized;
                        tdmsApp.Visible = false;
                        tdmsApp.Visible = true;

                        const string moduleName = "CMD_SYSLIB";
                        const string functionName = "CheckOutSelObj";

                        string pathName = tdmsApp.ExecuteScript(moduleName, functionName);
                        var extension = Path.GetExtension(pathName);

                        if (extension.ToLower() == ".jpg" | extension.ToLower() == ".tif" | extension.ToLower() == ".png" | extension.ToLower() == ".bmp")
                        {
                            Tdmsxrefimg(pathName);
                        }
                        else if (extension.ToLower() == ".dwg")
                        {
                            Tdmsxrefdwg(pathName);
                            ChangeFullPathOnRelativePath();
                        }
                        else
                        {
                            editor.WriteMessage("\n Некорректный формат выбранного файла.");
                        }
                    }
                    else
                    {
                        Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog("\n Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                    }
                }
                else
                {
                    Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog("\n Документ не принадлежит TDMS!");
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Вставка DWG как внешней ссылки
        /// </summary>
        public void Tdmsxrefdwg(string pathFile)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var docName = doc.Name;
            var editor = doc.Editor;
            try
            {
                var tdmsApp = new TDMSApplication();
                var pointOptions = new PromptPointOptions("\n УКАЖИТЕ ТОЧКУ: ");

                var db = doc.Database;

                var pathName = pathFile;
                editor.WriteMessage("\n Return GUID from ExecuteScript " + pathName);
                if (pathName != "")
                {
                    //Берём GUID из пути файла выгруженного на диск
                    var guidFromFile = docName;
                    string parseGuidFromFile = null;
                    var regFf = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                    var mcFf = regFf.Matches(guidFromFile);
                    foreach (Match mat in mcFf)
                    {
                        parseGuidFromFile += mat.Value.ToString();
                    }
                    parseGuidFromFile = parseGuidFromFile.Remove(0, 38);
                    //Берём GUID из выбранного в TDMS файла - внешней ссылки
                    var guidFromTdms = pathName;
                    string parseGuidFromTdms = null;
                    var regFt = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                    var mcFt = regFt.Matches(guidFromTdms);
                    foreach (Match mat in mcFt)
                    {
                        parseGuidFromTdms += mat.Value.ToString();
                    }
                    parseGuidFromTdms = parseGuidFromTdms.Remove(0, 38);
                    var tdmsObj = tdmsApp.GetObjectByGUID(parseGuidFromTdms);
                    if (parseGuidFromFile != parseGuidFromTdms)
                    {
                        var guid = pathName;
                        string parseGuid = null;
                        var reg = new Regex(".dwg", RegexOptions.IgnoreCase);
                        var mc = reg.Matches(guid);
                        foreach (Match mat in mc)
                        {
                            parseGuid += mat.Value;
                        }
                        if (parseGuid.ToLower() == ".dwg")
                        {
                            editor.WriteMessage("\n PathNameFromTDMS: " + parseGuidFromTdms);
                            editor.WriteMessage("\n docName: " + parseGuidFromFile);

                            var mainFile = tdmsObj.Files.Main;

                            mainFile.CheckOut(pathName);

                            Autodesk.AutoCAD.ApplicationServices.Core.Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Maximized;
                            Thread.Sleep(1000);
                            var pointResult = editor.GetPoint(pointOptions);
                            var pointFrame = pointResult.Value;
                            var pointX = pointFrame.X;
                            var pointY = pointFrame.Y;

                            var path = pathName;
                            var xRefId = db.AttachXref(Path.Combine(Environment.GetFolderPath(
                                Environment.SpecialFolder.MyDocuments), path), path);
                            using (var tr = db.TransactionManager.StartTransaction())
                            {
                                //Создаём экземпляр вхождения блока для подключенной ссылки
                                var br = new BlockReference(new Point3d(pointX, pointY, 0), xRefId);
                                var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                                //Вставлять вхождение блока в пространстве Model или пространство Paper в зависимости от переменной "tilemode"
                                var space = Autodesk.AutoCAD.ApplicationServices.Core.Application.GetSystemVariable("tilemode").ToString();
                                if (space != "1")
                                {
                                    var papperSpace = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
                                    papperSpace.AppendEntity(br);
                                    tr.AddNewlyCreatedDBObject(br, true);
                                }
                                else
                                {
                                    var modalSpace = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                                    modalSpace.AppendEntity(br);
                                    tr.AddNewlyCreatedDBObject(br, true);
                                }
                                tr.Commit();
                            }
                            XrefUpdate();
                        }
                        else
                        {
                            Autodesk.AutoCAD.ApplicationServices.Core.Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Maximized;
                            Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog("\n Несоответствие формата выбранного файла!");
                        }
                    }
                    else
                    {
                        Autodesk.AutoCAD.ApplicationServices.Core.Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Maximized;
                        editor.WriteMessage("\n PathNameFromTDMS: " + parseGuidFromTdms);
                        editor.WriteMessage("\n docName: " + parseGuidFromFile);
                        Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog("Нельзя использовать исходный чертёж как ссылку на самого себя!");
                    }
                }
                else
                {
                    Autodesk.AutoCAD.ApplicationServices.Core.Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Maximized;
                    editor.WriteMessage("\n Вставка внешней ссылки отменена!");
                }
                ////Внедряем содержимое ссылки в чертёж
                ////ВНИМАНИЕ!
                ////Следующий фрагмент кода должен находиться ПОСЛЕ блока "using", т.к. должен быть выполнен
                ////после строки кода tr.Commit(); (т.е. когда изменения будут зафиксированы в базе данных)
                ////AcDb.ObjectIdCollection ids = new AcDb.ObjectIdCollection();
                ////ids.Add(xRefId);
                ////db.BindXrefs(ids, true);
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Вставка внешней ссылки. В зависимости от типа файла вызывает ту или иную функцию.
        /// </summary>
        public void Tdmsxrefimg(string pathFile)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var editor = doc.Editor;
            try
            {
                var tdmsApp = new TDMSApplication();

                Database acCurDb;
                var pointOptions = new PromptPointOptions("\n УКАЖИТЕ ТОЧКУ: ");
                acCurDb = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;

                var pathName = pathFile;

                editor.WriteMessage("\n PathName: " + Path.GetExtension(pathName));

                if (pathName != "")
                {
                    TDMSObject tdmsObj;
                    //Берём GUID из пути файла выгруженного на диск
                    var docName = doc.Name;
                    var guidFromFile = docName;
                    string parseGuidFromFile = null;
                    var regFf = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                    var mcFf = regFf.Matches(guidFromFile);
                    foreach (Match mat in mcFf)
                    {
                        parseGuidFromFile += mat.Value.ToString();
                    }
                    parseGuidFromFile = parseGuidFromFile.Remove(0, 38);
                    //Берём GUID из выбранного в TDMS файла - внешней ссылки
                    var guidFromTdms = pathName;
                    string parseGuidFromTdms = null;
                    var regFt = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
                    var mcFt = regFt.Matches(guidFromTdms);
                    foreach (Match mat in mcFt)
                    {
                        parseGuidFromTdms += mat.Value.ToString();
                    }
                    parseGuidFromTdms = parseGuidFromTdms.Remove(0, 38);
                    tdmsObj = tdmsApp.GetObjectByGUID(parseGuidFromTdms);
                    var rastrName = tdmsObj.Files.Main.FileName;

                    if (parseGuidFromFile != parseGuidFromTdms)
                    {
                        string parseGuidJpeg = null;
                        string parseGuidTiff = null;
                        string parseGuidPng = null;
                        string parseGuidBmp = null;

                        string parseGuidJpegUpper = null;
                        string parseGuidTiffUpper = null;
                        string parseGuidPngUpper = null;
                        string parseGuidBmpUpper = null;

                        var regJpeg = new Regex(".jpg", RegexOptions.IgnoreCase);
                        var mcJpeg = regJpeg.Matches(pathName);
                        foreach (Match mat in mcJpeg)
                        {
                            parseGuidJpeg += mat.Value.ToString();
                        }

                        var regTiff = new Regex(".tif", RegexOptions.IgnoreCase);
                        var mcTiff = regTiff.Matches(pathName);
                        foreach (Match mat in mcTiff)
                        {
                            parseGuidTiff += mat.Value.ToString();
                        }

                        var regPng = new Regex(".png", RegexOptions.IgnoreCase);
                        var mcPng = regPng.Matches(pathName);
                        foreach (Match mat in mcPng)
                        {
                            parseGuidPng += mat.Value.ToString();
                        }

                        var regBmp = new Regex(".bmp", RegexOptions.IgnoreCase);
                        var mcBmp = regBmp.Matches(pathName);
                        foreach (Match mat in mcBmp)
                        {
                            parseGuidBmp += mat.Value.ToString();
                        }

                        if (parseGuidJpeg == ".jpg" | parseGuidTiff == ".tif" | parseGuidTiff == ".TIF" | parseGuidJpeg == ".JPG" | parseGuidPng == ".png" | parseGuidPngUpper == ".PNG" | parseGuidBmp == ".bmp" | parseGuidBmpUpper == ".BMP")
                        {
                            editor.WriteMessage("\n jpg: " + parseGuidJpeg);
                            editor.WriteMessage("\n tif: " + parseGuidTiff);
                            editor.WriteMessage("\n png: " + parseGuidPng);
                            editor.WriteMessage("\n bmp: " + parseGuidBmp);

                            editor.WriteMessage("\n JPG: " + parseGuidJpegUpper);
                            editor.WriteMessage("\n TIF: " + parseGuidTiffUpper);
                            editor.WriteMessage("\n png: " + parseGuidPngUpper);
                            editor.WriteMessage("\n bmp: " + parseGuidBmpUpper);

                            var mainFile = tdmsObj.Files.Main;
                            mainFile.CheckOut(pathName);

                            parseGuidFromTdms = null;

                            Autodesk.AutoCAD.ApplicationServices.Core.Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Maximized;

                            using (var acTrans = acCurDb.TransactionManager.StartTransaction())
                            {
                                var pointResult = editor.GetPoint(pointOptions);
                                var pointFrame = pointResult.Value;
                                var path = pathName;
                                var strImgName = rastrName;
                                RasterImageDef acRasterDef;
                                var bRasterDefCreated = false;
                                ObjectId acImgDefId;
                                var acImgDctId = RasterImageDef.GetImageDictionary(acCurDb);
                                if (acImgDctId.IsNull)
                                {
                                    acImgDctId = RasterImageDef.CreateImageDictionary(acCurDb);
                                }
                                var acImgDict = acTrans.GetObject(acImgDctId, OpenMode.ForRead) as DBDictionary;
                                if (acImgDict.Contains(strImgName))
                                {
                                    acImgDefId = acImgDict.GetAt(strImgName);
                                    acRasterDef = acTrans.GetObject(acImgDefId, OpenMode.ForWrite) as RasterImageDef;
                                }
                                else
                                {
                                    var acRasterDefNew = new RasterImageDef();

                                    var check = new Condition();

                                    var parseGuid = check.ParseGuid(Path.GetDirectoryName(path));

                                    var childname = Path.GetFileName(path);

                                    var relativePath = @"..\" + parseGuid + @"\" + childname;

                                    acRasterDefNew.SourceFileName = relativePath;
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
                                using (var acRaster = new RasterImage())
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
                                    var insPt = new Point3d(pointFrame.X, pointFrame.Y, 0.0);
                                    var coordinateSystem = new CoordinateSystem3d(insPt, width * 2, height * 2);
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
                        }
                        else
                        {
                            Autodesk.AutoCAD.ApplicationServices.Core.Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Maximized;
                            Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog("\n Несоответствие формата выбранного файла!");
                        }
                    }
                    else
                    {
                        Autodesk.AutoCAD.ApplicationServices.Core.Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Maximized;
                        editor.WriteMessage("\n PathNameFromTDMS: " + parseGuidFromTdms);
                        editor.WriteMessage("\n docName: " + parseGuidFromFile);
                        Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog("\n Нельзя использовать исходный чертёж как ссылку на самого себя!");
                    }
                }
                else
                {
                    Autodesk.AutoCAD.ApplicationServices.Core.Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Maximized;
                    editor.WriteMessage("\n Вставка внешней ссылки отменена!");
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Обновление всех атрибутов и внешних ссылок
        /// </summary>
        [CommandMethod("UpdateXrefAttr")]
        public void UpdateXrefAttr()
        {
            var update = new Commands();
            XrefUpdate();
            update.UpdateAttribute();
            RefreshXref();
        }

        [CommandMethod("XREFSYNCHRO")]
        public void XrefSynchro()
        {
            var acDoc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var editor = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage("Функция в разработке");

            //try
            //{
            //    Condition ChangePath = new Condition();
            //    string strDwgName = acDoc.Name;
            //    //{
            //    //    if (checkP.CheckPath() == true)
            //    //    {
            //    //TDMSApplication tdmsApp = new TDMSApplication();
            //    //TDMSObject tdmsObj = null;
            //    //object obj = Application.GetSystemVariable("DWGTITLED");
            //    List<string> masPath = new List<string>();
            //    Database db = acDoc.Database;
            //    //Возвращение пути к файлу чертежа
            //    string guid = strDwgName;
            //    string parseGuid = null;

            //    using (Transaction tr = db.TransactionManager.StartTransaction())
            //    {
            //        editor.WriteMessage("\n PATH XREFERENCES: ");
            //        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
            //        foreach (ObjectId id in bt)
            //        {
            //            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForWrite);
            //            //if (btr.IsFromExternalReference)
            //            //{
            //            if (btr.XrefStatus.ToString() == "Resolved")
            //            {
            //                if (ChangePath.StartCheckPath(btr.PathName) == false)
            //                {
            //                    editor.WriteMessage("\n PathName: {0}", btr.PathName);
            //                    editor.WriteMessage("\n Name: {0}", btr.Name);
            //                    editor.WriteMessage("\n Status: {0}", btr.XrefStatus.ToString());
            //                    masPath.Add(btr.PathName);
            //                }
            //            }
            //            //}
            //        }
            //    }

            //for (int j = 0; j < masPath.Count; j++)
            //{
            //вызываю функцию Лили, передаю ей путь каждой внешней ссылки или массив ссылок, получаю время изменения документа в TDMS, если время одинаковое,
            //то ничего не делаю, если разное, то объект должен выгрузиться и должно пройти обновление внешних ссылок.
            //    editor.WriteMessage("\n ArrayPath: {0}", masPath[j]);
            //}

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
            //}
            //catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Обновление внешних ссылок
        /// </summary>
        [CommandMethod("XREFUpdate")]
        public void XrefUpdate()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var editor = doc.Editor;
            var ids = new ObjectIdCollection();
            //Получение пути к объектам c:\temp\tdms\userGUID\
            //из пути к главному файлу вычитается GUID объекта ({39 символов + \})
            try
            {
                var check = new Condition();
                if (check.CheckPath())
                {
                    if (check.CheckTdmsProcess())
                    {
                        var pathMainFile = Path.GetDirectoryName(doc.Name).ToString();
                        //editor.WriteMessage("\n 1: "+ PathMainFile);
                        pathMainFile = pathMainFile.Remove(pathMainFile.Length - 39);
                        //editor.WriteMessage("\n 2: " + PathMainFile);
                        var db = doc.Database;

                        string parseGuid = null;

                        var tdmsApp = new TDMSApplication();
                        TDMSObject tdmsObj = null;

                        var reg = new Regex("{.*?}", RegexOptions.IgnoreCase);

                        using (var tr = db.TransactionManager.StartTransaction())
                        {
                            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
                            foreach (var id in bt)
                            {
                                var btr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForWrite);
                                if (btr.IsFromExternalReference)
                                {
                                    //принадлежность внешней ссылки TDMS с проверкой на относительность пути
                                    //editor.WriteMessage("\n 3: " + Check.StartCheckPath(btr.PathName));
                                    //editor.WriteMessage("\n 4: " + Check.StartCheckRelativePath(btr.PathName));
                                    //если внешняя ссылка из tdms
                                    if (check.CheckPath(btr.PathName) | check.CheckRelativePath(btr.PathName))
                                    {
                                        //editor.WriteMessage("\n 5: " + btr.PathName.ToString());
                                        try
                                        {
                                            //парсинг пути, получение актуального GUID объекта
                                            var mc = reg.Matches(btr.PathName);
                                            foreach (Match mat in mc)
                                            {
                                                parseGuid += mat.Value.ToString();
                                            }

                                            //editor.WriteMessage("\n 6: " + System.Convert.ToUInt32(parseGuid.Length));
                                            //editor.WriteMessage("\n 7: " + parseGuid.Length);

                                            switch (Convert.ToUInt32(parseGuid.Length))
                                            {
                                                case 38:
                                                    {
                                                        try
                                                        {
                                                            //editor.WriteMessage("\n Relative path.");
                                                            //editor.WriteMessage("\n 8: " + parseGuid.Length);

                                                            //Получаем объект ТДМС
                                                            if (tdmsApp.GetObjectByGUID(parseGuid) != null)
                                                            {
                                                                tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);

                                                                var mainFile = tdmsObj.Files.Main;

                                                                //имя файла из ТДМС
                                                                var fileNameInTdms = Path.GetFileName(mainFile.FileName).ToString();
                                                                //имя файла из ссылки, которая ссылается на файл ТДМС
                                                                var fileNameInAutoCadXref = Path.GetFileName(btr.PathName).ToString();

                                                                //editor.WriteMessage("\n 9  fileNameInTDMS: "        + fileNameInTDMS);
                                                                //editor.WriteMessage("\n 10 fileNameInAutoCADXref: " + fileNameInAutoCADXref);
                                                                //editor.WriteMessage("\n 11 RelativePath: "          + PathMainFile + @"\" + parseGuid + @"\" + fileNameInTDMS);

                                                                mainFile.CheckOut(pathMainFile + @"\" + parseGuid + @"\" + fileNameInAutoCadXref);
                                                                parseGuid = null;
                                                                ids.Add(id);
                                                                db.ReloadXrefs(ids);
                                                            }
                                                            else
                                                            {
                                                                parseGuid = null;

                                                                editor.WriteMessage("\n File " + btr.PathName + " not found in TDMS.");
                                                            }
                                                        }
                                                        catch (System.Exception ex)
                                                        {
                                                            parseGuid = null;
                                                            editor.WriteMessage("\n Exception caught: " + ex);
                                                        }
                                                        break;
                                                    }
                                                case 76:
                                                    {
                                                        try
                                                        {
                                                            //editor.WriteMessage("\n Full path.");
                                                            //editor.WriteMessage("\n 8: " + parseGuid.Length);

                                                            parseGuid = parseGuid.Remove(0, 38);
                                                            //editor.WriteMessage("\n 9: " + parseGuid);
                                                            if (tdmsApp.GetObjectByGUID(parseGuid) != null)
                                                            {
                                                                tdmsObj = tdmsApp.GetObjectByGUID(parseGuid);

                                                                var mainFile = tdmsObj.Files.Main;

                                                                //имя файла из ТДМС
                                                                var fileNameInTdms = Path.GetFileName(mainFile.FileName).ToString();
                                                                //имя файла из ссылки, которая ссылается на файл ТДМС
                                                                var fileNameInAutoCadXref = Path.GetFileName(btr.PathName).ToString();
                                                                //наименование директории
                                                                var directoryNameInAutoCadXref = Path.GetDirectoryName(btr.PathName);

                                                                //editor.WriteMessage("\n 10  fileNameInTDMS: " + fileNameInTDMS);
                                                                //editor.WriteMessage("\n 11 fileNameInAutoCADXref: " + fileNameInAutoCADXref);
                                                                //editor.WriteMessage("\n 12 directoryNameInAutoCADXref: " + directoryNameInAutoCADXref);

                                                                mainFile.CheckOut(directoryNameInAutoCadXref + @"\" + fileNameInAutoCadXref);
                                                                parseGuid = null;
                                                                ids.Add(id);
                                                                db.ReloadXrefs(ids);
                                                            }
                                                            else
                                                            {
                                                                parseGuid = null;
                                                                editor.WriteMessage("\n File " + btr.PathName + " not found in TDMS.");
                                                            }
                                                        }
                                                        catch (System.Exception ex)
                                                        {
                                                            parseGuid = null;
                                                            editor.WriteMessage("\n Exception caught: " + ex);
                                                        }
                                                        break;
                                                    }
                                            }
                                        }
                                        catch (System.Exception ex)
                                        {
                                            editor.WriteMessage("\n Link is out TDMS.");
                                            editor.WriteMessage("\n Path Link is out TDMS: " + btr.PathName);
                                            parseGuid = null;
                                        }
                                    }
                                    else
                                    {
                                        editor.WriteMessage("\n Path Link is out TDMS: " + btr.PathName);
                                        parseGuid = null;
                                    }
                                    RefreshXref();
                                }
                            }
                        }
                        FindImages();
                    }
                    else
                    {
                        editor.WriteMessage("\nНевозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                    }
                }
                else
                {
                    editor.WriteMessage("\nДокумент не принадлежит TDMS!");
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\nException caught: {0}\n", ex); }
        }

        private void Initialize()
        {
        }

        //////////////////////////////////
        //Обновление всех внешних ссылок//
        //////////////////////////////////
        private void Refremap(string pathXref, string newPath)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var editor = doc.Editor;
            var db = new Database(false, true);
            using (db)
            {
                try
                {
                    db.ReadDwgFile(pathXref, FileShare.ReadWrite, false, "");
                }
                catch (System.Exception ex)
                {
                    editor.WriteMessage("\n Unable to read the drawingfile: " + ex.Message + "\n" + ex.StackTrace);
                    return;
                }
                try
                {
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        editor.WriteMessage("\n --------Xrefs Details--------");
                        db.ResolveXrefs(true, false);

                        var xg = db.GetHostDwgXrefGraph(true);
                        var root = xg.RootNode;

                        var objcoll = new ObjectIdCollection();

                        var xrefcount = xg.NumNodes - 1;

                        if (xrefcount == 0)
                        {
                            editor.WriteMessage("\n No xrefs found in the drawing");
                            return;
                        }
                        else
                        {
                            for (var r = 1; r < (xrefcount + 1); r++)
                            {
                                var child = xg.GetXrefNode(r);
                                if (child.XrefStatus == XrefStatus.Resolved)
                                {
                                    var btr = (BlockTableRecord)tr.GetObject(child.BlockTableRecordId, OpenMode.ForWrite);

                                    db.XrefEditEnabled = true;

                                    var originalpath = btr.PathName;
                                    var childname = Path.GetFileName(originalpath);
                                    var newpath = newPath + childname;

                                    Refremap(originalpath, newpath);
                                    btr.PathName = newpath;

                                    editor.WriteMessage("\n xref old path: " + originalpath);
                                    editor.WriteMessage("\n xref new path: " + newpath + " xref fixed !!");
                                }
                            }
                            db.SaveAs(pathXref, true, DwgVersion.Current, doc.Database.SecurityParameters);
                        }
                        tr.Commit();
                    }
                }
                catch (System.Exception ex) { editor.WriteMessage("Work with base drawing Error: " + ex.Message + "\n" + ex.StackTrace); }
            }
        }

        /// <summary>
        /// Выгрузка из TDMS внешних ссылок, которые вставлены во внешние ссылки в чертеже
        /// </summary>
        private void RefremapFromTdms(string pathXref, string newPath)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var editor = doc.Editor;

            var db = new Database(false, true);
            using (db)
            {
                try
                {
                    db.ReadDwgFile(pathXref, FileShare.ReadWrite, false, "");
                }
                catch (System.Exception ex)
                {
                    return;
                }
                try
                {
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        db.ResolveXrefs(true, false);

                        var xg = db.GetHostDwgXrefGraph(true);
                        var root = xg.RootNode;

                        var objcoll = new ObjectIdCollection();

                        var xrefcount = xg.NumNodes - 1;

                        if (xrefcount == 0)
                        {
                            editor.WriteMessage("\nNo xrefs found in the drawing\n");
                            return;
                        }
                        else
                        {
                            for (var r = 1; r < (xrefcount + 1); r++)
                            {
                                var child = xg.GetXrefNode(r);
                                if (child.XrefStatus == XrefStatus.Resolved)
                                {
                                    var btr = (BlockTableRecord)tr.GetObject(child.BlockTableRecordId, OpenMode.ForWrite);

                                    db.XrefEditEnabled = true;

                                    var originalpath = btr.PathName;
                                    var childname = Path.GetFileName(originalpath);
                                    var newpath = newPath + childname;

                                    RefremapFromTdms(originalpath, newpath);
                                    //
                                    btr.PathName = @"..\" + childname;
                                    //
                                    editor.WriteMessage("\nXREF old path: " + originalpath + "\n");
                                    editor.WriteMessage("\nXREF new path: " + newpath + "XREF changed" + "\n");
                                }
                            }
                            db.SaveAs(pathXref, true, DwgVersion.Current, doc.Database.SecurityParameters);
                        }
                        tr.Commit();
                    }
                }
                catch (System.Exception ex) { editor.WriteMessage("\n Work with base drawing Error: " + ex.Message + "\n" + ex.StackTrace + "\n"); }
            }
        }

        private void Terminate()
        {
        }

        /// <summary>
        /// Cинхронизация внешних ссылок в чертеже AutoCAD с файлами расположенными в ТДМС.
        /// </summary>
        ///
    }
}